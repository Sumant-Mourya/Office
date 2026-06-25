"""Network connectivity detection – WiFi, LAN, and internet checks."""

import socket
import ipaddress

from logger_setup import get_logger

log = get_logger("network_monitor")


def is_network_connected() -> bool:
    """Check if the system has any active network connection (WiFi or LAN)."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.settimeout(2)
        s.connect(("8.8.8.8", 80))
        s.close()
        return True
    except Exception:
        return False


def is_internet_available() -> bool:
    """Check if Google APIs are reachable."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.settimeout(3)
        s.connect(("sheets.googleapis.com", 443))
        s.close()
        return True
    except Exception:
        return False


def get_all_local_ips() -> list[str]:
    """Return all usable private IPv4 addresses on this machine."""
    ips: list[str] = []
    seen: set[str] = set()

    def _add(ip: str):
        if ip in seen:
            return
        seen.add(ip)
        try:
            addr = ipaddress.ip_address(ip)
        except ValueError:
            return
        if (
            addr.version == 4
            and addr.is_private
            and not addr.is_loopback
            and not addr.is_link_local
        ):
            ips.append(ip)

    # UDP probe to common targets
    for target in ("8.8.8.8", "1.1.1.1", "192.168.1.1", "10.255.255.255"):
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.settimeout(1)
            s.connect((target, 80))
            _add(s.getsockname()[0])
            s.close()
        except Exception:
            continue

    # Hostname resolution
    try:
        hostname = socket.gethostname()
        for _, _, _, _, sockaddr in socket.getaddrinfo(
            hostname, None, socket.AF_INET, socket.SOCK_DGRAM
        ):
            _add(sockaddr[0])
    except Exception:
        pass

    # Sort: 192.168.x first, then 10.x, then 172.x
    def _rank(ip: str) -> int:
        if ip.startswith("192.168."):
            return 0
        if ip.startswith("10."):
            return 1
        if ip.startswith("172."):
            return 2
        return 3

    ips.sort(key=_rank)
    return ips


def get_best_local_ip() -> str:
    """Return the best available private IP, or 127.0.0.1 as fallback."""
    ips = get_all_local_ips()
    return ips[0] if ips else "127.0.0.1"
