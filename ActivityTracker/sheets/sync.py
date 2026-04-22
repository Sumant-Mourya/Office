"""Google Sheets sync – simple daily summary + config sheet."""

import socket
import ipaddress
from datetime import datetime
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from config import UI_PORT
from logger_setup import get_logger

log = get_logger("sheets.sync")

# Header columns for the PC data tab
_HEADERS = [
    "DATE", "TOTAL MOUSE CLICKS", "TOTAL KEY PRESSES",
    "WEBSITES (time spent)", "APPS (time spent)",
    "TOTAL ACTIVE TIME", "TOTAL IDLE TIME",
]


def _get_local_ip() -> str:
    """Get a usable private IPv4 LAN address for this machine."""

    def _is_usable_private_ipv4(ip: str) -> bool:
        try:
            addr = ipaddress.ip_address(ip)
        except ValueError:
            return False
        return (
            addr.version == 4
            and addr.is_private
            and not addr.is_loopback
            and not addr.is_link_local
        )

    candidates: list[str] = []
    seen: set[str] = set()

    def _add_candidate(ip: str):
        if ip in seen:
            return
        seen.add(ip)
        if _is_usable_private_ipv4(ip):
            candidates.append(ip)

    # UDP probe usually reveals the active interface IP.
    for target in ("8.8.8.8", "1.1.1.1", "192.168.1.1", "10.255.255.255"):
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect((target, 80))
            _add_candidate(s.getsockname()[0])
            s.close()
        except Exception:
            continue

    # Hostname resolution as a fallback source.
    try:
        hostname = socket.gethostname()
        for _, _, _, _, sockaddr in socket.getaddrinfo(
            hostname, None, socket.AF_INET, socket.SOCK_DGRAM
        ):
            _add_candidate(sockaddr[0])
    except Exception:
        pass

    def _rank(ip: str) -> int:
        if ip.startswith("192.168."):
            return 0
        if ip.startswith("10."):
            return 1
        if ip.startswith("172."):
            return 2
        return 3

    if candidates:
        candidates.sort(key=_rank)
        return candidates[0]

    return "127.0.0.1"


class SheetSync:
    """Simple daily-summary sync + config tab."""

    def __init__(self, creds, spreadsheet_id: str, sheet_name: str):
        self.spreadsheet_id = spreadsheet_id
        self.sheet_name = sheet_name
        self.service = build("sheets", "v4", credentials=creds)
        self.sheets = self.service.spreadsheets()

    # ── Tab helpers ───────────────────────────────────────────────────

    def _tab_exists(self, name: str) -> bool:
        meta = self.sheets.get(spreadsheetId=self.spreadsheet_id).execute()
        return any(
            s["properties"]["title"] == name for s in meta.get("sheets", [])
        )

    def sheet_exists(self) -> bool:
        try:
            return self._tab_exists(self.sheet_name)
        except HttpError as exc:
            log.error("sheet_exists check failed: %s", exc)
            raise

    def create_sheet(self):
        """Create the PC tab if it doesn't exist, then write headers."""
        if not self._tab_exists(self.sheet_name):
            self.sheets.batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body={"requests": [
                    {"addSheet": {"properties": {"title": self.sheet_name}}}
                ]},
            ).execute()
            log.info("Created sheet tab '%s'", self.sheet_name)

        # Always write headers on row 1
        self.sheets.values().update(
            spreadsheetId=self.spreadsheet_id,
            range=f"'{self.sheet_name}'!A1:G1",
            valueInputOption="RAW",
            body={"values": [_HEADERS]},
        ).execute()

    def log_start_time(self):
        """Append a row with the program start timestamp."""
        now_str = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
        row = [[f"▶ Program started: {now_str}", "", "", "", "", "", ""]]
        self.sheets.values().append(
            spreadsheetId=self.spreadsheet_id,
            range=f"'{self.sheet_name}'!A:G",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": row},
        ).execute()
        log.info("Logged start time to sheet")

    # ── Config sheet ──────────────────────────────────────────────────

    def ensure_config_sheet(self, pc_name: str):
        """Create / update a 'config' tab with pc_name + local WiFi IP."""
        tab = "config"
        if not self._tab_exists(tab):
            self.sheets.batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body={"requests": [
                    {"addSheet": {"properties": {"title": tab}}}
                ]},
            ).execute()
            self.sheets.values().update(
                spreadsheetId=self.spreadsheet_id,
                range=f"'{tab}'!A1:D1",
                valueInputOption="RAW",
                body={"values": [["PC_NAME", "LOCAL_IP", "APP_URL", "LAST_SEEN"]]},
            ).execute()

        result = self.sheets.values().get(
            spreadsheetId=self.spreadsheet_id,
            range=f"'{tab}'!A2:D200",
        ).execute()
        rows = result.get("values", [])

        row_idx = None
        for i, r in enumerate(rows, start=2):
            if r and r[0] == pc_name:
                row_idx = i
                break

        local_ip = _get_local_ip()
        if local_ip == "127.0.0.1" and row_idx:
            existing_row = rows[row_idx - 2]
            existing_ip = existing_row[1] if len(existing_row) > 1 else ""
            if existing_ip and existing_ip != "127.0.0.1":
                local_ip = existing_ip
        if local_ip == "127.0.0.1":
            log.warning(
                "Could not detect private LAN IP for %s; storing localhost.",
                pc_name,
            )
        app_url = f"http://{local_ip}:{UI_PORT}"
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        values = [[pc_name, local_ip, app_url, now_str]]

        if row_idx:
            self.sheets.values().update(
                spreadsheetId=self.spreadsheet_id,
                range=f"'{tab}'!A{row_idx}:D{row_idx}",
                valueInputOption="RAW",
                body={"values": values},
            ).execute()
        else:
            self.sheets.values().append(
                spreadsheetId=self.spreadsheet_id,
                range=f"'{tab}'!A:D",
                valueInputOption="RAW",
                insertDataOption="INSERT_ROWS",
                body={"values": values},
            ).execute()
        log.info("Config sheet updated for %s (url=%s)", pc_name, app_url)

    def pc_name_taken(self, pc_name: str) -> bool:
        """Check if a PC name is already claimed in config.

        A tab can exist without a config entry if a previous start attempt failed
        midway. Those partial cases are treated as not-taken so the same machine
        can retry setup successfully.
        """
        if not self._tab_exists(pc_name):
            return False

        tab = "config"
        if not self._tab_exists(tab):
            return False

        try:
            result = self.sheets.values().get(
                spreadsheetId=self.spreadsheet_id,
                range=f"'{tab}'!A2:A200",
            ).execute()
        except HttpError:
            # Be conservative when we cannot read ownership metadata.
            return True

        for row in result.get("values", []):
            if row and row[0] == pc_name:
                return True
        return False

    def read_config_pcs(self) -> list[dict]:
        """Return list of PCs from config tab."""
        tab = "config"
        if not self._tab_exists(tab):
            return []
        try:
            result = self.sheets.values().get(
                spreadsheetId=self.spreadsheet_id,
                range=f"'{tab}'!A2:D100",
            ).execute()
            out = []
            for row in result.get("values", []):
                if row:
                    out.append({
                        "pc_name": row[0] if len(row) > 0 else "",
                        "local_ip": row[1] if len(row) > 1 else "",
                        "app_url": row[2] if len(row) > 2 else "",
                        "last_seen": row[3] if len(row) > 3 else "",
                    })
            return out
        except HttpError:
            return []

    # ── Formatting helpers ────────────────────────────────────────────

    @staticmethod
    def _fmt_usage(usage_dict: dict) -> str:
        """Format usage dict as multiline: 'name = Xh Ym Zs' per line."""
        items = sorted(usage_dict.items(), key=lambda x: -x[1])
        lines = []
        for name, sec in items:
            h, rem = divmod(int(sec), 3600)
            m, s = divmod(rem, 60)
            if h > 0:
                lines.append(f"{name} = {h}h {m}m {s}s")
            elif m > 0:
                lines.append(f"{name} = {m}m {s}s")
            else:
                lines.append(f"{name} = {s}s")
        return "\n".join(lines) if lines else ""

    @staticmethod
    def _fmt_seconds(sec: float) -> str:
        h, rem = divmod(int(sec), 3600)
        m, s = divmod(rem, 60)
        return f"{h}h {m}m {s}s"

    # ── Sync – one row per date with daily totals ─────────────────────

    def sync(self, data: dict):
        """Upsert a single summary row for the given date."""
        target_date = data["date"]
        hours_data = data.get("hours", {})

        # Aggregate totals across all hours
        total_mouse = sum(
            h.get("mouse_clicks", 0) for h in hours_data.values())
        total_keys = sum(
            h.get("key_presses", 0) for h in hours_data.values())
        total_work = sum(
            h.get("work_seconds", 0) for h in hours_data.values())
        total_idle = sum(
            h.get("idle_seconds", 0) for h in hours_data.values())

        # Merge website usage across all hours
        all_websites: dict[str, float] = {}
        all_apps: dict[str, float] = {}
        for hdata in hours_data.values():
            for site, sec in hdata.get("websites", {}).items():
                all_websites[site] = all_websites.get(site, 0) + sec
            for app, sec in hdata.get("windows", {}).items():
                all_apps[app] = all_apps.get(app, 0) + sec

        row = [
            target_date,
            str(total_mouse),
            str(total_keys),
            self._fmt_usage(all_websites),
            self._fmt_usage(all_apps),
            self._fmt_seconds(total_work),
            self._fmt_seconds(total_idle),
        ]

        # Find existing date row (skip header and start-time rows)
        try:
            result = self.sheets.values().get(
                spreadsheetId=self.spreadsheet_id,
                range=f"'{self.sheet_name}'!A:A",
            ).execute()
            existing = result.get("values", [])
        except HttpError as exc:
            log.error("Failed to read sheet: %s", exc)
            return

        row_idx = None
        for i, r in enumerate(existing):
            if r and r[0] == target_date:
                row_idx = i + 1
                break

        try:
            if row_idx:
                self.sheets.values().update(
                    spreadsheetId=self.spreadsheet_id,
                    range=f"'{self.sheet_name}'!A{row_idx}:G{row_idx}",
                    valueInputOption="RAW",
                    body={"values": [row]},
                ).execute()
            else:
                self.sheets.values().append(
                    spreadsheetId=self.spreadsheet_id,
                    range=f"'{self.sheet_name}'!A:G",
                    valueInputOption="RAW",
                    insertDataOption="INSERT_ROWS",
                    body={"values": [row]},
                ).execute()
            log.debug("Synced summary row for %s", target_date)
        except HttpError as exc:
            log.error("Sheet sync failed: %s", exc)
            raise

    # ── Access validation ─────────────────────────────────────────────

    def validate_access(self) -> bool:
        try:
            self.sheets.get(spreadsheetId=self.spreadsheet_id).execute()
            return True
        except HttpError:
            return False
