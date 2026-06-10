"""Active window tracker using pywin32.

Only returns info for the foreground window that is NOT minimized.
"""

import re
import win32gui
import win32process
import psutil

from logger_setup import get_logger

log = get_logger("tracker.window")

# Browsers whose titles typically end with " - <Browser Name>"
_BROWSER_SUFFIXES_RE = re.compile(
    r"\s*[-–—]\s*"
    r"(?:Google Chrome|Microsoft\s?Edge|Mozilla Firefox|Brave|Opera|Opera GX"
    r"|Vivaldi|Arc|Waterfox|Chromium|Thorium)"
    r"\s*$",
    re.IGNORECASE,
)


def _resolve_uwp_child(pid: int) -> str | None:
    """If the foreground process is ApplicationFrameHost.exe, try to find
    the real hosted app by inspecting child processes."""
    try:
        parent = psutil.Process(pid)
        children = parent.children(recursive=False)
        for child in children:
            try:
                name = child.name()
                if name.lower() not in (
                    "runtimebroker.exe",
                    "applicationframehost.exe",
                ):
                    return name
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        pass
    return None


def get_active_window_info() -> dict | None:
    """Return {'app': str, 'title': str, 'minimized': bool} for the foreground window.

    Returns None if the window handle is invalid or the process cannot be read.
    """
    try:
        hwnd = win32gui.GetForegroundWindow()
        if not hwnd:
            return None

        title = win32gui.GetWindowText(hwnd)
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        try:
            proc = psutil.Process(pid)
            app_name = proc.name()
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            return None

        # Resolve UWP hosted apps
        if app_name.lower() == "applicationframehost.exe":
            real_app = _resolve_uwp_child(pid)
            if real_app:
                app_name = real_app

        minimized = bool(win32gui.IsIconic(hwnd))
        return {"app": app_name, "title": title, "minimized": minimized}
    except Exception:
        return None


def extract_browser_domain(title: str) -> str | None:
    """Extract the page/site portion from a browser window title.

    Handles Chrome, Edge, Firefox, Brave, Opera, Vivaldi, Arc, Waterfox,
    Chromium – they all append `` - <Browser Name>`` at the end.
    """
    if not title:
        return None
    cleaned = _BROWSER_SUFFIXES_RE.sub("", title)
    return cleaned.strip() if cleaned.strip() else None


# Keep old name as alias for backward compat
extract_chrome_domain = extract_browser_domain
