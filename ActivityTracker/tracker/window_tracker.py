"""Active window tracker using pywin32.

Only returns info for the foreground window that is NOT minimized.
"""

import re
import win32gui
import win32process
import psutil

from logger_setup import get_logger

log = get_logger("tracker.window")


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
        proc = psutil.Process(pid)
        app_name = proc.name()  # e.g. "chrome.exe"
        minimized = bool(win32gui.IsIconic(hwnd))

        return {"app": app_name, "title": title, "minimized": minimized}
    except Exception:
        return None


def extract_browser_domain(title: str) -> str | None:
    """Extract the page/site portion from a browser window title.

    Handles Chrome, Edge, Firefox, Brave, Opera – they all append
    `` - <Browser Name>`` at the end.
    """
    if not title:
        return None
    cleaned = re.sub(
        r"\s*[-–—]\s*(Google Chrome|Microsoft\s?Edge|Mozilla Firefox|Brave|Opera)\s*$",
        "",
        title,
        flags=re.IGNORECASE,
    )
    return cleaned.strip() if cleaned.strip() else None


# Keep old name as alias for backward compat
extract_chrome_domain = extract_browser_domain
