"""Global configuration constants."""

import os
import sys

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
    BUNDLE_DIR = getattr(sys, "_MEIPASS", BASE_DIR)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    BUNDLE_DIR = BASE_DIR


def _default_user_data_dir() -> str:
    """Return a writable per-user directory for runtime files."""
    local_app_data = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA")
    if local_app_data:
        return os.path.join(local_app_data, "ActivityTracker")
    return BASE_DIR


USER_DATA_DIR = (
    os.environ.get("ACTIVITY_TRACKER_DATA_DIR", "").strip()
    or _default_user_data_dir()
)
os.makedirs(USER_DATA_DIR, exist_ok=True)


def _legacy_path(name: str) -> str:
    return os.path.join(BASE_DIR, name)


def _runtime_path(name: str) -> str:
    return os.path.join(USER_DATA_DIR, name)

# Google OAuth
_credentials_in_app_dir = os.path.join(BASE_DIR, "credentials.json")
_credentials_in_bundle = os.path.join(BUNDLE_DIR, "credentials.json")
CREDENTIALS_FILE = (
    _credentials_in_app_dir
    if os.path.exists(_credentials_in_app_dir)
    else _credentials_in_bundle
)
TOKEN_FILE = _runtime_path("token.json")
LEGACY_TOKEN_FILE = _legacy_path("token.json")
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# Default Google Sheet ID. Can be overridden by UI and saved config.
# Optional env var makes first-run setup easier in managed deployments.
SHEET_ID = os.environ.get("ACTIVITY_TRACKER_SHEET_ID", "").strip()

# Local data storage
DATA_DIR = _runtime_path("data")
LEGACY_DATA_DIR = _legacy_path("data")

# Saved tracker config (sheet_id, pc_name) – encrypted on disk
TRACKER_CONFIG_FILE = _runtime_path("tracker_config.enc")
TRACKER_CONFIG_KEY_FILE = _runtime_path(".config_key")
LEGACY_TRACKER_CONFIG_FILE = _legacy_path("tracker_config.enc")
LEGACY_TRACKER_CONFIG_KEY_FILE = _legacy_path(".config_key")

# Logging
LOG_DIR = _runtime_path("logs")
LOG_FILE = os.path.join(LOG_DIR, "tracker.log")

# Tracking intervals (seconds)
SHEET_SYNC_INTERVAL = 60  # sync to Google Sheets every 1 minute
IDLE_THRESHOLD = 120       # seconds of no input before considered idle
TRACKER_POLL_INTERVAL = 1  # poll active window every 1 second

# Login credentials
DEFAULT_USER = "admin"
DEFAULT_PASS = "admin"

# NiceGUI
UI_PORT = 8580
UI_TITLE = "Activity Tracker"

# System / non-productive apps to exclude from work-time tracking.
# Only the foreground window is ever tracked; these are filtered out
# so antivirus popups, search UI, etc. don't count as "work".
SYSTEM_APPS = {
    # Windows shell / UI
    "SearchUI.exe", "SearchHost.exe", "SearchApp.exe",
    "ShellExperienceHost.exe", "StartMenuExperienceHost.exe",
    "RuntimeBroker.exe", "ApplicationFrameHost.exe",
    "TextInputHost.exe", "LockApp.exe", "SystemSettings.exe",
    "ctfmon.exe", "dwm.exe", "sihost.exe", "taskhostw.exe",
    "fontdrvhost.exe",
    # Security / Antivirus
    "MsMpEng.exe", "SecurityHealthSystray.exe",
    "SecurityHealthService.exe", "NisSrv.exe",
    "MpCmdRun.exe", "smartscreen.exe",
    "WindowsDefender.exe", "SecHealthUI.exe",
    # Other system processes
    "svchost.exe", "csrss.exe", "lsass.exe",
    "conhost.exe", "dllhost.exe", "WmiPrvSE.exe",
}
