"""Setup Activity Tracker to run at Windows startup using an EXE in Startup.

Run:  python setup_autostart.py           (copy app EXE into Startup folder)
    python setup_autostart.py --remove  (remove app EXE from Startup folder)
"""

import os
import shutil
import sys
from glob import glob

from logger_setup import get_logger

log = get_logger("autostart")

ROOT_DIR = (
    os.path.dirname(os.path.abspath(sys.executable))
    if getattr(sys, "frozen", False)
    else os.path.dirname(os.path.abspath(__file__))
)
STARTUP_EXE_NAME = "ActivityTracker_Autostart.exe"
LEGACY_STARTUP_BAT_NAME = "ActivityTracker_Autostart.bat"
PREFERRED_EXE_NAMES = ("ActivityTracker.exe", "main.exe")


def _startup_folder() -> str:
    appdata = os.environ.get("APPDATA", "")
    return os.path.join(appdata, "Microsoft", "Windows", "Start Menu", "Programs", "Startup")


def _startup_exe_path() -> str:
    return os.path.join(_startup_folder(), STARTUP_EXE_NAME)


def _legacy_startup_bat_path() -> str:
    return os.path.join(_startup_folder(), LEGACY_STARTUP_BAT_NAME)


def _resolve_source_exe() -> str | None:
    """Find the app EXE in the project/app directory."""
    candidates: list[str] = []

    env_exe = os.environ.get("ACTIVITY_TRACKER_EXE", "").strip()
    if env_exe:
        env_path = env_exe if os.path.isabs(env_exe) else os.path.join(ROOT_DIR, env_exe)
        candidates.append(env_path)

    # When app is already running as EXE, prefer that same binary.
    if getattr(sys, "frozen", False):
        candidates.append(sys.executable)

    for name in PREFERRED_EXE_NAMES:
        candidates.append(os.path.join(ROOT_DIR, name))

    for path in sorted(glob(os.path.join(ROOT_DIR, "*.exe"))):
        if os.path.basename(path).lower() == STARTUP_EXE_NAME.lower():
            continue
        candidates.append(path)

    seen: set[str] = set()
    for path in candidates:
        norm = os.path.normcase(os.path.abspath(path))
        if norm in seen:
            continue
        seen.add(norm)
        if os.path.isfile(path):
            return path
    return None


def install_startup_script() -> bool:
    """Copy app EXE to Startup folder for auto-run on login."""
    startup_dir = _startup_folder()
    if not startup_dir:
        log.error("APPDATA is unavailable; cannot locate Startup folder.")
        return False

    try:
        os.makedirs(startup_dir, exist_ok=True)
    except Exception as exc:
        log.error("Failed to create Startup folder path: %s", exc)
        return False

    source_exe = _resolve_source_exe()
    if not source_exe:
        log.error(
            "No source EXE found in %s. Build ActivityTracker.exe first.",
            ROOT_DIR,
        )
        return False

    try:
        shutil.copy2(source_exe, _startup_exe_path())
        legacy_bat = _legacy_startup_bat_path()
        if os.path.exists(legacy_bat):
            os.remove(legacy_bat)
        log.info("Startup EXE copied to: %s", _startup_exe_path())
        return True
    except Exception as exc:
        log.error("Failed to copy EXE to Startup folder: %s", exc)
        return False


def remove_startup_script() -> bool:
    """Remove startup EXE (and legacy BAT) from Startup folder."""
    removed_any = False
    startup_exe = _startup_exe_path()
    legacy_bat = _legacy_startup_bat_path()

    try:
        if os.path.exists(startup_exe):
            os.remove(startup_exe)
            removed_any = True
            log.info("Startup EXE removed from: %s", startup_exe)
        if os.path.exists(legacy_bat):
            os.remove(legacy_bat)
            removed_any = True
            log.info("Legacy startup BAT removed from: %s", legacy_bat)
        if not removed_any:
            log.info("Startup entry not present (already removed).")
        return True
    except Exception as exc:
        log.error("Failed to remove startup entry: %s", exc)
        return False


def startup_script_exists() -> bool:
    """Return True when startup EXE exists in Startup folder."""
    return os.path.exists(_startup_exe_path())


# Backward-compatible names used by existing UI code.
def setup_task_scheduler() -> bool:
    return install_startup_script()


def remove_task() -> bool:
    return remove_startup_script()


def task_exists() -> bool:
    return startup_script_exists()


if __name__ == "__main__":
    if "--remove" in sys.argv:
        ok = remove_startup_script()
        print("Removed." if ok else "Failed to remove.")
    else:
        ok = install_startup_script()
        if ok:
            print("Startup EXE installed. Tracker will start on next login.")
        else:
            print("Failed to install startup EXE.")
