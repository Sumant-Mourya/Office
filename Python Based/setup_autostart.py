"""Setup Activity Tracker to run at Windows startup and prevent multiple instances.

Uses:
1. Windows shortcut (.lnk) in the Startup folder pointing to the original path (resolves firewall prompts).
2. Task Scheduler task registered to start at logon.
3. Named system mutex to prevent multiple instances from running.
"""

import os
import sys
import shutil
import subprocess
import win32event
import win32api
import winerror
import psutil
import time

from logger_setup import get_logger
from config import USER_DATA_DIR, UI_PORT

log = get_logger("autostart")

ROOT_DIR = (
    os.path.dirname(os.path.abspath(sys.executable))
    if getattr(sys, "frozen", False)
    else os.path.dirname(os.path.abspath(__file__))
)
STARTUP_EXE_NAME = "ActivityTracker_Autostart.exe"
LEGACY_STARTUP_BAT_NAME = "ActivityTracker_Autostart.bat"

# Keep global reference to the mutex so it doesn't get garbage-collected
_mutex = None


def kill_other_instances():
    """Find and terminate any other running instances of the tracker or processes using our UI port.
    This frees up the single instance mutex and port so the new instance can gracefully start.
    """
    current_pid = os.getpid()
    pids_to_kill = set()
    
    # 1. Identify processes listening on UI_PORT (default 8580)
    try:
        for conn in psutil.net_connections(kind='inet'):
            if conn.laddr and conn.laddr.port == UI_PORT:
                if conn.pid and conn.pid != current_pid and conn.pid != os.getppid():
                    pids_to_kill.add(conn.pid)
    except Exception as exc:
        log.warning("Failed to query network connections: %s", exc)
        
    # 2. Identify processes by name or cmdline (ActivityTracker.exe, main.py)
    try:
        for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
            try:
                pid = proc.info['pid']
                if pid == current_pid or pid == os.getppid():
                    continue
                name = (proc.info['name'] or '').lower()
                cmdline = proc.info['cmdline'] or []
                cmd_str = " ".join(cmdline).lower()
                
                is_target = False
                if name in ("activitytracker.exe", "activitytracker_autostart.exe"):
                    is_target = True
                elif "python" in name and any("main.py" in arg.lower() for arg in cmdline):
                    is_target = True
                elif "main.py" in cmd_str:
                    is_target = True
                    
                if is_target:
                    pids_to_kill.add(pid)
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                continue
    except Exception as exc:
        log.warning("Failed to iterate processes: %s", exc)
        
    if pids_to_kill:
        log.info("Found other running instances to terminate: %s", pids_to_kill)
        for pid in pids_to_kill:
            try:
                p = psutil.Process(pid)
                p.kill()
                log.info("Successfully terminated process %d", pid)
            except Exception as exc:
                log.warning("Could not terminate process %d: %s", pid, exc)
        
        # Sleep to allow sockets and mutex handles to be released by Windows
        time.sleep(1.0)


def check_single_instance() -> bool:
    """Check if another instance is already running using a named system mutex.
    Returns True if this is the only instance, False otherwise.
    """
    kill_other_instances()
    global _mutex
    mutex_name = "Global\\ActivityTrackerSingleInstanceMutex"
    try:
        # Create mutex. If it already exists, GetLastError returns ERROR_ALREADY_EXISTS.
        _mutex = win32event.CreateMutex(None, False, mutex_name)
        if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
            _mutex = None
            log.warning("Another instance of Activity Tracker is already running. Exiting.")
            return False
        return True
    except Exception as exc:
        log.warning("Single instance mutex check encountered an issue: %s", exc)
        return True  # Fallback to allow startup if mutex creation fails


def _startup_folder() -> str:
    appdata = os.environ.get("APPDATA", "")
    return os.path.join(appdata, "Microsoft", "Windows", "Start Menu", "Programs", "Startup")


def _startup_shortcut_path() -> str:
    return os.path.join(_startup_folder(), "ActivityTracker.lnk")


def _legacy_startup_exe_path() -> str:
    return os.path.join(_startup_folder(), STARTUP_EXE_NAME)


def _legacy_startup_bat_path() -> str:
    return os.path.join(_startup_folder(), LEGACY_STARTUP_BAT_NAME)


def _resolve_target_info() -> tuple[str, str]:
    """Find the target path and arguments to run the application."""
    if getattr(sys, "frozen", False):
        return sys.executable, ""
    
    # Resolve pythonw.exe to start without a terminal window when starting from script
    exe = sys.executable
    if exe.lower().endswith("python.exe"):
        w_exe = exe[:-10] + "pythonw.exe"
        if os.path.exists(w_exe):
            exe = w_exe
    elif exe.lower().endswith("python"):
        w_exe = exe[:-6] + "pythonw"
        if os.path.exists(w_exe):
            exe = w_exe
            
    main_script = os.path.abspath(sys.argv[0])
    if os.path.basename(main_script) != "main.py":
        main_script = os.path.join(ROOT_DIR, "main.py")
    return exe, f'"{main_script}"'


def install_startup_script() -> bool:
    """Register startup shortcut in Startup folder AND logon task in Task Scheduler.
    Both point to the original executable path to avoid firewall prompting.
    """
    startup_dir = _startup_folder()
    if not startup_dir:
        log.error("APPDATA is unavailable; cannot locate Startup folder.")
        return False

    try:
        os.makedirs(startup_dir, exist_ok=True)
    except Exception as exc:
        log.error("Failed to create Startup folder: %s", exc)
        return False

    target_exe, target_args = _resolve_target_info()
    if not target_exe:
        log.error("Unable to resolve target executable path.")
        return False

    # 1. Create Startup folder shortcut (.lnk) via PowerShell
    shortcut_ok = False
    try:
        shortcut_path = os.path.normpath(_startup_shortcut_path())
        norm_target = os.path.normpath(target_exe)
        
        # Escape single quotes in fields for single-quoted PowerShell strings
        shortcut_path_esc = shortcut_path.replace("'", "''")
        norm_target_esc = norm_target.replace("'", "''")
        target_args_esc = target_args.replace("'", "''")
        working_dir_esc = os.path.dirname(norm_target).replace("'", "''")
        
        ps_cmd = (
            f'$WshShell = New-Object -ComObject WScript.Shell; '
            f'$Shortcut = $WshShell.CreateShortcut(\'{shortcut_path_esc}\'); '
            f'$Shortcut.TargetPath = \'{norm_target_esc}\'; '
            f'$Shortcut.Arguments = \'{target_args_esc}\'; '
            f'$Shortcut.WorkingDirectory = \'{working_dir_esc}\'; '
            f'$Shortcut.Save()'
        )
        cmd = ["powershell", "-Command", ps_cmd]
        res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, check=True)
        log.info("Startup shortcut created at: %s", shortcut_path)
        shortcut_ok = True
    except Exception as exc:
        log.error("Failed to create startup shortcut: %s", exc)

    # 2. Register task in Task Scheduler (run at logon with elevated RunAs privileges)
    task_ok = False
    temp_bat = None
    try:
        temp_bat = os.path.join(USER_DATA_DIR, "register_task.bat")
        
        main_script_path = target_args.strip('"')
        if main_script_path:
            tr_val = f'\\"{target_exe}\\" \\"{main_script_path}\\"'
        else:
            tr_val = f'\\"{target_exe}\\"'
            
        bat_content = f'@echo off\nschtasks /create /tn "ActivityTracker" /tr "{tr_val}" /sc onlogon /f\n'
        with open(temp_bat, "w", encoding="utf-8") as f:
            f.write(bat_content)
            
        ps_cmd = f"Start-Process '{temp_bat}' -Verb RunAs -WindowStyle Hidden -Wait"
        cmd = ["powershell", "-Command", ps_cmd]
        res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        if res.returncode == 0:
            log.info("Registered logon task in Windows Task Scheduler (requested elevation).")
            task_ok = True
        else:
            log.error("Failed to run task registration batch: %s", res.stderr)
    except Exception as exc:
        log.error("Failed to register Task Scheduler logon task: %s", exc)
    finally:
        if temp_bat and os.path.exists(temp_bat):
            try:
                os.remove(temp_bat)
            except Exception:
                pass

    # Clean up legacy copied exe or bat files in the startup folder
    for legacy_path in (_legacy_startup_exe_path(), _legacy_startup_bat_path()):
        if os.path.exists(legacy_path):
            try:
                os.remove(legacy_path)
            except Exception:
                pass

    return shortcut_ok or task_ok


def remove_startup_script() -> bool:
    """Remove startup shortcut and Task Scheduler task entry."""
    removed_any = False
    shortcut_path = _startup_shortcut_path()
    legacy_exe = _legacy_startup_exe_path()
    legacy_bat = _legacy_startup_bat_path()

    # Remove Startup folder files
    for path in (shortcut_path, legacy_exe, legacy_bat):
        if os.path.exists(path):
            try:
                os.remove(path)
                removed_any = True
                log.info("Removed startup file: %s", path)
            except Exception as exc:
                log.error("Failed to remove startup file %s: %s", path, exc)

    # Remove Task Scheduler task (requires elevated RunAs privileges)
    temp_bat = None
    try:
        temp_bat = os.path.join(USER_DATA_DIR, "delete_task.bat")
        
        bat_content = '@echo off\nschtasks /delete /tn "ActivityTracker" /f\n'
        with open(temp_bat, "w", encoding="utf-8") as f:
            f.write(bat_content)
            
        ps_cmd = f"Start-Process '{temp_bat}' -Verb RunAs -WindowStyle Hidden -Wait"
        cmd = ["powershell", "-Command", ps_cmd]
        res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        if res.returncode == 0:
            removed_any = True
            log.info("Removed ActivityTracker Task Scheduler entry.")
        else:
            log.error("Failed to run task deletion batch: %s", res.stderr)
    except Exception as exc:
        log.error("Failed to remove Task Scheduler entry: %s", exc)
    finally:
        if temp_bat and os.path.exists(temp_bat):
            try:
                os.remove(temp_bat)
            except Exception:
                pass

    return removed_any


def startup_script_exists() -> bool:
    """Return True if shortcut exists or Task Scheduler task is registered."""
    if os.path.exists(_startup_shortcut_path()):
        return True
    try:
        cmd = ["schtasks", "/query", "/tn", "ActivityTracker"]
        res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        return res.returncode == 0
    except Exception:
        return False


# Backward-compatible names for UI code integration
setup_task_scheduler = install_startup_script
remove_task = remove_startup_script
task_exists = startup_script_exists

if __name__ == "__main__":
    if "--remove" in sys.argv:
        ok = remove_startup_script()
        print("Removed." if ok else "Failed to remove.")
    else:
        ok = install_startup_script()
        if ok:
            print("Auto-start registered. Tracker will launch at next login.")
        else:
            print("Failed to register auto-start options.")
