"""Local JSON data store – per-user, per-second granularity.

Each user (pc_name) gets a directory under data/<pc_name>/ with one
JSON file per day: YYYY-MM-DD.json. Old data beyond MAX_DAYS is
pruned automatically.
"""

import glob
import json
import os
import shutil
from datetime import date, timedelta

from config import DATA_DIR, LEGACY_DATA_DIR
from logger_setup import get_logger

log = get_logger("local_store")

MAX_DAYS = 50


def _user_dir(pc_name: str) -> str:
    d = os.path.join(DATA_DIR, pc_name)
    os.makedirs(d, exist_ok=True)
    return d


def _candidate_user_dirs(pc_name: str) -> list[str]:
    dirs = [_user_dir(pc_name)]
    legacy_dir = os.path.join(LEGACY_DATA_DIR, pc_name)
    if legacy_dir != dirs[0] and os.path.isdir(legacy_dir):
        dirs.append(legacy_dir)
    return dirs


def _day_path(pc_name: str, day: str) -> str:
    return os.path.join(_user_dir(pc_name), f"{day}.json")


def _day_backup_path(pc_name: str, day: str) -> str:
    return os.path.join(_user_dir(pc_name), f"{day}.json.bak")


def _iter_day_files(pc_name: str) -> list[str]:
    files: list[str] = []
    seen: set[str] = set()
    for user_dir in _candidate_user_dirs(pc_name):
        for path in glob.glob(os.path.join(user_dir, "????-??-??.json")):
            day = os.path.basename(path).replace(".json", "")
            if day in seen:
                continue
            seen.add(day)
            files.append(path)
    return files


def list_pcs() -> list[str]:
    """Return merged list of known PC names from current + legacy stores."""
    names: set[str] = set()
    for base_dir in (DATA_DIR, LEGACY_DATA_DIR):
        if not os.path.isdir(base_dir):
            continue
        try:
            entries = os.listdir(base_dir)
        except OSError:
            continue
        for name in entries:
            if os.path.isdir(os.path.join(base_dir, name)):
                names.add(name)
    return sorted(names)


# -- Read / Write -----------------------------------------------------

def load_day(pc_name: str, day: str) -> dict | None:
    """Load a single day's data. Returns None if not found."""
    for user_dir in _candidate_user_dirs(pc_name):
        path = os.path.join(user_dir, f"{day}.json")
        if not os.path.exists(path):
            continue
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except json.JSONDecodeError as exc:
            bak_path = os.path.join(user_dir, f"{day}.json.bak")
            if os.path.exists(bak_path):
                try:
                    with open(bak_path, "r", encoding="utf-8") as f:
                        recovered = json.load(f)
                    log.warning(
                        "Primary day file invalid (%s), recovered from backup: %s",
                        exc,
                        bak_path,
                    )
                    return recovered
                except Exception as bak_exc:
                    log.warning("Backup load also failed %s: %s", bak_path, bak_exc)
            log.warning("Failed to load %s: %s", path, exc)
        except Exception as exc:
            log.warning("Failed to load %s: %s", path, exc)
    return None


def save_day(pc_name: str, day: str, data: dict) -> None:
    """Persist a day's data and prune old files."""
    path = _day_path(pc_name, day)
    tmp_path = f"{path}.tmp"
    bak_path = _day_backup_path(pc_name, day)
    try:
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(data, f, default=str)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp_path, path)
        shutil.copyfile(path, bak_path)
    except Exception as exc:
        log.error("Failed to save %s: %s", path, exc)
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except OSError:
            pass
    _prune(pc_name)


def list_days(pc_name: str) -> list[str]:
    """Return sorted list of date strings that have data."""
    days = []
    for path in _iter_day_files(pc_name):
        days.append(os.path.basename(path).replace(".json", ""))
    return sorted(days)


def load_all(pc_name: str) -> list[dict]:
    """Load all days, newest first."""
    days = list_days(pc_name)
    result = []
    for day in reversed(days):
        d = load_day(pc_name, day)
        if d:
            result.append(d)
    return result


# -- Pruning ----------------------------------------------------------

def _prune(pc_name: str):
    """Delete day-files older than MAX_DAYS from the active storage."""
    cutoff = (date.today() - timedelta(days=MAX_DAYS)).isoformat()
    d = _user_dir(pc_name)
    for f in glob.glob(os.path.join(d, "????-??-??.json")):
        name = os.path.basename(f).replace(".json", "")
        if name < cutoff:
            try:
                os.remove(f)
                log.debug("Pruned old data file: %s", f)
            except OSError:
                pass
