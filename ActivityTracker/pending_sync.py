"""Track which day-data has not been synced to Google Sheets yet."""

import json
import os

from config import PENDING_SYNC_FILE
from logger_setup import get_logger

log = get_logger("pending_sync")


def _load() -> dict:
    """Load the pending sync state from disk."""
    if not os.path.exists(PENDING_SYNC_FILE):
        return {}
    try:
        with open(PENDING_SYNC_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception as exc:
        log.warning("Failed to load pending sync file: %s", exc)
        return {}


def _save(data: dict) -> None:
    """Save the pending sync state to disk."""
    try:
        os.makedirs(os.path.dirname(PENDING_SYNC_FILE), exist_ok=True)
        tmp = f"{PENDING_SYNC_FILE}.tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        os.replace(tmp, PENDING_SYNC_FILE)
    except Exception as exc:
        log.error("Failed to save pending sync file: %s", exc)


def mark_pending(pc_name: str, day_str: str) -> None:
    """Mark a day as needing sync to Google Sheets."""
    data = _load()
    pending = data.get(pc_name, [])
    if day_str not in pending:
        pending.append(day_str)
        data[pc_name] = pending
        _save(data)
        log.debug("Marked pending sync: %s / %s", pc_name, day_str)


def mark_synced(pc_name: str, day_str: str) -> None:
    """Mark a day as successfully synced."""
    data = _load()
    pending = data.get(pc_name, [])
    if day_str in pending:
        pending.remove(day_str)
        data[pc_name] = pending
        _save(data)
        log.debug("Marked synced: %s / %s", pc_name, day_str)


def get_pending(pc_name: str) -> list[str]:
    """Return list of day strings that need sync."""
    data = _load()
    return list(data.get(pc_name, []))


def clear_all(pc_name: str) -> None:
    """Clear all pending entries for a PC."""
    data = _load()
    if pc_name in data:
        del data[pc_name]
        _save(data)
