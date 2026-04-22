"""Encrypted config storage for sheet_id and pc_name.

Uses Fernet symmetric encryption so the saved file is not plain-text.
A random key is generated on first use and stored in .config_key.
"""

import json
import os

from cryptography.fernet import Fernet

from config import (
    TRACKER_CONFIG_FILE,
    TRACKER_CONFIG_KEY_FILE,
    LEGACY_TRACKER_CONFIG_FILE,
    LEGACY_TRACKER_CONFIG_KEY_FILE,
    SHEET_ID,
)
from logger_setup import get_logger

log = get_logger("config_store")


def _get_or_create_key() -> bytes:
    """Return the Fernet key, creating it if it doesn't exist."""
    if os.path.exists(TRACKER_CONFIG_KEY_FILE):
        with open(TRACKER_CONFIG_KEY_FILE, "rb") as f:
            return f.read()

    if (
        LEGACY_TRACKER_CONFIG_KEY_FILE != TRACKER_CONFIG_KEY_FILE
        and os.path.exists(LEGACY_TRACKER_CONFIG_KEY_FILE)
    ):
        with open(LEGACY_TRACKER_CONFIG_KEY_FILE, "rb") as f:
            key = f.read()
        os.makedirs(os.path.dirname(TRACKER_CONFIG_KEY_FILE), exist_ok=True)
        with open(TRACKER_CONFIG_KEY_FILE, "wb") as f:
            f.write(key)
        return key

    key = Fernet.generate_key()
    os.makedirs(os.path.dirname(TRACKER_CONFIG_KEY_FILE), exist_ok=True)
    with open(TRACKER_CONFIG_KEY_FILE, "wb") as f:
        f.write(key)
    log.info("Generated new config encryption key.")
    return key


def _payload_candidates() -> list[tuple[str, str]]:
    candidates = [(TRACKER_CONFIG_FILE, TRACKER_CONFIG_KEY_FILE)]
    legacy_pair = (LEGACY_TRACKER_CONFIG_FILE, LEGACY_TRACKER_CONFIG_KEY_FILE)
    if legacy_pair != candidates[0]:
        candidates.append(legacy_pair)
    return candidates


def _load_payload() -> dict | None:
    for cfg_path, key_path in _payload_candidates():
        if not os.path.exists(cfg_path) or not os.path.exists(key_path):
            continue
        try:
            with open(key_path, "rb") as f:
                key = f.read()
            fernet = Fernet(key)
            with open(cfg_path, "rb") as f:
                encrypted = f.read()
            payload = fernet.decrypt(encrypted)
            data = json.loads(payload.decode())
            if not isinstance(data, dict):
                continue
            if cfg_path != TRACKER_CONFIG_FILE:
                _save_payload(data)
                log.info("Migrated legacy tracker config to user data directory.")
            return data
        except Exception as exc:
            log.warning("Failed to load encrypted config from %s: %s", cfg_path, exc)
    return None


def _save_payload(payload: dict) -> None:
    key = _get_or_create_key()
    fernet = Fernet(key)
    encrypted = fernet.encrypt(json.dumps(payload).encode())
    os.makedirs(os.path.dirname(TRACKER_CONFIG_FILE), exist_ok=True)
    with open(TRACKER_CONFIG_FILE, "wb") as f:
        f.write(encrypted)


def save_config(pc_name: str, sheet_id: str) -> None:
    """Encrypt and persist tracker config."""
    payload = _load_payload() or {}
    payload.update({"pc_name": pc_name, "sheet_id": sheet_id})
    _save_payload(payload)
    log.info("Tracker config saved (encrypted).")


def load_config() -> dict | None:
    """Load and decrypt tracker config. Returns None if missing or corrupt."""
    data = _load_payload()
    if not data:
        return None
    if data.get("pc_name"):
        if not data.get("sheet_id"):
            # Backward compatibility for older config files.
            data["sheet_id"] = SHEET_ID
        return data
    return None


def save_view_filter(start_text: str, end_text: str) -> None:
    """Persist public-view time filter in encrypted config."""
    payload = _load_payload() or {}
    payload.update({"view_from": start_text, "view_to": end_text})
    _save_payload(payload)


def load_view_filter() -> dict | None:
    """Load persisted public-view time filter if available."""
    data = _load_payload()
    if not data:
        return None
    start_text = data.get("view_from")
    end_text = data.get("view_to")
    if start_text and end_text:
        return {"from": str(start_text), "to": str(end_text)}
    return None


def delete_config() -> None:
    """Remove saved tracker setup while preserving saved view filter."""
    data = _load_payload() or {}
    view_from = data.get("view_from")
    view_to = data.get("view_to")

    if view_from and view_to:
        _save_payload({"view_from": view_from, "view_to": view_to})
        log.info("Tracker config deleted; preserved encrypted view filter.")
        return

    all_paths = {
        TRACKER_CONFIG_FILE,
        TRACKER_CONFIG_KEY_FILE,
        LEGACY_TRACKER_CONFIG_FILE,
        LEGACY_TRACKER_CONFIG_KEY_FILE,
    }
    for path in all_paths:
        if os.path.exists(path):
            os.remove(path)
    log.info("Tracker config deleted.")
