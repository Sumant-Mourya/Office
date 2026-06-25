"""NiceGUI-based dashboard – login, control panel & full-width graphs.

Redesigned for service-account auth, offline-first operation,
and configurable Google Sheet backup.
"""

import asyncio
from bisect import bisect_right
import os
import httpx
from datetime import date, datetime
from urllib.parse import urlparse

from googleapiclient.errors import HttpError
from nicegui import ui, app

from config import (
    FIREBASE_AUTH_URL, SHEET_ID, UI_PORT,
    SERVICE_ACCOUNT_FILE,
)
from config_store import (
    save_config,
    load_config,
    delete_config,
    save_view_filter,
    load_view_filter,
    save_uploaded_sa_filename,
    load_uploaded_sa_filename,
)
from logger_setup import get_logger
from auth.service_account_auth import ServiceAccountAuth
from sheets.sync import SheetSync
from tracker.engine import TrackerEngine
from setup_autostart import (
    install_startup_script,
    remove_startup_script,
    startup_script_exists,
)
from local_store import load_all, list_days, load_day, list_pcs
from app_ui.icons import (
    TRACKER_LOGO, CONSOLE_HEADER_ICON, SHEET_ICON, COMPUTER_ICON,
    KEYBOARD_ICON, MOUSE_ICON, CLOCK_ICON
)

log = get_logger("ui.dashboard")

# ── Shared state ──────────────────────────────────────────────────────
service_auth = ServiceAccountAuth()
tracker_engine: TrackerEngine | None = None
sheet_sync: SheetSync | None = None

_saved = load_config()
_has_saved_config = bool(
    _saved
    and _saved.get("pc_name")
)

# Auto-login if saved config exists to bypass login on start
_state = {
    "logged_in": _has_saved_config,
    "user_id": _saved.get("user_id", "") if _saved else "",
    "pc_name": _saved["pc_name"] if _saved else "",
    "sheet_id": (_saved.get("sheet_id") if _saved else "") or SHEET_ID,
    "sheet_backup_enabled": (_saved.get("sheet_backup_enabled") if _saved else False),
    "tracking": False,
    "config_locked": _has_saved_config,
}

# ── Common Styling (Premium Dark Mode with Google Loader Animation) ─
CSS_STYLES = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700&display=swap');

* {
    font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif !important;
}

body {
    background-color: #0b0c10 !important;
    background-image: radial-gradient(circle at 50% 0%, #171821 0%, #0b0c10 100%) !important;
    color: #f3f4f6 !important;
}

.glass-card {
    background-color: rgba(20, 22, 31, 0.6) !important;
    backdrop-filter: blur(20px) !important;
    -webkit-backdrop-filter: blur(20px) !important;
    border: 1px solid rgba(255, 255, 255, 0.05) !important;
    border-radius: 16px !important;
    box-shadow: 0 10px 40px 0 rgba(0, 0, 0, 0.6) !important;
    padding: 24px !important;
}

.glow-card-green {
    background-color: rgba(20, 22, 31, 0.7) !important;
    border: 1px solid rgba(16, 185, 129, 0.15) !important;
    border-radius: 16px !important;
    box-shadow: 0 4px 30px rgba(16, 185, 129, 0.05) !important;
    padding: 24px !important;
}

.glow-card-indigo {
    background-color: rgba(20, 22, 31, 0.7) !important;
    border: 1px solid rgba(99, 102, 241, 0.15) !important;
    border-radius: 16px !important;
    box-shadow: 0 4px 30px rgba(99, 102, 241, 0.05) !important;
    padding: 24px !important;
}

.app-header {
    background-color: rgba(11, 12, 16, 0.8) !important;
    backdrop-filter: blur(16px) !important;
    border-bottom: 1px solid rgba(255, 255, 255, 0.05) !important;
}

.q-field {
    background-color: rgba(255, 255, 255, 0.02) !important;
    border-radius: 10px !important;
    padding: 2px 10px !important;
    border: 1px solid rgba(255, 255, 255, 0.08) !important;
    transition: all 0.25s ease !important;
}

.q-field--focused {
    border-color: #6366f1 !important;
    box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.25) !important;
}

.q-field__control:before, .q-field__control:after {
    display: none !important;
}

.q-field__native, .q-field__input {
    color: #f3f4f6 !important;
}

.q-field__label {
    color: #9ca3af !important;
}

.btn-primary {
    background: linear-gradient(135deg, #6366f1 0%, #4f46e5 100%) !important;
    color: #ffffff !important;
    font-weight: 600 !important;
    border-radius: 10px !important;
    box-shadow: 0 4px 14px 0 rgba(99, 102, 241, 0.4) !important;
    transition: all 0.25s ease !important;
    text-transform: none !important;
}

.btn-primary:hover {
    box-shadow: 0 6px 20px 0 rgba(99, 102, 241, 0.6) !important;
    transform: translateY(-1px) !important;
}

.btn-success {
    background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
    color: #ffffff !important;
    font-weight: 600 !important;
    border-radius: 10px !important;
    box-shadow: 0 4px 14px 0 rgba(16, 185, 129, 0.4) !important;
    transition: all 0.25s ease !important;
    text-transform: none !important;
}

.btn-success:hover {
    box-shadow: 0 6px 20px 0 rgba(16, 185, 129, 0.6) !important;
    transform: translateY(-1px) !important;
}

.btn-danger {
    background: linear-gradient(135deg, #f43f5e 0%, #e11d48 100%) !important;
    color: #ffffff !important;
    font-weight: 600 !important;
    border-radius: 10px !important;
    box-shadow: 0 4px 14px 0 rgba(244, 63, 94, 0.4) !important;
    transition: all 0.25s ease !important;
    text-transform: none !important;
}

.btn-danger:hover {
    box-shadow: 0 6px 20px 0 rgba(244, 63, 94, 0.6) !important;
    transform: translateY(-1px) !important;
}

.gradient-title {
    background: linear-gradient(135deg, #a78bfa 0%, #6366f1 50%, #3b82f6 100%) !important;
    -webkit-background-clip: text !important;
    -webkit-text-fill-color: transparent !important;
    font-weight: 800 !important;
}

.chart-card {
    background-color: rgba(20, 22, 31, 0.4) !important;
    border: 1px solid rgba(255, 255, 255, 0.03) !important;
    border-radius: 16px !important;
    padding: 16px !important;
}

/* Custom scrollbars */
::-webkit-scrollbar {
    width: 8px;
    height: 8px;
}
::-webkit-scrollbar-track {
    background: #0b0c10;
}
::-webkit-scrollbar-thumb {
    background: #272a37;
    border-radius: 4px;
}
::-webkit-scrollbar-thumb:hover {
    background: #3e445b;
}

/* Google-colored loader animation */
.google-spin-container {
    width: 50px;
    height: 50px;
    position: relative;
    animation: rotate-loader 1.4s linear infinite;
}
.google-spin-circle {
    width: 100%;
    height: 100%;
    border: 4px solid transparent;
    border-radius: 50%;
    border-top-color: #4285F4;
    border-right-color: #EA4335;
    border-bottom-color: #FBBC05;
    border-left-color: #34A853;
}

/* Override Quasar's default Material Icons for dropdowns to handle offline mode */
.q-select__dropdown-icon {
    color: transparent !important;
    font-size: 0 !important;
    display: inline-block !important;
    overflow: hidden !important;
    white-space: nowrap !important;
    letter-spacing: -9999px !important;
    width: 24px !important;
    height: 24px !important;
    background: url("data:image/svg+xml;charset=utf8,%3Csvg viewBox='0 0 24 24' fill='%239ca3af' xmlns='http://www.w3.org/2000/svg'%3E%3Cpath d='M7 10l5 5 5-5z'/%3E%3C/svg%3E") no-repeat center center !important;
}

/* Hide text in q-icon when font is missing */
.q-icon {
    font-size: 0 !important;
}
.q-icon::before {
    font-size: 24px !important;
}

/* Fallback SVG for expansion item toggle (keyboard_arrow_down) */
.q-expansion-item__toggle-icon {
    color: transparent !important;
    font-size: 0 !important;
    display: inline-block !important;
    overflow: hidden !important;
    background: url("data:image/svg+xml;charset=utf8,%3Csvg viewBox='0 0 24 24' fill='%239ca3af' xmlns='http://www.w3.org/2000/svg'%3E%3Cpath d='M7.41 8.59L12 13.17l4.59-4.58L18 10l-6 6-6-6 1.41-1.41z'/%3E%3C/svg%3E") no-repeat center center !important;
    background-size: 24px 24px !important;
    width: 24px !important;
    height: 24px !important;
}
.q-expansion-item--expanded .q-expansion-item__toggle-icon {
    transform: rotate(180deg) !important;
}

@keyframes rotate-loader {
    100% { transform: rotate(360deg); }
}
</style>
"""

# ── Auto-start ────────────────────────────────────────────────────────
def try_auto_start():
    global sheet_sync, tracker_engine
    if not _state["config_locked"]:
        return
    if not _state["pc_name"]:
        return
    if tracker_engine and tracker_engine.is_running:
        return

    sync_cb = None
    backup_enabled = _state["sheet_backup_enabled"]

    if backup_enabled:
        if _state.get("user_id"):
            from firestore_sync import FirestoreSync
            fs_sync = FirestoreSync(_state["user_id"], _state["pc_name"])
            sync_cb = fs_sync.sync
        else:
            log.warning("Auto-start: Backup enabled but no user_id found. Please log in again.")

    try:
        tracker_engine = TrackerEngine(
            pc_name=_state["pc_name"],
            sync_callback=sync_cb,
            sheet_backup_enabled=backup_enabled and sync_cb is not None,
        )
        tracker_engine.start()
        _state["tracking"] = True
        log.info("Auto-started tracking (pc=%s, backup=%s).",
                 _state["pc_name"], backup_enabled)
    except Exception as exc:
        log.error("Auto-start failed: %s", exc)


# ── Remote helpers ────────────────────────────────────────────────────
def _fetch_remote_all(app_url: str, pc_name: str) -> list[dict] | None:
    urls_to_try = [app_url]
    for url in urls_to_try:
        try:
            r = httpx.get(
                f"{url}/api/all/{pc_name}",
                timeout=httpx.Timeout(4.0, connect=1.2),
            )
            if r.status_code == 200:
                return r.json()
        except Exception:
            continue
    return None


def _normalize_remote_url(app_url: str, local_ip: str) -> str:
    if not app_url and local_ip:
        return f"http://{local_ip}:{UI_PORT}"
    if not app_url:
        return ""
    try:
        parsed = urlparse(app_url)
    except Exception:
        return app_url

    host = (parsed.hostname or "").strip().lower()
    port = parsed.port or UI_PORT
    if host in {"", "localhost", "127.0.0.1", "0.0.0.0", "::1"}:
        if local_ip and local_ip not in {"127.0.0.1", "0.0.0.0", "::1"}:
            return f"http://{local_ip}:{port}"
    return app_url


def _read_config_pc_entries() -> list[dict]:
    if not _state["sheet_backup_enabled"]:
        return []
    try:
        if sheet_sync:
            rows = sheet_sync.read_config_pcs()
            if rows:
                return rows
    except Exception as exc:
        log.debug("Failed to read PCs from active sheet sync: %s", exc)

    cfg = load_config() or {}
    sheet_id = (cfg.get("sheet_id") or _state.get("sheet_id") or "").strip()
    if not sheet_id:
        return []

    creds = service_auth.get_credentials()
    if creds is None:
        return []

    sheet_name = (
        cfg.get("pc_name")
        or _state.get("pc_name")
        or "ActivityTracker"
    )
    try:
        temp_sync = SheetSync(creds, sheet_id, str(sheet_name))
        return temp_sync.read_config_pcs()
    except Exception as exc:
        log.debug("Failed to read PCs from config fallback: %s", exc)
        return []


def _build_pc_map(include_local_data_dirs: bool = True) -> dict[str, dict]:
    merged: dict[str, dict] = {}
    for row in _read_config_pc_entries():
        pc_name = str(row.get("pc_name", "")).strip()
        if not pc_name:
            continue
        merged[pc_name] = {
            "pc_name": pc_name,
            "local_ip": str(row.get("local_ip", "")).strip(),
            "app_url": str(row.get("app_url", "")).strip(),
            "last_seen": str(row.get("last_seen", "")).strip(),
        }

    if include_local_data_dirs:
        for pc_name in list_pcs():
            if pc_name not in merged:
                merged[pc_name] = {
                    "pc_name": pc_name,
                    "local_ip": "",
                    "app_url": "",
                    "last_seen": "",
                }

    local_pc = (_state.get("pc_name") or "").strip()
    if local_pc:
        local_row = merged.setdefault(
            local_pc,
            {
                "pc_name": local_pc,
                "local_ip": "127.0.0.1",
                "app_url": f"http://127.0.0.1:{UI_PORT}",
                "last_seen": "",
            },
        )
        if not local_row.get("app_url"):
            local_row["app_url"] = f"http://127.0.0.1:{UI_PORT}"
        if not local_row.get("local_ip"):
            local_row["local_ip"] = "127.0.0.1"
    return merged


def _parse_slot_minutes(slot: str) -> tuple[int, int] | None:
    try:
        start_str, end_str = [s.strip() for s in slot.split("-", 1)]
        start_dt = datetime.strptime(start_str.upper(), "%I:%M%p")
        end_dt = datetime.strptime(end_str.upper(), "%I:%M%p")
        start_min = start_dt.hour * 60 + start_dt.minute
        end_min = end_dt.hour * 60 + end_dt.minute
        if end_min <= start_min:
            end_min += 24 * 60
        return start_min, end_min
    except Exception:
        return None


def _slot_sort_key(slot: str) -> tuple[int, str]:
    parsed = _parse_slot_minutes(slot)
    if parsed:
        return parsed[0], slot
    return 24 * 60 + 1, slot


def _to_minutes(hour_text: str, minute_text: str, period_text: str) -> int:
    hour = int(hour_text)
    minute = int(minute_text)
    period = (period_text or "AM").upper()
    if hour == 12:
        hour = 0
    if period == "PM":
        hour += 12
    return hour * 60 + minute


def _parse_ampm_text(value: str) -> int | None:
    cleaned = (value or "").strip().lower().replace(" ", "")
    try:
        parsed = datetime.strptime(cleaned, "%I:%M%p")
    except Exception:
        return None
    return parsed.hour * 60 + parsed.minute


def _format_ampm(minutes: int) -> str:
    value = minutes % (24 * 60)
    hour24 = value // 60
    minute = value % 60
    period = "am" if hour24 < 12 else "pm"
    hour12 = hour24 % 12
    if hour12 == 0:
        hour12 = 12
    return f"{hour12}:{minute:02d}{period}"


def _build_segments(start_min: int, end_min: int) -> list[tuple[int, int, str]]:
    segments: list[tuple[int, int, str]] = []
    cursor = start_min
    while cursor < end_min:
        next_hour = ((cursor // 60) + 1) * 60
        seg_end = min(end_min, next_hour)
        segments.append((cursor, seg_end, f"{_format_ampm(cursor)}-{_format_ampm(seg_end)}"))
        cursor = seg_end
    return segments


def _aggregate_range(day_data: dict, start_min: int, end_min: int) -> dict:
    segments = _build_segments(start_min, end_min)
    mouse_vals = [0.0] * len(segments)
    key_vals = [0.0] * len(segments)
    work_vals = [0.0] * len(segments)
    idle_vals = [0.0] * len(segments)
    no_run_vals = [0.0] * len(segments)
    segment_ends = [seg_end for _, seg_end, _ in segments]
    segment_seconds = [(seg_end - seg_start) * 60.0 for seg_start, seg_end, _ in segments]
    has_ticks = False

    hours = day_data.get("hours", {})
    for hour_data in hours.values():
        ticks = hour_data.get("ticks", []) or []
        for tick in ticks:
            ts = tick.get("ts")
            if ts is None:
                continue
            try:
                dt = datetime.fromtimestamp(float(ts))
            except Exception:
                continue

            minute_point = dt.hour * 60 + dt.minute + (dt.second / 60.0)
            if minute_point < start_min or minute_point >= end_min:
                continue

            idx = bisect_right(segment_ends, minute_point)
            if idx >= len(segments):
                idx = len(segments) - 1

            mouse_vals[idx] += float(tick.get("mouse", 0) or 0)
            key_vals[idx] += float(tick.get("keyboard", 0) or 0)

            is_idle = bool(tick.get("idle", False))
            if is_idle:
                idle_vals[idx] += 1.0
            elif tick.get("app"):
                work_vals[idx] += 1.0
            else:
                idle_vals[idx] += 1.0
            has_ticks = True

    if not has_ticks:
        for slot, hour_data in hours.items():
            parsed = _parse_slot_minutes(slot)
            if not parsed:
                continue
            slot_start, slot_end = parsed
            span = max(slot_end - slot_start, 1)
            for idx, (seg_start, seg_end, _) in enumerate(segments):
                overlap = max(0, min(slot_end, seg_end) - max(slot_start, seg_start))
                if overlap <= 0:
                    continue
                ratio = overlap / span
                mouse_vals[idx] += float(hour_data.get("mouse_clicks", 0) or 0) * ratio
                key_vals[idx] += float(hour_data.get("key_presses", 0) or 0) * ratio
                work_vals[idx] += float(hour_data.get("work_seconds", 0) or 0) * ratio
                idle_vals[idx] += float(hour_data.get("idle_seconds", 0) or 0) * ratio

    for idx in range(len(segments)):
        tracked = min(segment_seconds[idx], work_vals[idx] + idle_vals[idx])
        no_run_vals[idx] = max(segment_seconds[idx] - tracked, 0.0)

    labels = [label for _, _, label in segments]
    mouse_int = [int(round(v)) for v in mouse_vals]
    key_int = [int(round(v)) for v in key_vals]

    return {
        "labels": labels,
        "mouse": mouse_int,
        "keys": key_int,
        "work_minutes": [round(v / 60.0, 1) for v in work_vals],
        "idle_minutes": [round(v / 60.0, 1) for v in idle_vals],
        "no_run_minutes": [round(v / 60.0, 1) for v in no_run_vals],
        "total_mouse": sum(mouse_int),
        "total_keys": sum(key_int),
        "total_work_seconds": sum(work_vals),
        "total_idle_seconds": sum(idle_vals),
        "total_no_run_seconds": sum(no_run_vals),
    }


# =====================================================================
#  LOGIN PAGE (Bypasses broken Quasar toggles with a text button)
# =====================================================================
@ui.page("/")
def login_page():
    if _state["logged_in"]:
        ui.navigate.to("/dashboard")
        return

    ui.dark_mode(True)
    ui.add_head_html(CSS_STYLES)

    with ui.column().classes("absolute-center items-center w-full max-w-md p-4 gap-6"):
        ui.html(TRACKER_LOGO)
        
        with ui.column().classes("items-center gap-0"):
            ui.label("ACTIVITY TRACKER").classes("text-3xl font-extrabold tracking-wider gradient-title")
            ui.label("Enterprise Productivity Engine").classes("text-xs tracking-widest text-gray-500 uppercase font-semibold")

        with ui.card().classes("glass-card w-full gap-5"):
            ui.label("System Access").classes("text-lg font-semibold text-white tracking-wide border-b border-gray-800 pb-2")
            
            # Login Form
            with ui.column().classes("w-full gap-4"):
                username_input = ui.input("Username").classes("w-full").props("dark label-color=gray-400")
                
                with ui.row().classes("w-full items-center gap-2"):
                    password_input = ui.input("Password", password=True).classes("flex-1").props("dark label-color=gray-400")
                    
                    visibility_state = {"visible": False}
                    
                    def toggle_visibility():
                        visibility_state["visible"] = not visibility_state["visible"]
                        is_visible = visibility_state["visible"]
                        password_input.props(f'type={"text" if is_visible else "password"}')
                        toggle_label.set_text("Hide" if is_visible else "Show")

                    with ui.button(on_click=toggle_visibility).props("flat color=primary").classes("text-xs font-semibold px-2"):
                        toggle_label = ui.label("Show")
                
            error_label = ui.label("").classes("text-red-400 text-sm font-semibold")

            async def handle_login():
                if not username_input.value or not password_input.value:
                    error_label.set_text("⚠ Please enter username and password.")
                    return
                try:
                    from config import FIREBASE_FIRESTORE_URL
                    async with httpx.AsyncClient() as client:
                        url = f"{FIREBASE_FIRESTORE_URL}:runQuery"
                        payload = {
                            "structuredQuery": {
                                "from": [{"collectionId": "users"}],
                                "where": {
                                    "compositeFilter": {
                                        "op": "AND",
                                        "filters": [
                                            {
                                                "fieldFilter": {
                                                    "field": {"fieldPath": "username"},
                                                    "op": "EQUAL",
                                                    "value": {"stringValue": username_input.value}
                                                }
                                            },
                                            {
                                                "fieldFilter": {
                                                    "field": {"fieldPath": "password"},
                                                    "op": "EQUAL",
                                                    "value": {"stringValue": password_input.value}
                                                }
                                            }
                                        ]
                                    }
                                }
                            }
                        }
                        resp = await client.post(url, json=payload)
                        if resp.status_code == 200:
                            data = resp.json()
                            if data and len(data) > 0 and "document" in data[0]:
                                doc_name = data[0]["document"]["name"]
                                user_doc_id = doc_name.split("/")[-1]
                                _state["logged_in"] = True
                                _state["user_id"] = user_doc_id
                                ui.navigate.to("/dashboard")
                            else:
                                error_label.set_text("⚠ Incorrect username or password.")
                                ui.notify("❌ Login failed", color="red")
                        else:
                            error_label.set_text("⚠ Error communicating with server.")
                            ui.notify("❌ Login failed", color="red")
                except Exception as e:
                    log.error(f"Login error: {e}")
                    error_label.set_text("⚠ Network error during login.")

            password_input.on("keydown.enter", handle_login)
            username_input.on("keydown.enter", handle_login)

            ui.button("Login to Console", on_click=handle_login).classes("w-full btn-primary py-2 mt-2")


# =====================================================================
#  DASHBOARD (Redesigned with Asynchronous Live Validation & Google Loading)
# =====================================================================
@ui.page("/dashboard")
def dashboard_page():
    if not _state["logged_in"]:
        ui.navigate.to("/")
        return

    ui.dark_mode(True)
    ui.add_head_html(CSS_STYLES)

    # ── Google-style Loading Animation Overlay ──
    with ui.element('div').classes('absolute top-0 left-0 w-full h-full bg-[#0b0c10] z-[99999] flex flex-col justify-center items-center gap-4').style('transition: opacity 0.8s ease, visibility 0.8s ease;') as loader_overlay:
        ui.html('<div class="google-spin-container"><div class="google-spin-circle"></div></div>')
        ui.label("Loading Activities Console...").classes("text-gray-400 font-semibold tracking-wider text-sm")
    
    # Auto fade out the Google loader overlay
    async def remove_loader():
        await asyncio.sleep(1.2)
        loader_overlay.style('opacity: 0; visibility: hidden; pointer-events: none;')
    ui.timer(0.05, remove_loader, once=True)

    # ── Header (Uses simple clean text without Quasar icons) ───────────────
    with ui.header().classes("app-header items-center justify-between px-6 py-3"):
        with ui.row().classes("items-center gap-3"):
            ui.html(CONSOLE_HEADER_ICON)
            ui.label("Activity Tracker Console").classes("text-lg font-bold text-white tracking-wide")
        with ui.row().classes("gap-3"):
            ui.button("Graphs", on_click=lambda: ui.navigate.to("/graphs")).props("flat color=white").classes("text-sm font-semibold")
            ui.button("Logout", on_click=_logout).props("flat color=red").classes("text-sm font-semibold")

    # ── Main Content Container ───────────────
    with ui.column().classes("w-full max-w-4xl mx-auto p-6 gap-6"):
        
        # 1. PC & Sheet Configuration Card
        with ui.card().classes("glass-card w-full gap-4"):
            ui.label("1. Sheet Sync & Computer Setup").classes("text-xl font-bold text-white tracking-wide border-b border-gray-800 pb-2")
            
            locked = _state["config_locked"]

            # Google Sheet Backup Checkbox (PLACED FIRST)
            backup_checkbox = ui.checkbox(
                "Enable Google Sheet Backup (Hourly Sync)",
                value=_state["sheet_backup_enabled"],
            ).classes("text-white font-medium")
            if locked:
                backup_checkbox.props("disable")

            # Google Sheet ID Input (conditionally visible)
            with ui.row().classes("w-full items-center gap-2") as sheet_row:
                ui.html(SHEET_ICON)
                sheet_input = ui.input(
                    "Google Sheet ID",
                    value=_state["sheet_id"],
                ).classes("flex-1").props("dark")
            if locked:
                sheet_input.props("readonly")

            # Initialize upload mode state
            upload_mode = {"active": not os.path.exists(SERVICE_ACCOUNT_FILE)}

            @ui.refreshable
            def render_sa_section():
                file_exists = os.path.exists(SERVICE_ACCOUNT_FILE)
                if file_exists and not upload_mode["active"]:
                    orig_fn = load_uploaded_sa_filename() or "service_account.json"
                    with ui.row().classes("items-center justify-between w-full bg-emerald-950/20 border border-emerald-500/20 p-3 rounded-lg"):
                        with ui.row().classes("items-center gap-2"):
                            ui.label("✓").classes("text-emerald-400 font-bold text-lg")
                            with ui.column().classes("gap-0"):
                                ui.label("Credentials File Configured").classes("text-xs text-gray-400")
                                ui.label(orig_fn).classes("text-sm font-semibold text-emerald-400")
                        
                        def click_change():
                            upload_mode["active"] = True
                            render_sa_section.refresh()
                            asyncio.create_task(validate_form())

                        ui.button("Change File", on_click=click_change).props("flat color=primary").classes("text-xs font-semibold")
                else:
                    ui.label("⚠️ Please add service_account.json file").classes("text-rose-400 text-sm font-semibold")
                    
                    async def handle_upload(e):
                        try:
                            if not e.file.name.endswith('.json'):
                                ui.notify("❌ Invalid file type. Please select a .json file.", color="red")
                                return
                            # Copy file to AppData
                            os.makedirs(os.path.dirname(SERVICE_ACCOUNT_FILE), exist_ok=True)
                            await e.file.save(SERVICE_ACCOUNT_FILE)
                            
                            # Store original filename
                            save_uploaded_sa_filename(e.file.name)
                            
                            # Reload auth
                            service_auth.reload()
                            ui.notify(f"✅ Uploaded and copied {e.file.name} to appdata!", color="green")
                            
                            upload_mode["active"] = False
                            render_sa_section.refresh()
                            await validate_form()
                        except Exception as exc:
                            log.error("File upload failed: %s", exc)
                            ui.notify(f"❌ Upload failed: {exc}", color="red")

                    # The invisible uploader
                    uploader = ui.upload(
                        auto_upload=True,
                        max_files=1,
                        on_upload=handle_upload,
                    ).classes("hidden").props('accept=.json')
                    # Sleek, functional text button that calls pickFiles programmatically
                    ui.button("Choose Credentials File", on_click=lambda: uploader.run_method("pickFiles")).classes("w-full btn-primary py-2")

            # Service Account credential container (Uses a hidden upload picker to bypass icon errors)
            with ui.column().classes("w-full gap-2") as sa_container:
                render_sa_section()

            # PC Name input
            with ui.row().classes("w-full items-center gap-2"):
                ui.html(COMPUTER_ICON)
                pc_input = ui.input(
                    "PC Name (Unique name for this computer)",
                    value=_state["pc_name"]
                ).classes("flex-1").props("dark")
            if locked:
                pc_input.props("readonly")

            # Live Verification Status
            validation_status = ui.markdown("Checking current setup...").classes("text-sm font-semibold px-3 py-2 rounded-lg bg-gray-900/50 border border-gray-800 w-full")

            # Save settings button
            save_btn = ui.button("Save Settings", on_click=lambda: on_save()).classes("w-full py-2 btn-success")
            save_btn.disable()
            
            # Helper for displaying save success messages
            save_msg = ui.label("").classes("text-sm font-semibold")

            # ── Form Validation Logic (Asynchronous & Reactive) ──
            async def validate_form():
                pc_val = pc_input.value.strip()
                sheet_val = sheet_input.value.strip()
                backup_on = backup_checkbox.value
                sa_exists = os.path.exists(SERVICE_ACCOUNT_FILE)

                # Reset UI state (Properly disables immediately on load/reset)
                save_btn.disable()
                validation_status.set_content("Validating form fields...")
                validation_status.classes("text-yellow-400", remove="text-red-400 text-green-400")

                if not pc_val:
                    validation_status.set_content("⚠ PC Name is required.")
                    validation_status.classes("text-red-400", remove="text-yellow-400 text-green-400")
                    save_btn.disable()
                    return

                # Validation when backup is disabled
                if not backup_on:
                    loop = asyncio.get_event_loop()
                    local_pcs = await loop.run_in_executor(None, list_pcs)
                    
                    saved_cfg = load_config()
                    current_owned_pc = saved_cfg.get("pc_name", "") if saved_cfg else ""

                    # Enforce that pc_name does not already exist in local database
                    if pc_val in local_pcs and pc_val != current_owned_pc:
                        validation_status.set_content(
                            f"⚠ PC Name `'{pc_val}'` already exists in the local database. Please use a unique name."
                        )
                        validation_status.classes("text-red-400", remove="text-yellow-400 text-green-400")
                        save_btn.disable()
                        return
                    else:
                        validation_status.set_content("✓ Validation passed. Ready to save settings.")
                        validation_status.classes("text-green-400", remove="text-yellow-400 text-red-400")
                        save_btn.enable()
                        return

                # If Backup is enabled, enforce sheet_id and service_account
                if not sa_exists:
                    validation_status.set_content("⚠ Please add service_account.json file.")
                    validation_status.classes("text-red-400", remove="text-yellow-400 text-green-400")
                    save_btn.disable()
                    return

                if not sheet_val:
                    validation_status.set_content("⚠ Please enter Google Sheet ID.")
                    validation_status.classes("text-red-400", remove="text-yellow-400 text-green-400")
                    save_btn.disable()
                    return

                validation_status.set_content("Checking spreadsheet connection & verifying PC name availability...")
                validation_status.classes("text-yellow-400", remove="text-red-400 text-green-400")

                try:
                    service_auth.reload()
                    creds = service_auth.get_credentials()
                    if not creds:
                        validation_status.set_content("⚠ Failed to load service account credentials.")
                        validation_status.classes("text-red-400", remove="text-yellow-400 text-green-400")
                        save_btn.disable()
                        return

                    # Run Sheets check in background thread
                    loop = asyncio.get_event_loop()
                    temp_sync = SheetSync(creds, sheet_val, pc_val)

                    access_ok = await loop.run_in_executor(None, temp_sync.validate_access)
                    if not access_ok:
                        email = service_auth.service_account_email or "service account email"
                        validation_status.set_content(
                            f"⚠ Cannot access Google Sheet.<br>Please ensure Sheet ID is correct "
                            f"and that you shared the sheet with:<br><b class='text-cyan-300'>{email}</b>"
                        )
                        validation_status.classes("text-red-400", remove="text-yellow-400 text-green-400")
                        save_btn.disable()
                        return

                    # Validate if PC Name is available (not taken by another PC)
                    is_taken = await loop.run_in_executor(None, temp_sync.pc_name_taken, pc_val)
                    if is_taken:
                        cfg = load_config()
                        if cfg and cfg.get("pc_name") == pc_val and cfg.get("sheet_id") == sheet_val:
                            # Resuming owned config
                            pass
                        else:
                            validation_status.set_content(f"⚠ PC Name `'{pc_val}'` is already claimed by another PC on this sheet.")
                            validation_status.classes("text-red-400", remove="text-yellow-400 text-green-400")
                            save_btn.disable()
                            return

                    # All validations passed!
                    validation_status.set_content("✓ Validation passed! Google Sheet is reachable and PC Name is available.")
                    validation_status.classes("text-green-400", remove="text-yellow-400 text-red-400")
                    save_btn.enable()
                except HttpError as exc:
                    email = service_auth.service_account_email or "service account email"
                    if exc.resp is not None and exc.resp.status == 403:
                        validation_status.set_content(
                            f"⚠ Access Denied (403). Share Google Sheet with:<br><b class='text-cyan-300'>{email}</b>"
                        )
                    else:
                        validation_status.set_content(f"⚠ Connection failed: {exc.reason or str(exc)}")
                    validation_status.classes("text-red-400", remove="text-yellow-400 text-green-400")
                    save_btn.disable()
                except Exception as exc:
                    validation_status.set_content(f"⚠ Verification error: {exc}")
                    validation_status.classes("text-red-400", remove="text-yellow-400 text-green-400")
                    save_btn.disable()

            # ── Render Service Account Upload / Status ──
            # (Now handled by the refreshable render_sa_section declared above)
            # ── Input Visibility and Event Listeners ──
            def update_sheet_visibility():
                if backup_checkbox.value:
                    sheet_row.set_visibility(True)
                    sa_container.set_visibility(True)
                else:
                    sheet_row.set_visibility(False)
                    sa_container.set_visibility(False)

            # Bind input change listeners to validate_form
            pc_input.on_value_change(lambda _: asyncio.create_task(validate_form()))
            sheet_input.on_value_change(lambda _: asyncio.create_task(validate_form()))
            
            async def toggle_backup(e):
                update_sheet_visibility()
                await validate_form()

            backup_checkbox.on_value_change(toggle_backup)

            # Render SA elements & Initial Checks
            update_sheet_visibility()
            
            # Perform initial validation on load
            if not locked:
                ui.timer(0.2, lambda: asyncio.create_task(validate_form()), once=True)

            # ── Save Settings Action ──
            async def on_save():
                pc_val = pc_input.value.strip()
                sheet_val = sheet_input.value.strip()
                backup_on = backup_checkbox.value

                _state["pc_name"] = pc_val
                _state["sheet_id"] = sheet_val
                _state["sheet_backup_enabled"] = backup_on
                save_config(pc_val, sheet_val, backup_on, _state.get("user_id", ""))
                _state["config_locked"] = True

                pc_input.props("readonly")
                sheet_input.props("readonly")
                backup_checkbox.props("disable")

                save_msg.set_text("Settings saved successfully ✓")
                save_msg.classes("text-emerald-400")
                ui.notify("✅ Settings saved!", color="green")

                # Lock the save button and update text
                save_btn.disable()
                upload_mode["active"] = False
                render_sa_section.refresh()

            if locked:
                validation_status.set_content("✓ Settings loaded from saved config (locked)")
                validation_status.classes("text-emerald-400", remove="text-yellow-400 text-red-400")
                save_btn.disable()

        # 2. Tracking Control Card
        with ui.card().classes("glass-card w-full gap-4"):
            ui.label("2. Tracking Engine Control").classes("text-xl font-bold text-white tracking-wide border-b border-gray-800 pb-2")
            
            tracking_label = ui.label()
            _update_tracking_label(tracking_label)
            
            startup_status = ui.label("")
            _update_startup_label(startup_status)

            async def on_start():
                global tracker_engine
                pc_val = _state["pc_name"]
                if not pc_val or not _state["config_locked"]:
                    ui.notify("⚠️ Please save setup configuration first.", color="orange")
                    return

                start_btn.props("loading disable")
                await asyncio.sleep(0.1)

                sync_cb = None
                backup_enabled = _state["sheet_backup_enabled"]

                try:
                    if backup_enabled:
                        if _state.get("user_id"):
                            from firestore_sync import FirestoreSync
                            fs_sync = FirestoreSync(_state["user_id"], pc_val)
                            sync_cb = fs_sync.sync
                        else:
                            ui.notify("❌ User ID not found. Please log in again.", color="red")
                            return

                    tracker_engine = TrackerEngine(
                        pc_name=pc_val,
                        sync_callback=sync_cb,
                        sheet_backup_enabled=backup_enabled and sync_cb is not None,
                    )
                    tracker_engine.start()
                    _state["tracking"] = True
                    _update_tracking_label(tracking_label)
                    
                    start_btn.props("disable")
                    stop_btn.props(remove="disable")

                    ok = await asyncio.to_thread(install_startup_script)
                    _update_startup_label(startup_status)

                    msg = "✅ Tracking engine active!"
                    if backup_enabled:
                        msg += " Uploading metrics to Firestore."
                    if ok:
                        msg += " Auto-run shortcuts registered."
                    ui.notify(msg, color="green")
                except Exception as exc:
                    log.error("Tracker failed to start: %s", exc)
                    ui.notify(f"❌ Engine start failed: {exc}", color="red")
                    start_btn.props(remove="loading disable")
                finally:
                    start_btn.props(remove="loading")

            async def on_stop():
                global tracker_engine
                if tracker_engine and tracker_engine.is_running:
                    tracker_engine.stop()
                _state["tracking"] = False
                _update_tracking_label(tracking_label)
                await asyncio.to_thread(remove_startup_script)
                _update_startup_label(startup_status)
                delete_config()
                _state["config_locked"] = False

                pc_input.props(remove="readonly")
                sheet_input.props(remove="readonly")
                backup_checkbox.props(remove="disable")
                
                stop_btn.props("disable")
                start_btn.props(remove="disable")
                
                ui.notify("ℹ️ Tracking stopped. Configuration unlocked.", color="blue")
                await validate_form()

            with ui.row().classes("gap-4 w-full"):
                start_btn = ui.button("Start Tracking", on_click=on_start).classes("btn-success flex-1 py-2 text-sm font-semibold")
                stop_btn = ui.button("Stop & Unlock Settings", on_click=on_stop).classes("btn-danger flex-1 py-2 text-sm font-semibold")
                
                if _state["tracking"]:
                    start_btn.props("disable")
                else:
                    stop_btn.props("disable")

        # 3. Live Activity Stats
        with ui.card().classes("glass-card w-full gap-4"):
            ui.label("Live Machine Activity").classes("text-xl font-bold text-white tracking-wide border-b border-gray-800 pb-2")
            status_area = ui.markdown("").classes("text-gray-300 w-full")

            def refresh_status():
                if tracker_engine and tracker_engine.is_running:
                    snap = tracker_engine.get_today_snapshot()
                    lines = [
                        f"**Current Date:** `{snap['date']}`  ",
                        f"**Work Time:** `{_fmt(snap['total_work'])}` · **Idle Time:** `{_fmt(snap['total_idle'])}`  ",
                        f"**Input Density:** Mouse Clicks: `{snap['total_mouse_clicks']}` | Key Presses: `{snap['total_key_presses']}`  ",
                        "",
                        "#### Hourly Aggregated Metric Slots",
                    ]
                    for slot in sorted(snap.get("hours", {}).keys(), key=_slot_sort_key):
                        h = snap["hours"][slot]
                        lines.append(
                            f"**{slot}** — "
                            f"Clicks: `{h['mouse_clicks']}`, "
                            f"Keys: `{h['key_presses']}`, "
                            f"Work: `{_fmt(h.get('work_seconds', 0))}`, "
                            f"Idle: `{_fmt(h.get('idle_seconds', 0))}`"
                        )
                    status_area.set_content("\n".join(lines))
                else:
                    status_area.set_content("*Tracker is currently inactive. Run the engine to see stats.*")

            ui.timer(5, refresh_status)
            refresh_status()


# =====================================================================
#  GRAPHS PAGE (Uses standard text controls to avoid broken CDN icons)
# =====================================================================
@ui.page("/graphs")
def graphs_page():
    if not _state["logged_in"]:
        ui.navigate.to("/")
        return

    ui.dark_mode(True)
    ui.add_head_html(CSS_STYLES)

    _saved_filter = load_view_filter() or {}
    time_options = [_format_ampm(m) for m in range(0, 24 * 60, 5)]
    picker_values = {
        "start_text": str(_saved_filter.get("from") or "9:00am"),
        "end_text": str(_saved_filter.get("to") or "5:30pm"),
    }
    picker_refs: dict[str, object] = {}
    last_saved_filter = {
        "from": picker_values["start_text"],
        "to": picker_values["end_text"],
    }

    def _selected_minutes() -> tuple[int, int] | None:
        start = _parse_ampm_text(str(picker_refs["start_text"].value))
        end = _parse_ampm_text(str(picker_refs["end_text"].value))
        if start is None or end is None:
            return None
        if end <= start:
            return None
        return start, end

    # ── Header ───────────────
    with ui.header().classes("app-header items-center justify-between px-6 py-3"):
        with ui.row().classes("items-center gap-4 flex-wrap"):
            with ui.row().classes("items-center gap-3"):
                ui.html(CONSOLE_HEADER_ICON)
                ui.label("Analytics Dashboard").classes("text-xl font-bold text-white tracking-wide")

            with ui.row().classes("items-center gap-2 bg-gray-900 border border-gray-800 rounded px-3 py-1 ml-4"):
                picker_refs["start_text"] = ui.select(
                    time_options,
                    label="From Time",
                    value=picker_values["start_text"],
                    on_change=lambda _: render_graphs(),
                ).classes("w-32").props("dark dense")
                picker_refs["end_text"] = ui.select(
                    time_options,
                    label="To Time",
                    value=picker_values["end_text"],
                    on_change=lambda _: render_graphs(),
                ).classes("w-32").props("dark dense")

        with ui.row().classes("gap-3 items-center"):
            ui.button("Control Panel", on_click=lambda: ui.navigate.to("/dashboard")).props("flat color=white").classes("text-sm font-semibold")
            ui.button("Logout", on_click=_logout).props("flat color=red").classes("text-sm font-semibold")

    pc_map = _build_pc_map(include_local_data_dirs=True)
    pc_names = list(pc_map.keys())
    if _state["pc_name"] in pc_map:
        pc_names = [_state["pc_name"]] + [n for n in pc_names if n != _state["pc_name"]]

    selected_pc = {"value": _state["pc_name"] or (pc_names[0] if pc_names else "")}
    last_good_data_by_pc: dict[str, list[dict]] = {}

    with ui.column().classes("w-full p-6 gap-6"):

        # Control Row
        with ui.card().classes("glass-card w-full row items-center justify-between gap-4 p-4"):
            with ui.row().classes("items-center gap-4"):
                if pc_names:
                    def on_pc_change(e):
                        selected_pc["value"] = e.value
                        render_graphs()

                    ui.select(
                        pc_names,
                        value=selected_pc["value"],
                        label="Source Computer",
                        on_change=on_pc_change,
                    ).classes("w-72").props("dark")

                status_chip = ui.label("").classes("text-sm font-bold px-3 py-1 rounded-full")
            
            ui.button("Force Refresh", on_click=lambda: render_graphs()).classes("btn-primary text-xs font-semibold py-1.5 px-3")

        data_container = ui.column().classes("w-full gap-4")

        def render_graphs():
            data_container.clear()
            
            selected = _selected_minutes()
            if selected is None:
                with data_container:
                    with ui.card().classes("w-full p-4 bg-rose-950/20 border border-rose-500/20"):
                        ui.label("Invalid Time Boundaries").classes("text-lg font-bold text-rose-400")
                        ui.label("Please check format (e.g. 9:00am) and make sure 'To' boundary occurs after 'From'.").classes("text-sm text-gray-300")
                return

            start_min, end_min = selected
            current_from = _format_ampm(start_min)
            current_to = _format_ampm(end_min)
            if (
                last_saved_filter["from"] != current_from
                or last_saved_filter["to"] != current_to
            ):
                save_view_filter(current_from, current_to)
                last_saved_filter["from"] = current_from
                last_saved_filter["to"] = current_to

            pc = selected_pc["value"]
            if not pc:
                with data_container:
                    ui.label("No source computer selected.").classes("text-gray-400 italic")
                return

            info = pc_map.get(pc, {})
            app_url = _normalize_remote_url(
                info.get("app_url", ""),
                info.get("local_ip", ""),
            )
            is_local = pc == _state["pc_name"]
            all_days: list[dict] = []
            online = is_local
            
            if is_local:
                all_days = load_all(pc)
            else:
                remote_days = _fetch_remote_all(app_url, pc) if app_url else None
                online = remote_days is not None
                if remote_days is not None:
                    all_days = remote_days
                    last_good_data_by_pc[pc] = remote_days
                else:
                    all_days = load_all(pc) or last_good_data_by_pc.get(pc, [])

            if online:
                status_chip.set_text("● Online Sync")
                status_chip.classes("text-emerald-400 bg-emerald-950/30 border border-emerald-500/20", remove="text-red-400 bg-red-950/30 border-red-500/20")
            else:
                status_chip.set_text("○ Offline Mode")
                status_chip.classes("text-rose-400 bg-rose-950/30 border border-rose-500/20", remove="text-emerald-400 bg-emerald-950/30 border-emerald-500/20")

            if not online and not is_local and all_days:
                with data_container:
                    with ui.card().classes("w-full p-3 bg-gray-900 border border-yellow-500/20"):
                        ui.label(f"⚠ App cannot reach {pc}. Displaying cached offline statistics.").classes("text-sm text-yellow-400 font-medium")

            if not all_days:
                with data_container:
                    ui.label("No productivity tracking logs found for this system.").classes("text-gray-500 italic p-4")
                return

            today_str = date.today().isoformat()

            with data_container:
                for day_data in reversed(all_days):
                    day_str = day_data.get("date", "")
                    hours = day_data.get("hours", {})
                    is_today = day_str == today_str

                    with ui.expansion(
                        f"📅 Log: {day_str}" + (" (Today)" if is_today else ""),
                        value=is_today,
                    ).classes("w-full bg-gray-900/60 border border-gray-800 rounded-xl overflow-hidden"):
                        
                        sliced = _aggregate_range(day_data, start_min, end_min)

                        total_work = sliced["total_work_seconds"]
                        total_idle = sliced["total_idle_seconds"]
                        total_no_run = sliced["total_no_run_seconds"]
                        total_mouse = sliced["total_mouse"]
                        total_keys = sliced["total_keys"]

                        with ui.row().classes("w-full items-center justify-between p-4 bg-gray-900/40 border-b border-gray-800"):
                            with ui.row().classes("gap-4"):
                                ui.markdown(f"**Productive Work:** `{_fmt(total_work)}`").classes("text-sm text-emerald-400 font-semibold")
                                ui.markdown(f"**System Idle:** `{_fmt(total_idle)}`").classes("text-sm text-rose-400 font-semibold")
                                ui.markdown(f"**Inactive (Off):** `{_fmt(total_no_run)}`").classes("text-sm text-gray-400 font-semibold")
                            with ui.row().classes("gap-4"):
                                ui.markdown(f"**Clicks:** `{total_mouse}`").classes("text-sm text-cyan-400 font-semibold")
                                ui.markdown(f"**Keystrokes:** `{total_keys}`").classes("text-sm text-indigo-400 font-semibold")

                        if not sliced["labels"]:
                            ui.label("No metrics gathered in this timeframe.").classes("text-gray-500 italic p-4")
                            continue

                        chart_labels = sliced["labels"]
                        mouse_vals = sliced["mouse"]
                        key_vals = sliced["keys"]
                        work_vals = sliced["work_minutes"]
                        idle_vals = sliced["idle_minutes"]
                        no_run_vals = sliced["no_run_minutes"]

                        with ui.row().classes("w-full gap-4 p-4"):
                            # Chart 1: Input density
                            with ui.card().classes("chart-card flex-1"):
                                with ui.row().classes("items-center gap-1 mb-2"):
                                    ui.html(KEYBOARD_ICON).classes("text-gray-400")
                                    ui.label("Clicks & Keystroke Density").classes("text-sm font-semibold text-gray-300")
                                ui.echart({
                                    "backgroundColor": "transparent",
                                    "tooltip": {"trigger": "axis"},
                                    "legend": {"data": ["Mouse Clicks", "Key Presses"], "textStyle": {"color": "#9ca3af"}, "bottom": 0},
                                    "grid": {"left": "3%", "right": "3%", "bottom": "25%", "top": "5%", "containLabel": True},
                                    "xAxis": {"type": "category", "data": chart_labels, "axisLabel": {"color": "#9ca3af", "rotate": 20}},
                                    "yAxis": {"type": "value", "axisLabel": {"color": "#9ca3af"}},
                                    "series": [
                                        {"name": "Mouse Clicks", "type": "bar", "stack": "input", "data": mouse_vals, "itemStyle": {"color": "#6366f1"}},
                                        {"name": "Key Presses", "type": "bar", "stack": "input", "data": key_vals, "itemStyle": {"color": "#10b981"}},
                                    ],
                                }).style("height: 280px")

                            # Chart 2: Time breakdown
                            with ui.card().classes("chart-card flex-1"):
                                with ui.row().classes("items-center gap-1 mb-2"):
                                    ui.html(CLOCK_ICON).classes("text-gray-400")
                                    ui.label("Productive Timeline (Minutes)").classes("text-sm font-semibold text-gray-300")
                                ui.echart({
                                    "backgroundColor": "transparent",
                                    "tooltip": {"trigger": "axis"},
                                    "legend": {"data": ["Work (min)", "Idle (min)", "System No Run (min)"], "textStyle": {"color": "#9ca3af"}, "bottom": 0},
                                    "grid": {"left": "3%", "right": "3%", "bottom": "25%", "top": "5%", "containLabel": True},
                                    "xAxis": {"type": "category", "data": chart_labels, "axisLabel": {"color": "#9ca3af", "rotate": 20}},
                                    "yAxis": {"type": "value", "axisLabel": {"color": "#9ca3af"}},
                                    "series": [
                                        {"name": "Work (min)", "type": "bar", "stack": "time", "data": work_vals, "itemStyle": {"color": "#10b981"}},
                                        {"name": "Idle (min)", "type": "bar", "stack": "time", "data": idle_vals, "itemStyle": {"color": "#f43f5e"}},
                                        {"name": "System No Run (min)", "type": "bar", "stack": "time", "data": no_run_vals, "itemStyle": {"color": "#4b5563"}},
                                    ],
                                }).style("height: 280px")

                        # Hour details lists filtered by selected interval
                        with ui.column().classes("w-full p-4 gap-2"):
                            ui.label("Detailed Window & Domain Logs").classes("text-sm font-semibold text-gray-400 mb-1")
                            sorted_slots = sorted(hours.keys(), key=_slot_sort_key)
                            for slot in sorted_slots:
                                parsed = _parse_slot_minutes(slot)
                                if not parsed:
                                    continue
                                slot_start, slot_end = parsed
                                # Check if slot overlaps with selected interval
                                overlap = max(0, min(slot_end, end_min) - max(slot_start, start_min))
                                if overlap <= 0:
                                    continue

                                h = hours[slot]
                                apps = h.get("windows", {})
                                sites = h.get("websites", {})
                                if not apps and not sites:
                                    continue
                                with ui.expansion(f"Breakdown: {slot}", value=False).classes("w-full bg-gray-900/30 rounded-lg"):
                                    with ui.row().classes("w-full gap-8 p-3"):
                                        if apps:
                                            with ui.column().classes("flex-1"):
                                                ui.label("Active Software Process usage").classes("text-xs font-bold text-gray-500 uppercase tracking-wide")
                                                for a, s in sorted(apps.items(), key=lambda x: -x[1])[:10]:
                                                    ui.label(f"• {a}: {_fmt(s)}").classes("text-xs text-gray-300")
                                        if sites:
                                            with ui.column().classes("flex-1"):
                                                ui.label("Visited Browser Domains").classes("text-xs font-bold text-gray-500 uppercase tracking-wide")
                                                for s, sec in sorted(sites.items(), key=lambda x: -x[1])[:10]:
                                                    ui.label(f"• {s}: {_fmt(sec)}").classes("text-xs text-gray-300")

        render_graphs()


# =====================================================================
#  PUBLIC VIEW PAGE
# =====================================================================
@ui.page("/view")
def public_view_page():
    ui.dark_mode(True)
    ui.add_head_html(CSS_STYLES)

    _saved_filter = load_view_filter() or {}
    time_options = [_format_ampm(m) for m in range(0, 24 * 60, 5)]
    picker_values = {
        "start_text": str(_saved_filter.get("from") or "9:00am"),
        "end_text": str(_saved_filter.get("to") or "5:30pm"),
    }

    picker_refs: dict[str, object] = {}
    last_saved_filter = {
        "from": picker_values["start_text"],
        "to": picker_values["end_text"],
    }

    def _selected_minutes() -> tuple[int, int] | None:
        start = _parse_ampm_text(str(picker_refs["start_text"].value))
        end = _parse_ampm_text(str(picker_refs["end_text"].value))
        if start is None or end is None:
            return None
        if end <= start:
            return None
        return start, end

    with ui.header().classes("app-header items-center justify-between px-6 py-3"):
        with ui.row().classes("items-center gap-4 flex-wrap"):
            ui.label("Activity Tracker — Live View").classes("text-xl font-bold text-white")

            with ui.row().classes("items-center gap-2 bg-gray-900 border border-gray-800 rounded px-3 py-1"):
                picker_refs["start_text"] = ui.select(
                    time_options,
                    label="From",
                    value=picker_values["start_text"],
                    on_change=lambda _: _render_public_graphs(),
                ).classes("w-32").props("dark dense")
                picker_refs["end_text"] = ui.select(
                    time_options,
                    label="To",
                    value=picker_values["end_text"],
                    on_change=lambda _: _render_public_graphs(),
                ).classes("w-32").props("dark dense")

        ui.label("Public Monitor View").classes("text-sm text-gray-500 font-semibold uppercase tracking-wider")

    # Build PC list
    pc_map = _build_pc_map(include_local_data_dirs=True)
    pc_names = list(pc_map.keys())
    if not pc_names:
        with ui.column().classes("absolute-center items-center"):
            ui.label("No tracking data available yet.").classes("text-xl text-gray-400 italic")
        return

    default_pc = _state["pc_name"] if _state["pc_name"] in pc_map else pc_names[0]
    selected_pc = {"value": default_pc}
    last_good_data_by_pc: dict[str, list[dict]] = {}

    with ui.column().classes("w-full p-6 gap-6"):

        with ui.card().classes("glass-card w-full row items-center justify-between gap-4 p-4"):
            with ui.row().classes("items-center gap-4"):
                if len(pc_names) > 1:
                    def on_pc_change(e):
                        selected_pc["value"] = e.value
                        _render_public_graphs()

                    ui.select(
                        pc_names,
                        value=selected_pc["value"],
                        label="Target PC",
                        on_change=on_pc_change,
                    ).classes("w-72").props("dark")
                else:
                    ui.label(f"Computer Node: {pc_names[0]}").classes("text-lg font-bold text-white")

                status_chip = ui.label("").classes("text-sm font-bold px-3 py-1 rounded-full")

        data_container = ui.column().classes("w-full gap-4")

        def _render_public_graphs():
            data_container.clear()
            selected = _selected_minutes()
            if selected is None:
                with data_container:
                    with ui.card().classes("w-full p-4 bg-rose-950/20 border border-rose-500/20"):
                        ui.label("Invalid Time Boundaries").classes("text-lg font-bold text-rose-400")
                        ui.label("Please check format (e.g. 9:00am) and make sure 'To' boundary occurs after 'From'.").classes("text-sm text-gray-300")
                return

            start_min, end_min = selected
            current_from = _format_ampm(start_min)
            current_to = _format_ampm(end_min)
            if (
                last_saved_filter["from"] != current_from
                or last_saved_filter["to"] != current_to
            ):
                save_view_filter(current_from, current_to)
                last_saved_filter["from"] = current_from
                last_saved_filter["to"] = current_to

            pc = selected_pc["value"]
            info = pc_map.get(pc, {})
            app_url = _normalize_remote_url(
                info.get("app_url", ""),
                info.get("local_ip", ""),
            )
            is_local = pc == _state["pc_name"]

            all_days: list[dict] = []
            online = is_local
            if is_local:
                all_days = load_all(pc)
            else:
                remote_days = _fetch_remote_all(app_url, pc) if app_url else None
                online = remote_days is not None
                if remote_days is not None:
                    all_days = remote_days
                    last_good_data_by_pc[pc] = remote_days
                else:
                    all_days = load_all(pc) or last_good_data_by_pc.get(pc, [])

            if online:
                status_chip.set_text("● Online Sync")
                status_chip.classes("text-emerald-400 bg-emerald-950/30 border border-emerald-500/20", remove="text-red-400 bg-rose-950/30 border-rose-500/20")
            else:
                status_chip.set_text("○ Offline Mode")
                status_chip.classes("text-rose-400 bg-rose-950/30 border border-rose-500/20", remove="text-emerald-400 bg-emerald-950/30 border-emerald-500/20")

            if not online and not is_local and all_days:
                with data_container:
                    with ui.card().classes("w-full p-3 bg-gray-900 border border-yellow-500/20"):
                        ui.label(f"⚠ App cannot reach {pc}. Showing cached offline logs.").classes("text-sm text-yellow-400 font-medium")

            if not all_days:
                with data_container:
                    ui.label("No metrics available for this configuration.").classes("text-gray-500 italic p-4")
                return

            today_str = date.today().isoformat()

            with data_container:
                for day_data in all_days:
                    day_str = day_data.get("date", "")
                    hours = day_data.get("hours", {})
                    is_today = day_str == today_str

                    with ui.expansion(
                        f"📅 Date: {day_str}" + (" (Today)" if is_today else ""),
                        value=is_today,
                    ).classes("w-full bg-gray-900/60 border border-gray-800 rounded-xl overflow-hidden"):

                        sliced = _aggregate_range(day_data, start_min, end_min)

                        total_work = sliced["total_work_seconds"]
                        total_idle = sliced["total_idle_seconds"]
                        total_no_run = sliced["total_no_run_seconds"]
                        total_mouse = sliced["total_mouse"]
                        total_keys = sliced["total_keys"]

                        with ui.row().classes("w-full items-center justify-between p-4 bg-gray-900/40 border-b border-gray-800"):
                            with ui.row().classes("gap-4"):
                                ui.markdown(f"**Productive Work:** `{_fmt(total_work)}`").classes("text-sm text-emerald-400 font-semibold")
                                ui.markdown(f"**System Idle:** `{_fmt(total_idle)}`").classes("text-sm text-rose-400 font-semibold")
                                ui.markdown(f"**Inactive (Off):** `{_fmt(total_no_run)}`").classes("text-sm text-gray-400 font-semibold")
                            with ui.row().classes("gap-4"):
                                ui.markdown(f"**Clicks:** `{total_mouse}`").classes("text-sm text-cyan-400")
                                ui.markdown(f"**Keystrokes:** `{total_keys}`").classes("text-sm text-indigo-400")

                        if not sliced["labels"]:
                            ui.label("No metrics gathered in this timeframe.").classes("text-gray-500 italic p-4")
                            continue

                        chart_labels = sliced["labels"]
                        mouse_vals = sliced["mouse"]
                        key_vals = sliced["keys"]
                        work_vals = sliced["work_minutes"]
                        idle_vals = sliced["idle_minutes"]
                        no_run_vals = sliced["no_run_minutes"]

                        with ui.row().classes("w-full gap-4 p-4"):
                            # Chart 1
                            with ui.card().classes("chart-card flex-1"):
                                with ui.row().classes("items-center gap-1 mb-2"):
                                    ui.html(MOUSE_ICON).classes("text-gray-400")
                                    ui.label("Input Frequencies").classes("text-sm font-semibold text-gray-300")
                                ui.echart({
                                    "backgroundColor": "transparent",
                                    "tooltip": {"trigger": "axis"},
                                    "legend": {"data": ["Mouse Clicks", "Key Presses"], "textStyle": {"color": "#9ca3af"}, "bottom": 0},
                                    "grid": {"left": "3%", "right": "3%", "bottom": "25%", "top": "5%", "containLabel": True},
                                    "xAxis": {"type": "category", "data": chart_labels, "axisLabel": {"color": "#9ca3af", "rotate": 20}},
                                    "yAxis": {"type": "value", "axisLabel": {"color": "#9ca3af"}},
                                    "series": [
                                        {"name": "Mouse Clicks", "type": "bar", "stack": "input", "data": mouse_vals, "itemStyle": {"color": "#6366f1"}},
                                        {"name": "Key Presses", "type": "bar", "stack": "input", "data": key_vals, "itemStyle": {"color": "#10b981"}},
                                    ],
                                }).style("height: 280px")

                            # Chart 2
                            with ui.card().classes("chart-card flex-1"):
                                with ui.row().classes("items-center gap-1 mb-2"):
                                    ui.html(CLOCK_ICON).classes("text-gray-400")
                                    ui.label("Productive Timeline (Minutes)").classes("text-sm font-semibold text-gray-300")
                                ui.echart({
                                    "backgroundColor": "transparent",
                                    "tooltip": {"trigger": "axis"},
                                    "legend": {"data": ["Work (min)", "Idle (min)", "System No Run (min)"], "textStyle": {"color": "#9ca3af"}, "bottom": 0},
                                    "grid": {"left": "3%", "right": "3%", "bottom": "25%", "top": "5%", "containLabel": True},
                                    "xAxis": {"type": "category", "data": chart_labels, "axisLabel": {"color": "#9ca3af", "rotate": 20}},
                                    "yAxis": {"type": "value", "axisLabel": {"color": "#9ca3af"}},
                                    "series": [
                                        {"name": "Work (min)", "type": "bar", "stack": "time", "data": work_vals, "itemStyle": {"color": "#10b981"}},
                                        {"name": "Idle (min)", "type": "bar", "stack": "time", "data": idle_vals, "itemStyle": {"color": "#f43f5e"}},
                                        {"name": "System No Run (min)", "type": "bar", "stack": "time", "data": no_run_vals, "itemStyle": {"color": "#4b5563"}},
                                    ],
                                }).style("height: 280px")

        _render_public_graphs()
        ui.timer(15, _render_public_graphs)


# ── Helpers ───────────────────────────────────────────────────────────
def _update_tracking_label(label):
    if _state["tracking"]:
        label.set_text("Tracking Engine Status: ACTIVE ✅")
        label.classes("text-emerald-400 font-bold", remove="text-rose-400 text-gray-400")
    else:
        label.set_text("Tracking Engine Status: INACTIVE ✕")
        label.classes("text-gray-400 font-semibold", remove="text-emerald-400 text-rose-400")


def _update_startup_label(label):
    if startup_script_exists():
        label.set_text("System Startup Launch: ENABLED ✓")
        label.classes("text-emerald-400 font-bold", remove="text-gray-400")
    else:
        label.set_text("System Startup Launch: disabled")
        label.classes("text-gray-400", remove="text-emerald-400")


def _logout():
    _state["logged_in"] = False
    ui.navigate.to("/")


def _fmt(seconds: float) -> str:
    h, rem = divmod(int(seconds), 3600)
    m, s = divmod(rem, 60)
    return f"{h}h {m}m {s}s"
