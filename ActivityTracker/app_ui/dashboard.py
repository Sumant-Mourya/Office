"""NiceGUI-based dashboard – login, control panel & full-width graphs."""

import asyncio
from bisect import bisect_right
import httpx
from datetime import date, datetime
from urllib.parse import urlparse

from googleapiclient.errors import HttpError
from nicegui import ui, app

from config import DEFAULT_USER, DEFAULT_PASS, SHEET_ID, UI_PORT
from config_store import (
    save_config,
    load_config,
    delete_config,
    save_view_filter,
    load_view_filter,
)
from logger_setup import get_logger
from auth.google_auth import GoogleAuth
from sheets.sync import SheetSync
from tracker.engine import TrackerEngine
from setup_autostart import (
    install_startup_script,
    remove_startup_script,
    startup_script_exists,
)
from local_store import load_all, list_days, load_day, list_pcs

log = get_logger("ui.dashboard")

# ── Shared state ──────────────────────────────────────────────────────
google_auth = GoogleAuth()
tracker_engine: TrackerEngine | None = None
sheet_sync: SheetSync | None = None

_saved = load_config()
_has_saved_config = bool(
    _saved
    and _saved.get("pc_name")
    and _saved.get("sheet_id")
)

_state = {
    "logged_in": False,
    "google_connected": google_auth.is_connected,
    "pc_name": _saved["pc_name"] if _saved else "",
    "sheet_id": (_saved.get("sheet_id") if _saved else "") or SHEET_ID,
    "tracking": False,
    "config_locked": _has_saved_config,
}


# ── Auto-start ────────────────────────────────────────────────────────
def try_auto_start():
    global sheet_sync, tracker_engine
    if not _state["config_locked"] or not _state["google_connected"]:
        return
    if not _state["sheet_id"] or not _state["pc_name"]:
        return
    if tracker_engine and tracker_engine.is_running:
        return
    try:
        creds = google_auth.get_credentials()
        if creds is None:
            _state["google_connected"] = False
            log.warning("Auto-start skipped: missing/expired Google credentials.")
            return
        sheet_sync = SheetSync(creds, _state["sheet_id"], _state["pc_name"])
        if not sheet_sync.sheet_exists():
            sheet_sync.create_sheet()
        sheet_sync.ensure_config_sheet(_state["pc_name"])
        tracker_engine = TrackerEngine(
            pc_name=_state["pc_name"], sync_callback=sheet_sync.sync,
        )
        tracker_engine.start()
        _state["tracking"] = True
        log.info("Auto-started tracking (tab=%s).", _state["pc_name"])
    except Exception as exc:
        log.error("Auto-start failed: %s", exc)


# ── Remote helpers ────────────────────────────────────────────────────


def _fetch_remote_all(app_url: str, pc_name: str) -> list[dict] | None:
    """Fetch all days data from a remote PC via WiFi."""
    try:
        r = httpx.get(
            f"{app_url}/api/all/{pc_name}",
            timeout=httpx.Timeout(4.0, connect=1.2),
        )
        if r.status_code == 200:
            return r.json()
        return None
    except Exception:
        return None


def _normalize_remote_url(app_url: str, local_ip: str) -> str:
    """Replace localhost-style URLs with LAN IPs for cross-device access."""
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
    """Best-effort read of known PCs from the shared sheet config tab."""
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

    creds = google_auth.get_credentials()
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
    """Merge PCs from sheet config + local storage, preserving known metadata."""
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
    """Convert slots like 09:00AM-10:00AM into minute bounds."""
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
    # Segment rules:
    # 1) First segment starts at selected start and ends at next hour boundary.
    # 2) Middle segments are full clock hours.
    # 3) Last segment can be partial if selected end is mid-hour.
    segments: list[tuple[int, int, str]] = []
    cursor = start_min
    while cursor < end_min:
        next_hour = ((cursor // 60) + 1) * 60
        seg_end = min(end_min, next_hour)
        segments.append((cursor, seg_end, f"{_format_ampm(cursor)}-{_format_ampm(seg_end)}"))
        cursor = seg_end
    return segments


def _aggregate_range(day_data: dict, start_min: int, end_min: int) -> dict:
    """Aggregate metrics for selected time range with partial-hour support."""
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

    # Fallback for older files that do not have per-second ticks.
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
#  LOGIN PAGE
# =====================================================================
@ui.page("/")
def login_page():
    if _state["logged_in"]:
        ui.navigate.to("/dashboard")
        return

    ui.dark_mode(True)
    with ui.column().classes("absolute-center items-center gap-4"):
        ui.label("Activity Tracker").classes("text-3xl font-bold text-white")
        ui.label("Sign in to continue").classes("text-gray-400")
        with ui.card().classes("w-80 p-6"):
            username = ui.input("Username").classes("w-full")
            password = ui.input("Password", password=True,
                                password_toggle_button=True).classes("w-full")
            error_label = ui.label("").classes("text-red-500 text-sm")

            def handle_login():
                if username.value == DEFAULT_USER and password.value == DEFAULT_PASS:
                    _state["logged_in"] = True
                    ui.navigate.to("/dashboard")
                else:
                    error_label.set_text("Invalid credentials.")

            ui.button("Login", on_click=handle_login).classes("w-full mt-2")


# =====================================================================
#  DASHBOARD (control panel)
# =====================================================================
@ui.page("/dashboard")
def dashboard_page():
    if not _state["logged_in"]:
        ui.navigate.to("/")
        return

    ui.dark_mode(True)

    with ui.header().classes("bg-gray-900 items-center justify-between px-6"):
        ui.label("Activity Tracker — Dashboard").classes("text-xl font-bold")
        with ui.row().classes("gap-2"):
            ui.button("Graphs", on_click=lambda: ui.navigate.to("/graphs")).props(
                "flat color=white")
            ui.button("Logout", on_click=_logout).props("flat color=red")

    with ui.column().classes("w-full max-w-3xl mx-auto p-6 gap-6"):

        # ── Google Auth ───────────────────────────────────────────────
        with ui.card().classes("w-full p-4"):
            ui.label("1. Google Authentication").classes("text-lg font-semibold")
            google_status = ui.label()
            _update_google_label(google_status)

            def on_connect_google():
                ui.notify("Opening browser for Google sign-in…", type="info")
                ok = google_auth.authenticate()
                _state["google_connected"] = google_auth.is_connected
                _update_google_label(google_status)
                if ok:
                    ui.notify("Google connected!", type="positive")
                else:
                    ui.notify("Google auth failed.", type="negative")

            connect_btn = ui.button("Connect Google Account",
                                    on_click=on_connect_google)
            if _state["google_connected"]:
                connect_btn.props("disable")

        # ── Sheet + PC Config ─────────────────────────────────────────
        with ui.card().classes("w-full p-4"):
            ui.label("2. Sheet & PC Configuration").classes(
                "text-lg font-semibold"
            )
            locked = _state["config_locked"]

            sheet_input = ui.input(
                "Google Sheet ID",
                value=_state["sheet_id"],
            ).classes("w-full")
            if locked:
                sheet_input.props("readonly")

            pc_input = ui.input("PC Name (unique name for this computer)",
                                value=_state["pc_name"]).classes("w-full")
            if locked:
                pc_input.props("readonly")
            pc_msg = ui.label("").classes("text-sm")
            if locked:
                pc_msg.set_text("Loaded from saved config (read-only)")
                pc_msg.classes("text-blue-400")

        # ── Tracking Control ──────────────────────────────────────────
        with ui.card().classes("w-full p-4"):
            ui.label("3. Tracking Control").classes("text-lg font-semibold")
            tracking_label = ui.label()
            _update_tracking_label(tracking_label)
            startup_status = ui.label("")
            _update_startup_label(startup_status)

            async def on_start():
                global sheet_sync, tracker_engine
                if not _state["google_connected"]:
                    ui.notify("Connect Google first.", type="warning")
                    return

                sheet_val = sheet_input.value.strip()
                if not sheet_val:
                    ui.notify("Enter a Google Sheet ID.", type="warning")
                    return

                pc_val = pc_input.value.strip()
                if not pc_val:
                    ui.notify("Enter a PC name.", type="warning")
                    return

                # Disable button & show loading
                start_btn.props("loading disable")
                await asyncio.sleep(0.1)  # yield so UI renders spinner

                try:
                    creds = google_auth.get_credentials()
                    if creds is None:
                        _state["google_connected"] = False
                        _update_google_label(google_status)
                        ui.notify(
                            "Google token expired. Reconnect your Google account.",
                            type="warning",
                            close_button=True,
                        )
                        return

                    _state["pc_name"] = pc_val
                    _state["sheet_id"] = sheet_val
                    sheet_sync = SheetSync(
                        creds,
                        _state["sheet_id"],
                        _state["pc_name"],
                    )

                    existing_pc_tab = False
                    try:
                        existing_pc_tab = sheet_sync.sheet_exists()
                    except HttpError as exc:
                        if exc.resp is not None and exc.resp.status == 403:
                            ui.notify(
                                "Sheet access denied (403). Use a Sheet ID you own "
                                "or share that sheet with your connected Google "
                                "account, then try again.",
                                type="negative",
                                close_button=True,
                            )
                            return
                        raise

                    # Check if PC name tab already exists and config is not locked
                    # (i.e. this is a new setup, not a resume)
                    if not _state["config_locked"]:
                        if existing_pc_tab and sheet_sync.pc_name_taken(_state["pc_name"]):
                            ui.notify(
                                f"PC name '{_state['pc_name']}' is already taken. "
                                "Please choose a different name.",
                                type="negative",
                                close_button=True,
                            )
                            return

                    if not existing_pc_tab:
                        sheet_sync.create_sheet()
                    sheet_sync.ensure_config_sheet(_state["pc_name"])

                    save_config(_state["pc_name"], _state["sheet_id"])
                    _state["config_locked"] = True
                    sheet_input.props("readonly")
                    pc_input.props("readonly")

                    tracker_engine = TrackerEngine(
                        pc_name=_state["pc_name"],
                        sync_callback=sheet_sync.sync,
                    )
                    tracker_engine.start()
                    _state["tracking"] = True
                    _update_tracking_label(tracking_label)
                    start_btn.props("disable")
                    stop_btn.props(remove="disable")

                    ok = install_startup_script()
                    _update_startup_label(startup_status)
                    if ok:
                        ui.notify("Tracking started & startup auto-launch enabled!",
                                  type="positive")
                    else:
                        ui.notify("Tracking started. (Startup auto-launch setup failed.)",
                                  type="warning")
                except Exception as exc:
                    log.error("on_start failed: %s", exc)
                    ui.notify(f"Start failed: {exc}", type="negative")
                    start_btn.props(remove="loading disable")
                finally:
                    start_btn.props(remove="loading")

            def on_stop():
                global tracker_engine
                if tracker_engine and tracker_engine.is_running:
                    tracker_engine.stop()
                _state["tracking"] = False
                _update_tracking_label(tracking_label)
                remove_startup_script()
                _update_startup_label(startup_status)
                delete_config()
                _state["config_locked"] = False
                sheet_input.props(remove="readonly")
                pc_input.props(remove="readonly")
                stop_btn.props("disable")
                start_btn.props(remove="disable")
                ui.notify("Tracking stopped & startup auto-launch removed.", type="info")

            with ui.row().classes("gap-4"):
                start_btn = ui.button("Start Tracking", on_click=on_start).props(
                    "color=green")
                stop_btn = ui.button("Stop Tracking", on_click=on_stop).props(
                    "color=red outline")
                if _state["tracking"]:
                    start_btn.props("disable")
                else:
                    stop_btn.props("disable")

        # ── Live Status ───────────────────────────────────────────────
        with ui.card().classes("w-full p-4"):
            ui.label("Live Status").classes("text-lg font-semibold")
            status_area = ui.markdown("")

            def refresh_status():
                if tracker_engine and tracker_engine.is_running:
                    snap = tracker_engine.get_today_snapshot()
                    lines = [
                        f"**Date:** {snap['date']}  ",
                        f"**Total Work:** {_fmt(snap['total_work'])}  ",
                        f"**Total Idle:** {_fmt(snap['total_idle'])}  ",
                        f"**Mouse Clicks:** {snap['total_mouse_clicks']}  ",
                        f"**Key Presses:** {snap['total_key_presses']}  ",
                        "",
                        "### Hourly Breakdown",
                    ]
                    for slot in sorted(snap.get("hours", {}).keys(), key=_slot_sort_key):
                        h = snap["hours"][slot]
                        lines.append(
                            f"**{slot}** — "
                            f"Mouse: {h['mouse_clicks']}, "
                            f"Keys: {h['key_presses']}, "
                            f"Work: {_fmt(h.get('work_seconds', 0))}, "
                            f"Idle: {_fmt(h.get('idle_seconds', 0))}"
                        )
                    status_area.set_content("\n".join(lines))
                else:
                    status_area.set_content("*Tracker not running.*")

            ui.timer(5, refresh_status)


# =====================================================================
#  GRAPHS PAGE – full-screen width, date collapsible, remote PC support
# =====================================================================
@ui.page("/graphs")
def graphs_page():
    if not _state["logged_in"]:
        ui.navigate.to("/")
        return

    ui.dark_mode(True)

    with ui.header().classes("bg-gray-900 items-center justify-between px-6"):
        ui.label("Activity Tracker — Graphs").classes("text-xl font-bold")
        with ui.row().classes("gap-2"):
            ui.button("Dashboard", on_click=lambda: ui.navigate.to("/dashboard")
                       ).props("flat color=white")
            ui.button("Logout", on_click=_logout).props("flat color=red")

    pc_map = _build_pc_map(include_local_data_dirs=True)
    pc_names = list(pc_map.keys())
    if _state["pc_name"] in pc_map:
        pc_names = [_state["pc_name"]] + [n for n in pc_names if n != _state["pc_name"]]

    selected_pc = {"value": _state["pc_name"] or (pc_names[0] if pc_names else "")}
    last_good_data_by_pc: dict[str, list[dict]] = {}

    with ui.column().classes("w-full p-4 gap-4"):  # full width, no max-w

        # ── PC selector row ───────────────────────────────────────────
        with ui.row().classes("w-full items-center gap-4"):
            if pc_names:
                def on_pc_change(e):
                    selected_pc["value"] = e.value
                    render_graphs()

                ui.select(
                    pc_names,
                    value=selected_pc["value"],
                    label="Select PC",
                    on_change=on_pc_change,
                ).classes("w-72")

            status_chip = ui.label("").classes("text-sm font-semibold")

        data_container = ui.column().classes("w-full gap-3")

        def render_graphs():
            data_container.clear()
            pc = selected_pc["value"]
            if not pc:
                with data_container:
                    ui.label("No PC selected.").classes("text-gray-400")
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
                status_chip.set_text(f"● {pc} — Online")
                status_chip.classes("text-green-400", remove="text-red-400 text-gray-400")
            else:
                status_chip.set_text(f"● {pc} — Offline")
                status_chip.classes("text-red-400", remove="text-green-400 text-gray-400")

            if not online and not is_local and all_days:
                with data_container:
                    with ui.card().classes("w-full p-3 bg-gray-900"):
                        ui.label(f"{pc} is offline. Showing last available data.").classes(
                            "text-sm text-yellow-300"
                        )

            if not all_days and not online and not is_local:
                with data_container:
                    with ui.card().classes("w-full p-6 bg-red-900"):
                        ui.label(f"⚠ {pc} is offline").classes(
                            "text-xl font-bold text-white")
                        ui.label(
                            "This PC is not reachable on the WiFi network. "
                            "Make sure the Activity Tracker app is running on "
                            "that computer and both PCs are on the same WiFi."
                        ).classes("text-gray-300")
                        if app_url:
                            ui.label(f"Expected at: {app_url}").classes(
                                "text-gray-400 text-sm")
                return

            if not all_days:
                with data_container:
                    ui.label("No data found for this PC.").classes("text-gray-400")
                return

            today_str = date.today().isoformat()

            with data_container:
                for day_data in all_days:  # already newest-first
                    day_str = day_data.get("date", "")
                    hours = day_data.get("hours", {})
                    is_today = day_str == today_str

                    with ui.expansion(
                        f"📅 {day_str}" + (" (today)" if is_today else ""),
                        value=is_today,
                    ).classes("w-full bg-gray-800 rounded"):

                        total_work = sum(
                            h.get("work_seconds", 0) for h in hours.values())
                        total_idle = sum(
                            h.get("idle_seconds", 0) for h in hours.values())
                        total_mouse = sum(
                            h.get("mouse_clicks", 0) for h in hours.values())
                        total_keys = sum(
                            h.get("key_presses", 0) for h in hours.values())

                        ui.markdown(
                            f"**Work:** {_fmt(total_work)} · "
                            f"**Idle:** {_fmt(total_idle)} · "
                            f"**Mouse:** {total_mouse} · "
                            f"**Keys:** {total_keys}"
                        ).classes("text-sm mb-2")

                        if not hours:
                            ui.label("No hourly data.").classes("text-gray-500")
                            continue

                        sorted_slots = sorted(hours.keys(), key=_slot_sort_key)
                        chart_labels = []
                        mouse_vals = []
                        key_vals = []
                        work_vals = []
                        idle_vals = []
                        for slot in sorted_slots:
                            h = hours[slot]
                            chart_labels.append(slot.split("-")[0])
                            mouse_vals.append(h.get("mouse_clicks", 0))
                            key_vals.append(h.get("key_presses", 0))
                            work_vals.append(round(
                                h.get("work_seconds", 0) / 60, 1))
                            idle_vals.append(round(
                                h.get("idle_seconds", 0) / 60, 1))

                        ui.label("Input Activity (Mouse & Keyboard)").classes(
                            "text-sm font-semibold mt-2")
                        ui.echart({
                            "tooltip": {"trigger": "axis"},
                            "legend": {"data": ["Mouse Clicks", "Key Presses"],
                                       "textStyle": {"color": "#ccc"}},
                            "grid": {"left": "3%", "right": "3%",
                                     "bottom": "3%", "containLabel": True},
                            "xAxis": {"type": "category",
                                      "data": chart_labels,
                                      "axisLabel": {"color": "#ccc",
                                                    "rotate": 0}},
                            "yAxis": {"type": "value",
                                      "axisLabel": {"color": "#ccc"}},
                            "series": [
                                {"name": "Mouse Clicks", "type": "bar",
                                 "data": mouse_vals,
                                 "itemStyle": {"color": "#4fc3f7"}},
                                {"name": "Key Presses", "type": "bar",
                                 "data": key_vals,
                                 "itemStyle": {"color": "#81c784"}},
                            ],
                        }).classes("w-full").style("height: 350px")

                        ui.label("Work / Idle Time (minutes)").classes(
                            "text-sm font-semibold mt-4")
                        ui.echart({
                            "tooltip": {"trigger": "axis"},
                            "legend": {"data": ["Work (min)", "Idle (min)"],
                                       "textStyle": {"color": "#ccc"}},
                            "grid": {"left": "3%", "right": "3%",
                                     "bottom": "3%", "containLabel": True},
                            "xAxis": {"type": "category",
                                      "data": chart_labels,
                                      "axisLabel": {"color": "#ccc",
                                                    "rotate": 0}},
                            "yAxis": {"type": "value",
                                      "axisLabel": {"color": "#ccc"}},
                            "series": [
                                {"name": "Work (min)", "type": "bar",
                                 "stack": "time", "data": work_vals,
                                 "itemStyle": {"color": "#66bb6a"}},
                                {"name": "Idle (min)", "type": "bar",
                                 "stack": "time", "data": idle_vals,
                                 "itemStyle": {"color": "#ef5350"}},
                            ],
                        }).classes("w-full").style("height: 350px")

                        # Per-hour detail
                        for slot in sorted_slots:
                            h = hours[slot]
                            apps = h.get("windows", {})
                            sites = h.get("websites", {})
                            if not apps and not sites:
                                continue
                            with ui.expansion(
                                f"🕐 {slot}", value=False,
                            ).classes("w-full"):
                                with ui.row().classes("w-full gap-8"):
                                    if apps:
                                        with ui.column():
                                            ui.label("Top Apps").classes(
                                                "text-xs font-semibold")
                                            for a, s in sorted(
                                                apps.items(),
                                                key=lambda x: -x[1],
                                            )[:8]:
                                                ui.label(
                                                    f"{a}: {_fmt(s)}"
                                                ).classes(
                                                    "text-xs text-gray-300")
                                    if sites:
                                        with ui.column():
                                            ui.label("Top Websites").classes(
                                                "text-xs font-semibold")
                                            for s, sec in sorted(
                                                sites.items(),
                                                key=lambda x: -x[1],
                                            )[:8]:
                                                ui.label(
                                                    f"{s}: {_fmt(sec)}"
                                                ).classes(
                                                    "text-xs text-gray-300")

        render_graphs()
        ui.timer(15, render_graphs)


# =====================================================================
#  PUBLIC VIEW – no login, same WiFi access
# =====================================================================
@ui.page("/view")
def public_view_page():
    """Public graphs page – accessible without login from any WiFi device."""
    ui.dark_mode(True)

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

    with ui.header().classes("bg-gray-900 items-center justify-between px-6"):
        with ui.row().classes("items-center gap-4 flex-wrap"):
            ui.label("Activity Tracker — Live View").classes("text-xl font-bold")

            with ui.row().classes("items-center gap-2 bg-gray-800 rounded px-3 py-2"):
                picker_refs["start_text"] = ui.select(
                    time_options,
                    label="From",
                    value=picker_values["start_text"],
                    on_change=lambda _: _render_public_graphs(),
                ).classes("w-36")
                picker_refs["end_text"] = ui.select(
                    time_options,
                    label="To",
                    value=picker_values["end_text"],
                    on_change=lambda _: _render_public_graphs(),
                ).classes("w-36")
                ui.label("Format: 9:30am").classes("text-xs text-gray-400")

        ui.label("Public read-only view").classes("text-sm text-gray-400")

    # Build PC list from shared config + local stores.
    pc_map = _build_pc_map(include_local_data_dirs=True)
    pc_names = list(pc_map.keys())
    if not pc_names:
        with ui.column().classes("absolute-center items-center"):
            ui.label("No tracking data available yet.").classes(
                "text-xl text-gray-400")
        return

    default_pc = _state["pc_name"] if _state["pc_name"] in pc_map else pc_names[0]
    selected_pc = {"value": default_pc}
    last_good_data_by_pc: dict[str, list[dict]] = {}

    with ui.column().classes("w-full p-4 gap-4"):

        with ui.row().classes("w-full items-center gap-4"):
            if len(pc_names) > 1:
                def on_pc_change(e):
                    selected_pc["value"] = e.value
                    _render_public_graphs()

                ui.select(
                    pc_names,
                    value=selected_pc["value"],
                    label="Select PC",
                    on_change=on_pc_change,
                ).classes("w-72")
            else:
                ui.label(f"PC: {pc_names[0]}").classes(
                    "text-lg font-semibold")

            status_chip = ui.label("").classes("text-sm font-semibold")

        data_container = ui.column().classes("w-full gap-3")

        def _render_public_graphs():
            data_container.clear()

            selected = _selected_minutes()
            if selected is None:
                with data_container:
                    with ui.card().classes("w-full p-4 bg-red-900"):
                        ui.label("Invalid time range").classes("text-lg font-semibold")
                        ui.label("Use time format like 9:30am, and keep To after From.").classes(
                            "text-sm text-gray-200"
                        )
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
                status_chip.set_text(f"● {pc} — Online")
                status_chip.classes("text-green-400", remove="text-red-400 text-gray-400")
            else:
                status_chip.set_text(f"● {pc} — Offline")
                status_chip.classes("text-red-400", remove="text-green-400 text-gray-400")

            if not online and not is_local and all_days:
                with data_container:
                    with ui.card().classes("w-full p-3 bg-gray-900"):
                        ui.label(f"{pc} is offline. Showing last available data.").classes(
                            "text-sm text-yellow-300"
                        )

            if not all_days:
                with data_container:
                    if not online and not is_local:
                        with ui.card().classes("w-full p-6 bg-red-900"):
                            ui.label(f"⚠ {pc} is offline").classes(
                                "text-xl font-bold text-white"
                            )
                            ui.label(
                                "No remote data is reachable right now. "
                                "Make sure the Activity Tracker app is running on "
                                "that computer and both PCs are on the same WiFi."
                            ).classes("text-gray-300")
                            if app_url:
                                ui.label(f"Expected at: {app_url}").classes(
                                    "text-gray-400 text-sm"
                                )
                    else:
                        ui.label("No data found right now. Retrying automatically...").classes(
                            "text-gray-400"
                        )
                return

            today_str = date.today().isoformat()

            with data_container:
                for day_data in all_days:
                    day_str = day_data.get("date", "")
                    hours = day_data.get("hours", {})
                    is_today = day_str == today_str

                    with ui.expansion(
                        f"📅 {day_str}" + (" (today)" if is_today else ""),
                        value=is_today,
                    ).classes("w-full bg-gray-800 rounded"):

                        sliced = _aggregate_range(day_data, start_min, end_min)

                        total_work = sliced["total_work_seconds"]
                        total_idle = sliced["total_idle_seconds"]
                        total_no_run = sliced["total_no_run_seconds"]
                        total_mouse = sliced["total_mouse"]
                        total_keys = sliced["total_keys"]

                        with ui.row().classes("w-full items-center gap-4 mb-1"):
                            ui.label(f"Work: {_fmt(total_work)}").classes(
                                "text-sm text-green-400 font-semibold"
                            )
                            ui.label(f"Idle: {_fmt(total_idle)}").classes(
                                "text-sm text-red-400 font-semibold"
                            )

                        ui.markdown(
                            f"**Range:** {_format_ampm(start_min)} - {_format_ampm(end_min)}  \n"
                            f"**Work:** {_fmt(total_work)}  \n"
                            f"**Idle:** {_fmt(total_idle)}  \n"
                            f"**System No Run:** {_fmt(total_no_run)}  \n"
                            f"**Mouse:** {total_mouse}  \n"
                            f"**Keys:** {total_keys}"
                        ).classes("text-sm mb-2")

                        if not sliced["labels"]:
                            ui.label("No hourly data.").classes(
                                "text-gray-500")
                            continue

                        chart_labels = sliced["labels"]
                        mouse_vals = sliced["mouse"]
                        key_vals = sliced["keys"]
                        work_vals = sliced["work_minutes"]
                        idle_vals = sliced["idle_minutes"]
                        no_run_vals = sliced["no_run_minutes"]

                        ui.label("Input Activity").classes(
                            "text-sm font-semibold mt-2")
                        ui.echart({
                            "tooltip": {"trigger": "axis"},
                            "legend": {"data": ["Mouse Clicks", "Key Presses"],
                                       "textStyle": {"color": "#ccc"}},
                            "grid": {"left": "3%", "right": "3%",
                                     "bottom": "15%", "containLabel": True},
                            "xAxis": {"type": "category",
                                      "data": chart_labels,
                                      "axisLabel": {"color": "#ccc",
                                                    "rotate": 20}},
                            "yAxis": {"type": "value",
                                      "axisLabel": {"color": "#ccc"}},
                            "series": [
                                {"name": "Mouse Clicks", "type": "bar",
                                 "stack": "input",
                                 "data": mouse_vals,
                                 "itemStyle": {"color": "#4fc3f7"}},
                                {"name": "Key Presses", "type": "bar",
                                 "stack": "input",
                                 "data": key_vals,
                                 "itemStyle": {"color": "#81c784"}},
                            ],
                        }).classes("w-full").style("height: 350px")

                        with ui.row().classes("w-full gap-2 flex-wrap mt-2"):
                            for idx, label in enumerate(chart_labels):
                                with ui.card().classes("p-2 bg-gray-900"):
                                    ui.label(label).classes("text-xs text-gray-300")
                                    ui.label(f"Mouse: {mouse_vals[idx]}").classes(
                                        "text-xs text-cyan-300"
                                    )
                                    ui.label(f"Keys: {key_vals[idx]}").classes(
                                        "text-xs text-green-300"
                                    )

                        ui.label("Work / Idle (minutes)").classes(
                            "text-sm font-semibold mt-4")
                        ui.echart({
                            "tooltip": {"trigger": "axis"},
                            "legend": {"data": ["Work (min)", "Idle (min)", "System No Run (min)"],
                                       "textStyle": {"color": "#ccc"}},
                            "grid": {"left": "3%", "right": "3%",
                                     "bottom": "15%", "containLabel": True},
                            "xAxis": {"type": "category",
                                      "data": chart_labels,
                                      "axisLabel": {"color": "#ccc",
                                                    "rotate": 20}},
                            "yAxis": {"type": "value",
                                      "axisLabel": {"color": "#ccc"}},
                            "series": [
                                {"name": "Work (min)", "type": "bar",
                                 "stack": "time", "data": work_vals,
                                 "itemStyle": {"color": "#66bb6a"}},
                                {"name": "Idle (min)", "type": "bar",
                                 "stack": "time", "data": idle_vals,
                                 "itemStyle": {"color": "#ef5350"}},
                                {"name": "System No Run (min)", "type": "bar",
                                 "stack": "time", "data": no_run_vals,
                                 "itemStyle": {"color": "#9e9e9e"}},
                            ],
                        }).classes("w-full").style("height: 350px")

                        with ui.row().classes("w-full gap-2 flex-wrap mt-2"):
                            for idx, label in enumerate(chart_labels):
                                with ui.card().classes("p-2 bg-gray-900"):
                                    ui.label(label).classes("text-xs text-gray-300")
                                    ui.label(f"Work: {work_vals[idx]} min").classes(
                                        "text-xs text-green-300"
                                    )
                                    ui.label(f"Idle: {idle_vals[idx]} min").classes(
                                        "text-xs text-red-300"
                                    )
                                    ui.label(f"No Run: {no_run_vals[idx]} min").classes(
                                        "text-xs text-gray-300"
                                    )

        _render_public_graphs()
        ui.timer(15, _render_public_graphs)


# ── Helpers ───────────────────────────────────────────────────────────

def _update_google_label(label):
    if _state["google_connected"]:
        label.set_text("Connected ✅")
        label.classes("text-green-500", remove="text-red-500")
    else:
        label.set_text("Not connected")
        label.classes("text-red-500", remove="text-green-500")


def _update_tracking_label(label):
    if _state["tracking"]:
        label.set_text("Tracking is ACTIVE ✅")
        label.classes("text-green-500", remove="text-red-500")
    else:
        label.set_text("Tracking is OFF")
        label.classes("text-gray-400", remove="text-green-500")


def _update_startup_label(label):
    if startup_script_exists():
        label.set_text("Auto-start on login (Startup folder EXE): ENABLED ✅")
        label.classes("text-green-500", remove="text-gray-400")
    else:
        label.set_text("Auto-start on login (Startup folder EXE): disabled")
        label.classes("text-gray-400", remove="text-green-500")


def _logout():
    _state["logged_in"] = False
    ui.navigate.to("/")


def _fmt(seconds: float) -> str:
    h, rem = divmod(int(seconds), 3600)
    m, s = divmod(rem, 60)
    return f"{h}h {m}m {s}s"
