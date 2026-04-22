"""Activity Tracker – main entry point.

On system startup this launches the NiceGUI web UI
on the configured port AND auto-starts tracking if saved config exists.
Binds to 0.0.0.0 so other PCs on the same WiFi can access the dashboard
and fetch data via the JSON API.
"""

import sys
import os

if sys.platform.startswith("win"):
    import asyncio

    # Proactor can emit noisy connection-reset callbacks on Windows when
    # clients disconnect abruptly; Selector is more stable for this app.
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

# Ensure project root is on sys.path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import UI_PORT, UI_TITLE
from logger_setup import get_logger

log = get_logger("main")


def run():
    from nicegui import ui, app
    from fastapi.responses import JSONResponse
    import app_ui.dashboard  # noqa: F401  (registers pages)

    from app_ui.dashboard import try_auto_start
    from local_store import load_day, list_days, load_all

    app.on_startup(try_auto_start)

    # ── JSON API endpoints (used by other PCs on WiFi) ────────────
    @app.get("/api/days/{pc_name}")
    def api_days(pc_name: str):
        """Return list of available date strings for a PC."""
        return JSONResponse({"days": list_days(pc_name)})

    @app.get("/api/data/{pc_name}/{day}")
    def api_day_data(pc_name: str, day: str):
        """Return a single day's data for a PC."""
        data = load_day(pc_name, day)
        if data is None:
            return JSONResponse({"error": "not found"}, status_code=404)
        # Strip ticks to reduce payload size for remote requests
        stripped = dict(data)
        hours = {}
        for slot, hdata in data.get("hours", {}).items():
            h = dict(hdata)
            h.pop("ticks", None)
            hours[slot] = h
        stripped["hours"] = hours
        return JSONResponse(stripped)

    @app.get("/api/all/{pc_name}")
    def api_all_data(pc_name: str):
        """Return all days' data (no ticks) for a PC."""
        all_data = load_all(pc_name)
        result = []
        for day_data in all_data:
            stripped = dict(day_data)
            hours = {}
            for slot, hdata in day_data.get("hours", {}).items():
                h = dict(hdata)
                h.pop("ticks", None)
                hours[slot] = h
            stripped["hours"] = hours
            result.append(stripped)
        return JSONResponse(result)

    @app.get("/api/ping")
    def api_ping():
        """Health check – used to detect if PC is online."""
        return JSONResponse({"status": "ok"})

    ui.run(
        title=UI_TITLE,
        port=UI_PORT,
        reload=False,
        show=False,
        host="0.0.0.0",  # allow access from other PCs on WiFi
    )


if __name__ in {"__main__", "__mp_main__"}:
    run()
