"""Master tracker – per-second local logging + hourly sheet sync."""

import threading
import time
from datetime import datetime, date, timedelta
from collections import defaultdict

from config import (
    TRACKER_POLL_INTERVAL,
    IDLE_THRESHOLD,
    SHEET_SYNC_INTERVAL,
    SYSTEM_APPS,
)
from logger_setup import get_logger
from tracker.window_tracker import get_active_window_info, extract_browser_domain
from tracker.keyboard_tracker import KeyboardTracker
from tracker.mouse_tracker import MouseTracker
from tracker.idle_detector import get_idle_seconds
from local_store import load_day, save_day

log = get_logger("tracker.engine")

_SYSTEM_APPS_LOWER = {a.lower() for a in SYSTEM_APPS}
_BROWSERS = {"chrome.exe", "msedge.exe", "firefox.exe", "brave.exe", "opera.exe"}


def _hour_slot(dt: datetime | None = None) -> str:
    if dt is None:
        dt = datetime.now()
    start = dt.replace(minute=0, second=0, microsecond=0)
    end = start + timedelta(hours=1)
    return f"{start.strftime('%I:%M%p')}-{end.strftime('%I:%M%p')}"


# ======================================================================
class HourlyBucket:
    def __init__(self):
        self.mouse_clicks: int = 0
        self.key_presses: int = 0
        self.window_usage: dict[str, float] = defaultdict(float)
        self.website_usage: dict[str, float] = defaultdict(float)
        self.work_seconds: float = 0.0
        self.idle_seconds: float = 0.0
        self.ticks: list[dict] = []

    def to_dict(self) -> dict:
        return {
            "mouse_clicks": self.mouse_clicks,
            "key_presses": self.key_presses,
            "windows": dict(self.window_usage),
            "websites": dict(self.website_usage),
            "work_seconds": round(self.work_seconds, 1),
            "idle_seconds": round(self.idle_seconds, 1),
            "ticks": self.ticks,
        }

    @classmethod
    def from_dict(cls, d: dict) -> "HourlyBucket":
        b = cls()
        b.mouse_clicks = d.get("mouse_clicks", 0)
        b.key_presses = d.get("key_presses", 0)
        b.window_usage = defaultdict(float, d.get("windows", {}))
        b.website_usage = defaultdict(float, d.get("websites", {}))
        b.work_seconds = d.get("work_seconds", 0.0)
        b.idle_seconds = d.get("idle_seconds", 0.0)
        b.ticks = d.get("ticks", [])
        return b


class DailyData:
    def __init__(self, day: str | None = None):
        self.date_str = day or date.today().isoformat()
        self.hourly: dict[str, HourlyBucket] = {}

    def bucket(self, hour_slot: str) -> HourlyBucket:
        if hour_slot not in self.hourly:
            self.hourly[hour_slot] = HourlyBucket()
        return self.hourly[hour_slot]

    def to_dict(self) -> dict:
        hours = {slot: b.to_dict() for slot, b in self.hourly.items()}
        return {
            "date": self.date_str,
            "hours": hours,
            "total_work": round(sum(b.work_seconds for b in self.hourly.values()), 1),
            "total_idle": round(sum(b.idle_seconds for b in self.hourly.values()), 1),
            "total_mouse_clicks": sum(b.mouse_clicks for b in self.hourly.values()),
            "total_key_presses": sum(b.key_presses for b in self.hourly.values()),
        }

    @classmethod
    def from_dict(cls, d: dict) -> "DailyData":
        dd = cls(d.get("date"))
        for slot, hd in d.get("hours", {}).items():
            dd.hourly[slot] = HourlyBucket.from_dict(hd)
        return dd


# ======================================================================
class TrackerEngine:
    def __init__(self, pc_name: str, sync_callback=None):
        self.pc_name = pc_name
        self.sync_callback = sync_callback
        self.keyboard = KeyboardTracker()
        self.mouse = MouseTracker()
        self._running = False
        self._thread: threading.Thread | None = None
        self._lock = threading.Lock()
        self._last_window: str | None = None
        self._last_window_ts: float = 0.0
        self._last_website: str | None = None
        self._last_website_ts: float = 0.0
        self._last_sync: float = 0.0
        self._last_save: float = 0.0

        # Try to resume from saved local data
        today_str = date.today().isoformat()
        existing = load_day(pc_name, today_str)
        if existing:
            self._today = DailyData.from_dict(existing)
            log.info("Resumed local data for %s", today_str)
        else:
            self._today = DailyData(today_str)

    @property
    def is_running(self) -> bool:
        return self._running

    def start(self):
        if self._running:
            return
        self._running = True
        self.keyboard.start()
        self.mouse.start()
        self._last_sync = time.time()
        self._last_save = time.time()
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()
        log.info("Tracker engine started.")

    def stop(self):
        self._running = False
        self.keyboard.stop()
        self.mouse.stop()
        if self._thread:
            self._thread.join(timeout=5)
        self._flush_sync()
        self._save_local()
        log.info("Tracker engine stopped.")

    def _loop(self):
        while self._running:
            try:
                self._tick()
            except Exception as exc:
                log.error("Tick error: %s", exc, exc_info=True)
            time.sleep(TRACKER_POLL_INTERVAL)

    def _tick(self):
        now = time.time()
        dt_now = datetime.now()
        hour_slot = _hour_slot(dt_now)
        today_str = dt_now.date().isoformat()

        # Day rollover
        with self._lock:
            if self._today.date_str != today_str:
                self._flush_sync()
                self._save_local()
                self._today = DailyData(today_str)
                log.info("Day rollover to %s", today_str)

        # Idle
        idle_sec = get_idle_seconds()
        is_idle = idle_sec >= IDLE_THRESHOLD

        # Active window
        is_productive = False
        info = get_active_window_info()
        current_app = ""
        current_title = ""

        if info and not info.get("minimized", False) and not is_idle:
            app = info["app"]
            title = info["title"]
            current_app = app
            current_title = title
            app_lower = app.lower()

            if app_lower not in _SYSTEM_APPS_LOWER:
                is_productive = True
                with self._lock:
                    b = self._today.bucket(hour_slot)

                    if self._last_window == app:
                        elapsed = now - self._last_window_ts
                        if elapsed < 10:
                            b.window_usage[app] += elapsed
                    self._last_window = app
                    self._last_window_ts = now

                    if app_lower in _BROWSERS:
                        domain = extract_browser_domain(title)
                        if domain:
                            if self._last_website == domain:
                                elapsed = now - self._last_website_ts
                                if elapsed < 10:
                                    b.website_usage[domain] += elapsed
                            self._last_website = domain
                            self._last_website_ts = now

        # Work / idle
        with self._lock:
            b = self._today.bucket(hour_slot)
            if is_idle:
                b.idle_seconds += TRACKER_POLL_INTERVAL
            elif is_productive:
                b.work_seconds += TRACKER_POLL_INTERVAL
            else:
                b.idle_seconds += TRACKER_POLL_INTERVAL

        # Mouse / keyboard
        kb = self.keyboard.snapshot_and_reset()
        ms = self.mouse.snapshot_and_reset()

        with self._lock:
            b = self._today.bucket(hour_slot)
            b.key_presses += kb["press_count"]
            b.mouse_clicks += ms["click_count"]

            # Per-second tick for local storage
            b.ticks.append({
                "ts": int(now),
                "app": current_app,
                "title": current_title[:120],
                "idle": is_idle,
                "mouse": ms["click_count"],
                "keyboard": kb["press_count"],
            })

        # Save local every 10 seconds
        if now - self._last_save >= 10:
            self._save_local()
            self._last_save = now

        # Sheet sync
        if now - self._last_sync >= SHEET_SYNC_INTERVAL:
            self._flush_sync()
            self._last_sync = now

    def _save_local(self):
        with self._lock:
            data = self._today.to_dict()
        save_day(self.pc_name, data["date"], data)

    def _flush_sync(self):
        with self._lock:
            data = self._today.to_dict()
        if self.sync_callback:
            try:
                self.sync_callback(data)
            except Exception as exc:
                log.error("Sync callback failed: %s", exc)

    def get_today_snapshot(self) -> dict:
        with self._lock:
            return self._today.to_dict()
