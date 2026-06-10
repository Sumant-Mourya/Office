"""Master tracker – per-second local logging + periodic sheet sync.

Uses monotonic clock for accurate elapsed-time measurement.
Supports offline operation: marks data as pending when internet is unavailable.
Auto-restarts dead keyboard/mouse listeners.
"""

import threading
import time
from datetime import datetime, date, timedelta
from collections import defaultdict

from config import (
    TRACKER_POLL_INTERVAL,
    IDLE_THRESHOLD,
    SHEET_SYNC_INTERVAL,
    LOCAL_SAVE_INTERVAL,
    SYSTEM_APPS,
)
from logger_setup import get_logger
from tracker.window_tracker import get_active_window_info, extract_browser_domain
from tracker.keyboard_tracker import KeyboardTracker
from tracker.mouse_tracker import MouseTracker
from tracker.idle_detector import get_idle_seconds
from local_store import load_day, save_day
from network_monitor import is_internet_available
from pending_sync import mark_pending, mark_synced, get_pending

log = get_logger("tracker.engine")

_SYSTEM_APPS_LOWER = {a.lower() for a in SYSTEM_APPS}
_BROWSERS = {
    "chrome.exe", "msedge.exe", "firefox.exe", "brave.exe", "opera.exe",
    "vivaldi.exe", "arc.exe", "waterfox.exe", "chromium.exe", "thorium.exe",
}

# Minimum mouse movement (pixels) per tick to consider user active
_MOUSE_MOVE_THRESHOLD = 50.0


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
        self.scroll_count: int = 0
        self.window_usage: dict[str, float] = defaultdict(float)
        self.website_usage: dict[str, float] = defaultdict(float)
        self.work_seconds: float = 0.0
        self.idle_seconds: float = 0.0
        self.ticks: list[dict] = []

    def to_dict(self) -> dict:
        return {
            "mouse_clicks": self.mouse_clicks,
            "key_presses": self.key_presses,
            "scroll_count": self.scroll_count,
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
        b.scroll_count = d.get("scroll_count", 0)
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
    def __init__(self, pc_name: str, sync_callback=None,
                 sheet_backup_enabled: bool = False):
        self.pc_name = pc_name
        self.sync_callback = sync_callback
        self.sheet_backup_enabled = sheet_backup_enabled
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
        self._last_tick_mono: float = 0.0

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
        now = time.monotonic()
        self._last_sync = now
        self._last_save = now
        self._last_tick_mono = now
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
        now_mono = time.monotonic()
        now_wall = time.time()
        dt_now = datetime.now()
        hour_slot = _hour_slot(dt_now)
        today_str = dt_now.date().isoformat()

        # Compute actual elapsed since last tick (handles sleep drift)
        elapsed = now_mono - self._last_tick_mono if self._last_tick_mono else TRACKER_POLL_INTERVAL
        elapsed = min(elapsed, 10.0)  # cap at 10s to avoid inflation from system sleep
        self._last_tick_mono = now_mono

        # Day rollover
        with self._lock:
            if self._today.date_str != today_str:
                self._flush_sync()
                self._save_local()
                self._today = DailyData(today_str)
                log.info("Day rollover to %s", today_str)

        # Auto-restart dead listeners
        self.keyboard.ensure_alive()
        self.mouse.ensure_alive()

        # Idle detection
        idle_sec = get_idle_seconds()
        is_idle = idle_sec >= IDLE_THRESHOLD

        # Mouse / keyboard snapshots
        kb = self.keyboard.snapshot_and_reset()
        ms = self.mouse.snapshot_and_reset()

        has_input = (
            kb["press_count"] > 0
            or ms["click_count"] > 0
            or ms["scroll_count"] > 0
            or ms["move_distance"] > _MOUSE_MOVE_THRESHOLD
        )

        # If user has recent input but GetLastInputInfo says idle (possible
        # edge case with certain input types), trust our own counters.
        if has_input and is_idle:
            is_idle = False

        # Active window
        is_productive = False
        current_app = ""
        current_title = ""

        info = get_active_window_info()

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
                        win_elapsed = now_wall - self._last_window_ts
                        if win_elapsed < 10:
                            b.window_usage[app] += win_elapsed
                    self._last_window = app
                    self._last_window_ts = now_wall

                    if app_lower in _BROWSERS:
                        domain = extract_browser_domain(title)
                        if domain:
                            if self._last_website == domain:
                                site_elapsed = now_wall - self._last_website_ts
                                if site_elapsed < 10:
                                    b.website_usage[domain] += site_elapsed
                            self._last_website = domain
                            self._last_website_ts = now_wall
            elif has_input:
                # System app but user is actively interacting → still work time
                is_productive = True

        # Work / idle accounting using actual elapsed time
        with self._lock:
            b = self._today.bucket(hour_slot)
            if is_idle:
                b.idle_seconds += elapsed
            elif is_productive:
                b.work_seconds += elapsed
            else:
                # No active window / screen locked / etc.
                b.idle_seconds += elapsed

        # Record input counts
        with self._lock:
            b = self._today.bucket(hour_slot)
            b.key_presses += kb["press_count"]
            b.mouse_clicks += ms["click_count"]
            b.scroll_count += ms["scroll_count"]

            # Per-second tick for local storage
            b.ticks.append({
                "ts": int(now_wall),
                "app": current_app,
                "title": current_title[:120],
                "idle": is_idle,
                "mouse": ms["click_count"],
                "keyboard": kb["press_count"],
                "scroll": ms["scroll_count"],
            })

        # Save local periodically
        if now_mono - self._last_save >= LOCAL_SAVE_INTERVAL:
            self._save_local()
            self._last_save = now_mono

        # Sheet sync periodically (if backup enabled)
        if now_mono - self._last_sync >= SHEET_SYNC_INTERVAL:
            self._flush_sync()
            self._last_sync = now_mono

    def _save_local(self):
        with self._lock:
            data = self._today.to_dict()
        save_day(self.pc_name, data["date"], data)

    def _flush_sync(self):
        """Attempt to sync to Google Sheets.  If offline, mark as pending."""
        if not self.sheet_backup_enabled or not self.sync_callback:
            return

        with self._lock:
            data = self._today.to_dict()

        if not is_internet_available():
            mark_pending(self.pc_name, data["date"])
            log.debug("Offline – marked %s as pending sync", data["date"])
            return

        # Sync today's data
        try:
            self.sync_callback(data)
            mark_synced(self.pc_name, data["date"])
        except Exception as exc:
            log.error("Sync callback failed: %s", exc)
            mark_pending(self.pc_name, data["date"])
            return

        # Also flush any pending days from previous offline periods
        self._sync_pending()

    def _sync_pending(self):
        """Sync any days that were queued during offline periods."""
        pending = get_pending(self.pc_name)
        today_str = date.today().isoformat()
        for day_str in pending:
            if day_str == today_str:
                continue  # already synced above
            day_data = load_day(self.pc_name, day_str)
            if day_data and self.sync_callback:
                try:
                    self.sync_callback(day_data)
                    mark_synced(self.pc_name, day_str)
                    log.info("Synced pending day: %s", day_str)
                except Exception as exc:
                    log.error("Failed to sync pending day %s: %s", day_str, exc)
                    break  # stop on first failure to avoid hammering

    def get_today_snapshot(self) -> dict:
        with self._lock:
            return self._today.to_dict()
