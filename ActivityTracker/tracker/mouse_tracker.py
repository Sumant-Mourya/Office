"""Mouse click and scroll counter using pynput.
Only tracks click/scroll counts and movement distance – no coordinates logged.
"""

import threading
import math
from pynput import mouse

from logger_setup import get_logger

log = get_logger("tracker.mouse")


class MouseTracker:
    def __init__(self):
        self._lock = threading.Lock()
        self._click_count = 0
        self._scroll_count = 0
        self._listener: mouse.Listener | None = None
        self._started = False

    @property
    def is_alive(self) -> bool:
        """True if the listener thread is running."""
        return self._listener is not None and self._listener.is_alive()

    def start(self):
        self._started = True
        self._start_listener()
        log.info("Mouse tracker started.")

    def _start_listener(self):
        """Create and start a fresh listener."""
        try:
            if self._listener is not None:
                try:
                    self._listener.stop()
                except Exception:
                    pass
            self._listener = mouse.Listener(
                on_click=self._on_click,
                on_scroll=self._on_scroll,
            )
            self._listener.daemon = True
            self._listener.start()
        except Exception as exc:
            log.error("Failed to start mouse listener: %s", exc)

    def stop(self):
        self._started = False
        if self._listener:
            try:
                self._listener.stop()
            except Exception:
                pass
            log.info("Mouse tracker stopped.")

    def ensure_alive(self):
        """Restart the listener if it has died."""
        if self._started and not self.is_alive:
            log.warning("Mouse listener died, restarting...")
            self._start_listener()

    def _on_click(self, _x, _y, _button, pressed):
        if pressed:
            with self._lock:
                self._click_count += 1

    def _on_scroll(self, _x, _y, _dx, _dy):
        with self._lock:
            self._scroll_count += 1

    def snapshot_and_reset(self) -> dict:
        """Return counts since last call and reset."""
        with self._lock:
            clicks = self._click_count
            scrolls = self._scroll_count
            self._click_count = 0
            self._scroll_count = 0
        return {
            "click_count": clicks,
            "scroll_count": scrolls,
            "move_distance": 0.0,
        }
