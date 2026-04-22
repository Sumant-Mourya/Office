"""Mouse click counter using pynput.
Only tracks click counts – no coordinates or movement logged.
"""

import threading
from pynput import mouse

from logger_setup import get_logger

log = get_logger("tracker.mouse")


class MouseTracker:
    def __init__(self):
        self._lock = threading.Lock()
        self._click_count = 0
        self._listener: mouse.Listener | None = None

    def start(self):
        self._listener = mouse.Listener(on_click=self._on_click)
        self._listener.daemon = True
        self._listener.start()
        log.info("Mouse tracker started.")

    def stop(self):
        if self._listener:
            self._listener.stop()
            log.info("Mouse tracker stopped.")

    def _on_click(self, _x, _y, _button, pressed):
        if pressed:
            with self._lock:
                self._click_count += 1

    def snapshot_and_reset(self) -> dict:
        """Return click count since last call and reset."""
        with self._lock:
            count = self._click_count
            self._click_count = 0
        return {"click_count": count}
