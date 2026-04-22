"""Keyboard press counter using pynput.
Only tracks press counts – NO actual keystrokes are logged.
"""

import threading
from pynput import keyboard

from logger_setup import get_logger

log = get_logger("tracker.keyboard")


class KeyboardTracker:
    def __init__(self):
        self._lock = threading.Lock()
        self._press_count = 0
        self._listener: keyboard.Listener | None = None

    def start(self):
        self._listener = keyboard.Listener(on_press=self._on_press)
        self._listener.daemon = True
        self._listener.start()
        log.info("Keyboard tracker started.")

    def stop(self):
        if self._listener:
            self._listener.stop()
            log.info("Keyboard tracker stopped.")

    def _on_press(self, _key):
        with self._lock:
            self._press_count += 1

    def snapshot_and_reset(self) -> dict:
        """Return press count since last call and reset."""
        with self._lock:
            count = self._press_count
            self._press_count = 0
        return {"press_count": count}
