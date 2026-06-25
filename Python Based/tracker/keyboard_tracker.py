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
        self._started = False

    @property
    def is_alive(self) -> bool:
        """True if the listener thread is running."""
        return self._listener is not None and self._listener.is_alive()

    def start(self):
        self._started = True
        self._start_listener()
        log.info("Keyboard tracker started.")

    def _start_listener(self):
        """Create and start a fresh listener."""
        try:
            if self._listener is not None:
                try:
                    self._listener.stop()
                except Exception:
                    pass
            self._listener = keyboard.Listener(on_press=self._on_press)
            self._listener.daemon = True
            self._listener.start()
        except Exception as exc:
            log.error("Failed to start keyboard listener: %s", exc)

    def stop(self):
        self._started = False
        if self._listener:
            try:
                self._listener.stop()
            except Exception:
                pass
            log.info("Keyboard tracker stopped.")

    def ensure_alive(self):
        """Restart the listener if it has died."""
        if self._started and not self.is_alive:
            log.warning("Keyboard listener died, restarting...")
            self._start_listener()

    def _on_press(self, _key):
        with self._lock:
            self._press_count += 1

    def snapshot_and_reset(self) -> dict:
        """Return press count since last call and reset."""
        with self._lock:
            count = self._press_count
            self._press_count = 0
        return {"press_count": count}
