"""Clipboard monitoring utilities.

This module reads text from the system clipboard without writing anything back to
it. It intentionally avoids network access, persistent storage, and clipboard
mutation.
"""

from __future__ import annotations

import threading
import time
from dataclasses import dataclass
from typing import Callable, Optional

import pyperclip

import config


@dataclass(frozen=True)
class ClipboardEvent:
    """Represents newly detected text clipboard content."""

    text: str
    length: int
    detected_at: float


class ClipboardMonitor:
    """Polls the system clipboard and reports newly copied text."""

    def __init__(
        self,
        interval_seconds: float = config.CLIPBOARD_POLL_INTERVAL,
        on_text: Optional[Callable[[ClipboardEvent], None]] = None,
        on_error: Optional[Callable[[str], None]] = None,
        on_status: Optional[Callable[[str], None]] = None,
    ) -> None:
        self.interval_seconds = max(0.10, interval_seconds)
        self.on_text = on_text
        self.on_error = on_error
        self.on_status = on_status
        self._stop_event = threading.Event()
        self._thread: Optional[threading.Thread] = None
        self._last_text: Optional[str] = None
        self._last_error_message: Optional[str] = None

    @property
    def is_running(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    def start(self, ignore_current_clipboard: bool = True) -> None:
        """Start monitoring in a background thread.

        If ignore_current_clipboard is True, the current clipboard value becomes
        the baseline. This prevents old text from being typed immediately when
        monitoring starts.
        """
        if self.is_running:
            self._emit_status("Clipboard monitor is already running.")
            return
        if ignore_current_clipboard:
            try:
                self._last_text = self.get_current_text()
            except Exception:
                self._last_text = None
        self._stop_event.clear()
        self._thread = threading.Thread(target=self._run, name="ClipboardMonitor", daemon=True)
        self._thread.start()
        self._emit_status("Clipboard monitoring started.")

    def stop(self) -> None:
        """Stop monitoring."""
        self._stop_event.set()
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=2.0)
        self._emit_status("Clipboard monitoring stopped.")

    def reset_last_seen(self) -> None:
        """Forget the last observed clipboard text so the current text can be reprocessed."""
        self._last_text = None

    @staticmethod
    def get_current_text() -> str:
        """Return current clipboard text without modifying the clipboard."""
        value = pyperclip.paste()
        if value is None:
            return ""
        if not isinstance(value, str):
            raise TypeError("Clipboard content is not text.")
        return value

    def _run(self) -> None:
        while not self._stop_event.is_set():
            try:
                text = self.get_current_text()
                self._last_error_message = None
            except Exception as exc:  # pyperclip raises platform-specific exceptions.
                message = f"Unable to read clipboard text: {exc}"
                if message != self._last_error_message:
                    self._emit_error(message)
                    self._last_error_message = message
                time.sleep(self.interval_seconds)
                continue

            if text and text != self._last_text:
                self._last_text = text
                if self.on_text:
                    self.on_text(ClipboardEvent(text=text, length=len(text), detected_at=time.time()))
            elif text == "" and self._last_text is None:
                self._last_text = ""

            time.sleep(self.interval_seconds)

    def _emit_error(self, message: str) -> None:
        if self.on_error:
            self.on_error(message)

    def _emit_status(self, message: str) -> None:
        if self.on_status:
            self.on_status(message)
