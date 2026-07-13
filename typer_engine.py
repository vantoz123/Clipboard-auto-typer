"""Typing engine for inserting clipboard text into Microsoft Word."""

from __future__ import annotations

import threading
import time
import unicodedata
from dataclasses import dataclass
from typing import Callable, Dict, Iterable, Optional

import config
from word_detector import WordDetector


LogCallback = Callable[[str], None]
DoneCallback = Callable[[bool, str], None]
ProgressCallback = Callable[[int, int], None]


@dataclass(frozen=True)
class TypingOptions:
    characters_per_second: int
    initial_delay_seconds: float
    insertion_mode: str = config.DEFAULT_INSERTION_MODE
    require_word_active: bool = True


class AutoCorrectGuard:
    """Temporarily disables common Word AutoFormat-as-you-type mutations.

    Word can change straight quotes, fractions, list markers, and borders while
    text is being typed. This guard uses best-effort COM property access and
    restores any properties that were successfully changed.
    """

    OPTION_NAMES: Iterable[str] = (
        "AutoFormatAsYouTypeReplaceQuotes",
        "AutoFormatAsYouTypeReplaceSymbols",
        "AutoFormatAsYouTypeReplaceOrdinals",
        "AutoFormatAsYouTypeReplaceFractions",
        "AutoFormatAsYouTypeReplacePlainTextEmphasis",
        "AutoFormatAsYouTypeApplyHeadings",
        "AutoFormatAsYouTypeApplyBorders",
        "AutoFormatAsYouTypeApplyBulletedLists",
        "AutoFormatAsYouTypeApplyNumberedLists",
        "AutoFormatAsYouTypeDefineStyles",
        "AutoFormatAsYouTypeFormatListItemBeginning",
    )

    AUTOCORRECT_NAMES: Iterable[str] = ("ReplaceText",)

    def __init__(self, word_app) -> None:
        self.word_app = word_app
        self._saved_options: Dict[str, object] = {}
        self._saved_autocorrect: Dict[str, object] = {}

    def __enter__(self):
        options = getattr(self.word_app, "Options", None)
        if options is not None:
            for name in self.OPTION_NAMES:
                try:
                    current = getattr(options, name)
                    self._saved_options[name] = current
                    setattr(options, name, False)
                except Exception:
                    continue

        autocorrect = getattr(self.word_app, "AutoCorrect", None)
        if autocorrect is not None:
            for name in self.AUTOCORRECT_NAMES:
                try:
                    current = getattr(autocorrect, name)
                    self._saved_autocorrect[name] = current
                    setattr(autocorrect, name, False)
                except Exception:
                    continue
        return self

    def __exit__(self, exc_type, exc, traceback) -> None:
        options = getattr(self.word_app, "Options", None)
        if options is not None:
            for name, value in self._saved_options.items():
                try:
                    setattr(options, name, value)
                except Exception:
                    continue

        autocorrect = getattr(self.word_app, "AutoCorrect", None)
        if autocorrect is not None:
            for name, value in self._saved_autocorrect.items():
                try:
                    setattr(autocorrect, name, value)
                except Exception:
                    continue


class TyperEngine:
    """Types or inserts text into the active Word selection in a background thread."""

    def __init__(
        self,
        on_log: Optional[LogCallback] = None,
        on_done: Optional[DoneCallback] = None,
        on_progress: Optional[ProgressCallback] = None,
    ) -> None:
        self.on_log = on_log
        self.on_done = on_done
        self.on_progress = on_progress
        self._stop_event = threading.Event()
        self._pause_event = threading.Event()
        self._pause_event.set()
        self._thread: Optional[threading.Thread] = None
        self._lock = threading.Lock()

    @property
    def is_typing(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    @property
    def is_paused(self) -> bool:
        return not self._pause_event.is_set()

    def type_async(self, text: str, options: TypingOptions) -> bool:
        """Start typing text asynchronously. Returns False if already busy."""
        with self._lock:
            if self.is_typing:
                self._log("Typing is already in progress. The latest clipboard text can be queued by the GUI.")
                return False
            self._stop_event.clear()
            self._pause_event.set()
            self._thread = threading.Thread(
                target=self._run_typing,
                args=(text, options),
                name="WordTyperEngine",
                daemon=True,
            )
            self._thread.start()
            return True

    def stop(self) -> None:
        self._stop_event.set()
        self._pause_event.set()
        self._log("Stop requested.")

    def pause(self) -> None:
        if self.is_typing:
            self._pause_event.clear()
            self._log("Typing paused.")

    def resume(self) -> None:
        if self.is_typing:
            self._pause_event.set()
            self._log("Typing resumed.")

    def toggle_pause(self) -> None:
        if self.is_paused:
            self.resume()
        else:
            self.pause()

    def _run_typing(self, text: str, options: TypingOptions) -> None:
        try:
            import pythoncom
        except Exception as exc:
            self._finish(False, f"pywin32 is required for Word automation: {exc}")
            return

        pythoncom.CoInitialize()
        try:
            self._validate_text(text)
            cps = max(config.MIN_CPS, min(config.MAX_CPS, int(options.characters_per_second)))

            if not WordDetector.is_word_running():
                self._finish(False, "Microsoft Word is not running. Open Word and a document first.")
                return

            self._log(
                f"Typing will begin after {options.initial_delay_seconds:.1f} second(s). "
                "Click inside the target Word document now."
            )
            if not self._wait_for_initial_delay(options.initial_delay_seconds):
                self._finish(False, "Typing stopped before it started.")
                return

            if options.require_word_active and not WordDetector.is_word_active():
                self._finish(False, "Microsoft Word is not active. Click inside Word and try again.")
                return

            word_app = WordDetector.get_active_word_application()
            selection = getattr(word_app, "Selection", None)
            if selection is None:
                self._finish(False, "No active Word selection was found. Click inside a Word document and try again.")
                return

            mode = options.insertion_mode if options.insertion_mode in config.INSERTION_MODES else config.DEFAULT_INSERTION_MODE
            if mode == config.INSERTION_MODE_SAFE:
                self._safe_insert(selection, text)
            else:
                self._human_type(word_app, selection, text, cps, options)
        except Exception as exc:
            self._finish(False, f"Typing failed: {exc}")
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def _safe_insert(self, selection, text: str) -> None:
        """Insert text in one Word COM operation for maximum Unicode fidelity."""
        total = len(text)
        self._log(
            f"Safe Insert mode: inserting {total} character(s) directly through Word. "
            "This is usually more exact than simulated keystrokes."
        )
        if self.on_progress:
            self.on_progress(0, total)
        if self._stop_event.is_set():
            self._finish(False, "Typing stopped before insertion.")
            return
        self._pause_event.wait()
        word_text = normalize_line_breaks_for_word(text)
        # Setting Selection.Range.Text replaces selected text or inserts at the caret.
        selection.Range.Text = word_text
        # Collapse to the end of the inserted text. Word constant wdCollapseEnd = 0.
        try:
            selection.Collapse(0)
        except Exception:
            pass
        if self.on_progress:
            self.on_progress(total, total)
        self._finish(True, f"Safe Insert completed successfully. {total} character(s) processed.")

    def _human_type(self, word_app, selection, text: str, cps: int, options: TypingOptions) -> None:
        total = len(text)
        typed = 0
        delay_per_character = 1.0 / cps
        self._log(f"Human Type mode: typing {total} character(s) at {cps} character(s) per second.")

        with AutoCorrectGuard(word_app):
            index = 0
            while index < total:
                if self._stop_event.is_set():
                    self._finish(False, f"Typing stopped after {typed} of {total} character(s).")
                    return

                self._pause_event.wait()

                if options.require_word_active and typed % config.WORD_ACTIVE_CHECK_INTERVAL_CHARS == 0:
                    if not WordDetector.is_word_active():
                        self._log("Microsoft Word lost focus. Typing is paused until Word is active again.")
                        while not self._stop_event.is_set() and not WordDetector.is_word_active():
                            time.sleep(0.20)
                        if self._stop_event.is_set():
                            self._finish(False, f"Typing stopped after {typed} of {total} character(s).")
                            return
                        self._log("Microsoft Word is active again. Typing resumed.")

                character = text[index]
                if character == "\r":
                    if index + 1 < total and text[index + 1] == "\n":
                        index += 1
                        typed += 1
                    selection.TypeParagraph()
                elif character == "\n":
                    selection.TypeParagraph()
                elif character == "\t":
                    selection.TypeText("\t")
                else:
                    selection.TypeText(character)

                typed += 1
                index += 1
                if self.on_progress:
                    self.on_progress(typed, total)
                time.sleep(delay_per_character)

        self._finish(True, f"Typing completed successfully. {typed} character(s) processed.")

    def _wait_for_initial_delay(self, seconds: float) -> bool:
        end_at = time.time() + max(0.0, seconds)
        while time.time() < end_at:
            if self._stop_event.is_set():
                return False
            self._pause_event.wait(timeout=0.05)
            time.sleep(0.05)
        return True

    @staticmethod
    def _validate_text(text: str) -> None:
        if not isinstance(text, str):
            raise ValueError("Clipboard content is not text.")
        if text == "":
            raise ValueError("Clipboard is empty.")
        unsupported = []
        for character in text:
            if character in "\r\n\t":
                continue
            # Other control characters are usually not representable as typed text.
            if unicodedata.category(character).startswith("C"):
                unsupported.append(repr(character))
                if len(unsupported) >= 5:
                    break
        if unsupported:
            raise ValueError("Clipboard contains unsupported control characters: " + ", ".join(unsupported))

    def _log(self, message: str) -> None:
        if self.on_log:
            self.on_log(message)

    def _finish(self, success: bool, message: str) -> None:
        if self.on_done:
            self.on_done(success, message)
        else:
            self._log(message)


def normalize_line_breaks_for_word(text: str) -> str:
    """Convert common text newlines to Word paragraph marks.

    Word stores paragraph breaks as carriage returns. The visible result in Word
    matches copied lines and blank lines, while avoiding inconsistent CRLF/LF
    handling across Word COM versions.
    """
    return text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r")
