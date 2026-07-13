"""Microsoft Word detection and readiness checks for Windows."""

from __future__ import annotations

import platform
from dataclasses import dataclass
from typing import Optional, Tuple

import psutil


@dataclass(frozen=True)
class ActiveWindowInfo:
    hwnd: int
    title: str
    process_id: int
    process_name: str


class WordDetector:
    """Detects whether Microsoft Word is running and active."""

    WORD_PROCESS_NAME = "winword.exe"

    @staticmethod
    def is_windows() -> bool:
        return platform.system().lower() == "windows"

    @classmethod
    def is_word_running(cls) -> bool:
        """Return True when WINWORD.EXE is running."""
        if not cls.is_windows():
            return False
        for process in psutil.process_iter(attrs=["name"]):
            try:
                name = (process.info.get("name") or "").lower()
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
            if name == cls.WORD_PROCESS_NAME:
                return True
        return False

    @classmethod
    def active_window_info(cls) -> Optional[ActiveWindowInfo]:
        """Return information about the foreground window, if available."""
        if not cls.is_windows():
            return None
        try:
            import win32gui
            import win32process

            hwnd = win32gui.GetForegroundWindow()
            if not hwnd:
                return None
            title = win32gui.GetWindowText(hwnd)
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            process_name = ""
            try:
                process_name = psutil.Process(pid).name()
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                process_name = ""
            return ActiveWindowInfo(hwnd=hwnd, title=title, process_id=pid, process_name=process_name)
        except Exception:
            return None

    @classmethod
    def is_word_active(cls) -> bool:
        """Return True when Microsoft Word is the foreground application."""
        info = cls.active_window_info()
        if not info:
            return False
        name_matches = info.process_name.lower() == cls.WORD_PROCESS_NAME
        title_matches = "word" in info.title.lower()
        return bool(name_matches or title_matches)

    @classmethod
    def get_active_word_application(cls):
        """Return the running Word COM application object.

        The caller must call pythoncom.CoInitialize() in the current thread before
        using this method.
        """
        if not cls.is_windows():
            raise RuntimeError("Microsoft Word integration is available only on Windows.")
        try:
            import win32com.client

            return win32com.client.GetActiveObject("Word.Application")
        except Exception as exc:
            raise RuntimeError("Microsoft Word is not running or cannot be accessed.") from exc

    @classmethod
    def active_document_name(cls) -> str:
        """Return the active Word document name when available."""
        if not cls.is_windows() or not cls.is_word_running():
            return ""
        try:
            import pythoncom

            pythoncom.CoInitialize()
            try:
                word_app = cls.get_active_word_application()
                doc = getattr(word_app, "ActiveDocument", None)
                return str(getattr(doc, "Name", "") or "")
            finally:
                pythoncom.CoUninitialize()
        except Exception:
            return ""

    @classmethod
    def readiness_message(cls) -> Tuple[bool, str]:
        """Return readiness state and a user-friendly message."""
        if not cls.is_windows():
            return False, "This application requires Windows for Microsoft Word automation."
        if not cls.is_word_running():
            return False, "Microsoft Word is not running. Open Word and a document first."
        if not cls.is_word_active():
            doc_name = cls.active_document_name()
            if doc_name:
                return False, f"Word is running ({doc_name}) but not active. Click inside the target document."
            return False, "Microsoft Word is running but not active. Click inside the target Word document."
        doc_name = cls.active_document_name()
        if doc_name:
            return True, f"Microsoft Word is active: {doc_name}"
        return True, "Microsoft Word is active and ready."
