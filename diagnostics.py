"""Beginner-friendly diagnostics for Clipboard Auto Typer."""

from __future__ import annotations

import importlib
import platform
import sys


def check_module(name: str) -> str:
    try:
        module = importlib.import_module(name)
        version = getattr(module, "__version__", "installed")
        return f"OK - {name}: {version}"
    except Exception as exc:
        return f"MISSING/ERROR - {name}: {exc}"


def main() -> None:
    print("Clipboard Auto Typer Diagnostics")
    print("=" * 36)
    print(f"Python: {sys.version}")
    print(f"Executable: {sys.executable}")
    print(f"Platform: {platform.platform()}")
    print()
    for module_name in ["tkinter", "pyperclip", "psutil", "keyboard", "pystray", "PIL"]:
        print(check_module(module_name))
    if platform.system().lower() == "windows":
        print(check_module("pythoncom"))
        print(check_module("win32com.client"))
        try:
            from word_detector import WordDetector

            print(f"Word running: {WordDetector.is_word_running()}")
            ready, message = WordDetector.readiness_message()
            print(f"Word ready: {ready} - {message}")
        except Exception as exc:
            print(f"Word check failed: {exc}")
    else:
        print("Word automation check skipped: this feature requires Windows.")
    print()
    print("If anything says MISSING/ERROR, run install_dependencies.bat again.")


if __name__ == "__main__":
    main()
