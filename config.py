"""Application configuration for Clipboard Auto Typer."""

from __future__ import annotations

APP_NAME = "Clipboard Auto Typer"
APP_VERSION = "2.1.0"
APP_ID = "ClipboardAutoTyper"

# Clipboard polling interval in seconds. Lower values react faster but use more CPU.
CLIPBOARD_POLL_INTERVAL = 0.50

# Default typing controls.
DEFAULT_DELAY_SECONDS = 3.0
SPEED_PRESETS = {
    "Slow": 5,
    "Normal": 25,
    "Fast": 60,
}
DEFAULT_SPEED_MODE = "Normal"
MIN_CPS = 1
MAX_CPS = 200

# Insertion modes.
# SAFE_INSERT uses Word COM Range.Text in one operation for best exactness.
# HUMAN_TYPE types character by character and uses CPS for visible human-like typing.
INSERTION_MODE_SAFE = "Safe Insert"
INSERTION_MODE_HUMAN = "Human Type"
INSERTION_MODES = (INSERTION_MODE_SAFE, INSERTION_MODE_HUMAN)
DEFAULT_INSERTION_MODE = INSERTION_MODE_SAFE

# Shortcuts. Global shortcuts use the keyboard package and may require permission
# in some Windows environments. Local shortcuts work when the app window is focused.
EMERGENCY_STOP_HOTKEY = "ctrl+alt+esc"
GLOBAL_SHORTCUTS = {
    "start": "ctrl+shift+s",
    "stop": "ctrl+shift+x",
    "pause_resume": "ctrl+shift+p",
    "type_current": "ctrl+shift+t",
    "show_hide": "ctrl+shift+h",
}

LOCAL_SHORTCUTS = {
    "start": "<Control-Shift-S>",
    "stop": "<Control-Shift-X>",
    "pause_resume": "<Control-Shift-P>",
    "type_current": "<Control-Shift-T>",
    "show_hide": "<Control-Shift-H>",
    "emergency_stop": "<Escape>",
}

# Preview/log limits for the GUI.
MAX_PREVIEW_CHARS = 5000
MAX_LOG_LINES = 500

# During Human Type mode, periodically confirm that Word is still active.
WORD_ACTIVE_CHECK_INTERVAL_CHARS = 10

# Queue behavior: if new clipboard content appears while typing, keep the most
# recent text and type that next after the current job completes.
QUEUE_LATEST_WHILE_BUSY = True

# Settings are stored under the user's profile and never contain clipboard text.
SETTINGS_FILENAME = "settings.json"

# Tray behavior.
DEFAULT_MINIMIZE_TO_TRAY = True
