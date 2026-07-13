"""Small settings helper for non-sensitive user preferences.

Only UI preferences are stored. Clipboard text is never written to disk.
"""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any, Dict

import config


DEFAULT_SETTINGS: Dict[str, Any] = {
    "speed_mode": config.DEFAULT_SPEED_MODE,
    "custom_cps": config.SPEED_PRESETS[config.DEFAULT_SPEED_MODE],
    "delay_seconds": config.DEFAULT_DELAY_SECONDS,
    "insertion_mode": config.DEFAULT_INSERTION_MODE,
    "ignore_current_clipboard_on_start": True,
    "minimize_to_tray": config.DEFAULT_MINIMIZE_TO_TRAY,
}


class SettingsManager:
    """Loads and saves non-sensitive settings."""

    def __init__(self) -> None:
        appdata = os.environ.get("APPDATA")
        if appdata:
            self.settings_dir = Path(appdata) / config.APP_ID
        else:
            self.settings_dir = Path.home() / f".{config.APP_ID}"
        self.settings_path = self.settings_dir / config.SETTINGS_FILENAME

    def load(self) -> Dict[str, Any]:
        settings = dict(DEFAULT_SETTINGS)
        try:
            if self.settings_path.exists():
                raw = json.loads(self.settings_path.read_text(encoding="utf-8"))
                if isinstance(raw, dict):
                    settings.update(raw)
        except Exception:
            # Corrupt settings should not prevent the application from starting.
            return settings
        return self._sanitize(settings)

    def save(self, settings: Dict[str, Any]) -> None:
        clean = self._sanitize(settings)
        self.settings_dir.mkdir(parents=True, exist_ok=True)
        self.settings_path.write_text(json.dumps(clean, indent=2), encoding="utf-8")

    @staticmethod
    def _sanitize(settings: Dict[str, Any]) -> Dict[str, Any]:
        clean = dict(DEFAULT_SETTINGS)
        speed_mode = str(settings.get("speed_mode", clean["speed_mode"]))
        clean["speed_mode"] = speed_mode if speed_mode in (*config.SPEED_PRESETS.keys(), "Custom") else config.DEFAULT_SPEED_MODE

        try:
            cps = int(settings.get("custom_cps", clean["custom_cps"]))
        except (TypeError, ValueError):
            cps = int(DEFAULT_SETTINGS["custom_cps"])
        clean["custom_cps"] = max(config.MIN_CPS, min(config.MAX_CPS, cps))

        try:
            delay = float(settings.get("delay_seconds", clean["delay_seconds"]))
        except (TypeError, ValueError):
            delay = float(DEFAULT_SETTINGS["delay_seconds"])
        clean["delay_seconds"] = max(0.0, delay)

        insertion_mode = str(settings.get("insertion_mode", clean["insertion_mode"]))
        clean["insertion_mode"] = insertion_mode if insertion_mode in config.INSERTION_MODES else config.DEFAULT_INSERTION_MODE
        clean["ignore_current_clipboard_on_start"] = bool(settings.get("ignore_current_clipboard_on_start", True))
        clean["minimize_to_tray"] = bool(settings.get("minimize_to_tray", config.DEFAULT_MINIMIZE_TO_TRAY))
        return clean
