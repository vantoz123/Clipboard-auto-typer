"""Tkinter GUI for Clipboard Auto Typer."""

from __future__ import annotations

import queue
import threading
import tkinter as tk
from datetime import datetime
from tkinter import messagebox, scrolledtext, ttk
from typing import Callable, Optional

import config
from clipboard_manager import ClipboardEvent, ClipboardMonitor
from settings_manager import SettingsManager
from typer_engine import TyperEngine, TypingOptions
from word_detector import WordDetector

try:
    import keyboard  # type: ignore
except Exception:  # The app still works without optional global hotkeys.
    keyboard = None

try:
    import pystray  # type: ignore
    from PIL import Image, ImageDraw  # type: ignore
except Exception:  # Tray behavior is optional; the app still runs normally.
    pystray = None
    Image = None
    ImageDraw = None


class ClipboardAutoTyperApp:
    """Main GUI application."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(f"{config.APP_NAME} {config.APP_VERSION}")
        self.root.minsize(940, 720)

        self.ui_queue: "queue.Queue[tuple[str, object]]" = queue.Queue()
        self.current_clipboard_text = ""
        self.queued_clipboard_text: Optional[str] = None
        self.hotkey_registered_names: list[str] = []
        self.tray_icon = None
        self.tray_thread: Optional[threading.Thread] = None
        self.exiting = False
        self._minimize_event_in_progress = False

        self.settings_manager = SettingsManager()
        self.settings = self.settings_manager.load()

        self.monitor = ClipboardMonitor(
            on_text=self._on_clipboard_text,
            on_error=lambda message: self._enqueue("log", f"ERROR: {message}"),
            on_status=lambda message: self._enqueue("log", message),
        )
        self.typer = TyperEngine(
            on_log=lambda message: self._enqueue("log", message),
            on_done=lambda success, message: self._enqueue("done", (success, message)),
            on_progress=lambda current, total: self._enqueue("progress", (current, total)),
        )

        self.status_var = tk.StringVar(value="Stopped")
        self.word_status_var = tk.StringVar(value="Word status unknown")
        self.speed_mode_var = tk.StringVar(value=str(self.settings["speed_mode"]))
        self.custom_cps_var = tk.StringVar(value=str(self.settings["custom_cps"]))
        self.delay_var = tk.StringVar(value=str(self.settings["delay_seconds"]))
        self.insertion_mode_var = tk.StringVar(value=str(self.settings["insertion_mode"]))
        self.ignore_current_var = tk.BooleanVar(value=bool(self.settings["ignore_current_clipboard_on_start"]))
        self.minimize_to_tray_var = tk.BooleanVar(value=bool(self.settings["minimize_to_tray"]))
        self.progress_var = tk.StringVar(value="Progress: idle")
        self.queue_var = tk.StringVar(value="Queue: empty")
        self.shortcut_status_var = tk.StringVar(value="Shortcuts: initializing")
        self.tray_status_var = tk.StringVar(value="Tray: initializing")

        self._build_widgets()
        self._register_shortcuts()
        self._setup_tray_icon()
        self._schedule_queue_processing()
        self._schedule_word_status_update()
        self.root.protocol("WM_DELETE_WINDOW", self.confirm_exit)
        self.root.bind("<Unmap>", self._on_window_unmap)
        self._log("Application started. Clipboard text is never saved to disk or sent online.")

    def _build_widgets(self) -> None:
        outer = ttk.Frame(self.root, padding=12)
        outer.pack(fill=tk.BOTH, expand=True)

        title = ttk.Label(outer, text=config.APP_NAME, font=("Segoe UI", 18, "bold"))
        title.pack(anchor=tk.W)
        subtitle = ttk.Label(
            outer,
            text="Reads copied text locally and inserts it into the active Microsoft Word document.",
        )
        subtitle.pack(anchor=tk.W, pady=(0, 10))

        controls = ttk.LabelFrame(outer, text="Controls", padding=10)
        controls.pack(fill=tk.X)

        button_row = ttk.Frame(controls)
        button_row.pack(fill=tk.X)
        ttk.Button(button_row, text="Start Monitoring", command=self.start_monitoring).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(button_row, text="Stop", command=self.stop_all).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(button_row, text="Pause/Resume", command=self.toggle_pause_typing).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(button_row, text="Type Current Clipboard", command=self.type_current_clipboard).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(button_row, text="Hide to Tray", command=self.hide_to_tray).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(button_row, text="Exit", command=self.confirm_exit).pack(side=tk.RIGHT)

        status_row = ttk.Frame(controls)
        status_row.pack(fill=tk.X, pady=(10, 0))
        ttk.Label(status_row, text="Monitoring:").pack(side=tk.LEFT)
        ttk.Label(status_row, textvariable=self.status_var, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=(5, 20))
        ttk.Label(status_row, text="Word:").pack(side=tk.LEFT)
        ttk.Label(status_row, textvariable=self.word_status_var, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=(5, 20))
        ttk.Label(status_row, textvariable=self.progress_var).pack(side=tk.LEFT, padx=(5, 20))
        ttk.Label(status_row, textvariable=self.queue_var).pack(side=tk.LEFT, padx=(5, 20))

        speed_frame = ttk.LabelFrame(outer, text="Typing Speed, Delay, and Mode", padding=10)
        speed_frame.pack(fill=tk.X, pady=(10, 0))

        for mode in ("Slow", "Normal", "Fast", "Custom"):
            ttk.Radiobutton(
                speed_frame,
                text=mode,
                value=mode,
                variable=self.speed_mode_var,
                command=self._save_settings_safely,
            ).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(speed_frame, text="Custom CPS:").pack(side=tk.LEFT, padx=(8, 4))
        custom_entry = ttk.Entry(speed_frame, textvariable=self.custom_cps_var, width=8)
        custom_entry.pack(side=tk.LEFT, padx=(0, 16))
        custom_entry.bind("<FocusOut>", lambda _event: self._save_settings_safely())

        ttk.Label(speed_frame, text="Delay before typing (seconds):").pack(side=tk.LEFT, padx=(0, 4))
        delay_entry = ttk.Entry(speed_frame, textvariable=self.delay_var, width=8)
        delay_entry.pack(side=tk.LEFT, padx=(0, 16))
        delay_entry.bind("<FocusOut>", lambda _event: self._save_settings_safely())

        ttk.Label(speed_frame, text="Insertion mode:").pack(side=tk.LEFT, padx=(0, 4))
        insertion_combo = ttk.Combobox(
            speed_frame,
            textvariable=self.insertion_mode_var,
            values=config.INSERTION_MODES,
            state="readonly",
            width=14,
        )
        insertion_combo.pack(side=tk.LEFT)
        insertion_combo.bind("<<ComboboxSelected>>", lambda _event: self._save_settings_safely())

        options_frame = ttk.LabelFrame(outer, text="Options", padding=10)
        options_frame.pack(fill=tk.X, pady=(10, 0))
        ttk.Checkbutton(
            options_frame,
            text="Ignore current clipboard when monitoring starts",
            variable=self.ignore_current_var,
            command=self._save_settings_safely,
        ).pack(side=tk.LEFT, padx=(0, 18))
        ttk.Checkbutton(
            options_frame,
            text="Minimize to system tray instead of taskbar",
            variable=self.minimize_to_tray_var,
            command=self._save_settings_safely,
        ).pack(side=tk.LEFT, padx=(0, 18))
        ttk.Label(options_frame, textvariable=self.shortcut_status_var).pack(side=tk.LEFT, padx=(0, 18))
        ttk.Label(options_frame, textvariable=self.tray_status_var).pack(side=tk.LEFT)

        shortcuts_frame = ttk.LabelFrame(outer, text="Shortcut Keys", padding=10)
        shortcuts_frame.pack(fill=tk.X, pady=(10, 0))
        shortcuts_text = (
            "Start: Ctrl+Shift+S   |   Stop: Ctrl+Shift+X   |   Pause/Resume: Ctrl+Shift+P   |   "
            "Type Clipboard: Ctrl+Shift+T   |   Show/Hide: Ctrl+Shift+H   |   Emergency Stop: Ctrl+Alt+Esc or Esc in app"
        )
        ttk.Label(shortcuts_frame, text=shortcuts_text, wraplength=900).pack(anchor=tk.W)

        preview_frame = ttk.LabelFrame(outer, text="Current Clipboard Preview", padding=10)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        self.preview_text = scrolledtext.ScrolledText(preview_frame, height=10, wrap=tk.WORD)
        self.preview_text.pack(fill=tk.BOTH, expand=True)
        self.preview_text.configure(state=tk.DISABLED)

        log_frame = ttk.LabelFrame(outer, text="Activity Log", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        self.log_text = scrolledtext.ScrolledText(log_frame, height=10, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.configure(state=tk.DISABLED)

        help_note = ttk.Label(
            outer,
            text="Workflow: open Word, click inside the target document, copy text, then wait for the delay or press Type Current Clipboard.",
            wraplength=900,
        )
        help_note.pack(anchor=tk.W, pady=(8, 0))

    def start_monitoring(self) -> None:
        self._save_settings_safely()
        self.monitor.start(ignore_current_clipboard=bool(self.ignore_current_var.get()))
        self.status_var.set("Monitoring")
        self._log("Monitoring enabled. Copy text, then click inside the target Word document before the delay ends.")

    def stop_all(self) -> None:
        self.monitor.stop()
        self.typer.stop()
        self.queued_clipboard_text = None
        self.status_var.set("Stopped")
        self.progress_var.set("Progress: stopped")
        self.queue_var.set("Queue: empty")

    def pause_typing(self) -> None:
        self.typer.pause()
        if self.typer.is_typing:
            self.progress_var.set("Progress: paused")

    def resume_typing(self) -> None:
        self.typer.resume()

    def toggle_pause_typing(self) -> None:
        if not self.typer.is_typing:
            self._log("Nothing is currently typing.")
            return
        self.typer.toggle_pause()
        self.progress_var.set("Progress: paused" if self.typer.is_paused else "Progress: resumed")

    def type_current_clipboard(self) -> None:
        try:
            text = ClipboardMonitor.get_current_text()
        except Exception as exc:
            self._log(f"ERROR: Unable to read clipboard: {exc}")
            messagebox.showerror(config.APP_NAME, f"Unable to read clipboard:\n{exc}")
            return
        if not text:
            self._log("WARNING: Clipboard is empty.")
            messagebox.showwarning(config.APP_NAME, "Clipboard is empty. Copy text first.")
            return
        self.current_clipboard_text = text
        self._update_preview(text)
        self._request_typing(text)

    def hide_to_tray(self) -> None:
        if pystray is None:
            self._log("System tray support is unavailable. Install pystray and Pillow, then restart the app.")
            self.root.iconify()
            return
        self._log("Window hidden to system tray. Use the tray icon or Ctrl+Shift+H to restore it.")
        self.root.withdraw()

    def show_window(self) -> None:
        self.root.after(0, self._show_window_on_ui_thread)

    def _show_window_on_ui_thread(self) -> None:
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()
        try:
            self.root.attributes("-topmost", True)
            self.root.after(250, lambda: self.root.attributes("-topmost", False))
        except Exception:
            pass

    def toggle_window_visibility(self) -> None:
        if self.root.state() == "withdrawn" or not self.root.winfo_viewable():
            self.show_window()
        else:
            self.hide_to_tray()

    def confirm_exit(self) -> None:
        if self.typer.is_typing:
            ok = messagebox.askyesno(config.APP_NAME, "Typing is in progress. Stop typing and exit?")
            if not ok:
                return
        self.exit_app()

    def exit_app(self) -> None:
        self.exiting = True
        self._save_settings_safely()
        self.stop_all()
        self._unregister_shortcuts()
        if self.tray_icon is not None:
            try:
                self.tray_icon.stop()
            except Exception:
                pass
        self.root.destroy()

    def _on_clipboard_text(self, event: ClipboardEvent) -> None:
        self.current_clipboard_text = event.text
        self._enqueue("preview", event.text)
        self._enqueue("log", f"New text detected on clipboard: {event.length} character(s).")
        self._request_typing(event.text)

    def _request_typing(self, text: str) -> None:
        cps = self._get_characters_per_second()
        if cps is None:
            return
        delay = self._get_delay_seconds()
        if delay is None:
            return

        if self.typer.is_typing:
            if config.QUEUE_LATEST_WHILE_BUSY:
                self.queued_clipboard_text = text
                self._enqueue("queue", f"Queue: 1 latest item ({len(text)} chars)")
                self._enqueue("log", "Typing is busy. The latest clipboard text was queued and will run next.")
            else:
                self._enqueue("log", "Typing is busy. New clipboard text was ignored.")
            return

        if not WordDetector.is_word_running():
            self._enqueue("log", "WARNING: Microsoft Word is not running. Open Word and a document first.")
            return

        if not WordDetector.is_word_active():
            self._enqueue(
                "log",
                "WARNING: Word is not active yet. Click inside the target Word document before the delay ends.",
            )

        options = TypingOptions(
            characters_per_second=cps,
            initial_delay_seconds=delay,
            insertion_mode=self.insertion_mode_var.get(),
        )
        started = self.typer.type_async(text, options)
        if started:
            self._enqueue("progress", (0, len(text)))

    def _get_characters_per_second(self) -> Optional[int]:
        mode = self.speed_mode_var.get()
        if mode in config.SPEED_PRESETS:
            return config.SPEED_PRESETS[mode]
        try:
            cps = int(self.custom_cps_var.get())
        except ValueError:
            self._log("ERROR: Custom characters-per-second must be a whole number.")
            return None
        if cps < config.MIN_CPS or cps > config.MAX_CPS:
            self._log(f"ERROR: Custom CPS must be between {config.MIN_CPS} and {config.MAX_CPS}.")
            return None
        return cps

    def _get_delay_seconds(self) -> Optional[float]:
        try:
            delay = float(self.delay_var.get())
        except ValueError:
            self._log("ERROR: Delay must be a number of seconds.")
            return None
        if delay < 0:
            self._log("ERROR: Delay cannot be negative.")
            return None
        return delay

    def _register_shortcuts(self) -> None:
        # Local shortcuts work while the Tk window has focus.
        self.root.bind(config.LOCAL_SHORTCUTS["start"], lambda _event: self.start_monitoring())
        self.root.bind(config.LOCAL_SHORTCUTS["stop"], lambda _event: self.stop_all())
        self.root.bind(config.LOCAL_SHORTCUTS["pause_resume"], lambda _event: self.toggle_pause_typing())
        self.root.bind(config.LOCAL_SHORTCUTS["type_current"], lambda _event: self.type_current_clipboard())
        self.root.bind(config.LOCAL_SHORTCUTS["show_hide"], lambda _event: self.toggle_window_visibility())
        self.root.bind(config.LOCAL_SHORTCUTS["emergency_stop"], lambda _event: self.stop_all())

        if keyboard is None:
            self.shortcut_status_var.set("Shortcuts: local only")
            self._log("Global shortcut library is unavailable. Local shortcuts work when the app window is focused.")
            return

        callbacks: dict[str, Callable[[], None]] = {
            config.GLOBAL_SHORTCUTS["start"]: lambda: self.root.after(0, self.start_monitoring),
            config.GLOBAL_SHORTCUTS["stop"]: lambda: self.root.after(0, self.stop_all),
            config.GLOBAL_SHORTCUTS["pause_resume"]: lambda: self.root.after(0, self.toggle_pause_typing),
            config.GLOBAL_SHORTCUTS["type_current"]: lambda: self.root.after(0, self.type_current_clipboard),
            config.GLOBAL_SHORTCUTS["show_hide"]: lambda: self.root.after(0, self.toggle_window_visibility),
            config.EMERGENCY_STOP_HOTKEY: lambda: self.root.after(0, self.stop_all),
        }
        registered = 0
        for hotkey, callback in callbacks.items():
            try:
                keyboard.add_hotkey(hotkey, callback)
                self.hotkey_registered_names.append(hotkey)
                registered += 1
            except Exception as exc:
                self._log(f"WARNING: Could not register global shortcut {hotkey.upper()}: {exc}")
        self.shortcut_status_var.set(f"Shortcuts: {registered} global registered")
        if registered:
            self._log("Global shortcuts registered. Use Ctrl+Alt+Esc for emergency stop.")

    def _unregister_shortcuts(self) -> None:
        if keyboard is None:
            return
        for hotkey in self.hotkey_registered_names:
            try:
                keyboard.remove_hotkey(hotkey)
            except Exception:
                pass
        self.hotkey_registered_names.clear()

    def _setup_tray_icon(self) -> None:
        if pystray is None or Image is None or ImageDraw is None:
            self.tray_status_var.set("Tray: unavailable")
            self._log("Tray support unavailable. Install pystray and Pillow for minimize-to-tray.")
            return

        image = self._create_tray_image()
        menu = pystray.Menu(
            pystray.MenuItem("Show Clipboard Auto Typer", lambda _icon, _item: self.root.after(0, self.show_window)),
            pystray.MenuItem("Start Monitoring", lambda _icon, _item: self.root.after(0, self.start_monitoring)),
            pystray.MenuItem("Stop", lambda _icon, _item: self.root.after(0, self.stop_all)),
            pystray.MenuItem("Pause/Resume", lambda _icon, _item: self.root.after(0, self.toggle_pause_typing)),
            pystray.MenuItem("Type Current Clipboard", lambda _icon, _item: self.root.after(0, self.type_current_clipboard)),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Exit", lambda _icon, _item: self.root.after(0, self.confirm_exit)),
        )
        self.tray_icon = pystray.Icon(config.APP_ID, image, config.APP_NAME, menu)
        self.tray_thread = threading.Thread(target=self.tray_icon.run, name="TrayIcon", daemon=True)
        self.tray_thread.start()
        self.tray_status_var.set("Tray: active")
        self._log("System tray icon started. Minimize the window to hide it from the taskbar.")

    @staticmethod
    def _create_tray_image():
        size = 64
        image = Image.new("RGB", (size, size), "white")
        draw = ImageDraw.Draw(image)
        draw.rectangle((8, 8, 56, 56), outline="black", width=4)
        draw.rectangle((18, 16, 46, 22), fill="black")
        draw.rectangle((18, 30, 46, 36), fill="black")
        draw.rectangle((18, 44, 38, 50), fill="black")
        return image

    def _on_window_unmap(self, _event) -> None:
        if self.exiting or not self.minimize_to_tray_var.get():
            return
        if pystray is None:
            return
        # Tk reports iconic state after the minimize event has settled.
        if self._minimize_event_in_progress:
            return
        self._minimize_event_in_progress = True
        self.root.after(150, self._hide_if_minimized)

    def _hide_if_minimized(self) -> None:
        try:
            if not self.exiting and self.root.state() == "iconic" and self.minimize_to_tray_var.get():
                self.root.withdraw()
                self._log("Window minimized to system tray.")
        finally:
            self._minimize_event_in_progress = False

    def _schedule_queue_processing(self) -> None:
        self._process_queue()
        self.root.after(100, self._schedule_queue_processing)

    def _process_queue(self) -> None:
        while True:
            try:
                item_type, value = self.ui_queue.get_nowait()
            except queue.Empty:
                break
            if item_type == "log":
                self._log(str(value))
            elif item_type == "preview":
                self._update_preview(str(value))
            elif item_type == "progress":
                current, total = value  # type: ignore[misc]
                self.progress_var.set(f"Progress: {current}/{total}")
            elif item_type == "queue":
                self.queue_var.set(str(value))
            elif item_type == "done":
                success, message = value  # type: ignore[misc]
                self._log(("SUCCESS: " if success else "WARNING: ") + str(message))
                self.progress_var.set("Progress: complete" if success else "Progress: stopped")
                self._start_queued_text_if_any()

    def _start_queued_text_if_any(self) -> None:
        if not self.queued_clipboard_text:
            self.queue_var.set("Queue: empty")
            return
        next_text = self.queued_clipboard_text
        self.queued_clipboard_text = None
        self.queue_var.set("Queue: starting latest item")
        self._log("Starting queued clipboard text.")
        self._request_typing(next_text)

    def _schedule_word_status_update(self) -> None:
        ready, message = WordDetector.readiness_message()
        self.word_status_var.set("Ready" if ready else message)
        self.root.after(1000, self._schedule_word_status_update)

    def _enqueue(self, item_type: str, value: object) -> None:
        self.ui_queue.put((item_type, value))

    def _update_preview(self, text: str) -> None:
        display_text = text
        if len(display_text) > config.MAX_PREVIEW_CHARS:
            display_text = display_text[: config.MAX_PREVIEW_CHARS] + "\n\n[Preview truncated]"
        self.preview_text.configure(state=tk.NORMAL)
        self.preview_text.delete("1.0", tk.END)
        self.preview_text.insert(tk.END, display_text)
        self.preview_text.configure(state=tk.DISABLED)

    def _log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self._trim_log()
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _trim_log(self) -> None:
        line_count = int(self.log_text.index("end-1c").split(".")[0])
        if line_count > config.MAX_LOG_LINES:
            self.log_text.delete("1.0", f"{line_count - config.MAX_LOG_LINES}.0")

    def _save_settings_safely(self) -> None:
        settings = {
            "speed_mode": self.speed_mode_var.get(),
            "custom_cps": self.custom_cps_var.get(),
            "delay_seconds": self.delay_var.get(),
            "insertion_mode": self.insertion_mode_var.get(),
            "ignore_current_clipboard_on_start": self.ignore_current_var.get(),
            "minimize_to_tray": self.minimize_to_tray_var.get(),
        }
        try:
            self.settings_manager.save(settings)
        except Exception as exc:
            self._log(f"WARNING: Could not save settings: {exc}")


def run_app() -> None:
    root = tk.Tk()
    ClipboardAutoTyperApp(root)
    root.mainloop()


if __name__ == "__main__":
    run_app()
