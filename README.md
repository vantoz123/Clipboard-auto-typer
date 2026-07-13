# Clipboard Auto Typer v2.1

Clipboard Auto Typer is a Python-based Windows desktop application that reads text copied to the clipboard and inserts it into the active Microsoft Word document.

This v2.1 upgrade adds shortcut keys, system-tray minimization, saved settings, better queue handling, diagnostics, and beginner-friendly batch files.

## Important Safety and Privacy Notes

- The app runs locally and offline.
- Clipboard text is not uploaded, transmitted, or saved to disk.
- Only non-sensitive preferences such as speed, delay, insertion mode, and tray behavior are saved.
- The app is designed for legitimate automation, accessibility, testing, and productivity workflows.
- Use the app only where you have permission to use typing automation.

## Main Features

- Continuous clipboard monitoring.
- Manual `Type Current Clipboard` button.
- Microsoft Word readiness detection.
- Types only into Microsoft Word when Word is active.
- Preserves letters, spaces, line breaks, punctuation, symbols, Unicode characters, and emojis where Word supports them.
- Two insertion modes:
  - **Safe Insert**: direct Word COM insertion for best exactness and Unicode support.
  - **Human Type**: character-by-character typing using the selected CPS speed.
- Slow, Normal, Fast, and Custom CPS speed modes.
- Delay before typing so the user can click the Word destination.
- Pause/resume and emergency stop.
- Queue latest clipboard item while typing.
- Minimize to Windows system tray instead of staying on the taskbar.
- Tray menu with Show, Start, Stop, Pause/Resume, Type Current Clipboard, and Exit.
- Local and global shortcut keys.
- Settings saved between sessions.
- Beginner-friendly install, run, build, and diagnostics batch files.

## Project Structure

```text
Clipboard_Auto_Typer_v2/
├── main.py
├── clipboard_manager.py
├── typer_engine.py
├── word_detector.py
├── gui.py
├── config.py
├── settings_manager.py
├── diagnostics.py
├── requirements.txt
├── install_dependencies.bat
├── run_app.bat
├── run_diagnostics.bat
├── build_exe.bat
├── README.md
├── USER_MANUAL.pdf
└── screenshots/
```

## Requirements

- Windows 10 or Windows 11.
- Python 3.10 or newer recommended.
- Microsoft Word installed.
- A Word document open when typing begins.

## Beginner Installation

1. Extract the ZIP file.
2. Open the extracted `Clipboard_Auto_Typer_v2` folder.
3. Double-click `install_dependencies.bat`.
4. Wait until it says installation is complete.
5. Double-click `run_app.bat`.

## Manual Installation from VS Code or Terminal

Open a terminal in the project folder and run:

```powershell
py -m venv .venv
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
python main.py
```

If PowerShell blocks activation, use:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.venv\Scripts\Activate.ps1
```

Or run without activating:

```powershell
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe main.py
```

## Shortcut Keys

| Action | Shortcut |
|---|---|
| Start monitoring | Ctrl+Shift+S |
| Stop monitoring and stop typing | Ctrl+Shift+X |
| Pause/resume typing | Ctrl+Shift+P |
| Type current clipboard | Ctrl+Shift+T |
| Show/hide app window | Ctrl+Shift+H |
| Emergency stop | Ctrl+Alt+Esc |
| Emergency stop while app window is focused | Esc |

Global shortcuts use the `keyboard` library. On some locked-down Windows systems, global shortcuts may require administrator permission or may be blocked by security policy. The buttons and local shortcuts still work.

## Minimize to System Tray

When **Minimize to system tray instead of taskbar** is enabled, minimizing the window hides it from the taskbar and keeps a tray icon near the Windows clock.

Use one of these methods to restore it:

- Click the tray icon menu and choose **Show Clipboard Auto Typer**.
- Press `Ctrl+Shift+H`.

Use the tray menu or the app's Exit button to close the app completely.

## Using the App

1. Open Microsoft Word.
2. Open or create a document.
3. Click inside the document where the text should go.
4. Open Clipboard Auto Typer.
5. Choose speed, delay, and insertion mode.
6. Click **Start Monitoring**.
7. Copy text from any source.
8. During the delay, click back into Word if needed.
9. The app inserts the copied text into Word.

For the most accurate Unicode and line-break preservation, use **Safe Insert**. For visible character-by-character typing, use **Human Type**.

## Test Examples

Copy each example, let the app type it into Word, and compare the Word output with the original text.

### Example 1

```text
Hello World!
This is a clipboard typing test.
```

### Example 2

```text
Numbers: 123456789
Symbols: @#$%^&*()
Special characters: ä ö ü ñ
```

### Example 3

```text
Paragraph one.

Paragraph two with spacing preserved.
```

## Building a Windows EXE

Install dependencies first, then double-click:

```text
build_exe.bat
```

The executable will be created in the `dist` folder.

Manual command:

```powershell
.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean --windowed --onefile --name "Clipboard Auto Typer" main.py
```

## Troubleshooting

### PowerShell says running scripts is disabled

Run this in the current PowerShell window:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

Then activate again:

```powershell
.venv\Scripts\Activate.ps1
```

### Word is not active

Open Word, open a document, and click where the text should appear. The app only types into Word when Word is active.

### Tray icon does not appear

Run:

```powershell
pip install pystray Pillow
```

Or run `install_dependencies.bat` again.

### Global shortcuts do not work

Some Windows policies block global keyboard hooks. Use the app buttons or run VS Code/terminal as administrator only if your environment permits it.

### Text changes unexpectedly in Word

Use **Safe Insert** mode. It is designed to preserve copied text more exactly than simulated keystrokes. Also check Word AutoCorrect settings.

### Nothing happens when copying text

Check that monitoring says **Monitoring**, Microsoft Word is running, and Word is active. You can also use **Type Current Clipboard**.

## Diagnostics

Double-click:

```text
run_diagnostics.bat
```

This checks Python, required modules, and Microsoft Word readiness.
