# Backrest Watcher

Backrest Watcher is a lightweight Windows system tray app that monitors your Backrest log and alerts when new log content contains the keyword `"logger":`.

## Who is this for?
- You are running Backrest and want to know immediately when something goes wrong.
- You do not want to keep checking logs manually.

## Get Started in 2 Minutes
1. Run `BackrestTrayWatcher.exe`.
2. Find the app icon in the system tray (bottom-right corner of the taskbar).
3. Right-click the icon and choose `Set log file path...` (eg. C:\Users\\[YOUR_USERNAME]\AppData\Roaming\backrest\data\processlogs\backrest.log)
4. Select your Backrest log file.
5. Done. The app will monitor in the background.

## Daily Usage
- Double-left-click tray icon: open the log file quickly.
- `Open log file`: open the log to inspect details.
- `Open log folder`: open the folder containing the log file.
- `Set check interval...`: change how often the app checks the log.
- `Acknowledge warning/error`: mark the current alert as handled.
- `Exit`: close the app.

## Tray Status Meaning
- `OK`: no new `"logger":` keyword found.
- `WARNING/ERROR` (blinking): a new `"logger":` keyword was detected.

## What to do when an alert appears
1. Double-left-click the tray icon (or right-click and choose `Open log file`).
2. Check the latest lines containing `"logger":`.
3. Fix the issue in Backrest.
4. Return to the app and click `Acknowledge warning/error` from menu.

## Usage Tips
- Keep the app running during your Windows session so you do not miss alerts.
- If your log file location changes, just set it again with `Set log file path...`.
- For faster alerts, use a lower check interval; for lighter resource usage, use a higher interval.

## Quick FAQ
### 1. I opened the app but no window appears?
That is expected. The app runs in the system tray and has no main window.

### 2. I cannot find the tray icon?
Click the `^` arrow in the tray area to show hidden icons.

### 3. How do I reset the current alert state?
Use `Acknowledge warning/error` from the tray menu.

## GitHub Auto Release
This repo includes GitHub Actions workflow at `.github/workflows/release.yml`.

- Trigger: push a Git tag starting with `v` (example: `v1.0.0`).
- Action: build on `windows-latest` using `build.bat`.
- Output: upload `BackrestTrayWatcher.exe` as artifact and publish it to GitHub Release with the same tag name.

Example:
```bash
git tag v1.0.0
git push origin v1.0.0
```

## Note
The config file `backrest_tray_watcher.ini` is created/updated automatically in the same folder as the executable.
