# Activity Tracker

Windows desktop background activity tracker with NiceGUI dashboard and Google Sheets sync.

## Features
- Active window & Chrome website tracking
- Keyboard / mouse activity intervals (privacy-safe – no keystrokes logged)
- Idle detection via Windows API
- Daily data synced to Google Sheets
- Local JSON backup when offline
- OAuth 2.0 Google authentication

## Setup

### 1. Install dependencies
```bash
pip install -r requirements.txt
```

### 2. Google Cloud credentials
Place your `credentials.json` (OAuth Desktop app) in the project root.  
Enable the **Google Sheets API** in your Google Cloud project.

### 3. Run
```bash
python main.py
```
Open `http://localhost:8580` in your browser (auto-opens).

### 3b. Build EXE (recommended for startup)
```bash
python build_exe.py
```
This creates `ActivityTracker.exe` in the project root.

Runtime files (OAuth token, encrypted config, logs, local JSON data) are
stored in a per-user writable directory by default:

- `%LOCALAPPDATA%\\ActivityTracker`

This avoids permission issues when the app is installed under protected
locations such as `C:\\Program Files`.

Optional override:

- Set `ACTIVITY_TRACKER_DATA_DIR` to a custom writable folder before launch.

- **Login:** `admin` / `admin`
- Connect your Google account
- Enter the Sheet ID and a unique PC Name
- Click **Start Tracking**

### 4. Auto-start on boot (optional)
```bash
python setup_autostart.py
```

Or use the dashboard buttons:
- Click **Start Tracking** to enable auto-start (copies `ActivityTracker.exe` to the Startup folder)
- Click **Stop Tracking** to disable auto-start (removes startup entry)

## Project Structure
```
ActivityTracker/
├── main.py                # Entry point
├── config.py              # Global settings
├── logger_setup.py        # Logging config
├── credentials.json       # Google OAuth client (you provide)
├── requirements.txt
├── build_exe.py          # Build ActivityTracker.exe from main.py
├── setup_autostart.py     # Optional: add to Windows startup
├── auth/
│   └── google_auth.py     # OAuth 2.0 flow
├── tracker/
│   ├── engine.py          # Master tracker orchestrator
│   ├── window_tracker.py  # Active window detection
│   ├── keyboard_tracker.py
│   ├── mouse_tracker.py
│   └── idle_detector.py   # Windows idle time via ctypes
├── sheets/
│   └── sync.py            # Google Sheets read/write
├── app_ui/
│   └── dashboard.py       # NiceGUI login + dashboard pages
└── Runtime data (auto-created in %LOCALAPPDATA%\ActivityTracker)
	├── data/              # Local JSON data (+ .bak files)
	├── logs/              # Application logs
	├── token.json         # Google OAuth token
	├── tracker_config.enc # Encrypted app config
	└── .config_key        # Encryption key
```

## Google Sheet Output

Each PC Name creates a tab. One row per day:

| DATE | WEBSITE | WINDOW | MOUSE | KEYBOARD | TOTAL_WORK_TIME | TOTAL_IDLE_TIME |
|------|---------|--------|-------|----------|-----------------|-----------------|
| 2026-04-06 | google.com - 120 sec | Chrome - 500 sec | 10:00 - 10:10 | 10:02 - 10:08 | 2h 30m 0s | 0h 45m 0s |
