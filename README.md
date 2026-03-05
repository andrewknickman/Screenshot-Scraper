# SCREENSHOT SCRAPER BATCH RUNNER

A PySide6 desktop app for working through batches of screenshot jobs from spreadsheets or manual entry, opening links in a controlled browser session, and saving consistently named captures.

## What it does

- Loads capture rows from Excel workbooks
- Drives a browser through Playwright launch mode or CDP attach mode
- Supports viewport captures and optional screen-region captures
- Keeps a local config for repeatable capture settings
- Exports a CSV report of capture activity
- Helps with domain-by-domain warmups and relaunch workflows when sites are finicky

## Stack

- Python
- PySide6
- Playwright
- openpyxl
- mss
- Pillow

## Project files

- `app.py` - main desktop application
- `requirements.txt` - Python dependencies
- `.gitignore` - ignores local cache, browser profile, screenshots, and config state

## Requirements

- Python 3.10+
- Chromium installed through Playwright for launch mode
- A Chromium-based browser with remote debugging enabled if using CDP attach mode

## Install

### Windows PowerShell

```powershell
cd <REPO_FOLDER>
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
playwright install chromium
```

## Run

```powershell
python app.py
```

## Basic workflow

1. Launch the app.
2. Choose an output folder if you do not want to use the default `screenshots` directory.
3. Import your workbook or enter jobs directly.
4. Open each row in the browser.
5. Capture the screenshot once the page is ready.
6. Export the CSV report when the batch is done.

## Notes

- Local settings are written to `config.json` next to the app.
- Persistent browser data is stored in `profile/` when that mode is enabled.
- Saved captures go to `screenshots/` by default.
- If you change browser channel, executable path, viewport, HTTP/2, or QUIC settings, relaunch the browser from the app.

## GitHub upload checklist

1. Create a new GitHub repository.
2. Upload the contents of this folder to the repo root.
3. Commit `app.py`, `requirements.txt`, `README.md`, and `.gitignore`.
4. Do not commit `config.json`, `profile/`, `screenshots/`, or `__pycache__/`.
