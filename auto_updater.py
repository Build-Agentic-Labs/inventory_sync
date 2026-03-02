"""
Auto-updater for Inventory Sync.
Checks GitHub Releases (public repo) for new versions and applies updates.
"""

import os
import sys
import json
import tempfile
import threading
import subprocess
import time
from pathlib import Path

try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

# ─── CONFIGURATION ──────────────────────────────────────────────────────────
# Set GITHUB_REPO to "your-username/your-repo-name"
# No token needed — public repo.
# ────────────────────────────────────────────────────────────────────────────

GITHUB_REPO = "Build-Agentic-Labs/inventory_sync"
UPDATE_CHECK_DELAY = 10   # Seconds to wait after startup before checking


def is_frozen():
    """Check if running as a compiled exe."""
    return getattr(sys, 'frozen', False)


def get_current_exe():
    """Get path of the currently running executable."""
    if is_frozen():
        return Path(sys.executable)
    return None


def parse_version(version_str):
    """Parse version string like 'v1.2.3' or '1.2.3' into a tuple of ints."""
    v = version_str.strip().lstrip('v')
    try:
        return tuple(int(x) for x in v.split('.'))
    except (ValueError, AttributeError):
        return (0, 0, 0)


def check_for_update(current_version):
    """
    Check GitHub Releases API for a newer version.

    Returns:
        (has_update, latest_version, download_url) or (False, None, None) on error.
    """
    if not HAS_REQUESTS:
        print("Auto-updater: requests library not available")
        return False, None, None

    if not GITHUB_REPO:
        print("Auto-updater: GitHub repo not configured")
        return False, None, None

    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
    headers = {"Accept": "application/vnd.github+json"}

    try:
        resp = requests.get(api_url, headers=headers, timeout=15)

        if resp.status_code == 404:
            print("Auto-updater: No releases found yet")
            return False, None, None

        resp.raise_for_status()
        release = resp.json()

        latest_version = release.get("tag_name", "")
        if not latest_version:
            return False, None, None

        current_tuple = parse_version(current_version)
        latest_tuple = parse_version(latest_version)

        if latest_tuple <= current_tuple:
            print(f"Auto-updater: Up to date (current={current_version}, latest={latest_version})")
            return False, None, None

        # Find the .exe asset — use browser_download_url (public, no auth needed)
        download_url = None
        for asset in release.get("assets", []):
            if asset["name"].lower().endswith(".exe"):
                download_url = asset["browser_download_url"]
                break

        if not download_url:
            print("Auto-updater: No .exe asset found in latest release")
            return False, None, None

        print(f"Auto-updater: Update available! {current_version} -> {latest_version}")
        return True, latest_version, download_url

    except Exception as e:
        print(f"Auto-updater: Error checking for updates: {e}")
        return False, None, None


def download_update(download_url, progress_callback=None):
    """
    Download the updated exe from a GitHub Release asset URL.

    Args:
        download_url: The browser download URL for the asset.
        progress_callback: Optional function(bytes_downloaded, total_bytes).

    Returns:
        Path to the downloaded temp file, or None on failure.
    """
    try:
        resp = requests.get(download_url, stream=True, timeout=300)
        resp.raise_for_status()

        total_size = int(resp.headers.get('content-length', 0))

        # Download to system temp dir (NOT the install dir) to avoid conflicts
        temp_dir = Path(os.environ.get('TEMP', os.environ.get('TMP', '')))
        temp_path = temp_dir / "InventorySync_update.exe"

        downloaded = 0
        with open(temp_path, 'wb') as f:
            for chunk in resp.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
                    downloaded += len(chunk)
                    if progress_callback and total_size:
                        progress_callback(downloaded, total_size)

        print(f"Auto-updater: Downloaded update to {temp_path} ({downloaded} bytes)")
        return temp_path

    except Exception as e:
        print(f"Auto-updater: Error downloading update: {e}")
        return None


def apply_update(temp_exe_path):
    """
    Apply the update by launching a batch script that:
    1. Force-kills all InventorySync processes
    2. Cleans up PyInstaller _MEI temp dirs
    3. Deletes the old exe and renames the new one in place
    4. Restarts the app

    Args:
        temp_exe_path: Path to the downloaded new exe.
    """
    current_exe = get_current_exe()
    if not current_exe:
        print("Auto-updater: Cannot apply update - not running as exe")
        return False

    install_dir = current_exe.parent
    installed_exe = install_dir / "InventorySync.exe"
    staged_exe = install_dir / "InventorySync_new.exe"
    updater_bat = install_dir / "_updater.bat"

    temp_dir = os.environ.get('TEMP', os.environ.get('TMP', ''))
    bat_content = f'''@echo off
title Inventory Sync Updater
echo Updating Inventory Sync...

REM --- Step 1: Stage the update file into install dir ---
echo Staging update...
copy /Y "{temp_exe_path}" "{staged_exe}" >NUL
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Could not stage update file.
    pause
    goto cleanup
)

REM --- Step 2: Force-kill ALL InventorySync processes ---
echo Stopping application...
taskkill /F /IM InventorySync.exe >NUL 2>&1
taskkill /F /IM InventorySync_new.exe >NUL 2>&1

REM Wait until no InventorySync processes remain
:waitloop
timeout /t 2 /nobreak >NUL
tasklist /FI "IMAGENAME eq InventorySync.exe" 2>NUL | find /I "InventorySync.exe" >NUL
if %ERRORLEVEL%==0 (
    taskkill /F /IM InventorySync.exe >NUL 2>&1
    goto waitloop
)

REM --- Step 3: Clean up PyInstaller _MEI temp extraction dirs ---
echo Cleaning up temp files...
for /d %%i in ("{temp_dir}\\_MEI*") do (
    rmdir /s /q "%%i" >NUL 2>&1
)

REM --- Step 4: Wait for Windows to fully release file handles ---
timeout /t 5 /nobreak >NUL

REM --- Step 5: Replace exe using delete + rename (atomic on same drive) ---
echo Applying update...
del /F /Q "{installed_exe}" >NUL 2>&1

REM Retry delete if file is still locked
if exist "{installed_exe}" (
    timeout /t 3 /nobreak >NUL
    del /F /Q "{installed_exe}" >NUL 2>&1
)
if exist "{installed_exe}" (
    timeout /t 5 /nobreak >NUL
    del /F /Q "{installed_exe}" >NUL 2>&1
)
if exist "{installed_exe}" (
    echo ERROR: Could not remove old executable. It may still be in use.
    pause
    goto cleanup
)

rename "{staged_exe}" "InventorySync.exe"
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Could not rename update file.
    pause
    goto cleanup
)

REM --- Step 6: Launch updated app ---
echo Update complete! Starting new version...
timeout /t 2 /nobreak >NUL
"{installed_exe}"

:cleanup
REM Clean up temp download
if exist "{temp_exe_path}" del /Q "{temp_exe_path}"

REM Self-delete
(goto) 2>nul & del "%~f0"
'''

    try:
        with open(updater_bat, 'w') as f:
            f.write(bat_content)

        # Launch the updater script - CREATE_NO_WINDOW only (DETACHED_PROCESS breaks start)
        subprocess.Popen(
            ['cmd', '/c', str(updater_bat)],
            creationflags=subprocess.CREATE_NO_WINDOW,
            cwd=str(install_dir)
        )
        print("Auto-updater: Updater script launched, exiting for update...")
        return True

    except Exception as e:
        print(f"Auto-updater: Error launching updater: {e}")
        return False


def run_update_check(current_version, on_update_available=None):
    """
    Run the update check in a background thread.
    Waits UPDATE_CHECK_DELAY seconds after startup before checking.

    Args:
        current_version: The current app version string (e.g. "1.0.0").
        on_update_available: Callback function(latest_version, download_url)
                            called on the background thread when an update is found.
                            If None, the update is printed to console only.
    """
    def _check():
        time.sleep(UPDATE_CHECK_DELAY)
        has_update, latest_version, download_url = check_for_update(current_version)
        if has_update and on_update_available:
            on_update_available(latest_version, download_url)

    thread = threading.Thread(target=_check, daemon=True)
    thread.start()
    return thread
