# PrintSwitch

Lightweight Windows system tray app that automatically switches WiFi bands for printing.

**Problem:** Printer only connects on 2.4GHz, but 5GHz is better for everything else (especially video calls).

**Solution:** Stays on 5GHz by default. Detects print jobs, switches to 2.4GHz, waits for the job to finish, switches back. That's it.

## Download

Go to [Releases](../../releases/latest) and download **PrintSwitch.exe**. No login required, no Python needed.

## Auto-start on boot

1. Press `Win+R`
2. Type `shell:startup`
3. Drop `PrintSwitch.exe` into that folder

## Configuration

SSIDs are hardcoded in `print_wifi_switcher.py`:
- 5GHz: `MeadSuper`
- 2.4GHz: `MeadWifi`

Edit and push to rebuild.

## How it works

- Polls Windows Print Spooler every 2 seconds for active jobs
- Switches via `netsh wlan connect`
- 3-minute timeout forces return to 5GHz if a job hangs
- System tray icon shows status
