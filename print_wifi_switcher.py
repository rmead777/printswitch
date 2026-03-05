"""
PrintSwitch - Automatic WiFi Band Switcher for Printing
========================================================
Sits in the system tray on 5GHz by default.
When a print job is detected, switches to 2.4GHz,
waits for the job to finish, then switches back.

Shows Windows toast notifications for every state change
and changes tray icon color to reflect current status.

Setup:
  1. pip install pystray Pillow wmi pywin32
  2. Run: python print_wifi_switcher.py

To build as .exe:
  pip install pyinstaller
  pyinstaller --onefile --noconsole --name PrintSwitch print_wifi_switcher.py
"""

import subprocess
import threading
import time
import logging
import sys
import os
from datetime import datetime

# ============================================================
# CONFIG
# ============================================================
SSID_5GHZ = "MeadSuper"            # 5GHz network (home base)
SSID_24GHZ = "MeadWifi"            # 2.4GHz network (printer band)
WIFI_INTERFACE = "Wi-Fi"            # Usually "Wi-Fi" on Windows

PRINT_TIMEOUT_SECONDS = 180         # Force return to 5GHz after 3 min
POLL_INTERVAL = 2                   # Check every 2 seconds

LOG_FILE = os.path.join(os.path.expanduser("~"), "PrintSwitch.log")
# ============================================================

# Colors for tray icon states
COLOR_IDLE = "#00AA00"       # Green  = chilling on 5GHz
COLOR_PRINTING = "#FF8800"   # Orange = switched to 2.4GHz for printing
COLOR_ERROR = "#DD0000"      # Red    = something went wrong

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("PrintSwitch")


# ============================================================
# Windows Toast Notifications
# ============================================================
def notify(title, message):
    """Show a Windows toast notification. Falls back silently if unavailable."""
    log.info(f"NOTIFY: {title} - {message}")
    try:
        # Use PowerShell to trigger a native Windows toast notification
        # This works on Windows 10/11 without any extra packages
        ps_script = f"""
        [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
        [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom, ContentType = WindowsRuntime] | Out-Null

        $template = @"
        <toast duration="short">
            <visual>
                <binding template="ToastGeneric">
                    <text>{title}</text>
                    <text>{message}</text>
                </binding>
            </visual>
            <audio silent="true"/>
        </toast>
"@

        $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
        $xml.LoadXml($template)
        $notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("PrintSwitch")
        $toast = New-Object Windows.UI.Notifications.ToastNotification $xml
        $notifier.Show($toast)
        """
        subprocess.Popen(
            ["powershell", "-WindowStyle", "Hidden", "-Command", ps_script],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception as e:
        # Toast failed, no big deal -- we still have the tray icon and log
        log.debug(f"Toast notification failed (non-critical): {e}")


# ============================================================
# WiFi Management
# ============================================================
def get_current_ssid():
    """Get the SSID of the currently connected WiFi network."""
    try:
        result = subprocess.run(
            ["netsh", "wlan", "show", "interfaces"],
            capture_output=True, text=True, timeout=10,
        )
        for line in result.stdout.splitlines():
            line = line.strip()
            if line.startswith("SSID") and not line.startswith("BSSID"):
                return line.split(":", 1)[1].strip()
    except Exception as e:
        log.error(f"Failed to get current SSID: {e}")
    return None


def switch_wifi(target_ssid):
    """Switch to the specified WiFi network."""
    current = get_current_ssid()
    if current == target_ssid:
        log.info(f"Already on {target_ssid}, no switch needed")
        return True

    log.info(f"Switching from '{current}' to '{target_ssid}'...")
    try:
        result = subprocess.run(
            [
                "netsh", "wlan", "connect",
                f"name={target_ssid}",
                f"interface={WIFI_INTERFACE}",
            ],
            capture_output=True, text=True, timeout=15,
        )
        if "successfully" in result.stdout.lower() or result.returncode == 0:
            time.sleep(3)
            new_ssid = get_current_ssid()
            if new_ssid == target_ssid:
                log.info(f"Successfully connected to {target_ssid}")
                return True
            else:
                log.warning(f"Switch command ran but connected to '{new_ssid}' instead")
                return False
        else:
            log.error(f"Switch failed: {result.stdout} {result.stderr}")
            return False
    except Exception as e:
        log.error(f"Switch exception: {e}")
        return False


# ============================================================
# Print Job Detection
# ============================================================
def get_active_print_jobs():
    """Check for active print jobs using WMI."""
    try:
        import wmi
        c = wmi.WMI()
        jobs = c.Win32_PrintJob()
        return jobs
    except Exception as e:
        log.error(f"WMI print job query failed: {e}")
        return []


def get_active_print_jobs_fallback():
    """Fallback: check print queue via PowerShell."""
    try:
        result = subprocess.run(
            [
                "powershell", "-Command",
                "Get-PrintJob -PrinterName * -ErrorAction SilentlyContinue | "
                "Where-Object { $_.JobStatus -ne 'Complete' } | "
                "Measure-Object | Select-Object -ExpandProperty Count",
            ],
            capture_output=True, text=True, timeout=10,
        )
        count = int(result.stdout.strip() or "0")
        return count > 0
    except Exception:
        return False


# ============================================================
# Main Controller
# ============================================================
class PrintSwitcher:
    """State machine: IDLE on 5GHz, PRINTING on 2.4GHz."""

    def __init__(self):
        self.running = True
        self.state = "IDLE"
        self.switch_time = None
        self.use_wmi = True
        self.icon_ref = None
        self.make_icon_fn = None
        self.jobs_detected = 0

    def set_icon(self, icon, make_icon_fn):
        """Store references so we can update the tray icon."""
        self.icon_ref = icon
        self.make_icon_fn = make_icon_fn

    def update_tray(self, color, tooltip):
        """Update the tray icon color and tooltip text."""
        if self.icon_ref and self.make_icon_fn:
            try:
                self.icon_ref.icon = self.make_icon_fn(color)
                self.icon_ref.title = tooltip
            except Exception:
                pass

    def has_active_jobs(self):
        """Check if there are active print jobs."""
        if self.use_wmi:
            try:
                jobs = get_active_print_jobs()
                return len(jobs) > 0
            except Exception:
                log.info("WMI unavailable, falling back to PowerShell")
                self.use_wmi = False
                return get_active_print_jobs_fallback()
        else:
            return get_active_print_jobs_fallback()

    def run(self):
        """Main loop."""
        log.info("=" * 50)
        log.info("PrintSwitch started")
        log.info(f"5GHz SSID:   {SSID_5GHZ}")
        log.info(f"2.4GHz SSID: {SSID_24GHZ}")
        log.info(f"Timeout:     {PRINT_TIMEOUT_SECONDS}s")
        log.info(f"Log file:    {LOG_FILE}")
        log.info("=" * 50)

        # Ensure we start on 5GHz
        current = get_current_ssid()
        if current != SSID_5GHZ:
            log.info(f"Currently on '{current}', switching to 5GHz home base")
            switch_wifi(SSID_5GHZ)

        self.update_tray(COLOR_IDLE, f"PrintSwitch - Idle on {SSID_5GHZ}")
        notify(
            "PrintSwitch is running",
            f"Connected to {SSID_5GHZ}. Will auto-switch to {SSID_24GHZ} when you print.",
        )

        while self.running:
            try:
                self._tick()
            except Exception as e:
                log.error(f"Error in main loop: {e}")
                self.update_tray(COLOR_ERROR, f"PrintSwitch - Error: {e}")
            time.sleep(POLL_INTERVAL)

    def _tick(self):
        """Single iteration of the monitor loop."""
        has_jobs = self.has_active_jobs()

        if self.state == "IDLE":
            if has_jobs:
                self.jobs_detected += 1
                log.info(f">> Print job #{self.jobs_detected} detected! Switching to 2.4GHz...")
                self.state = "PRINTING"
                self.switch_time = datetime.now()

                self.update_tray(
                    COLOR_PRINTING,
                    f"PrintSwitch - PRINTING (switching to {SSID_24GHZ})",
                )
                notify(
                    "Switching to 2.4GHz for printing",
                    f"Print job detected. Connecting to {SSID_24GHZ}...",
                )
                switch_wifi(SSID_24GHZ)

        elif self.state == "PRINTING":
            elapsed = (datetime.now() - self.switch_time).total_seconds()

            if not has_jobs:
                log.info(f"<< Print job complete ({elapsed:.0f}s). Switching back to 5GHz...")
                self.state = "IDLE"
                self.switch_time = None

                self.update_tray(COLOR_IDLE, f"PrintSwitch - Idle on {SSID_5GHZ}")
                notify(
                    "Print done! Back on 5GHz",
                    f"Reconnected to {SSID_5GHZ}. Job took {elapsed:.0f} seconds.",
                )
                switch_wifi(SSID_5GHZ)

            elif elapsed > PRINT_TIMEOUT_SECONDS:
                log.warning(f"!! Print timeout after {PRINT_TIMEOUT_SECONDS}s. Forcing 5GHz.")
                self.state = "IDLE"
                self.switch_time = None

                self.update_tray(COLOR_IDLE, f"PrintSwitch - Idle on {SSID_5GHZ} (timed out)")
                notify(
                    "Print timed out -- back on 5GHz",
                    f"Print job exceeded {PRINT_TIMEOUT_SECONDS}s. Returned to {SSID_5GHZ}.",
                )
                switch_wifi(SSID_5GHZ)

    def stop(self):
        """Graceful shutdown."""
        self.running = False
        log.info("PrintSwitch stopping... returning to 5GHz")
        notify("PrintSwitch shutting down", f"Returning to {SSID_5GHZ}. Goodbye!")
        switch_wifi(SSID_5GHZ)


# ============================================================
# System Tray Icon
# ============================================================
def make_icon(color="#00AA00"):
    """Generate a tray icon: colored circle with 'P'."""
    from PIL import Image, ImageDraw, ImageFont

    size = 64
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.ellipse([4, 4, 60, 60], fill=color, outline="white", width=2)
    try:
        font = ImageFont.truetype("arial.ttf", 30)
    except Exception:
        font = ImageFont.load_default()
    draw.text((size // 2, size // 2), "P", fill="white", font=font, anchor="mm")
    return img


def create_tray_icon(switcher):
    """Create system tray icon with status, log viewer, and quit."""
    try:
        import pystray
    except ImportError:
        log.warning("pystray/Pillow not installed. Running without tray icon.")
        return None

    def on_quit(icon, item):
        switcher.stop()
        icon.stop()

    def on_open_log(icon, item):
        """Open the log file in Notepad."""
        try:
            os.startfile(LOG_FILE)
        except Exception:
            subprocess.Popen(["notepad.exe", LOG_FILE])

    def on_test_notify(icon, item):
        """Send a test notification so user knows it's working."""
        current = get_current_ssid()
        notify(
            "PrintSwitch is alive!",
            f"Currently on {current}. Jobs handled: {switcher.jobs_detected}.",
        )

    def get_status(item):
        current = get_current_ssid()
        if switcher.state == "PRINTING":
            elapsed = (datetime.now() - switcher.switch_time).total_seconds()
            return f"PRINTING on {current} ({elapsed:.0f}s)"
        return f"Idle on {current}"

    def get_job_count(item):
        return f"Jobs handled: {switcher.jobs_detected}"

    menu = pystray.Menu(
        pystray.MenuItem("PrintSwitch", None, enabled=False),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem(get_status, None, enabled=False),
        pystray.MenuItem(get_job_count, None, enabled=False),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Test notification", on_test_notify),
        pystray.MenuItem("Open log file", on_open_log),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Quit", on_quit),
    )

    icon = pystray.Icon(
        "PrintSwitch",
        make_icon(COLOR_IDLE),
        f"PrintSwitch - Idle on {SSID_5GHZ}",
        menu,
    )

    switcher.set_icon(icon, make_icon)
    return icon


# ============================================================
# Entry Point
# ============================================================
def main():
    switcher = PrintSwitcher()
    icon = create_tray_icon(switcher)

    if icon:
        monitor_thread = threading.Thread(target=switcher.run, daemon=True)
        monitor_thread.start()
        log.info("System tray icon active. Right-click for options.")
        icon.run()
    else:
        log.info("Running in console mode. Press Ctrl+C to stop.")
        try:
            switcher.run()
        except KeyboardInterrupt:
            switcher.stop()

    log.info("PrintSwitch exited.")


if __name__ == "__main__":
    main()
