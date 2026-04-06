#!/usr/bin/env python3
"""
Daemon Setup - US Economic Tracker
-----------------------------------
Automatically configures the tracker to run daily in the background.
Detects your operating system and sets up the appropriate method.

Usage:
    py -3 setup_daemon.py
    py -3 setup_daemon.py --time 08:30
    py -3 setup_daemon.py --remove

Supports: Windows, macOS, Linux
"""
import os, sys, platform, argparse, shutil

PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPT_PATH = os.path.join(PROJECT_DIR, "economic_tracker.py")
SYSTEM = platform.system()


def find_python():
    """Find the Python executable path."""
    if shutil.which("py"):
        return "py -3"
    if shutil.which("python3"):
        return "python3"
    return sys.executable


def setup_windows(time_str):
    """Create a VBS launcher and place it in the Windows Startup folder."""
    vbs_path = os.path.join(PROJECT_DIR, "automation", "start_daemon.vbs")
    startup_dir = os.path.join(
        os.environ.get("APPDATA", ""),
        "Microsoft", "Windows", "Start Menu", "Programs", "Startup"
    )
    startup_vbs = os.path.join(startup_dir, "US Economic Tracker Daemon.vbs")

    vbs_content = (
        'Set WshShell = CreateObject("WScript.Shell")\r\n'
        f'sDir = "{PROJECT_DIR}"\r\n'
        'WshShell.CurrentDirectory = sDir\r\n'
        f'WshShell.Run "py -3 """ & sDir & "\\economic_tracker.py"" --daemon {time_str}", 0, False\r\n'
    )

    with open(vbs_path, "w") as f:
        f.write(vbs_content)
    print(f"  Created: {vbs_path}")

    if os.path.isdir(startup_dir):
        shutil.copy2(vbs_path, startup_vbs)
        print(f"  Installed to Startup folder: {startup_vbs}")
        print(f"\n  The daemon will start automatically on your next login.")
        print(f"  To start it now, double-click: automation/start_daemon.vbs")
    else:
        print(f"  [WARN] Startup folder not found: {startup_dir}")
        print(f"  Manually copy start_daemon.vbs to your Startup folder.")


def remove_windows():
    """Remove the daemon from Windows Startup."""
    startup_vbs = os.path.join(
        os.environ.get("APPDATA", ""),
        "Microsoft", "Windows", "Start Menu", "Programs", "Startup",
        "US Economic Tracker Daemon.vbs"
    )
    local_vbs = os.path.join(PROJECT_DIR, "automation", "start_daemon.vbs")

    removed = False
    if os.path.isfile(startup_vbs):
        os.remove(startup_vbs)
        print(f"  Removed: {startup_vbs}")
        removed = True
    if os.path.isfile(local_vbs):
        os.remove(local_vbs)
        print(f"  Removed: {local_vbs}")
        removed = True
    if not removed:
        print("  No daemon found to remove.")
    else:
        print("  Daemon will no longer start on login.")


def setup_macos(time_str):
    """Create a launchd plist for macOS."""
    python = find_python().split()[0]
    if python == "py":
        python = shutil.which("python3") or "python3"

    plist_name = "com.economic-tracker.daemon"
    plist_dir = os.path.expanduser("~/Library/LaunchAgents")
    plist_path = os.path.join(plist_dir, f"{plist_name}.plist")

    hour, minute = time_str.split(":")

    plist_content = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>{plist_name}</string>
    <key>ProgramArguments</key>
    <array>
        <string>{python}</string>
        <string>{SCRIPT_PATH}</string>
        <string>--daemon</string>
        <string>{time_str}</string>
    </array>
    <key>WorkingDirectory</key>
    <string>{PROJECT_DIR}</string>
    <key>RunAtLoad</key>
    <true/>
    <key>KeepAlive</key>
    <true/>
    <key>StandardOutPath</key>
    <string>{os.path.join(PROJECT_DIR, "daemon.log")}</string>
    <key>StandardErrorPath</key>
    <string>{os.path.join(PROJECT_DIR, "daemon.log")}</string>
</dict>
</plist>
"""

    os.makedirs(plist_dir, exist_ok=True)
    with open(plist_path, "w") as f:
        f.write(plist_content)
    print(f"  Created: {plist_path}")

    os.system(f"launchctl load {plist_path}")
    print(f"  Loaded into launchd. The daemon is now running.")
    print(f"  It will start automatically on login.")
    print(f"  Logs: {os.path.join(PROJECT_DIR, 'daemon.log')}")


def remove_macos():
    """Remove the launchd plist."""
    plist_name = "com.economic-tracker.daemon"
    plist_path = os.path.expanduser(f"~/Library/LaunchAgents/{plist_name}.plist")

    if os.path.isfile(plist_path):
        os.system(f"launchctl unload {plist_path}")
        os.remove(plist_path)
        print(f"  Removed: {plist_path}")
        print("  Daemon stopped and will no longer start on login.")
    else:
        print("  No daemon found to remove.")


def setup_linux(time_str):
    """Create a systemd user service for Linux."""
    python = find_python().split()[0]
    if python == "py":
        python = shutil.which("python3") or "python3"

    service_name = "economic-tracker"
    service_dir = os.path.expanduser("~/.config/systemd/user")
    service_path = os.path.join(service_dir, f"{service_name}.service")

    service_content = f"""[Unit]
Description=US Economic Tracker Daemon
After=network-online.target

[Service]
Type=simple
WorkingDirectory={PROJECT_DIR}
ExecStart={python} {SCRIPT_PATH} --daemon {time_str}
Restart=on-failure
RestartSec=60

[Install]
WantedBy=default.target
"""

    os.makedirs(service_dir, exist_ok=True)
    with open(service_path, "w") as f:
        f.write(service_content)
    print(f"  Created: {service_path}")

    os.system("systemctl --user daemon-reload")
    os.system(f"systemctl --user enable {service_name}")
    os.system(f"systemctl --user start {service_name}")
    print(f"  Service enabled and started.")
    print(f"  It will start automatically on login.")
    print(f"  Check status: systemctl --user status {service_name}")
    print(f"  View logs: journalctl --user -u {service_name} -f")


def remove_linux():
    """Remove the systemd user service."""
    service_name = "economic-tracker"
    service_path = os.path.expanduser(f"~/.config/systemd/user/{service_name}.service")

    if os.path.isfile(service_path):
        os.system(f"systemctl --user stop {service_name}")
        os.system(f"systemctl --user disable {service_name}")
        os.remove(service_path)
        os.system("systemctl --user daemon-reload")
        print(f"  Removed: {service_path}")
        print("  Daemon stopped and will no longer start on login.")
    else:
        print("  No daemon found to remove.")


def main():
    parser = argparse.ArgumentParser(
        description="Set up the US Economic Tracker to run daily in the background"
    )
    parser.add_argument("--time", default="07:00", metavar="HH:MM",
                        help="Time to run daily (default: 07:00)")
    parser.add_argument("--remove", action="store_true",
                        help="Remove the daemon and stop auto-running")
    args = parser.parse_args()

    print()
    print(f"  US Economic Tracker - Daemon Setup")
    print(f"  ===================================")
    print(f"  OS: {SYSTEM} ({platform.platform()})")
    print(f"  Project: {PROJECT_DIR}")
    print()

    if args.remove:
        print("  Removing daemon...")
        if SYSTEM == "Windows":
            remove_windows()
        elif SYSTEM == "Darwin":
            remove_macos()
        elif SYSTEM == "Linux":
            remove_linux()
        else:
            print(f"  Unsupported OS: {SYSTEM}")
    else:
        print(f"  Setting up daily run at {args.time}...")
        if SYSTEM == "Windows":
            setup_windows(args.time)
        elif SYSTEM == "Darwin":
            setup_macos(args.time)
        elif SYSTEM == "Linux":
            setup_linux(args.time)
        else:
            print(f"  Unsupported OS: {SYSTEM}")
            print(f"  You can run manually: python3 economic_tracker.py --daemon {args.time}")

    print()


if __name__ == "__main__":
    main()
