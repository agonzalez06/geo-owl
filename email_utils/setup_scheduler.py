#!/usr/bin/env python3
"""
Setup Script for Email Scheduler
Configures launchd jobs for automated email sending on macOS.
"""

import os
import json
import shutil
import subprocess
from pathlib import Path
from datetime import datetime

PLIST_DIR = Path.home() / "Library" / "LaunchAgents"
PROJECT_DIR = Path(__file__).parent.parent
EMAIL_UTILS_DIR = Path(__file__).parent

def create_launchd_plist(
    job_name: str,
    script_path: str,
    schedule_hour: int = 6,
    schedule_minute: int = 0,
    schedule_days: list = None,
    args: list = None
) -> str:
    """Create a launchd plist for a scheduled job."""

    plist_name = f"com.temple.{job_name}"
    plist_path = PLIST_DIR / f"{plist_name}.plist"

    # Build calendar interval (schedule)
    calendar_interval = {
        'Hour': schedule_hour,
        'Minute': schedule_minute
    }

    # If specific days, create multiple intervals
    if schedule_days:
        day_map = {'Sun': 0, 'Mon': 1, 'Tue': 2, 'Wed': 3, 'Thu': 4, 'Fri': 5, 'Sat': 6}
        calendar_intervals = []
        for day in schedule_days:
            if day in day_map:
                interval = calendar_interval.copy()
                interval['Weekday'] = day_map[day]
                calendar_intervals.append(interval)
        calendar_key = 'StartCalendarInterval'
        calendar_value = calendar_intervals if len(calendar_intervals) > 1 else calendar_intervals[0]
    else:
        calendar_key = 'StartCalendarInterval'
        calendar_value = calendar_interval

    # Build program arguments
    program_args = ['/usr/bin/python3', script_path]
    if args:
        program_args.extend(args)

    # Create plist content
    plist_content = f'''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>{plist_name}</string>

    <key>ProgramArguments</key>
    <array>
'''
    for arg in program_args:
        plist_content += f'        <string>{arg}</string>\n'

    plist_content += f'''    </array>

    <key>{calendar_key}</key>
'''

    if isinstance(calendar_value, list):
        plist_content += '    <array>\n'
        for interval in calendar_value:
            plist_content += '        <dict>\n'
            for key, value in interval.items():
                plist_content += f'            <key>{key}</key>\n'
                plist_content += f'            <integer>{value}</integer>\n'
            plist_content += '        </dict>\n'
        plist_content += '    </array>\n'
    else:
        plist_content += '    <dict>\n'
        for key, value in calendar_value.items():
            plist_content += f'        <key>{key}</key>\n'
            plist_content += f'        <integer>{value}</integer>\n'
        plist_content += '    </dict>\n'

    plist_content += f'''
    <key>WorkingDirectory</key>
    <string>{PROJECT_DIR}</string>

    <key>StandardOutPath</key>
    <string>{PROJECT_DIR}/logs/{job_name}.log</string>

    <key>StandardErrorPath</key>
    <string>{PROJECT_DIR}/logs/{job_name}.error.log</string>

    <key>RunAtLoad</key>
    <false/>
</dict>
</plist>
'''

    return plist_path, plist_content

def setup_anc_job():
    """Set up the daily ANC email job."""

    # Ensure logs directory exists
    logs_dir = PROJECT_DIR / "logs"
    logs_dir.mkdir(exist_ok=True)

    # Load config for schedule
    config_path = EMAIL_UTILS_DIR / "config.json"
    if not config_path.exists():
        print("Error: config.json not found. Run setup first.")
        return False

    with open(config_path) as f:
        config = json.load(f)

    anc_config = config.get('jobs', {}).get('anc_daily', {})
    schedule = anc_config.get('schedule', {})

    # Parse schedule time
    time_str = schedule.get('time', '06:00')
    hour, minute = map(int, time_str.split(':'))
    days = schedule.get('days', ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'])

    # Create plist
    script_path = str(EMAIL_UTILS_DIR / "send_anc.py")
    plist_path, plist_content = create_launchd_plist(
        job_name='anc_daily',
        script_path=script_path,
        schedule_hour=hour,
        schedule_minute=minute,
        schedule_days=days
    )

    # Write plist
    PLIST_DIR.mkdir(parents=True, exist_ok=True)
    with open(plist_path, 'w') as f:
        f.write(plist_content)

    print(f"Created launchd plist: {plist_path}")

    # Load the job
    subprocess.run(['launchctl', 'unload', str(plist_path)], capture_output=True)
    result = subprocess.run(['launchctl', 'load', str(plist_path)], capture_output=True)

    if result.returncode == 0:
        print(f"Scheduled job loaded successfully!")
        print(f"  Schedule: {time_str} on {', '.join(days)}")
        print(f"  Logs: {logs_dir}/")
    else:
        print(f"Warning: Could not load job: {result.stderr.decode()}")

    return True

def unload_job(job_name: str):
    """Unload a scheduled job."""
    plist_path = PLIST_DIR / f"com.temple.{job_name}.plist"
    if plist_path.exists():
        subprocess.run(['launchctl', 'unload', str(plist_path)], capture_output=True)
        print(f"Unloaded job: {job_name}")
    else:
        print(f"Job not found: {job_name}")

def list_jobs():
    """List all Temple scheduled jobs."""
    print("Scheduled Temple jobs:")
    for plist in PLIST_DIR.glob("com.temple.*.plist"):
        job_name = plist.stem.replace("com.temple.", "")
        # Check if loaded
        result = subprocess.run(
            ['launchctl', 'list', plist.stem],
            capture_output=True
        )
        status = "loaded" if result.returncode == 0 else "not loaded"
        print(f"  - {job_name} ({status})")

def run_setup_wizard():
    """Interactive setup wizard."""
    print("=" * 60)
    print("Email Scheduler Setup")
    print("=" * 60)
    print()

    # Check for config file
    config_path = EMAIL_UTILS_DIR / "config.json"
    example_path = EMAIL_UTILS_DIR / "config.example.json"

    if not config_path.exists():
        print("Step 1: Creating configuration file...")
        if example_path.exists():
            # Copy example to config
            shutil.copy(example_path, config_path)
            print(f"  Created: {config_path}")
            print(f"  Please edit this file with your email and recipient settings.")
        else:
            print("  Error: config.example.json not found")
            return

    print()
    print("Step 2: Configure your settings in config.json:")
    print(f"  File: {config_path}")
    print()
    print("  Required settings:")
    print("    - email.sender_email: Your Temple email address")
    print("    - email.sender_name: Your name")
    print("    - jobs.anc_daily.recipients: List of recipient emails")
    print("    - jobs.anc_daily.schedule.time: Time to send (HH:MM)")
    print()

    # Check for keyring
    print("Step 3: Save your email password (for SMTP method):")
    print("  Run this command (optional, only needed for SMTP):")
    print(f"  cd '{EMAIL_UTILS_DIR}'")
    print(f"  python3 email_sender.py save-password --email YOUR_EMAIL --password YOUR_PASSWORD")
    print()

    print("Step 4: Test email sending:")
    print(f"  cd '{EMAIL_UTILS_DIR}'")
    print(f"  python3 send_anc.py --test")
    print()

    print("Step 5: Set up scheduled job:")
    print(f"  python3 setup_scheduler.py install")
    print()
    print("=" * 60)

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Setup email scheduler")
    parser.add_argument('command', nargs='?', default='wizard',
                       choices=['wizard', 'install', 'uninstall', 'list', 'run-now'],
                       help='Command to run')
    parser.add_argument('--job', default='anc_daily', help='Job name (default: anc_daily)')

    args = parser.parse_args()

    if args.command == 'wizard':
        run_setup_wizard()
    elif args.command == 'install':
        setup_anc_job()
    elif args.command == 'uninstall':
        unload_job(args.job)
    elif args.command == 'list':
        list_jobs()
    elif args.command == 'run-now':
        # Run the job immediately
        script_path = EMAIL_UTILS_DIR / "send_anc.py"
        subprocess.run(['/usr/bin/python3', str(script_path)])
