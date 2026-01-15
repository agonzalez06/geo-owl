#!/usr/bin/env python3
"""
Send ANC Sheet via Email
Generates the ANC sheet and emails it to configured recipients.
Can be run manually or via scheduled job.
"""

import sys
import os
from datetime import datetime, timedelta
from pathlib import Path

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from anc_generator import generate_anc_for_date
from email_utils.email_sender import send_email, load_config, format_template

def send_anc_sheet(
    date: datetime = None,
    recipients: list[str] = None,
    cc: list[str] = None,
    subject: str = None,
    body: str = None
) -> bool:
    """
    Generate and email the ANC sheet.

    Args:
        date: Date for ANC sheet (defaults to tomorrow)
        recipients: Override recipients from config
        cc: Override CC from config
        subject: Override subject template
        body: Override body template

    Returns:
        True if successful, False otherwise
    """

    # Load config
    try:
        config = load_config()
        job_config = config.get('jobs', {}).get('anc_daily', {})
    except FileNotFoundError as e:
        print(f"Error: {e}")
        return False

    # Default to tomorrow's date (ANC sheets are typically for next day)
    if date is None:
        date = datetime.now() + timedelta(days=1)

    date_str = date.strftime('%A, %B %d, %Y')
    date_short = date.strftime('%m/%d/%y')

    # Get recipients
    recipients = recipients or job_config.get('recipients', [])
    cc = cc or job_config.get('cc', [])

    if not recipients:
        print("Error: No recipients configured")
        return False

    # Format subject and body
    subject_template = subject or job_config.get('subject_template', 'ANC Sheet - {date}')
    body_template = body or job_config.get('body_template',
        'Please find attached the ANC sheet for {date}.\n\nThis is an automated message.')

    subject = format_template(subject_template, date=date_str, date_short=date_short)
    body = format_template(body_template, date=date_str, date_short=date_short)

    # Generate ANC sheet
    print(f"Generating ANC sheet for {date_str}...")
    output_dir = Path(__file__).parent.parent

    try:
        # Generate as Word doc (most compatible for email attachment)
        output_path = generate_anc_for_date(date, output_dir=str(output_dir), output_format='docx')
    except Exception as e:
        print(f"Error generating ANC sheet: {e}")
        return False

    if not os.path.exists(output_path):
        print(f"Error: Generated file not found: {output_path}")
        return False

    # Send email
    print(f"Sending to: {', '.join(recipients)}")
    if cc:
        print(f"CC: {', '.join(cc)}")

    success = send_email(
        to=recipients,
        subject=subject,
        body=body,
        attachments=[output_path],
        cc=cc
    )

    if success:
        print("ANC sheet sent successfully!")
    else:
        print("Failed to send ANC sheet")

    return success

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Generate and email ANC sheet")
    parser.add_argument('--date', help='Date for ANC sheet (MM-DD-YYYY), defaults to tomorrow')
    parser.add_argument('--to', nargs='+', help='Override recipients')
    parser.add_argument('--cc', nargs='+', help='Override CC recipients')
    parser.add_argument('--test', action='store_true', help='Send to yourself for testing')

    args = parser.parse_args()

    # Parse date if provided
    target_date = None
    if args.date:
        try:
            target_date = datetime.strptime(args.date, '%m-%d-%Y')
        except ValueError:
            print(f"Invalid date format: {args.date}. Use MM-DD-YYYY")
            exit(1)

    # For test mode, send to sender
    recipients = args.to
    if args.test:
        config = load_config()
        sender = config.get('email', {}).get('sender_email')
        if sender:
            recipients = [sender]
            print(f"Test mode: sending to {sender}")
        else:
            print("Error: No sender_email configured for test mode")
            exit(1)

    success = send_anc_sheet(
        date=target_date,
        recipients=recipients,
        cc=args.cc
    )

    exit(0 if success else 1)
