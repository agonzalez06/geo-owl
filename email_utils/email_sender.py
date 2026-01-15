#!/usr/bin/env python3
"""
Email Sender Utility
Generalizable email sending module for scheduled and on-demand emails.
Supports Microsoft 365 via Outlook app or SMTP.
"""

import os
import json
import subprocess
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
from pathlib import Path
from typing import Optional
import keyring

# Service name for keyring (secure credential storage)
KEYRING_SERVICE = "temple_email_sender"

def get_config_path() -> Path:
    """Get path to config file."""
    return Path(__file__).parent / "config.json"

def load_config() -> dict:
    """Load configuration from JSON file."""
    config_path = get_config_path()
    if not config_path.exists():
        raise FileNotFoundError(
            f"Config file not found: {config_path}\n"
            f"Copy config.example.json to config.json and fill in your settings."
        )
    with open(config_path) as f:
        return json.load(f)

def save_password(email: str, password: str) -> None:
    """Securely save email password to macOS Keychain."""
    keyring.set_password(KEYRING_SERVICE, email, password)
    print(f"Password saved securely for {email}")

def get_password(email: str) -> Optional[str]:
    """Retrieve email password from macOS Keychain."""
    return keyring.get_password(KEYRING_SERVICE, email)

def send_via_outlook(
    to: list[str],
    subject: str,
    body: str,
    attachments: list[str] = None,
    cc: list[str] = None
) -> bool:
    """Send email using Microsoft Outlook via AppleScript."""

    # Escape special characters for AppleScript
    subject_escaped = subject.replace('\\', '\\\\').replace('"', '\\"')
    body_escaped = body.replace('\\', '\\\\').replace('"', '\\"')

    # Build recipient commands
    recipient_commands = ""
    for addr in to:
        recipient_commands += f'make new to recipient at newMessage with properties {{email address:{{address:"{addr}"}}}}\n'

    # Build CC commands
    if cc:
        for addr in cc:
            recipient_commands += f'make new cc recipient at newMessage with properties {{email address:{{address:"{addr}"}}}}\n'

    # Build attachment commands
    attachment_commands = ""
    if attachments:
        for att_path in attachments:
            abs_path = os.path.abspath(att_path)
            attachment_commands += f'make new attachment at newMessage with properties {{file:POSIX file "{abs_path}"}}\n'

    # AppleScript to send email via Outlook
    script = f'''
tell application "Microsoft Outlook"
    set newMessage to make new outgoing message with properties {{subject:"{subject_escaped}", plain text content:"{body_escaped}"}}
    {recipient_commands}
    {attachment_commands}
    send newMessage
end tell
'''

    try:
        result = subprocess.run(
            ['osascript', '-e', script],
            capture_output=True,
            text=True,
            timeout=60
        )
        if result.returncode == 0:
            print(f"Email sent successfully via Outlook to {', '.join(to)}")
            return True
        else:
            print(f"Outlook error: {result.stderr}")
            return False
    except subprocess.TimeoutExpired:
        print("Outlook email timed out")
        return False
    except Exception as e:
        print(f"Error sending via Outlook: {e}")
        return False

def send_via_smtp(
    to: list[str],
    subject: str,
    body: str,
    attachments: list[str] = None,
    cc: list[str] = None,
    sender_email: str = None,
    sender_name: str = None,
    smtp_server: str = "smtp.office365.com",
    smtp_port: int = 587
) -> bool:
    """Send email using SMTP (Microsoft 365)."""

    config = load_config()
    email_config = config.get('email', {})

    sender_email = sender_email or email_config.get('sender_email')
    sender_name = sender_name or email_config.get('sender_name', '')
    smtp_server = smtp_server or email_config.get('smtp_server', 'smtp.office365.com')
    smtp_port = smtp_port or email_config.get('smtp_port', 587)

    if not sender_email:
        print("Error: sender_email not configured")
        return False

    # Get password from keychain
    password = get_password(sender_email)
    if not password:
        print(f"Error: No password found for {sender_email}")
        print(f"Run: python3 -c \"from email_sender import save_password; save_password('{sender_email}', 'YOUR_PASSWORD')\"")
        return False

    # Create message
    msg = MIMEMultipart()
    msg['From'] = f"{sender_name} <{sender_email}>" if sender_name else sender_email
    msg['To'] = ', '.join(to)
    if cc:
        msg['Cc'] = ', '.join(cc)
    msg['Subject'] = subject

    # Add body
    msg.attach(MIMEText(body, 'plain'))

    # Add attachments
    if attachments:
        for file_path in attachments:
            if os.path.exists(file_path):
                with open(file_path, 'rb') as f:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                filename = os.path.basename(file_path)
                part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
                msg.attach(part)
            else:
                print(f"Warning: Attachment not found: {file_path}")

    # Send email
    try:
        context = ssl.create_default_context()
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls(context=context)
            server.login(sender_email, password)
            all_recipients = to + (cc or [])
            server.sendmail(sender_email, all_recipients, msg.as_string())
        print(f"Email sent successfully via SMTP to {', '.join(to)}")
        return True
    except smtplib.SMTPAuthenticationError:
        print("SMTP Authentication failed. Check your password or app password.")
        print("If you have MFA enabled, you may need to create an app password.")
        return False
    except Exception as e:
        print(f"Error sending via SMTP: {e}")
        return False

def send_email(
    to: list[str],
    subject: str,
    body: str,
    attachments: list[str] = None,
    cc: list[str] = None,
    method: str = None
) -> bool:
    """
    Send an email using the configured method.

    Args:
        to: List of recipient email addresses
        subject: Email subject
        body: Email body text
        attachments: List of file paths to attach
        cc: List of CC email addresses
        method: 'outlook' or 'smtp' (defaults to config setting)

    Returns:
        True if email sent successfully, False otherwise
    """
    config = load_config()
    method = method or config.get('email', {}).get('method', 'outlook')

    if method == 'outlook':
        return send_via_outlook(to, subject, body, attachments, cc)
    elif method == 'smtp':
        return send_via_smtp(to, subject, body, attachments, cc)
    else:
        print(f"Unknown email method: {method}")
        return False

def format_template(template: str, **kwargs) -> str:
    """Format a template string with provided variables."""
    return template.format(**kwargs)

# CLI interface
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Send emails via Microsoft 365")
    subparsers = parser.add_subparsers(dest='command', help='Commands')

    # Send command
    send_parser = subparsers.add_parser('send', help='Send an email')
    send_parser.add_argument('--to', required=True, nargs='+', help='Recipient email(s)')
    send_parser.add_argument('--subject', required=True, help='Email subject')
    send_parser.add_argument('--body', required=True, help='Email body')
    send_parser.add_argument('--attach', nargs='+', help='File(s) to attach')
    send_parser.add_argument('--cc', nargs='+', help='CC recipient(s)')
    send_parser.add_argument('--method', choices=['outlook', 'smtp'], help='Send method')

    # Save password command
    pw_parser = subparsers.add_parser('save-password', help='Save email password securely')
    pw_parser.add_argument('--email', required=True, help='Email address')
    pw_parser.add_argument('--password', required=True, help='Password or app password')

    # Test command
    test_parser = subparsers.add_parser('test', help='Send a test email')
    test_parser.add_argument('--to', required=True, help='Recipient email')

    args = parser.parse_args()

    if args.command == 'send':
        success = send_email(
            to=args.to,
            subject=args.subject,
            body=args.body,
            attachments=args.attach,
            cc=args.cc,
            method=args.method
        )
        exit(0 if success else 1)

    elif args.command == 'save-password':
        save_password(args.email, args.password)

    elif args.command == 'test':
        success = send_email(
            to=[args.to],
            subject="Test Email from Email Sender Utility",
            body="This is a test email sent from the email_sender.py utility.\n\nIf you received this, the email system is working correctly!"
        )
        exit(0 if success else 1)

    else:
        parser.print_help()
