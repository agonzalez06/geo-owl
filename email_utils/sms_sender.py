#!/usr/bin/env python3
"""
SMS Sender Utility
Send text messages via email-to-SMS gateways or Twilio.

Supports major US carriers via email gateways (no API key needed):
- AT&T, Verizon, T-Mobile, Sprint, etc.

For more reliable delivery, can use Twilio (requires account).
"""

import os
import json
import subprocess
from pathlib import Path
from typing import Optional

# =============================================================================
# iMESSAGE SEND (bypasses email entirely)
# =============================================================================

def send_via_imessage(phone: str, message: str) -> bool:
    """
    Send SMS/iMessage via macOS Messages app.
    Requires iPhone linked to Mac for SMS, or sends as iMessage if recipient has it.
    """
    # Format phone number
    clean_phone = format_phone(phone)
    if len(clean_phone) == 10:
        # Format as +1XXXXXXXXXX for Messages
        clean_phone = f"+1{clean_phone}"

    # Escape for AppleScript
    message_escaped = message.replace('\\', '\\\\').replace('"', '\\"')

    script = f'''
tell application "Messages"
    set targetService to 1st account whose service type = SMS
    set targetBuddy to participant "{clean_phone}" of targetService
    send "{message_escaped}" to targetBuddy
end tell
'''

    try:
        result = subprocess.run(
            ['osascript', '-e', script],
            capture_output=True,
            text=True,
            timeout=30
        )
        if result.returncode == 0:
            return True
        else:
            # Try alternate approach - direct send
            script_alt = f'''
tell application "Messages"
    send "{message_escaped}" to buddy "{clean_phone}" of (service 1 whose service type is SMS)
end tell
'''
            result2 = subprocess.run(
                ['osascript', '-e', script_alt],
                capture_output=True,
                text=True,
                timeout=30
            )
            return result2.returncode == 0
    except Exception as e:
        print(f"Error sending via iMessage: {e}")
        return False

# =============================================================================
# CARRIER EMAIL-TO-SMS GATEWAYS
# =============================================================================
# Format: phone@gateway sends SMS to that phone number

CARRIER_GATEWAYS = {
    # Major US Carriers
    'att': 'txt.att.net',
    'verizon': 'vtext.com',
    'tmobile': 'tmomail.net',
    't-mobile': 'tmomail.net',
    'sprint': 'messaging.sprintpcs.com',
    'boost': 'sms.myboostmobile.com',
    'cricket': 'sms.cricketwireless.net',
    'metro': 'mymetropcs.com',
    'metropcs': 'mymetropcs.com',
    'uscellular': 'email.uscc.net',
    'virgin': 'vmobl.com',
    'googlefi': 'msg.fi.google.com',
    'fi': 'msg.fi.google.com',

    # MMS gateways (for longer messages/pictures)
    'att_mms': 'mms.att.net',
    'verizon_mms': 'vzwpix.com',
    'tmobile_mms': 'tmomail.net',
    'sprint_mms': 'pm.sprint.com',
}

def get_sms_config_path() -> Path:
    """Get path to SMS config file."""
    return Path(__file__).parent / "sms_config.json"

def load_sms_config() -> dict:
    """Load SMS configuration."""
    config_path = get_sms_config_path()
    if config_path.exists():
        with open(config_path) as f:
            return json.load(f)
    return {}

def save_sms_config(config: dict) -> None:
    """Save SMS configuration."""
    config_path = get_sms_config_path()
    with open(config_path, 'w') as f:
        json.dump(config, f, indent=2)
    print(f"SMS config saved to {config_path}")

def format_phone(phone: str) -> str:
    """Clean phone number to digits only."""
    return ''.join(c for c in phone if c.isdigit())

def get_sms_email(phone: str, carrier: str) -> str:
    """
    Get the email address for SMS gateway.

    Args:
        phone: Phone number (any format)
        carrier: Carrier name (e.g., 'verizon', 'att', 'tmobile')

    Returns:
        Email address for SMS gateway
    """
    clean_phone = format_phone(phone)

    # Ensure 10 digits (US)
    if len(clean_phone) == 11 and clean_phone.startswith('1'):
        clean_phone = clean_phone[1:]

    if len(clean_phone) != 10:
        raise ValueError(f"Invalid phone number: {phone} (need 10 digits)")

    carrier_lower = carrier.lower().replace(' ', '').replace('-', '')
    gateway = CARRIER_GATEWAYS.get(carrier_lower)

    if not gateway:
        available = ', '.join(sorted(set(CARRIER_GATEWAYS.keys())))
        raise ValueError(f"Unknown carrier: {carrier}. Available: {available}")

    return f"{clean_phone}@{gateway}"

def send_text(phone: str, message: str) -> bool:
    """
    Send a text message to any phone number (carrier-agnostic).

    Uses iMessage/SMS via macOS Messages app - works for any carrier.
    Requires iPhone linked to Mac for SMS to non-iMessage users.

    Args:
        phone: Recipient phone number (any format)
        message: Message to send

    Returns:
        True if sent successfully
    """
    clean_phone = format_phone(phone)

    if len(clean_phone) == 11 and clean_phone.startswith('1'):
        clean_phone = clean_phone[1:]

    if len(clean_phone) != 10:
        print(f"Error: Invalid phone number: {phone}")
        return False

    print(f"Sending text to {clean_phone}...")
    success = send_via_imessage(clean_phone, message)

    if success:
        print(f"Text sent successfully to {clean_phone}")
    else:
        print(f"Failed to send text to {clean_phone}")

    return success


def send_sms_via_gateway(
    phone: str,
    message: str,
    carrier: str = None,
    method: str = 'outlook'
) -> bool:
    """
    Send SMS via email-to-SMS gateway (legacy - requires carrier).

    For carrier-agnostic sending, use send_text() instead.
    """
    # Just use iMessage now - carrier not needed
    return send_text(phone, message)

def send_sms(
    phone: str = None,
    message: str = None,
    carrier: str = None  # Kept for backwards compatibility, but ignored
) -> bool:
    """
    Send SMS using configured settings (carrier-agnostic via iMessage).

    Args:
        phone: Recipient phone (uses config default if not specified)
        message: Message to send
        carrier: IGNORED - no longer needed with iMessage

    Returns:
        True if sent successfully
    """
    sms_config = load_sms_config()

    phone = phone or sms_config.get('default_phone')

    if not phone:
        print("Error: No phone number specified")
        return False

    if not message:
        print("Error: No message specified")
        return False

    return send_text(phone, message)

def setup_sms():
    """Interactive setup for SMS configuration."""
    print("=" * 50)
    print("SMS Configuration Setup")
    print("=" * 50)
    print()

    # Show available carriers
    print("Available carriers:")
    carriers = sorted(set(CARRIER_GATEWAYS.keys()))
    for i, carrier in enumerate(carriers, 1):
        print(f"  {carrier}")
    print()

    # Get phone number
    phone = input("Enter your phone number (10 digits): ").strip()
    phone = format_phone(phone)
    if len(phone) != 10:
        print("Error: Please enter a 10-digit phone number")
        return

    # Get carrier
    carrier = input("Enter your carrier (e.g., verizon, att, tmobile): ").strip().lower()
    if carrier not in CARRIER_GATEWAYS:
        print(f"Warning: Unknown carrier '{carrier}'")
        confirm = input("Continue anyway? (y/n): ").strip().lower()
        if confirm != 'y':
            return

    # Test message?
    test = input("Send a test message? (y/n): ").strip().lower()

    # Save config
    config = {
        'default_phone': phone,
        'default_carrier': carrier,
        'method': 'outlook'
    }
    save_sms_config(config)

    if test == 'y':
        print()
        print("Sending test message...")
        success = send_sms(
            phone=phone,
            message="Test from ANC Generator SMS system. If you received this, SMS is working!",
            carrier=carrier
        )
        if success:
            print("Check your phone for the test message!")
        else:
            print("Test failed. Check your carrier setting.")

# =============================================================================
# CLI
# =============================================================================

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Send text messages via iMessage/SMS")
    subparsers = parser.add_subparsers(dest='command', help='Commands')

    # Setup command
    setup_parser = subparsers.add_parser('setup', help='Configure default phone number')

    # Send command
    send_parser = subparsers.add_parser('send', help='Send a text message')
    send_parser.add_argument('--phone', '-p', help='Phone number (uses default if not specified)')
    send_parser.add_argument('--message', '-m', required=True, help='Message to send')

    # Test command
    test_parser = subparsers.add_parser('test', help='Send a test message')
    test_parser.add_argument('--phone', '-p', help='Phone number (uses default if not specified)')

    args = parser.parse_args()

    if args.command == 'setup':
        print("=" * 50)
        print("SMS Setup (via iMessage - works with any carrier)")
        print("=" * 50)
        print()
        phone = input("Enter default phone number (10 digits): ").strip()
        phone = format_phone(phone)
        if len(phone) != 10:
            print("Error: Please enter a 10-digit phone number")
            exit(1)
        config = {'default_phone': phone, 'use_imessage': True}
        save_sms_config(config)

        test = input("Send a test message? (y/n): ").strip().lower()
        if test == 'y':
            send_text(phone, "Test from ANC system. Texts are working!")

    elif args.command == 'send':
        success = send_sms(phone=args.phone, message=args.message)
        exit(0 if success else 1)

    elif args.command == 'test':
        success = send_sms(
            phone=args.phone,
            message="Test message from ANC Generator. Texts working!"
        )
        exit(0 if success else 1)

    else:
        parser.print_help()
