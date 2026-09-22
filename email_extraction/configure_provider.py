#!/usr/bin/env python3
"""
Email Provider Configuration Wizard
Helps configure IMAP settings for various email providers
"""

import os
import sys
from pathlib import Path
from dotenv import load_dotenv

# Known email providers and their IMAP settings
PROVIDERS = {
    'gmail': {
        'name': 'Gmail',
        'imap_server': 'imap.gmail.com',
        'imap_port': 993,
        'notes': 'Use App Password (not regular password). Enable 2FA first, then create app password at myaccount.google.com/apppasswords'
    },
    'outlook': {
        'name': 'Microsoft Outlook / Hotmail',
        'imap_server': 'outlook.office365.com',
        'imap_port': 993,
        'notes': 'May require app password if 2FA enabled. Check outlook.live.com settings.'
    },
    'yahoo': {
        'name': 'Yahoo Mail',
        'imap_server': 'imap.mail.yahoo.com',
        'imap_port': 993,
        'notes': 'Generate app password at account.yahoo.com/account/security'
    },
    'turbify': {
        'name': 'Turbify / Yahoo Business',
        'imap_server': 'imap.mail.yahoo.com',
        'imap_port': 993,
        'notes': 'Use your Turbify email and password. Generate app password if needed.'
    },
    'aol': {
        'name': 'AOL Mail',
        'imap_server': 'imap.aol.com',
        'imap_port': 993,
        'notes': 'Generate app password at account.aol.com/security'
    },
    'office365': {
        'name': 'Microsoft Office 365',
        'imap_server': 'outlook.office365.com',
        'imap_port': 993,
        'notes': 'Use your work email. May need to enable modern auth and generate app password.'
    },
    'custom': {
        'name': 'Custom Server',
        'imap_server': '',
        'imap_port': 993,
        'notes': 'Enter your IMAP server details'
    }
}


def print_header(text):
    """Print formatted header"""
    print("\n" + "=" * 70)
    print(text.center(70))
    print("=" * 70 + "\n")


def print_providers():
    """Print available providers"""
    print("Available Email Providers:")
    print()
    for idx, (key, info) in enumerate(PROVIDERS.items(), 1):
        print(f"  {idx}. {info['name']}")
        print(f"     Server: {info['imap_server'] or 'Custom'}")
        print(f"     Port: {info['imap_port']}")
        if info['notes']:
            print(f"     Note: {info['notes']}")
        print()


def get_provider():
    """Get provider choice from user"""
    print_providers()
    
    while True:
        try:
            choice = input("Select provider (1-7) or 'q' to quit: ").strip().lower()
            
            if choice == 'q':
                return None
            
            choice_idx = int(choice) - 1
            providers_list = list(PROVIDERS.items())
            
            if 0 <= choice_idx < len(providers_list):
                return providers_list[choice_idx][0]
            else:
                print("Invalid choice. Please try again.")
        except ValueError:
            print("Invalid input. Please enter a number.")


def configure_provider(provider_key):
    """Configure provider settings"""
    provider = PROVIDERS[provider_key]
    
    print_header(f"Configuring: {provider['name']}")
    
    config = {}
    
    # Get email
    while True:
        email = input("Enter your email address: ").strip()
        if '@' in email:
            config['email'] = email
            break
        else:
            print("Invalid email. Please try again.")
    
    # Get password
    while True:
        password = input("Enter your password (or app password): ").strip()
        if password:
            config['password'] = password
            break
        else:
            print("Password cannot be empty.")
    
    # Get IMAP server
    if provider['imap_server']:
        config['imap_server'] = provider['imap_server']
        print(f"IMAP Server: {provider['imap_server']}")
    else:
        while True:
            imap_server = input("Enter IMAP server address: ").strip()
            if imap_server:
                config['imap_server'] = imap_server
                break
            else:
                print("Server address cannot be empty.")
    
    # Get IMAP port
    while True:
        try:
            port_input = input(f"Enter IMAP port (default {provider['imap_port']}): ").strip()
            if not port_input:
                config['imap_port'] = provider['imap_port']
                break
            else:
                port = int(port_input)
                if 1 <= port <= 65535:
                    config['imap_port'] = port
                    break
                else:
                    print("Port must be between 1 and 65535.")
        except ValueError:
            print("Invalid port. Please enter a number.")
    
    # Get batch size
    while True:
        try:
            batch_size = input("Enter batch size (default 20): ").strip()
            if not batch_size:
                config['batch_size'] = 20
                break
            else:
                size = int(batch_size)
                if size > 0:
                    config['batch_size'] = size
                    break
                else:
                    print("Batch size must be greater than 0.")
        except ValueError:
            print("Invalid batch size. Please enter a number.")
    
    return config


def save_env(config):
    """Save configuration to .env file"""
    env_file = Path('.env')
    
    # Build env content
    content = """# Email Configuration
EMAIL={email}
PASSWORD={password}
IMAP_SERVER={imap_server}
IMAP_PORT={imap_port}
BATCH_SIZE={batch_size}

# Client Sharepoint Configuration (optional)
# DEMO_CLIENT_SHAREPOINT_SITE_URL=
# DEMO_CLIENT_SHAREPOINT_SITE_HOST=
# DEMO_CLIENT_SHAREPOINT_SITE_PATH=
# DEMO_CLIENT_SHAREPOINT_CLIENT_ID=
# DEMO_CLIENT_SHAREPOINT_CLIENT_SECRET=
# DEMO_CLIENT_SHAREPOINT_TENANT_ID=
# DEMO_CLIENT_SHAREPOINT_LIBRARY_NAME=
""".format(**config)
    
    # Load existing if exists to preserve sharepoint config
    if env_file.exists():
        existing = env_file.read_text()
        # Extract sharepoint lines
        sharepoint_lines = [line for line in existing.split('\n') if 'SHAREPOINT' in line or 'TENANT' in line or 'CLIENT_SECRET' in line]
        if sharepoint_lines:
            content += '\n'.join(sharepoint_lines) + '\n'
    
    env_file.write_text(content)
    print(f"\n✓ Configuration saved to {env_file}")


def verify_config(config):
    """Verify configuration works"""
    print("\nVerifying configuration...")
    
    import imaplib
    
    try:
        mail = imaplib.IMAP4_SSL(
            config['imap_server'],
            config['imap_port'],
            timeout=10
        )
        mail.login(config['email'], config['password'])
        mail.close()
        mail.logout()
        
        print("✓ Configuration verified successfully!")
        return True
    except imaplib.IMAP4.error as e:
        print(f"✗ Authentication failed: {e}")
        print("Please check your email and password.")
        return False
    except Exception as e:
        print(f"✗ Connection failed: {e}")
        print("Please check your IMAP server and port.")
        return False


def main():
    """Main execution"""
    print_header("Email Provider Configuration Wizard")
    
    # Get provider
    provider = get_provider()
    if not provider:
        print("Cancelled.")
        return
    
    # Configure
    config = configure_provider(provider)
    
    # Show summary
    print_header("Configuration Summary")
    print(f"Email: {config['email']}")
    print(f"IMAP Server: {config['imap_server']}")
    print(f"IMAP Port: {config['imap_port']}")
    print(f"Batch Size: {config['batch_size']}")
    print()
    
    # Confirm
    confirm = input("Save this configuration? (y/n): ").strip().lower()
    if confirm != 'y':
        print("Cancelled.")
        return
    
    # Save
    save_env(config)
    
    # Verify
    verify = input("\nVerify configuration now? (y/n): ").strip().lower()
    if verify == 'y':
        if verify_config(config):
            print("\n✓ Ready to use! Run: python email_extractor_v2.py")
        else:
            print("\nConfiguration verification failed. Please check your credentials.")
    else:
        print("\n✓ Configuration saved. You can verify later by running: python test_imap_connection_v2.py")


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nCancelled by user.")
        sys.exit(0)

