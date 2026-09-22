"""
Turbify IMAP Connection Tester
Tests IMAP connectivity with your actual credentials
"""

import imaplib
import os
from dotenv import load_dotenv
import sys

# Load environment variables
load_dotenv()

EMAIL_ADDRESS = os.getenv('EMAIL')
PASSWORD = os.getenv('PASSWORD')
IMAP_SERVER = os.getenv('IMAP_SERVER')
IMAP_PORT = int(os.getenv('IMAP_PORT', '993'))

def test_imap_connection():
    """Test IMAP connection with actual credentials"""
    
    print("\n" + "="*60)
    print("TURBIFY IMAP CONNECTION TESTER")
    print("="*60)
    
    # Validate configuration
    print("\n[1] Validating Configuration...")
    if not EMAIL_ADDRESS or not PASSWORD or not IMAP_SERVER:
        print("❌ Missing credentials in .env file")
        print("Required: EMAIL, PASSWORD, IMAP_SERVER")
        return False
    
    print(f"✓ Email: {EMAIL_ADDRESS}")
    print(f"✓ Server: {IMAP_SERVER}:{IMAP_PORT}")
    print(f"✓ Password: {'*' * len(PASSWORD)}")
    
    # Test connection
    print("\n[2] Testing IMAP Connection...")
    try:
        print(f"   Connecting to {IMAP_SERVER}:{IMAP_PORT}...")
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT, timeout=10)
        print("✓ SSL/TLS Connection established")
    except Exception as e:
        print(f"❌ Connection failed: {e}")
        return False
    
    # Test authentication
    print("\n[3] Testing Authentication...")
    try:
        mail.login(EMAIL_ADDRESS, PASSWORD)
        print(f"✓ Successfully authenticated as {EMAIL_ADDRESS}")
    except imaplib.IMAP4.error as e:
        print(f"❌ Authentication failed: {e}")
        print("   Verify email and password are correct")
        mail.close()
        return False
    except Exception as e:
        print(f"❌ Authentication error: {e}")
        mail.close()
        return False
    
    # Get mailbox list
    print("\n[4] Listing Available Folders...")
    try:
        status, mailboxes = mail.list()
        if status == 'OK':
            print(f"✓ Found {len(mailboxes)} folders:")
            for mailbox in mailboxes[:10]:  # Show first 10
                folder_name = mailbox.decode().split('"') [-1]
                print(f"   • {folder_name}")
            if len(mailboxes) > 10:
                print(f"   ... and {len(mailboxes) - 10} more")
    except Exception as e:
        print(f"⚠ Could not list folders: {e}")
    
    # Get INBOX statistics
    print("\n[5] Checking INBOX...")
    try:
        mail.select('INBOX')
        status, messages = mail.search(None, 'ALL')
        
        if status == 'OK':
            email_count = len(messages[0].split())
            print(f"✓ INBOX has {email_count} emails")
            
            # Show last few email subjects
            if email_count > 0:
                print("\n   Recent emails:")
                latest_ids = messages[0].split()[-3:]  # Get last 3
                
                for email_id in latest_ids:
                    try:
                        status, msg_data = mail.fetch(email_id, '(BODY[HEADER.FIELDS (SUBJECT FROM)])')
                        if status == 'OK':
                            msg_str = msg_data[0][1].decode()
                            subject_line = [line for line in msg_str.split('\r\n') if 'Subject:' in line]
                            from_line = [line for line in msg_str.split('\r\n') if 'From:' in line]
                            
                            if subject_line:
                                print(f"      • {subject_line[0]}")
                            if from_line:
                                print(f"        {from_line[0]}")
                    except:
                        pass
    except Exception as e:
        print(f"⚠ Could not access INBOX: {e}")
    
    # Close connection
    print("\n[6] Closing Connection...")
    try:
        mail.close()
        mail.logout()
        print("✓ Connection closed successfully")
    except Exception as e:
        print(f"⚠ Error closing connection: {e}")
    
    # Summary
    print("\n" + "="*60)
    print("✅ CONNECTION TEST SUCCESSFUL!")
    print("="*60)
    print("\nYou can now run the email extractor:")
    print("  python General\\extractContacts_Turbify.py")
    print("\n" + "="*60 + "\n")
    
    return True

if __name__ == "__main__":
    success = test_imap_connection()
    sys.exit(0 if success else 1)

