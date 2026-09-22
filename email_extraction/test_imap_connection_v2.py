#!/usr/bin/env python3
"""
IMAP Connection Tester
Tests IMAP connection to verify credentials and server settings
"""

import os
import sys
import logging
from pathlib import Path
from dotenv import load_dotenv
import imaplib

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Load environment
load_dotenv()

EMAIL = os.getenv('EMAIL', '')
PASSWORD = os.getenv('PASSWORD', '')
IMAP_SERVER = os.getenv('IMAP_SERVER', '')
IMAP_PORT = int(os.getenv('IMAP_PORT', '993'))


def test_imap_connection():
    """Test IMAP connection"""
    
    print("=" * 70)
    print("IMAP CONNECTION TESTER")
    print("=" * 70)
    
    # Validate credentials
    if not EMAIL or not PASSWORD or not IMAP_SERVER:
        logger.error("Missing credentials in .env file")
        logger.error("Required: EMAIL, PASSWORD, IMAP_SERVER")
        return False
    
    logger.info(f"Configuration:")
    logger.info(f"  Email: {EMAIL}")
    logger.info(f"  Server: {IMAP_SERVER}")
    logger.info(f"  Port: {IMAP_PORT}")
    
    try:
        # Test connection
        logger.info("\nTesting connection...")
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT, timeout=10)
        logger.info("✓ SSL connection established")
        
        # Test login
        logger.info("Testing login...")
        mail.login(EMAIL, PASSWORD)
        logger.info("✓ Authentication successful")
        
        # List folders
        logger.info("\nFolders available:")
        status, mailboxes = mail.list()
        if status == 'OK':
            for mailbox in mailboxes:
                folder_name = mailbox.decode('utf-8')
                logger.info(f"  {folder_name}")
        
        # Check INBOX
        logger.info("\nChecking INBOX...")
        status, messages = mail.select('INBOX')
        if status == 'OK':
            email_count = int(messages[0])
            logger.info(f"✓ INBOX has {email_count} emails")
            
            # Try to fetch first email
            if email_count > 0:
                logger.info("\nFetching first email header...")
                status, msg_data = mail.fetch('1', '(RFC822)')
                if status == 'OK':
                    logger.info("✓ Successfully fetched first email")
        else:
            logger.warning("Failed to select INBOX")
        
        # Disconnect
        mail.close()
        mail.logout()
        logger.info("\n✓ Connection test successful!")
        
        return True
        
    except imaplib.IMAP4.error as e:
        logger.error(f"✗ IMAP error: {e}")
        return False
    except Exception as e:
        logger.error(f"✗ Connection error: {e}")
        return False


if __name__ == '__main__':
    success = test_imap_connection()
    sys.exit(0 if success else 1)

