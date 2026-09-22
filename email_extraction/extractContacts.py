import asyncio
import imaplib
import email
from email.header import decode_header
import os
from dotenv import load_dotenv
from bs4 import BeautifulSoup
import re
from openpyxl import Workbook
import base64
import logging
from concurrent.futures import ThreadPoolExecutor

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Load environment variables from .env file
load_dotenv()

# Get credentials from environment variables
EMAIL_ADDRESS = os.getenv('EMAIL')
PASSWORD = os.getenv('PASSWORD')
IMAP_SERVER = os.getenv('IMAP_SERVER')
IMAP_PORT = int(os.getenv('IMAP_PORT', '993'))
BATCH_SIZE = int(os.getenv('BATCH_SIZE', '50'))

def connect_to_imap():
    """Connect to IMAP server"""
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(EMAIL_ADDRESS, PASSWORD)
        logger.info(f"Connected to {IMAP_SERVER} as {EMAIL_ADDRESS}")
        return mail
    except Exception as e:
        logger.error(f"Failed to connect to IMAP: {e}")
        return None

def get_email_list(mail, folder='INBOX'):
    """Get list of email IDs from folder"""
    try:
        mail.select(folder)
        status, messages = mail.search(None, 'ALL')
        if status == 'OK':
            email_ids = messages[0].split()
            logger.info(f"Found {len(email_ids)} emails in {folder}")
            return email_ids
        return []
    except Exception as e:
        logger.error(f"Failed to get email list: {e}")
        return []

def fetch_email_body(mail, email_id):
    """Fetch full email message by ID"""
    try:
        status, msg_data = mail.fetch(email_id, '(RFC822)')
        if status == 'OK':
            return email.message_from_bytes(msg_data[0][1])
        return None
    except Exception as e:
        logger.error(f"Failed to fetch email {email_id}: {e}")
        return None

def decode_email_header(header_value):
    """Decode email header (handles encoding)"""
    if not header_value:
        return ""
    decoded_parts = []
    for part, encoding in decode_header(header_value):
        if isinstance(part, bytes):
            decoded_parts.append(part.decode(encoding or 'utf-8', errors='ignore'))
        else:
            decoded_parts.append(str(part))
    return ''.join(decoded_parts)

def extract_email_addresses(email_string):
    """Extract email addresses from email string (comma-separated or From: format)"""
    if not email_string:
        return []
    
    # Extract emails using regex
    email_pattern = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b')
    emails = email_pattern.findall(email_string)
    return list(set(emails))  # Remove duplicates

def parse_signature(body_content):
    """Parse email signature for name, mobile, email, and logo"""
    if not body_content:
        return None, None, None, None
    
    try:
        soup = BeautifulSoup(body_content, 'html.parser')
        
        # Extract logo: look for img tags, prefer those with 'logo' in src or alt
        logo_data = None
        imgs = soup.find_all('img')
        for img in imgs:
            src = img.get('src', '')
            alt = img.get('alt', '')
            if src.startswith('data:image'):
                logo_data = src
                break
            elif 'logo' in src.lower() or 'logo' in alt.lower():
                logo_data = src
                break
        
        # Get text content
        text = soup.get_text(separator='\n')
        
        # Regex patterns
        phone_pattern = re.compile(r'\b(?:\+?\d{1,3}[-.\s]?)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}\b')
        email_pattern = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b')
        
        phones = phone_pattern.findall(text)
        mobile = phones[0] if phones else None
        
        emails = email_pattern.findall(text)
        sig_email = emails[0] if emails else None
        
        # Name: assume first non-empty line without numbers or emails
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        name = None
        for line in lines:
            if not re.search(r'\d|@', line) and len(line.split()) <= 5 and len(line) < 50:
                name = line
                break
        
        return name, mobile, sig_email, logo_data
    
    except Exception as e:
        logger.warning(f"Error parsing signature: {e}")
        return None, None, None, None

def download_or_save_logo(logo_data, email_id, logo_dir):
    """Save logo image (base64 or URL)"""
    if not logo_data:
        return None
    
    try:
        os.makedirs(logo_dir, exist_ok=True)
        filename = os.path.join(logo_dir, f"{email_id.decode() if isinstance(email_id, bytes) else email_id}.png")
        
        if logo_data.startswith('data:image'):
            # Base64 encoded image
            try:
                header, encoded = logo_data.split(',', 1)
                data = base64.b64decode(encoded)
                with open(filename, 'wb') as f:
                    f.write(data)
                return filename
            except Exception as e:
                logger.warning(f"Failed to decode base64 image: {e}")
                return None
        else:
            # URL - skip for now as we'd need async requests
            return None
    
    except Exception as e:
        logger.warning(f"Failed to save logo: {e}")
        return None

def get_email_body(msg):
    """Extract email body (text or html)"""
    body = ""
    try:
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                try:
                    if content_type == 'text/html':
                        body = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                        break
                    elif content_type == 'text/plain' and not body:
                        body = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                except:
                    pass
        else:
            body = msg.get_payload(decode=True).decode('utf-8', errors='ignore')
    except Exception as e:
        logger.warning(f"Error extracting body: {e}")
    
    return body

def process_email(mail, email_id):
    """Process single email and extract contact info"""
    try:
        msg = fetch_email_body(mail, email_id)
        if not msg:
            return None
        
        # Extract headers
        from_email = decode_email_header(msg.get('From', ''))
        cc_raw = decode_email_header(msg.get('Cc', ''))
        bcc_raw = decode_email_header(msg.get('Bcc', ''))
        
        # Extract email addresses
        cc_emails = extract_email_addresses(cc_raw)
        bcc_emails = extract_email_addresses(bcc_raw)
        
        # Get body content
        body_content = get_email_body(msg)
        
        # Parse signature
        name, mobile, sig_email, logo_data = parse_signature(body_content)
        
        # Save logo
        logo_path = download_or_save_logo(logo_data, email_id, 'logos')
        
        return {
            'from_email': from_email,
            'cc_emails': ', '.join(cc_emails),
            'bcc_emails': ', '.join(bcc_emails),
            'name': name,
            'mobile': mobile,
            'sig_email': sig_email,
            'logo_path': logo_path
        }
    
    except Exception as e:
        logger.error(f"Error processing email {email_id}: {e}")
        return None

async def fetch_emails_async(mail, email_ids, batch_size=50):
    """Fetch emails asynchronously using thread pool"""
    results = []
    loop = asyncio.get_event_loop()
    
    # Process emails in batches
    for i in range(0, len(email_ids), batch_size):
        batch = email_ids[i:i + batch_size]
        logger.info(f"Processing batch {i//batch_size + 1} ({len(batch)} emails)")
        
        # Use thread pool for I/O-bound IMAP operations
        with ThreadPoolExecutor(max_workers=5) as executor:
            tasks = [
                loop.run_in_executor(executor, process_email, mail, email_id)
                for email_id in batch
            ]
            batch_results = await asyncio.gather(*tasks)
            results.extend([r for r in batch_results if r is not None])
    
    return results

async def main():
    try:
        # Validate credentials
        if not EMAIL_ADDRESS or not PASSWORD or not IMAP_SERVER:
            logger.error("Missing credentials in .env file. Please set EMAIL, PASSWORD, and IMAP_SERVER")
            return
        
        # Connect to IMAP
        mail = connect_to_imap()
        if not mail:
            return
        
        # Get email list
        email_ids = get_email_list(mail, 'INBOX')
        if not email_ids:
            logger.warning("No emails found")
            mail.close()
            mail.logout()
            return
        
        logger.info(f"Processing {len(email_ids)} emails with batch size {BATCH_SIZE}")
        
        # Fetch and process emails asynchronously
        contacts = await fetch_emails_async(mail, email_ids, BATCH_SIZE)
        
        # Close IMAP connection
        mail.close()
        mail.logout()
        
        logger.info(f"Successfully processed {len(contacts)} emails")
        
        # Save to Excel
        if contacts:
            wb = Workbook()
            ws = wb.active
            ws.title = "Contacts"
            ws.append(['From Email', 'CC Emails', 'BCC Emails', 'Name', 'Mobile', 'Email from Sig', 'Logo Path'])
            
            for contact in contacts:
                ws.append([
                    contact['from_email'],
                    contact['cc_emails'],
                    contact['bcc_emails'],
                    contact['name'],
                    contact['mobile'],
                    contact['sig_email'],
                    contact['logo_path']
                ])
            
            wb.save('contacts.xlsx')
            logger.info("Data saved to contacts.xlsx")
        else:
            logger.warning("No contacts extracted")
    
    except Exception as e:
        logger.error(f"Error in main: {e}")

if __name__ == "__main__":
    asyncio.run(main())

