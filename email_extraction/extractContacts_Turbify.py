"""
Turbify Email Contact Extractor
Asynchronous email extraction tool for Turbify business emails
Extracts: From, CC, BCC, Name, Mobile, Logo, and signature email
Saves results to Excel file
"""

import asyncio
import imaplib
import email
from email.header import decode_header
import os
from dotenv import load_dotenv
from bs4 import BeautifulSoup
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import base64
import logging
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
import sys
import urllib.request
import urllib.error
import hashlib

# Fix Windows console encoding issue for Unicode characters
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# Setup logging with encoding fix
logging.basicConfig(
    level=logging.INFO, 
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('email_extraction.log', encoding='utf-8'),
        logging.StreamHandler(stream=sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# Load environment variables from .env file
load_dotenv()

# Get credentials from environment variables
EMAIL_ADDRESS = os.getenv('EMAIL')
PASSWORD = os.getenv('PASSWORD')
IMAP_SERVER = os.getenv('IMAP_SERVER')
IMAP_PORT = int(os.getenv('IMAP_PORT', '993'))
BATCH_SIZE = int(os.getenv('BATCH_SIZE', '20'))

class TurbifyEmailExtractor:
    """Handles Turbify email extraction and contact parsing"""
    
    def __init__(self):
        self.mail = None
        self.extracted_count = 0
        self.failed_count = 0
        self.image_hashes = {}  # Track downloaded images by hash to avoid duplicates
    
    def connect_to_imap(self):
        """Connect to IMAP server"""
        try:
            logger.info(f"Connecting to {IMAP_SERVER}:{IMAP_PORT}")
            self.mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT, timeout=30)  # Increased timeout from 10 to 30 seconds
            self.mail.login(EMAIL_ADDRESS, PASSWORD)
            logger.info(f"[OK] Successfully connected as {EMAIL_ADDRESS}")
            return True
        except imaplib.IMAP4.error as e:
            logger.error(f"IMAP Authentication Error: {e}")
            return False
        except Exception as e:
            logger.error(f"Failed to connect to IMAP: {e}")
            return False
    
    def reconnect_to_imap(self):
        """Reconnect to IMAP server after connection failure"""
        try:
            logger.info(f"[DEBUG] Reconnecting to IMAP...")
            # Try to close old connection
            try:
                self.mail.close()
            except:
                pass
            try:
                self.mail.logout()
            except:
                pass
            # Create new connection
            return self.connect_to_imap()
        except Exception as e:
            logger.error(f"Failed to reconnect: {e}")
            return False
    
    def get_email_list(self, folder='INBOX'):
        """Get list of email sequence numbers from folder, ensure IDs are strings"""
        try:
            status, _ = self.mail.select(folder)
            logger.info(f"[DEBUG] select({folder}) status: {status}")
            status, messages = self.mail.search(None, 'ALL')
            logger.info(f"[DEBUG] search(None, 'ALL') status: {status}")
            if status == 'OK':
                email_ids = messages[0].split()
                # Convert bytes to string
                email_ids = [eid.decode('utf-8') if isinstance(eid, bytes) else str(eid) for eid in email_ids]
                logger.info(f"[OK] Found {len(email_ids)} emails in {folder}")
                logger.info(f"[DEBUG] First 3 sequence numbers: {email_ids[:3]}")
                return email_ids
            logger.warning(f"[WARN] IMAP search did not return OK: {status}, messages: {messages}")
            return []
        except Exception as e:
            logger.error(f"Failed to get email list: {e}")
            return []
    
    def fetch_email_body(self, email_id, retries=3, delay=3, debug_limit=1):
        """Fetch full email message by sequence number, with retry and timeout recovery."""
        import time
        for attempt in range(retries):
            try:
                logger.info(f"[DEBUG] Fetching sequence number: {email_id} (attempt {attempt+1})")
                status, msg_data = self.mail.fetch(email_id, '(RFC822)')
                
                # Debug: log full raw response for first email only
                if hasattr(self, '_debug_count'):
                    self._debug_count += 1
                else:
                    self._debug_count = 1
                
                if self._debug_count <= debug_limit:
                    logger.info(f"[DEBUG] fetch status: {status}, type(msg_data): {type(msg_data)}")
                    logger.info(f"[DEBUG] Full msg_data: {repr(msg_data)}")
                else:
                    logger.info(f"[DEBUG] fetch status: {status}, type(msg_data): {type(msg_data)}, msg_data length: {len(repr(msg_data))}")
                
                if status == 'OK' and msg_data and isinstance(msg_data[0], tuple):
                    if isinstance(msg_data[0][1], bytes):
                        return email.message_from_bytes(msg_data[0][1])
                    else:
                        logger.warning(f"Fetch for email {email_id} did not return bytes, got: {type(msg_data[0][1])}")
                        return None
                else:
                    logger.warning(f"Fetch for email {email_id} returned unexpected response. Status: {status}, msg_data type: {type(msg_data)}")
                    # Attempt to recover connection
                    if attempt < retries - 1:
                        logger.info(f"[DEBUG] Attempting to recover connection...")
                        try:
                            self.mail.select('INBOX')
                        except Exception as e:
                            logger.warning(f"[DEBUG] Connection recovery failed: {e}")
                    return None
            except Exception as e:
                error_str = str(e).lower()
                # Check for SSL/connection/timeout errors
                is_connection_error = ('ssl' in error_str or 'tls' in error_str or 'bad length' in error_str or 
                                      'eof' in error_str or 'socket error' in error_str or 'timed out' in error_str or
                                      'timeout' in error_str or 'closed' in error_str)
                
                logger.warning(f"Failed to fetch email {email_id} (attempt {attempt+1}): {e}")
                
                if attempt < retries - 1:
                    if is_connection_error:
                        logger.info(f"[DEBUG] Connection/timeout error detected. Attempting full reconnect...")
                        # Increase delay for connection errors
                        logger.info(f"[DEBUG] Waiting {delay * 3} seconds before retry...")
                        time.sleep(delay * 3)
                        # Attempt full reconnection
                        if self.reconnect_to_imap():
                            logger.info(f"[DEBUG] Full reconnection successful, trying fetch again...")
                        else:
                            logger.warning(f"[DEBUG] Full reconnection failed")
                    else:
                        logger.info(f"[DEBUG] Waiting {delay} seconds before retry...")
                        time.sleep(delay)
                        # Try to recover connection
                        try:
                            self.mail.select('INBOX')
                        except Exception as e2:
                            logger.warning(f"[DEBUG] Connection recovery failed: {e2}")
                else:
                    logger.info(f"[WARN] Skipping email {email_id} after {retries} failed attempts")
                    return None
    
    @staticmethod
    def decode_email_header(header_value):
        """Decode email header (handles encoding)"""
        if not header_value:
            return ""
        decoded_parts = []
        try:
            for part, encoding in decode_header(header_value):
                if isinstance(part, bytes):
                    decoded_parts.append(part.decode(encoding or 'utf-8', errors='ignore'))
                else:
                    decoded_parts.append(str(part))
            return ''.join(decoded_parts)
        except Exception as e:
            logger.warning(f"Error decoding header: {e}")
            return header_value
    
    @staticmethod
    def extract_email_addresses(email_string):
        """Extract email addresses from email string"""
        if not email_string:
            return []
        
        email_pattern = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b')
        emails = email_pattern.findall(email_string)
        return list(set(emails))  # Remove duplicates
    
    @staticmethod
    def parse_signature(body_content):
        """Parse email signature for name, mobile, email, and logo"""
        if not body_content:
            return None, None, None, None
        
        try:
            soup = BeautifulSoup(body_content, 'html.parser')
            
            # Extract logo: look for img tags in signature area
            logo_data = None
            imgs = soup.find_all('img')
            
            # Look for images - prefer base64, then company logos, then first reasonable-sized image
            if imgs:
                for idx, img in enumerate(imgs):
                    src = img.get('src', '')
                    alt = img.get('alt', '')
                    width = img.get('width', '')
                    height = img.get('height', '')
                    
                    # Skip tracking pixels and very small images
                    try:
                        w = int(width) if width else 0
                        h = int(height) if height else 0
                        if w < 20 and h < 20:
                            continue
                    except:
                        pass
                    
                    # Priority 1: Base64 encoded images (embedded images)
                    if src.startswith('data:image'):
                        logo_data = src
                        break
                    
                    # Priority 2: URLs with 'logo' or 'brand' in them
                    elif 'logo' in src.lower() or 'logo' in alt.lower() or 'brand' in src.lower():
                        logo_data = src
                        break
                    
                    # Priority 3: First external URL that isn't from a known tracking service
                    elif (src.startswith('http') and 
                          not any(x in src.lower() for x in ['pixel', 'beacon', 'track', 'analytics', 'doubleclick', 'googleads'])):
                        # Skip if too small
                        if not (len(src) > 200 and src.endswith(('.png', '.jpg', '.jpeg', '.gif', '.svg'))):
                            continue
                        logo_data = src
                        break
                
                # If no suitable image found, try the first image if it's reasonable
                if not logo_data and len(imgs) > 0:
                    first_img = imgs[0]
                    src = first_img.get('src', '')
                    if src and not any(x in src.lower() for x in ['pixel', 'beacon', 'track', 'analytics']):
                        logo_data = src
            
            # Get text content (compatibility: join lines manually)
            text = '\n'.join(soup.stripped_strings)
            
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
    
    def download_or_save_logo(self, logo_data, email_id, company='', logo_dir='output/logos'):
        """Save logo image (base64 or URL) to output/logos directory, with deduplication based on image hash"""
        if not logo_data:
            return None
        
        try:
            os.makedirs(logo_dir, exist_ok=True)
            # Ensure email_id is a string
            if isinstance(email_id, bytes):
                email_id_str = email_id.decode()
            else:
                email_id_str = str(email_id)
            
            # Sanitize company name for use in filename
            safe_company = ''
            if company:
                safe_company = re.sub(r'[^a-zA-Z0-9_-]', '', company.lower())[:30]  # Remove special chars, limit length
            
            if logo_data.startswith('data:image'):
                # Base64 encoded image - decode and save
                try:
                    # Extract the base64 part after the comma
                    header, encoded = logo_data.split(',', 1)
                    data = base64.b64decode(encoded)
                    
                    # Calculate hash of the image data
                    image_hash = hashlib.md5(data).hexdigest()
                    logger.info(f"[DEBUG] Base64 image hash: {image_hash}")
                    
                    # Check if we've already saved this image
                    if image_hash in self.image_hashes:
                        saved_path = self.image_hashes[image_hash]
                        logger.info(f"[OK] Duplicate image detected (hash: {image_hash}), using existing: {saved_path}")
                        return saved_path
                    
                    # Determine file extension (default to .png for base64)
                    if 'jpeg' in header.lower():
                        ext = '.jpg'
                    elif 'gif' in header.lower():
                        ext = '.gif'
                    elif 'svg' in header.lower():
                        ext = '.svg'
                    else:
                        ext = '.png'
                    
                    # Use company name if available, otherwise use hash
                    if safe_company:
                        filename = os.path.join(logo_dir, f"{safe_company}{ext}")
                    else:
                        filename = os.path.join(logo_dir, f"{image_hash}{ext}")
                    
                    # Handle duplicate filenames (different files with same company name)
                    base_filename = filename
                    counter = 1
                    while os.path.exists(filename):
                        name_parts = base_filename.rsplit('.', 1)
                        filename = f"{name_parts[0]}_{counter}.{name_parts[1]}"
                        counter += 1
                    
                    # Save the file
                    with open(filename, 'wb') as f:
                        f.write(data)
                    
                    # Store in hash dictionary for future deduplication
                    self.image_hashes[image_hash] = filename
                    
                    logger.info(f"[OK] Saved base64 logo to: {filename} ({len(data)} bytes, hash: {image_hash})")
                    return filename
                except Exception as e:
                    logger.warning(f"Failed to decode base64 image: {e}")
                    return None
            elif logo_data.startswith('http'):
                # External URL - download it
                try:
                    # Determine file extension from URL or default to .png
                    ext = '.png'
                    if '.jpg' in logo_data.lower():
                        ext = '.jpg'
                    elif '.jpeg' in logo_data.lower():
                        ext = '.jpeg'
                    elif '.gif' in logo_data.lower():
                        ext = '.gif'
                    elif '.svg' in logo_data.lower():
                        ext = '.svg'
                    
                    logger.info(f"[DEBUG] Downloading logo from URL: {logo_data[:100]}")
                    
                    # Set a timeout and add a user-agent to avoid being blocked
                    req = urllib.request.Request(
                        logo_data,
                        headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
                    )
                    
                    with urllib.request.urlopen(req, timeout=10) as response:
                        data = response.read()
                    
                    # Calculate hash of the downloaded image
                    image_hash = hashlib.md5(data).hexdigest()
                    logger.info(f"[DEBUG] Downloaded image hash: {image_hash}")
                    
                    # Check if we've already saved this image
                    if image_hash in self.image_hashes:
                        saved_path = self.image_hashes[image_hash]
                        logger.info(f"[OK] Duplicate image detected (hash: {image_hash}), using existing: {saved_path}")
                        return saved_path
                    
                    # Use company name if available, otherwise use hash
                    if safe_company:
                        filename = os.path.join(logo_dir, f"{safe_company}{ext}")
                    else:
                        filename = os.path.join(logo_dir, f"{image_hash}{ext}")
                    
                    # Handle duplicate filenames (different files with same company name)
                    base_filename = filename
                    counter = 1
                    while os.path.exists(filename):
                        name_parts = base_filename.rsplit('.', 1)
                        filename = f"{name_parts[0]}_{counter}.{name_parts[1]}"
                        counter += 1
                    
                    with open(filename, 'wb') as f:
                        f.write(data)
                    
                    # Store in hash dictionary for future deduplication
                    self.image_hashes[image_hash] = filename
                    
                    logger.info(f"[OK] Downloaded and saved logo to: {filename} ({len(data)} bytes, hash: {image_hash})")
                    return filename
                except urllib.error.URLError as e:
                    logger.warning(f"Failed to download logo from URL: {e}")
                    # Still return the URL as reference
                    logger.info(f"[OK] Logo URL reference stored: {logo_data}")
                    return logo_data
                except Exception as e:
                    logger.warning(f"Failed to download logo: {e}")
                    # Still return the URL as reference
                    logger.info(f"[OK] Logo URL reference stored: {logo_data}")
                    return logo_data
            else:
                # Other format - store as text reference
                logger.info(f"[OK] Logo reference stored: {logo_data}")
                return logo_data
        
        except Exception as e:
            logger.warning(f"Failed to process logo: {e}")
            return None
    
    @staticmethod
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
    
    def process_email(self, email_id):
        """Process single email and extract from, cc, bcc, subject, and signature info."""
        try:
            # Ensure email_id is a string
            if isinstance(email_id, bytes):
                email_id = email_id.decode('utf-8')
            else:
                email_id = str(email_id)
            
            logger.info(f"Processing email {email_id}...")
            msg = self.fetch_email_body(email_id)
            if not msg:
                self.failed_count += 1
                return None
            # Extract headers
            from_email_raw = self.decode_email_header(msg.get('From', ''))
            cc_raw = self.decode_email_header(msg.get('Cc', ''))
            bcc_raw = self.decode_email_header(msg.get('Bcc', ''))
            subject = self.decode_email_header(msg.get('Subject', ''))
            # Extract email addresses
            cc_emails = self.extract_email_addresses(cc_raw)
            bcc_emails = self.extract_email_addresses(bcc_raw)
            
            # Parse From field to extract email and name
            from_email = ''
            from_name = ''
            # Try to extract name and email from From header (e.g., "John Doe <john@company.com>")
            from_match = re.search(r'^(.*?)\s*<(.+?)>$', from_email_raw.strip())
            if from_match:
                from_name = from_match.group(1).strip().strip('"')  # Remove quotes if present
                from_email = from_match.group(2).strip()
            else:
                # If no angle brackets, assume it's just an email
                from_email = from_email_raw.strip()
                from_name = ''
            
            # Extract signature info
            body = self.get_email_body(msg)
            name, phone, sig_email, logo_data = self.parse_signature(body)
            
            # Use From field name if available, otherwise use signature name
            contact_name = from_name or name or ''
            
            # Extract company name from From header email domain
            company = ""
            if from_email:
                domain_match = re.search(r'@(\S+)', from_email)
                if domain_match:
                    domain = domain_match.group(1)
                    company = domain.split('.')[0] if domain else ""
            
            # Save logo if available
            logo_path = None
            if logo_data:
                logger.info(f"[DEBUG] Email {email_id} has logo data, attempting to save...")
                # email_id is already a string at this point
                # Pass company name for meaningful filename
                logo_path = self.download_or_save_logo(logo_data, email_id, company=company)
                if logo_path:
                    logger.info(f"[OK] Email {email_id}: Logo saved to {logo_path}")
            else:
                logger.info(f"[INFO] Email {email_id}: No logo found in signature")
            
            self.extracted_count += 1
            logger.info(f"[OK] Email {email_id}: Extracted - Email: {from_email}, Name: {contact_name}, Phone: {phone}, Company: {company}, Logo: {'Yes' if logo_path else 'No'}")
            
            return {
                'email': from_email,
                'name': contact_name,
                'company': company or '',
                'cc_emails': ', '.join(cc_emails),
                'bcc_emails': ', '.join(bcc_emails),
                'subject': subject,
                'phone': phone or '',
                'logo': logo_path or ''
            }
        except Exception as e:
            logger.error(f"Error processing email {email_id}: {e}")
            self.failed_count += 1
            return None
    
    async def fetch_emails_async(self, email_ids, batch_size=20):
        """Fetch emails asynchronously using thread pool (lower concurrency for Yahoo/Turbify stability)"""
        import time
        results = []
        loop = asyncio.get_event_loop()
        # Process emails in batches
        for i in range(0, len(email_ids), batch_size):
            batch = email_ids[i:i + batch_size]
            batch_num = (i // batch_size) + 1
            total_batches = (len(email_ids) + batch_size - 1) // batch_size
            logger.info(f"Processing batch {batch_num}/{total_batches} ({len(batch)} emails)")
            # Use thread pool with single worker for stability (SSL connection issues)
            with ThreadPoolExecutor(max_workers=1) as executor:
                tasks = [
                    loop.run_in_executor(executor, self.process_email, email_id)
                    for email_id in batch
                ]
                batch_results = await asyncio.gather(*tasks)
                results.extend([r for r in batch_results if r is not None])
                logger.info(f"  [OK] Batch {batch_num} complete. Progress: {self.extracted_count}/{len(email_ids)}")
            
            # Add delay between batches to avoid SSL connection timeouts
            if batch_num < total_batches:
                logger.info(f"[DEBUG] Waiting 2 seconds before next batch...")
                time.sleep(2)

        return results
    
    def save_to_excel(self, contacts, filename='contacts.xlsx'):
        """Save extracted contacts to Excel file with new column order (without Logo column)"""
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Contacts"
            # Define headers in new order: Email, Name, Company, CC, BCC, Subject, Phone (no Logo)
            headers = [
                'Email', 'Name', 'Company', 'CC Emails', 'BCC Emails', 'Subject', 'Phone'
            ]
            # Add headers with formatting
            ws.append(headers)
            header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
            # Add data in new order (without Logo)
            for contact in contacts:
                ws.append([
                    contact.get('email', ''),
                    contact.get('name', ''),
                    contact.get('company', ''),
                    contact.get('cc_emails', ''),
                    contact.get('bcc_emails', ''),
                    contact.get('subject', ''),
                    contact.get('phone', '')
                ])
            # Auto-adjust column widths
            for column in ws.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2)
                ws.column_dimensions[column_letter].width = min(adjusted_width, 50)
            wb.save(filename)
            logger.info(f"[OK] Data saved to {filename}")
            return True
        except Exception as e:
            logger.error(f"Failed to save to Excel: {e}")
            return False
    
    def disconnect(self):
        """Close IMAP connection"""
        try:
            if self.mail:
                try:
                    self.mail.close()
                except Exception as e:
                    # Suppress errors on close (common if already closed)
                    logger.info(f"[INFO] IMAP close() warning: {e}")
                try:
                    self.mail.logout()
                except Exception as e:
                    # Suppress SSL errors on logout
                    if 'ssl' in str(e).lower() or 'bad length' in str(e).lower():
                        logger.info(f"[INFO] IMAP logout() SSL warning: {e}")
                    else:
                        logger.error(f"Error disconnecting: {e}")
                logger.info("[OK] Disconnected from IMAP server")
        except Exception as e:
            logger.error(f"Error disconnecting: {e}")

async def main():
    """Main execution function"""
    try:
        logger.info("="*60)
        logger.info("TURBIFY EMAIL CONTACT EXTRACTOR")
        logger.info("="*60)
        
        # Validate credentials
        if not EMAIL_ADDRESS or not PASSWORD or not IMAP_SERVER:
            logger.error("[FAIL] Missing credentials in .env file")
            logger.error("Please set: EMAIL, PASSWORD, and IMAP_SERVER")
            return
        # Debug: Log loaded credentials (except password)
        logger.info(f"Loaded EMAIL: {EMAIL_ADDRESS}")
        logger.info(f"Loaded IMAP_SERVER: {IMAP_SERVER}")
        
        # Initialize extractor
        extractor = TurbifyEmailExtractor()
        
        # Connect to IMAP
        if not extractor.connect_to_imap():
            logger.error("[FAIL] Failed to connect to IMAP server")
            return
        
        # Get email list
        email_ids = extractor.get_email_list('INBOX')
        if not email_ids:
            logger.warning("[WARN] No emails found in INBOX")
            extractor.disconnect()
            return
        # Process all emails (no limit)
        logger.info(f"Starting extraction of {len(email_ids)} emails with batch size {BATCH_SIZE}")
        
        # Fetch and process emails asynchronously
        start_time = datetime.now()
        contacts = await extractor.fetch_emails_async(email_ids, BATCH_SIZE)
        elapsed_time = datetime.now() - start_time
        
        # Disconnect
        extractor.disconnect()
        
        # Log results
        logger.info("="*60)
        logger.info("EXTRACTION COMPLETE")
        logger.info("="*60)
        logger.info(f"[OK] Successfully extracted: {extractor.extracted_count} emails")
        logger.info(f"[FAIL] Failed: {extractor.failed_count} emails")
        logger.info(f"[TIME] Time elapsed: {elapsed_time.total_seconds():.2f} seconds")
        
        # Save to Excel in output directory
        if contacts:
            # Create output directory if it doesn't exist
            os.makedirs('output', exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f'output/contacts_{timestamp}.xlsx'
            extractor.save_to_excel(contacts, filename)
            logger.info(f"[OK] Excel file saved to: {filename}")
            logger.info(f"[OK] Logos saved to: output/logos/")
        else:
            logger.warning("[WARN] No contacts extracted")
    
    except Exception as e:
        logger.error(f"Fatal error: {e}", exc_info=True)

if __name__ == "__main__":
    asyncio.run(main())

