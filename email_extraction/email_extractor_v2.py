"""
Email Contact Extractor v2 - Production Ready
============================================
Asynchronous email extraction from Outlook/Turbify business emails.

Features:
- Batch processing with configurable concurrency
- Async/await pattern for non-blocking I/O
- Resilient SSL connection handling and reconnection logic
- Signature parsing: name, phone, email, logo image
- Duplicate image detection via MD5 hashing
- Excel export with formatted headers
- Comprehensive logging and error handling

Usage:
    python email_extractor_v2.py [--limit N] [--batch-size N]

Environment Variables (.env):
    EMAIL           - Email address for IMAP login
    PASSWORD        - IMAP password (or app password)
    IMAP_SERVER     - IMAP server address (e.g., imap.mail.yahoo.com)
    IMAP_PORT       - IMAP port (default: 993 for SSL)
    BATCH_SIZE      - Emails per batch (default: 20)
"""

import asyncio
import imaplib
import email
import os
import sys
import logging
import re
import base64
import hashlib
import urllib.request
import urllib.error
from email.header import decode_header
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict, List, Tuple
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, asdict

# Third-party imports
from dotenv import load_dotenv
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

# Fix Windows console encoding
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# ============================================================================
# CONFIGURATION & LOGGING
# ============================================================================

load_dotenv()

# Logging configuration
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('email_extraction_v2.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# Environment variables
EMAIL_ADDRESS = os.getenv('EMAIL', '')
PASSWORD = os.getenv('PASSWORD', '')
IMAP_SERVER = os.getenv('IMAP_SERVER', '')
IMAP_PORT = int(os.getenv('IMAP_PORT', '993'))
BATCH_SIZE = int(os.getenv('BATCH_SIZE', '20'))

# Directories
OUTPUT_DIR = Path('output')
LOGOS_DIR = OUTPUT_DIR / 'logos'
OUTPUT_DIR.mkdir(exist_ok=True)
LOGOS_DIR.mkdir(parents=True, exist_ok=True)

# ============================================================================
# DATA MODEL
# ============================================================================

@dataclass
class Contact:
    """Email contact information extracted from an email"""
    email: str
    name: str = ''
    company: str = ''
    cc_emails: str = ''
    bcc_emails: str = ''
    subject: str = ''
    phone: str = ''
    signature_email: str = ''
    logo_file: str = ''

    def to_dict(self) -> Dict:
        """Convert to dictionary"""
        return asdict(self)


@dataclass
class ExtractionStats:
    """Statistics from extraction run"""
    total_emails: int = 0
    successful: int = 0
    failed: int = 0
    elapsed_seconds: float = 0.0

    def __str__(self) -> str:
        return (
            f"Total: {self.total_emails}, "
            f"Successful: {self.successful}, "
            f"Failed: {self.failed}, "
            f"Time: {self.elapsed_seconds:.1f}s"
        )


# ============================================================================
# EMAIL PARSER
# ============================================================================

class EmailParser:
    """Parses email messages and extracts contact information"""

    @staticmethod
    def decode_header(header_value: Optional[str]) -> str:
        """Decode email header with proper encoding handling"""
        if not header_value:
            return ''

        try:
            decoded_parts = []
            for part, encoding in decode_header(header_value):
                if isinstance(part, bytes):
                    decoded_parts.append(part.decode(encoding or 'utf-8', errors='ignore'))
                else:
                    decoded_parts.append(str(part))
            return ''.join(decoded_parts)
        except Exception as e:
            logger.warning(f"Error decoding header: {e}")
            return str(header_value)

    @staticmethod
    def extract_emails(text: str) -> List[str]:
        """Extract unique email addresses from text"""
        if not text:
            return []
        pattern = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b')
        return list(set(pattern.findall(text)))

    @staticmethod
    def parse_from_header(from_header: str) -> Tuple[str, str]:
        """
        Parse From header to extract email and name.
        
        Examples:
            "John Doe <john@example.com>" -> ("john@example.com", "John Doe")
            "john@example.com" -> ("john@example.com", "")
        """
        from_header = from_header.strip()
        
        # Try to extract name and email from "Name <email>" format
        match = re.search(r'^(.*?)\s*<(.+?)>$', from_header)
        if match:
            name = match.group(1).strip().strip('"\'')
            email_addr = match.group(2).strip()
            return email_addr, name

        # Try to extract from "email" format
        match = re.search(r'([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,})', from_header)
        if match:
            return match.group(1), ''

        return from_header, ''

    @staticmethod
    def extract_company_from_email(email_addr: str) -> str:
        """Extract company name from email domain"""
        if not email_addr:
            return ''
        match = re.search(r'@([^.]+)', email_addr)
        return match.group(1) if match else ''

    @staticmethod
    def get_email_body(msg) -> str:
        """Extract email body (preferring HTML over plain text)"""
        body = ''
        try:
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    try:
                        if content_type == 'text/html':
                            body = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                            return body
                        elif content_type == 'text/plain' and not body:
                            body = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                    except Exception:
                        pass
            else:
                body = msg.get_payload(decode=True).decode('utf-8', errors='ignore')
        except Exception as e:
            logger.warning(f"Error extracting body: {e}")
        return body

    @staticmethod
    def parse_signature(html_content: str) -> Tuple[Optional[str], Optional[str], Optional[str], Optional[str]]:
        """
        Parse email signature for name, phone, email, and logo.
        
        Returns:
            (name, phone, email, logo_data)
        """
        if not html_content:
            return None, None, None, None

        try:
            soup = BeautifulSoup(html_content, 'html.parser')

            # Extract logo from img tags
            logo_data = None
            imgs = soup.find_all('img')

            if imgs:
                for img in imgs:
                    src = img.get('src', '')
                    alt = img.get('alt', '')
                    width = img.get('width', '')
                    height = img.get('height', '')

                    # Skip tracking pixels
                    try:
                        w = int(width) if width else 0
                        h = int(height) if height else 0
                        if w < 20 and h < 20:
                            continue
                    except Exception:
                        pass

                    # Priority 1: Base64 embedded images
                    if src.startswith('data:image'):
                        logo_data = src
                        break

                    # Priority 2: Logo/brand URLs
                    elif any(x in src.lower() for x in ['logo', 'brand']):
                        logo_data = src
                        break

                    # Priority 3: External URLs (not tracking)
                    elif src.startswith('http'):
                        if not any(x in src.lower() for x in ['pixel', 'beacon', 'track', 'analytics', 'doubleclick']):
                            logo_data = src
                            break

            # Extract text content
            text = '\n'.join(soup.stripped_strings)

            # Extract phone (common formats)
            phone_pattern = re.compile(r'\b(?:\+?\d{1,3}[-.\s]?)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}\b')
            phones = phone_pattern.findall(text)
            phone = phones[0] if phones else None

            # Extract email addresses
            emails = EmailParser.extract_emails(text)
            sig_email = emails[0] if emails else None

            # Extract name (first non-email, non-phone line)
            name = None
            for line in text.split('\n'):
                line = line.strip()
                if (line and 
                    not re.search(r'\d{3}.*\d{4}', line) and  # Not phone
                    '@' not in line and  # Not email
                    len(line.split()) <= 5 and  # Reasonable length
                    len(line) < 50):
                    name = line
                    break

            return name, phone, sig_email, logo_data

        except Exception as e:
            logger.warning(f"Error parsing signature: {e}")
            return None, None, None, None


# ============================================================================
# IMAGE HANDLER
# ============================================================================

class ImageHandler:
    """Handles image downloads and deduplication"""

    def __init__(self):
        self.image_hashes: Dict[str, str] = {}  # hash -> filepath mapping

    def sanitize_filename(self, name: str, max_len: int = 30) -> str:
        """Sanitize filename for filesystem"""
        sanitized = re.sub(r'[^a-zA-Z0-9_-]', '', name.lower())
        return sanitized[:max_len]

    def save_logo(self, logo_data: str, company: str = '', email_id: str = '') -> Optional[str]:
        """
        Save logo from URL or base64 data.
        
        Returns:
            Path to saved file, or URL if download failed
        """
        if not logo_data:
            return None

        try:
            # Handle base64 embedded images
            if logo_data.startswith('data:image'):
                return self._save_base64_image(logo_data, company)

            # Handle HTTP(S) URLs
            elif logo_data.startswith('http'):
                return self._download_image(logo_data, company)

            else:
                logger.debug(f"Unknown logo format, storing reference: {logo_data[:50]}")
                return logo_data

        except Exception as e:
            logger.warning(f"Failed to process logo: {e}")
            return None

    def _save_base64_image(self, data_uri: str, company: str) -> Optional[str]:
        """Save base64 encoded image"""
        try:
            # Parse data URI
            header, encoded = data_uri.split(',', 1)
            image_data = base64.b64decode(encoded)

            # Calculate hash for deduplication
            image_hash = hashlib.md5(image_data).hexdigest()

            # Check if already saved
            if image_hash in self.image_hashes:
                logger.debug(f"Duplicate image detected (hash: {image_hash})")
                return self.image_hashes[image_hash]

            # Determine extension
            ext = '.png'
            if 'jpeg' in header.lower():
                ext = '.jpg'
            elif 'gif' in header.lower():
                ext = '.gif'
            elif 'svg' in header.lower():
                ext = '.svg'

            # Generate filename
            safe_company = self.sanitize_filename(company)
            if safe_company:
                filename = LOGOS_DIR / f"{safe_company}{ext}"
            else:
                filename = LOGOS_DIR / f"{image_hash}{ext}"

            # Handle filename collisions
            counter = 1
            original_filename = filename
            while filename.exists():
                filename = original_filename.parent / f"{original_filename.stem}_{counter}{ext}"
                counter += 1

            # Save file
            filename.write_bytes(image_data)
            self.image_hashes[image_hash] = str(filename)

            logger.debug(f"Saved base64 image: {filename.name} ({len(image_data)} bytes)")
            return str(filename)

        except Exception as e:
            logger.warning(f"Failed to save base64 image: {e}")
            return None

    def _download_image(self, url: str, company: str) -> Optional[str]:
        """Download image from URL"""
        try:
            # Determine extension from URL
            ext = '.png'
            if '.jpg' in url.lower():
                ext = '.jpg'
            elif '.jpeg' in url.lower():
                ext = '.jpeg'
            elif '.gif' in url.lower():
                ext = '.gif'
            elif '.svg' in url.lower():
                ext = '.svg'

            logger.debug(f"Downloading logo from URL: {url[:80]}")

            # Download with timeout and user-agent
            req = urllib.request.Request(
                url,
                headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
            )

            with urllib.request.urlopen(req, timeout=10) as response:
                image_data = response.read()

            # Calculate hash
            image_hash = hashlib.md5(image_data).hexdigest()

            # Check if already saved
            if image_hash in self.image_hashes:
                logger.debug(f"Duplicate image detected (hash: {image_hash})")
                return self.image_hashes[image_hash]

            # Generate filename
            safe_company = self.sanitize_filename(company)
            if safe_company:
                filename = LOGOS_DIR / f"{safe_company}{ext}"
            else:
                filename = LOGOS_DIR / f"{image_hash}{ext}"

            # Handle collisions
            counter = 1
            original_filename = filename
            while filename.exists():
                filename = original_filename.parent / f"{original_filename.stem}_{counter}{ext}"
                counter += 1

            # Save file
            filename.write_bytes(image_data)
            self.image_hashes[image_hash] = str(filename)

            logger.debug(f"Downloaded and saved logo: {filename.name} ({len(image_data)} bytes)")
            return str(filename)

        except urllib.error.URLError as e:
            logger.debug(f"Failed to download logo from {url[:50]}: {e}")
            # Return URL as reference
            return url
        except Exception as e:
            logger.warning(f"Failed to download image: {e}")
            return None


# ============================================================================
# IMAP CLIENT
# ============================================================================

class IMAPClient:
    """Manages IMAP connection with resilience"""

    def __init__(self, server: str, port: int, email: str, password: str):
        self.server = server
        self.port = port
        self.email = email
        self.password = password
        self.mail: Optional[imaplib.IMAP4_SSL] = None

    def connect(self) -> bool:
        """Connect to IMAP server"""
        try:
            logger.info(f"Connecting to {self.server}:{self.port}")
            self.mail = imaplib.IMAP4_SSL(self.server, self.port, timeout=30)
            self.mail.login(self.email, self.password)
            logger.info(f"Successfully logged in as {self.email}")
            return True
        except imaplib.IMAP4.error as e:
            logger.error(f"IMAP authentication failed: {e}")
            return False
        except Exception as e:
            logger.error(f"Failed to connect: {e}")
            return False

    def reconnect(self) -> bool:
        """Reconnect to IMAP server"""
        try:
            self.disconnect()
            return self.connect()
        except Exception as e:
            logger.error(f"Reconnection failed: {e}")
            return False

    def disconnect(self):
        """Disconnect from IMAP server"""
        if not self.mail:
            return

        try:
            try:
                self.mail.close()
            except Exception:
                pass
            try:
                self.mail.logout()
            except Exception:
                pass
            logger.info("Disconnected from IMAP server")
        except Exception as e:
            logger.error(f"Error disconnecting: {e}")

    def get_email_ids(self, folder: str = 'INBOX') -> List[str]:
        """Get list of email IDs from folder"""
        try:
            status, _ = self.mail.select(folder)
            if status != 'OK':
                logger.error(f"Failed to select folder {folder}")
                return []

            status, messages = self.mail.search(None, 'ALL')
            if status != 'OK':
                logger.error("Failed to search emails")
                return []

            email_ids = messages[0].split()
            email_ids = [
                eid.decode('utf-8') if isinstance(eid, bytes) else str(eid)
                for eid in email_ids
            ]

            logger.info(f"Found {len(email_ids)} emails in {folder}")
            return email_ids

        except Exception as e:
            logger.error(f"Failed to get email IDs: {e}")
            return []

    def fetch_email(self, email_id: str, retries: int = 3, delay: int = 2) -> Optional:
        """Fetch single email by ID with retry logic"""
        import time

        for attempt in range(retries):
            try:
                status, msg_data = self.mail.fetch(email_id, '(RFC822)')

                if status == 'OK' and msg_data and isinstance(msg_data[0], tuple):
                    if isinstance(msg_data[0][1], bytes):
                        return email.message_from_bytes(msg_data[0][1])
                    else:
                        logger.warning(f"Unexpected response type for email {email_id}")
                        return None
                else:
                    logger.warning(f"Failed to fetch email {email_id}: unexpected response")
                    return None

            except Exception as e:
                error_msg = str(e).lower()
                is_connection_error = any(
                    x in error_msg for x in ['ssl', 'tls', 'eof', 'socket', 'timeout', 'bad length']
                )

                logger.warning(f"Fetch attempt {attempt + 1}/{retries} failed: {e}")

                if attempt < retries - 1:
                    wait_time = delay * 3 if is_connection_error else delay
                    logger.info(f"Waiting {wait_time}s before retry...")
                    time.sleep(wait_time)

                    if is_connection_error:
                        logger.info("Attempting reconnection...")
                        if not self.reconnect():
                            logger.warning("Reconnection failed")
                else:
                    logger.warning(f"Giving up on email {email_id} after {retries} attempts")

        return None


# ============================================================================
# CONTACT EXTRACTOR
# ============================================================================

class ContactExtractor:
    """Main contact extraction engine"""

    def __init__(self, imap_client: IMAPClient):
        self.imap_client = imap_client
        self.image_handler = ImageHandler()
        self.parser = EmailParser()
        self.stats = ExtractionStats()

    def process_email(self, email_id: str) -> Optional[Contact]:
        """Process single email and extract contact information"""
        try:
            # Fetch email
            msg = self.imap_client.fetch_email(email_id)
            if not msg:
                self.stats.failed += 1
                return None

            # Extract basic headers
            from_header = self.parser.decode_header(msg.get('From', ''))
            cc_header = self.parser.decode_header(msg.get('Cc', ''))
            bcc_header = self.parser.decode_header(msg.get('Bcc', ''))
            subject = self.parser.decode_header(msg.get('Subject', ''))

            # Parse From header
            email_addr, from_name = self.parser.parse_from_header(from_header)

            # Extract email addresses
            cc_emails = self.parser.extract_emails(cc_header)
            bcc_emails = self.parser.extract_emails(bcc_header)

            # Extract company from email domain
            company = self.parser.extract_company_from_email(email_addr)

            # Parse signature
            body = self.parser.get_email_body(msg)
            name, phone, sig_email, logo_data = self.parser.parse_signature(body)

            # Use From name if available, otherwise signature name
            contact_name = from_name or name or ''

            # Save logo if available
            logo_file = None
            if logo_data:
                logo_file = self.image_handler.save_logo(logo_data, company, email_id)

            # Create contact object
            contact = Contact(
                email=email_addr,
                name=contact_name,
                company=company,
                cc_emails=', '.join(cc_emails),
                bcc_emails=', '.join(bcc_emails),
                subject=subject,
                phone=phone or '',
                signature_email=sig_email or '',
                logo_file=logo_file or ''
            )

            self.stats.successful += 1
            logger.info(f"Extracted: {email_addr} ({contact_name})")
            return contact

        except Exception as e:
            logger.error(f"Error processing email {email_id}: {e}")
            self.stats.failed += 1
            return None

    async def extract_batch(self, email_ids: List[str], batch_num: int, total_batches: int) -> List[Contact]:
        """Process batch of emails asynchronously"""
        logger.info(f"Processing batch {batch_num}/{total_batches} ({len(email_ids)} emails)")

        loop = asyncio.get_event_loop()
        contacts = []

        # Use single worker to maintain IMAP connection stability
        with ThreadPoolExecutor(max_workers=1) as executor:
            tasks = [
                loop.run_in_executor(executor, self.process_email, eid)
                for eid in email_ids
            ]
            results = await asyncio.gather(*tasks)
            contacts = [r for r in results if r is not None]

        logger.info(f"Batch {batch_num} complete ({self.stats.successful}/{self.stats.total_emails})")
        return contacts

    async def extract_all(self, email_ids: List[str], batch_size: int) -> List[Contact]:
        """Extract all emails in batches"""
        self.stats.total_emails = len(email_ids)
        all_contacts = []

        # Process in batches
        for i in range(0, len(email_ids), batch_size):
            batch = email_ids[i:i + batch_size]
            batch_num = (i // batch_size) + 1
            total_batches = (len(email_ids) + batch_size - 1) // batch_size

            contacts = await self.extract_batch(batch, batch_num, total_batches)
            all_contacts.extend(contacts)

            # Delay between batches
            if batch_num < total_batches:
                logger.debug("Waiting 2s before next batch...")
                await asyncio.sleep(2)

        return all_contacts


# ============================================================================
# EXCEL EXPORTER
# ============================================================================

class ExcelExporter:
    """Exports contacts to Excel file"""

    HEADERS = [
        'Email', 'Name', 'Company', 'CC Emails', 'BCC Emails', 'Subject', 'Phone'
    ]

    @staticmethod
    def export(contacts: List[Contact], filename: Optional[str] = None) -> Optional[str]:
        """Export contacts to Excel"""
        try:
            if not contacts:
                logger.warning("No contacts to export")
                return None

            if not filename:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = str(OUTPUT_DIR / f"contacts_{timestamp}.xlsx")

            # Create workbook
            wb = Workbook()
            ws = wb.active
            ws.title = "Contacts"

            # Add headers
            ws.append(ExcelExporter.HEADERS)
            header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)

            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            # Add data
            for contact in contacts:
                ws.append([
                    contact.email,
                    contact.name,
                    contact.company,
                    contact.cc_emails,
                    contact.bcc_emails,
                    contact.subject,
                    contact.phone
                ])

            # Auto-adjust column widths
            for column in ws.columns:
                max_length = 0
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except Exception:
                        pass
                ws.column_dimensions[column[0].column_letter].width = min(max_length + 2, 50)

            # Save file
            wb.save(filename)
            logger.info(f"Exported {len(contacts)} contacts to {filename}")
            return filename

        except Exception as e:
            logger.error(f"Failed to export to Excel: {e}")
            return None


# ============================================================================
# MAIN
# ============================================================================

async def main(limit: Optional[int] = None, batch_size: int = BATCH_SIZE):
    """Main execution function"""
    logger.info("=" * 70)
    logger.info("EMAIL CONTACT EXTRACTOR v2")
    logger.info("=" * 70)

    # Validate credentials
    if not EMAIL_ADDRESS or not PASSWORD or not IMAP_SERVER:
        logger.error("Missing credentials in .env file")
        logger.error("Required: EMAIL, PASSWORD, IMAP_SERVER")
        return

    try:
        # Connect to IMAP
        imap = IMAPClient(IMAP_SERVER, IMAP_PORT, EMAIL_ADDRESS, PASSWORD)
        if not imap.connect():
            logger.error("Failed to connect to IMAP server")
            return

        # Get email list
        email_ids = imap.get_email_ids('INBOX')
        if not email_ids:
            logger.warning("No emails found in INBOX")
            imap.disconnect()
            return

        # Apply limit if specified
        if limit:
            email_ids = email_ids[:limit]
            logger.info(f"Processing first {limit} emails")

        # Start extraction
        logger.info(f"Starting extraction of {len(email_ids)} emails")
        start_time = datetime.now()

        extractor = ContactExtractor(imap)
        contacts = await extractor.extract_all(email_ids, batch_size)

        elapsed = datetime.now() - start_time
        extractor.stats.elapsed_seconds = elapsed.total_seconds()

        # Disconnect
        imap.disconnect()

        # Export results
        logger.info("=" * 70)
        logger.info("EXTRACTION COMPLETE")
        logger.info("=" * 70)
        logger.info(f"Statistics: {extractor.stats}")

        if contacts:
            excel_file = ExcelExporter.export(contacts)
            if excel_file:
                logger.info(f"Excel: {excel_file}")
                logger.info(f"Logos: {LOGOS_DIR}")
        else:
            logger.warning("No contacts extracted")

    except Exception as e:
        logger.error(f"Fatal error: {e}", exc_info=True)
    finally:
        logger.info("=" * 70)


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='Extract email contacts')
    parser.add_argument('--limit', type=int, help='Limit number of emails to process')
    parser.add_argument('--batch-size', type=int, default=BATCH_SIZE, help='Emails per batch')

    args = parser.parse_args()

    asyncio.run(main(limit=args.limit, batch_size=args.batch_size))

