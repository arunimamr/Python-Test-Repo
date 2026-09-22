"""
Outlook Email Contact Extractor (pywin32)
Extracts: From, CC, BCC, Name, Mobile, Logo, and signature email
Saves results to Excel file
"""

import win32com.client
import re
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import base64
import os
from datetime import datetime
import pytesseract
from PIL import Image
import io

# Helper functions for parsing

def extract_email_addresses(email_string):
    if not email_string:
        return []
    email_pattern = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b')
    emails = email_pattern.findall(email_string)
    return list(set(emails))

def parse_signature(body_content):
    if not body_content:
        return None, None, None, None
    try:
        soup = BeautifulSoup(body_content, 'html.parser')
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
        text = soup.get_text(separator='\n')
        phone_pattern = re.compile(r'\b(?:\+?\d{1,3}[-.\s]?)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}\b')
        email_pattern = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b')
        phones = phone_pattern.findall(text)
        mobile = phones[0] if phones else None
        emails = email_pattern.findall(text)
        sig_email = emails[0] if emails else None
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        name = None
        for line in lines:
            if not re.search(r'\d|@', line) and len(line.split()) <= 5 and len(line) < 50:
                name = line
                break
        return name, mobile, sig_email, logo_data
    except Exception:
        return None, None, None, None

def download_or_save_logo(logo_data, mail_id, logo_dir='logos'):
    if not logo_data:
        return None
    try:
        os.makedirs(logo_dir, exist_ok=True)
        filename = os.path.join(logo_dir, f"{mail_id}_logo.png")
        if logo_data.startswith('data:image'):
            try:
                header, encoded = logo_data.split(',', 1)
                data = base64.b64decode(encoded)
                with open(filename, 'wb') as f:
                    f.write(data)
                return filename
            except Exception:
                return None
        else:
            return logo_data
    except Exception:
        return None

def extract_emails_from_images(soup, mail_id, logo_dir='logos'):
    """Download all images, run OCR, and extract emails from image text."""
    email_pattern = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b')
    found_emails = set()
    imgs = soup.find_all('img')
    os.makedirs(logo_dir, exist_ok=True)
    for idx, img in enumerate(imgs):
        src = img.get('src', '')
        if src.startswith('data:image'):
            try:
                header, encoded = src.split(',', 1)
                img_bytes = base64.b64decode(encoded)
                img_filename = os.path.join(logo_dir, f"{mail_id}_img_{idx}.png")
                with open(img_filename, 'wb') as f:
                    f.write(img_bytes)
                # OCR
                image = Image.open(io.BytesIO(img_bytes))
                text = pytesseract.image_to_string(image)
                emails = email_pattern.findall(text)
                found_emails.update(emails)
            except Exception as e:
                print(f"[WARN] OCR failed for image {idx}: {e}")
    return list(found_emails)

def save_to_excel(contacts, filename='contacts.xlsx'):
    wb = Workbook()
    ws = wb.active
    ws.title = "Contacts"
    headers = [
        'From Email', 'CC Emails', 'BCC Emails', 'Subject',
        'Name', 'Mobile', 'Email from Signature', 'Logo Path', 'Email from Image OCR'
    ]
    ws.append(headers)
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for contact in contacts:
        ws.append([
            contact['from_email'],
            contact['cc_emails'],
            contact['bcc_emails'],
            contact['subject'],
            contact['name'],
            contact['mobile'],
            contact['sig_email'],
            contact['logo_path'],
            contact.get('ocr_emails', '')
        ])
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
    print(f"[OK] Data saved to {filename}")

def main():
    print("="*60)
    print("OUTLOOK EMAIL CONTACT EXTRACTOR (pywin32)")
    print("="*60)
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
    messages = inbox.Items
    print(f"Total messages in Inbox: {messages.Count}")
    contacts = []
    max_count = 100  # Change as needed
    count = 0
    for message in messages:
        try:
            # Only process MailItem (class 43)
            if hasattr(message, 'Class') and message.Class != 43:
                continue
            mail_id = message.EntryID
            from_email = message.SenderEmailAddress
            cc_emails = extract_email_addresses(message.CC) if hasattr(message, 'CC') else []
            bcc_emails = extract_email_addresses(message.BCC) if hasattr(message, 'BCC') else []
            subject = message.Subject
            print(f"Processing: {subject}")
            body_content = message.HTMLBody if hasattr(message, 'HTMLBody') else message.Body
            name, mobile, sig_email, logo_data = parse_signature(body_content)
            logo_path = download_or_save_logo(logo_data, mail_id)
            # OCR for emails in images
            soup = BeautifulSoup(body_content, 'html.parser')
            ocr_emails = extract_emails_from_images(soup, mail_id)
            contacts.append({
                'from_email': from_email,
                'cc_emails': ', '.join(cc_emails),
                'bcc_emails': ', '.join(bcc_emails),
                'subject': subject,
                'name': name,
                'mobile': mobile,
                'sig_email': sig_email,
                'logo_path': logo_path,
                'ocr_emails': ', '.join(ocr_emails)
            })
            count += 1
            if count >= max_count:
                break
        except Exception as e:
            print(f"[WARN] Failed to process message: {e}")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f'contacts_outlook_{timestamp}.xlsx'
    save_to_excel(contacts, filename)
    print(f"[OK] Extracted {len(contacts)} emails.")

if __name__ == "__main__":
    main()
