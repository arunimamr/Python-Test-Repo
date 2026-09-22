# Turbify Email Contact Extractor

A Python-based asynchronous email extraction tool specifically designed for Turbify business email accounts.

## Features

✅ **Asynchronous Email Processing** - Fetches emails in batches for optimal performance  
✅ **Contact Information Extraction** - Extracts From, CC, BCC, subject, and signature details  
✅ **Advanced Parsing** - Extracts name, mobile number, and email from signatures  
✅ **Logo Extraction** - Downloads and saves company logos from email signatures  
✅ **Excel Export** - Saves all extracted data in a formatted Excel file  
✅ **Error Handling** - Comprehensive logging and error recovery  
✅ **Batch Processing** - Configurable batch sizes for stable performance  

---

## Turbify Email Accessibility

### Available Connection Methods

#### 1. **IMAP Protocol (RECOMMENDED)** ✅
- **Server**: `abluae.com` (for @abluae.com email addresses)
- **Port**: 993 (SSL/TLS)
- **Authentication**: Email address + Password
- **Status**: ✅ Verified working
- **Best for**: Batch processing and automation

#### 2. **Alternative Turbify Servers**
If you have a different Turbify domain:
- Test with: `mail.turbify.com` (may timeout in some regions)
- Use the diagnostic tool to find the correct server

#### 3. **POP3 Protocol** (Not recommended)
- **Server**: `pop.abluae.com` (if available)
- **Port**: 995 (SSL/TLS)
- **Limitation**: Download-only, doesn't maintain sync

#### 4. **Webmail Interface**
- Access directly via browser (not suitable for automation)
- Some Turbify instances support Office 365/Outlook web access

---

## Installation

### Step 1: Install Python 3.8+
Ensure you have Python 3.8 or higher installed.

### Step 2: Create Virtual Environment
```bash
python -m venv .venv
```

### Step 3: Activate Virtual Environment
**Windows:**
```bash
.\.venv\Scripts\Activate.ps1
```

**Linux/Mac:**
```bash
source .venv/bin/activate
```

### Step 4: Install Dependencies
```bash
pip install python-dotenv openpyxl beautifulsoup4
```

---

## Configuration

### Step 1: Create `.env` File
Create a `.env` file in the project directory with your credentials:

```dotenv
# Turbify Email Configuration
EMAIL=arunima@abluae.com
PASSWORD=dummy-password
IMAP_SERVER=abluae.com
IMAP_PORT=993
BATCH_SIZE=50
```

### Step 2: Security Best Practices

⚠️ **Important**: Never commit `.env` to version control!

Add to `.gitignore`:
```
.env
*.log
logos/
contacts_*.xlsx
```

**For Enhanced Security:**
- Some Turbify instances support App-Specific Passwords
- Contact your IT administrator for password policies
- Use a dedicated account for automation if possible

---

## Usage

### Run the Email Extractor

**Enhanced Version (Recommended):**
```bash
python General\extractContacts_Turbify.py
```

**Original Version:**
```bash
python General\extractContacts.py
```

### Output Files

- **`contacts_YYYYMMDD_HHMMSS.xlsx`** - Extracted contact data
- **`logos/`** - Directory containing downloaded company logos
- **`email_extraction.log`** - Detailed execution logs

### Excel Output Structure

| From Email | CC Emails | BCC Emails | Subject | Name | Mobile | Email from Signature | Logo Path |
|-----------|-----------|-----------|---------|------|--------|----------------------|-----------|
| sender@abluae.com | cc1@example.com | - | Meeting | John Doe | +971-12-345-6789 | john@abluae.com | logos/123_logo.png |

---

## Troubleshooting

### Issue 1: Connection Timeout
**Error**: `TimeoutError: A connection attempt failed`

**Solution**: 
- Your network/firewall may be blocking port 993
- Run the diagnostic tool: `python turbify_diagnostic.py`
- Contact IT for firewall configuration

### Issue 2: Authentication Failed
**Error**: `IMAP Authentication Error`

**Solution**:
- Verify email and password are correct
- Check if account uses 2FA (may need app password)
- Ensure account has IMAP access enabled

### Issue 3: No Emails Found
**Error**: `No emails found in INBOX`

**Solution**:
- Check if INBOX folder exists
- Verify email account has received emails
- Try accessing another folder in configuration

### Issue 4: Signature Parsing Issues
**Error**: Name/mobile not extracted correctly

**Solution**:
- Different email clients format signatures differently
- Manual review may be needed for unusual formats
- Check `email_extraction.log` for details

---

## Running the Diagnostic Tool

To identify the correct Turbify server for your domain:

```bash
python turbify_diagnostic.py
```

**Output**: Recommended IMAP server configuration

---

## Advanced Configuration

### Adjust Batch Size
```dotenv
BATCH_SIZE=100  # Larger batches = faster but more memory
BATCH_SIZE=10   # Smaller batches = slower but more stable
```

**Recommended**: 
- 50 for standard performance
- 100+ for large mailboxes (>10,000 emails)
- 10-25 for slower connections

### Process Multiple Folders
Edit the script to change folder:
```python
email_ids = extractor.get_email_list('Sent')  # or 'Drafts', 'Archive', etc.
```

### Customize Excel Output
Edit `save_to_excel()` method to add more columns or formatting.

---

## Performance Notes

- **Processing Speed**: ~100-200 emails/minute depending on:
  - Email size (large attachments slow down processing)
  - Network speed
  - Batch size
  
- **Memory Usage**: Minimal (~50-100MB for typical batches)

- **Best Practices**:
  - Run during off-peak hours for large extractions
  - Start with smaller batch sizes if experiencing timeouts
  - Check logs regularly for errors

---

## Turbify Email Features Supported

✅ IMAP Folder Navigation  
✅ SSL/TLS Encryption  
✅ Multipart Email Parsing  
✅ HTML Signature Extraction  
✅ Base64-Encoded Images  
✅ Character Encoding Handling  

❌ OAuth/SSO (Use password authentication)  
❌ Gmail-style Labels  
❌ Custom IMAP Extensions  

---

## Common Turbify Domains

| Domain | Email Format | IMAP Server | Status |
|--------|-------------|-----------|--------|
| abluae.com | user@abluae.com | abluae.com:993 | ✅ Verified |
| turbify.com | user@turbify.com | mail.turbify.com:993 | ⚠️ May timeout |
| Other | user@domain.com | mail.domain.com:993 | Test required |

---

## Support & Documentation

- **Turbify Official**: Contact your IT administrator
- **Python IMAP**: https://docs.python.org/3/library/imaplib.html
- **BeautifulSoup**: https://www.crummy.com/software/BeautifulSoup/
- **OpenPyXL**: https://openpyxl.readthedocs.io/

---

## License

This tool is provided as-is for business email automation.

---

## Security Disclaimer

⚠️ **Important**:
- Store credentials securely
- Never share `.env` files
- Use dedicated service accounts when possible
- Review logs for unauthorized access attempts
- Comply with your organization's email policies

---

## Version History

- **v2.0** (Enhanced Turbify Edition)
  - Verified IMAP server configuration for Turbify
  - Improved error handling and logging
  - Better Excel formatting
  - Diagnostic tool included

- **v1.0** (Original)
  - Basic IMAP email extraction
  - Signature parsing
  - Excel export

---

## Contact

For issues or questions about Turbify email access:
1. Run `turbify_diagnostic.py` to identify the server
2. Check `email_extraction.log` for detailed errors
3. Contact your IT administrator for Turbify-specific support

