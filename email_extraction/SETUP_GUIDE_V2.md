# Email Contact Extractor v2 - Complete Setup Guide

## Quick Start (3 Steps)

### Step 1: Run Configuration Wizard

```bash
python configure_provider.py
```

This will:
- ✓ List available email providers
- ✓ Guide you through configuration
- ✓ Test your credentials
- ✓ Save to `.env` file

### Step 2: Verify Connection

```bash
python test_imap_connection_v2.py
```

This should show:
```
✓ SSL connection established
✓ Authentication successful
✓ INBOX has XXX emails
✓ Connection test successful!
```

### Step 3: Extract Contacts

```bash
# Process all emails
python email_extractor_v2.py

# Or process first 10 emails to test
python email_extractor_v2.py --limit 10
```

---

## Provider-Specific Setup

### Gmail

**Steps:**
1. Enable 2-Factor Authentication (2FA)
2. Go to https://myaccount.google.com/apppasswords
3. Select "Mail" and "Windows Computer"
4. Generate app password
5. Copy the 16-character password

**Configuration:**
```env
EMAIL=your.email@gmail.com
PASSWORD=dummy-password
IMAP_SERVER=imap.gmail.com
IMAP_PORT=993
```

**Note:** Use the app password, NOT your Gmail password.

---

### Microsoft Outlook / Hotmail

**Steps:**
1. Check if 2FA is enabled: https://account.microsoft.com/security-settings
2. If 2FA enabled:
   - Go to https://account.microsoft.com/security-settings/app-passwords
   - Generate app password for "Mail" on "Windows"
   - Copy the password

**Configuration:**
```env
EMAIL=your.email@outlook.com
PASSWORD=dummy-password
IMAP_SERVER=outlook.office365.com
IMAP_PORT=993
```

---

### Yahoo Mail / Turbify

**Steps:**
1. Go to https://account.yahoo.com/account/security
2. Click "Generate app password"
3. Select "Other App" and enter "Email"
4. Copy the 16-character password

**Configuration:**
```env
EMAIL=your.email@yahoo.com
PASSWORD=dummy-password
IMAP_SERVER=imap.mail.yahoo.com
IMAP_PORT=993
```

**For Turbify Specifically:**
- Use your Turbify email (usually @turbify.com or custom domain)
- Use the app password generated above

Example:
```env
EMAIL=abl.quotations@abluae.com
PASSWORD=dummy-password
IMAP_SERVER=imap.mail.yahoo.com
IMAP_PORT=993
```

---

### AOL Mail

**Steps:**
1. Go to https://account.aol.com/security
2. Click "Generate app password"
3. Select your app and device
4. Copy the password

**Configuration:**
```env
EMAIL=your.email@aol.com
PASSWORD=dummy-password
IMAP_SERVER=imap.aol.com
IMAP_PORT=993
```

---

### Office 365 (Corporate)

**Steps:**
1. Ensure IMAP is enabled: https://outlook.office365.com -> Settings
2. If 2FA enabled:
   - Go to https://account.activedirectory.windowsazure.com/r/#/profile
   - Generate app password
3. Copy the password

**Configuration:**
```env
EMAIL=firstname.lastname@company.onmicrosoft.com
PASSWORD=dummy-password
IMAP_SERVER=outlook.office365.com
IMAP_PORT=993
```

---

## Configuration Files

### .env (Required)

Location: `D:\Projects\PythonPractice\.env`

**Minimal configuration:**
```env
EMAIL=your-email@domain.com
PASSWORD=dummy-password
IMAP_SERVER=imap.mail.yahoo.com
IMAP_PORT=993
```

**Full configuration:**
```env
# Email Configuration
EMAIL=your-email@domain.com
PASSWORD=dummy-password
IMAP_SERVER=imap.mail.yahoo.com
IMAP_PORT=993
BATCH_SIZE=20

# Optional: Client Sharepoint Configuration
DEMO_CLIENT_SHAREPOINT_SITE_URL=https://company.sharepoint.com/sites/Docs
DEMO_CLIENT_SHAREPOINT_SITE_HOST=company.sharepoint.com
DEMO_CLIENT_SHAREPOINT_SITE_PATH=/sites/Docs
DEMO_CLIENT_SHAREPOINT_CLIENT_ID=00000000-0000-0000-0000-000000000000
DEMO_CLIENT_SHAREPOINT_CLIENT_SECRET=dummy-client-secret
DEMO_CLIENT_SHAREPOINT_TENANT_ID=00000000-0000-0000-0000-000000000000
DEMO_CLIENT_SHAREPOINT_LIBRARY_NAME=Documents
```

---

## Troubleshooting

### Authentication Error: "Invalid credentials"

**Possible causes:**
1. ❌ Wrong email or password
2. ❌ Using regular password instead of app password (Gmail, Yahoo, Outlook)
3. ❌ Account has 2FA enabled but not using app password
4. ❌ IMAP not enabled in account settings

**Solutions:**
- Verify email is exactly as it appears in account settings
- Ensure you're using an app password if 2FA is enabled
- Check email provider's account security settings for IMAP access
- Use `configure_provider.py` to re-verify setup

### Connection Error: "SSL: CERTIFICATE_VERIFY_FAILED"

**Possible causes:**
1. ❌ Old Python version without updated SSL certificates
2. ❌ Antivirus/proxy intercepting SSL

**Solutions:**
```bash
# Update certificates (macOS/Linux)
/Applications/Python\ 3.x/Install\ Certificates.command

# For Windows, update Python
python -m pip install --upgrade certifi
```

### Connection Timeout: "socket error: timed out"

**Possible causes:**
1. ❌ Network connectivity issue
2. ❌ Batch size too large (causing server to slow down)
3. ❌ Too many concurrent connections

**Solutions:**
1. Test basic connectivity:
   ```bash
   python test_imap_connection_v2.py
   ```

2. Reduce batch size in `.env`:
   ```env
   BATCH_SIZE=10
   ```

3. Increase timeout in code:
   ```python
   # Edit email_extractor_v2.py, line ~550
   timeout=60  # Increase from 30
   ```

### "IMAP server not found"

**Possible causes:**
1. ❌ Wrong IMAP server address
2. ❌ ISP blocking port 993
3. ❌ Firewall blocking connection

**Solutions:**
- Verify IMAP server using: `python test_imap_connection_v2.py`
- Check provider's official IMAP documentation
- Try alternative port (usually 143 for non-SSL, but NOT RECOMMENDED)
- Check firewall/ISP settings

### No Emails Extracted

**Check in this order:**
1. Verify connection:
   ```bash
   python test_imap_connection_v2.py
   ```
   
   Should show email count in INBOX

2. Check logs:
   ```
   email_extraction_v2.log
   ```
   
   Look for ERROR or FAIL lines

3. Test with limit:
   ```bash
   python email_extractor_v2.py --limit 1
   ```
   
   Process single email to see detailed errors

4. Check email access:
   - Open email client and verify INBOX is accessible
   - Ensure folder name is exactly "INBOX" (case-sensitive on some servers)

---

## Performance Optimization

### Default Settings
- **Batch Size:** 20 emails
- **Workers:** 1 (per batch)
- **Delay between batches:** 2 seconds

### For Faster Processing

If connection is stable, increase batch size:

```bash
python email_extractor_v2.py --batch-size 50
```

Or edit `.env`:
```env
BATCH_SIZE=50
```

**Note:** Larger batches may cause SSL/timeout errors. Start with 20-30.

### For Slower/Unstable Connections

Decrease batch size:

```bash
python email_extractor_v2.py --batch-size 5
```

---

## Output Files

### Excel File
- **Location:** `output/contacts_YYYYMMDD_HHMMSS.xlsx`
- **Contains:** Email, Name, Company, CC, BCC, Subject, Phone
- **Format:** Headers in blue with white text, auto-width columns

Example output:
```
Email                    Name              Company    CC              BCC    Subject           Phone
john@amazon.com         John Smith        amazon     jane@amazon.com         Project Update   +1-206-555-0100
```

### Logos Directory
- **Location:** `output/logos/`
- **Format:** Automatically downloaded and deduplicated
- **Naming:** By company name (e.g., `amazon.png`, `google.jpg`)
- **Duplicate handling:** MD5-based, only one copy per unique image

### Logs
- **Location:** `email_extraction_v2.log`
- **Format:** Timestamp - Level - Message
- **Retention:** All runs are appended

---

## Command Line Reference

### Basic Usage

```bash
# Process all emails with default batch size (20)
python email_extractor_v2.py

# Process first 100 emails
python email_extractor_v2.py --limit 100

# Process all emails with batch size of 10
python email_extractor_v2.py --batch-size 10

# Process first 50 emails with batch size of 5
python email_extractor_v2.py --limit 50 --batch-size 5
```

### PowerShell (Windows)

```powershell
# Process all emails
./run_extractor_v2.ps1

# Process first 100 emails
./run_extractor_v2.ps1 -limit 100 -batch_size 10
```

### Batch File (Windows)

```cmd
# Process all emails
run_extractor_v2.bat

# Process first 100 emails with batch size 10
run_extractor_v2.bat 100 10
```

---

## API Usage (Advanced)

```python
import asyncio
from email_extractor_v2 import IMAPClient, ContactExtractor, ExcelExporter

async def main():
    # Create IMAP client
    imap = IMAPClient(
        server='imap.mail.yahoo.com',
        port=993,
        email='your-email@yahoo.com',
        password='your-password'
    )
    
    # Connect
    if not imap.connect():
        return
    
    # Get email IDs
    email_ids = imap.get_email_ids('INBOX')
    
    # Extract contacts
    extractor = ContactExtractor(imap)
    contacts = await extractor.extract_all(email_ids, batch_size=20)
    
    # Export to Excel
    ExcelExporter.export(contacts, 'my_contacts.xlsx')
    
    # Disconnect
    imap.disconnect()

if __name__ == '__main__':
    asyncio.run(main())
```

---

## Support Resources

- **Gmail Help:** https://support.google.com/mail/answer/7126229
- **Outlook Help:** https://support.microsoft.com/outlook
- **Yahoo Help:** https://help.yahoo.com/kb
- **Office 365 Help:** https://support.microsoft.com/office365

---

## Version History

- **v2.0** - Complete rewrite with OOP architecture
- **v1.0** - Original Turbify-specific implementation

---

## FAQ

**Q: Can I use my regular password?**
A: Only if 2FA is disabled. If 2FA is enabled, you MUST use an app password.

**Q: How many emails can I process?**
A: Tested up to 1,000+ emails. Performance depends on email size and connection speed.

**Q: Does it delete emails?**
A: No. It only reads emails. All messages remain in your inbox.

**Q: Can I process multiple folders?**
A: Currently only INBOX. Future versions may support multiple folders.

**Q: Are credentials stored securely?**
A: Credentials are stored in `.env` file. Make sure to:
- Add `.env` to `.gitignore`
- Never commit `.env` to version control
- Use app passwords instead of real passwords when possible

**Q: How do I handle international phone numbers?**
A: The regex pattern supports basic international formats. Check `EmailParser.parse_signature()` for custom patterns.

**Q: Can I customize which columns are exported?**
A: Yes, edit the `Contact` dataclass and `ExcelExporter.HEADERS` in `email_extractor_v2.py`

---

For more help, check `email_extraction_v2.log` for detailed error messages.

