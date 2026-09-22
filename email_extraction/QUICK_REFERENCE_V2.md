# Email Extractor v2 - Quick Reference Card

## Installation & Setup

```bash
# Step 1: Create virtual environment (optional)
python -m venv venv
venv\Scripts\Activate.ps1

# Step 2: Install dependencies
pip install -r requirements.txt

# Step 3: Configure email provider
python configure_provider.py

# Step 4: Test connection
python test_imap_connection_v2.py
```

## Usage

### Run Extraction

```bash
# All emails
python email_extractor_v2.py

# Limit to first 10
python email_extractor_v2.py --limit 10

# Custom batch size
python email_extractor_v2.py --batch-size 5
```

### Windows Scripts

```powershell
# PowerShell
./run_extractor_v2.ps1

# Command Prompt
run_extractor_v2.bat
```

## Configuration

### .env File (Required)

```env
EMAIL=your-email@domain.com
PASSWORD=dummy-password
IMAP_SERVER=imap.mail.yahoo.com
IMAP_PORT=993
BATCH_SIZE=20
```

### Provider Settings

| Provider | Server | Port | App Password? |
|----------|--------|------|---------------|
| Gmail | imap.gmail.com | 993 | YES (if 2FA) |
| Outlook | outlook.office365.com | 993 | YES (if 2FA) |
| Yahoo | imap.mail.yahoo.com | 993 | YES |
| Turbify | imap.mail.yahoo.com | 993 | YES |
| AOL | imap.aol.com | 993 | YES |

## Output

### Excel File
- Location: `output/contacts_YYYYMMDD_HHMMSS.xlsx`
- Columns: Email, Name, Company, CC Emails, BCC Emails, Subject, Phone

### Logos
- Location: `output/logos/`
- Format: Named by company (amazon.png, google.jpg, etc.)
- Duplicates: Removed via MD5 hash

### Logs
- Location: `email_extraction_v2.log`
- Contains: All extraction events, errors, warnings

## Troubleshooting

### Connection Issues

**Error:** IMAP Authentication Error  
**Fix:** Check credentials, ensure using app password (not regular password) if 2FA enabled

**Error:** Connection Timeout  
**Fix:** Reduce `BATCH_SIZE` in .env to 5-10

**Error:** SSL Certificate Verify Failed  
**Fix:** `python -m pip install --upgrade certifi`

### Extraction Issues

**No Emails Found:**
1. Run `python test_imap_connection_v2.py`
2. Check if INBOX has emails
3. Verify with `--limit 1` to test single email

**Logos Not Downloaded:**
- Check network connectivity
- Verify URLs are accessible
- Check `output/logos/` directory

## Key Features

- ✅ Async/concurrent processing
- ✅ Batch support with configurable sizes
- ✅ Auto-reconnection on failures
- ✅ Image deduplication via MD5
- ✅ Comprehensive error logging
- ✅ Excel export with formatting
- ✅ Works with Gmail, Outlook, Yahoo, Office 365, etc.

## Common Commands

```bash
# Test connection
python test_imap_connection_v2.py

# Configure email provider
python configure_provider.py

# Extract 10 emails with batch size 5
python email_extractor_v2.py --limit 10 --batch-size 5

# Check logs
tail -f email_extraction_v2.log

# View extracted emails
start output\contacts_*.xlsx
```

## Tips

1. **First Run:** Use `--limit 10` to test configuration
2. **Large Inbox:** Process in chunks using `--limit` parameter
3. **Slow Connection:** Reduce `--batch-size` to 5
4. **Multiple Accounts:** Create separate .env files or run multiple times
5. **App Passwords:** Always use app passwords with 2FA (more secure)

## Environment Variables Reference

| Variable | Required | Default | Example |
|----------|----------|---------|---------|
| EMAIL | Yes | - | user@domain.com |
| PASSWORD | Yes | - | app-password |
| IMAP_SERVER | Yes | - | imap.mail.yahoo.com |
| IMAP_PORT | No | 993 | 993 |
| BATCH_SIZE | No | 20 | 20 |

## Command Line Options

```
python email_extractor_v2.py [OPTIONS]

Options:
  --limit N          Process only first N emails
  --batch-size N     Emails per batch (default: 20)
  --help             Show help message
```

## Data Extracted per Email

- **From:** Email sender address
- **Name:** Sender name (from header or signature)
- **Company:** Extracted from email domain
- **CC:** All CC recipients
- **BCC:** All BCC recipients  
- **Subject:** Email subject line
- **Phone:** From signature (if found)
- **Logo:** Downloaded and deduplicated

## Performance Benchmarks

| Scenario | Time | Notes |
|----------|------|-------|
| Connection | 2-3s | Initial SSL handshake |
| Per Email | 1-2s | Includes retries |
| 100 emails | 2-3 min | Batch size 20 |
| 1000 emails | 20-30 min | Batch size 20 |

## Files Generated

```
output/
├── contacts_20260504_150000.xlsx    # Main Excel file
└── logos/
    ├── amazon.png
    ├── google.jpg
    └── company.gif
    
email_extraction_v2.log             # Log file
```

## Getting Help

1. Check `email_extraction_v2.log` for errors
2. Run `python test_imap_connection_v2.py`
3. Re-run `python configure_provider.py` to update settings
4. Check email provider's IMAP documentation

## Version

Current: **v2.0**

Last Updated: 2026-05-04

---

For detailed information, see:
- `README_V2.md` - Full documentation
- `SETUP_GUIDE_V2.md` - Detailed setup for each provider
- `email_extraction_v2.log` - Debug logs

