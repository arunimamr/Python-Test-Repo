# Email Contact Extractor v2

A production-ready asynchronous Python tool for extracting contact information from Outlook/Turbify business emails.

## Features

✅ **Asynchronous Processing** - Non-blocking async/await pattern for fast batch processing  
✅ **Batch Processing** - Configurable batch sizes with per-batch delays  
✅ **Email Extraction** - From, CC, BCC, Subject  
✅ **Signature Parsing** - Name, phone number, email, company  
✅ **Logo Download** - Automatic logo extraction from emails  
✅ **Deduplication** - MD5-based image deduplication  
✅ **Resilient Connections** - Automatic reconnection with exponential backoff  
✅ **Excel Export** - Formatted output with headers and auto-width columns  
✅ **Comprehensive Logging** - Detailed logs to file and console  

## Installation

### 1. Create Virtual Environment (optional but recommended)

```powershell
python -m venv venv
venv\Scripts\Activate.ps1
```

### 2. Install Dependencies

```bash
pip install -r requirements.txt
```

### 3. Configure Environment

Create/update `.env` file in the project root:

```env
# Email Configuration
EMAIL=your-email@company.com
PASSWORD=dummy-password
IMAP_SERVER=imap.mail.yahoo.com
IMAP_PORT=993
BATCH_SIZE=20
```

### Configuration Examples

**For Turbify/Yahoo Business Email:**
```env
IMAP_SERVER=imap.mail.yahoo.com
IMAP_PORT=993
```

**For Gmail:**
```env
IMAP_SERVER=imap.gmail.com
IMAP_PORT=993
# Note: Use app password, not regular password
```

**For Microsoft 365:**
```env
IMAP_SERVER=outlook.office365.com
IMAP_PORT=993
```

## Usage

### Basic Usage

```bash
python email_extractor_v2.py
```

### With Options

```bash
# Process only first 10 emails
python email_extractor_v2.py --limit 10

# Process with batch size of 5
python email_extractor_v2.py --batch-size 5

# Combined
python email_extractor_v2.py --limit 100 --batch-size 10
```

### PowerShell Script (Windows)

```powershell
./run_extractor.ps1
```

## Output

### Files Generated

- **Excel Sheet**: `output/contacts_YYYYMMDD_HHMMSS.xlsx`
  - Contains: Email, Name, Company, CC, BCC, Subject, Phone
  
- **Logos**: `output/logos/`
  - Automatically downloaded company logos
  - Deduplicated by MD5 hash
  - Named by company (e.g., `amazon.png`, `google.png`)

- **Logs**: `email_extraction_v2.log`
  - Detailed extraction logs
  - Error messages and warnings

### Excel Columns

| Column | Description |
|--------|-------------|
| Email | Sender's email address |
| Name | Sender's name (from From header or signature) |
| Company | Domain name extracted from email |
| CC Emails | All CC recipients (comma-separated) |
| BCC Emails | All BCC recipients (comma-separated) |
| Subject | Email subject line |
| Phone | Phone number from signature (if found) |

## Architecture

### Classes

- **`IMAPClient`** - IMAP connection management with reconnection logic
- **`EmailParser`** - Email parsing and signature extraction
- **`ImageHandler`** - Logo download and deduplication
- **`ContactExtractor`** - Main extraction engine
- **`ExcelExporter`** - Excel file generation
- **`Contact`** - Data model for extracted contact

### Key Features

#### Connection Resilience
- Automatic reconnection on SSL/timeout errors
- Exponential backoff retry strategy (2s → 6s)
- Per-batch delays to prevent connection pool exhaustion

#### Signature Parsing
- Regex patterns for phone numbers (multiple formats)
- Email extraction from signature
- Name detection (first non-numeric line)
- Logo priority: embedded base64 > logo URLs > company logos > first external URL

#### Image Deduplication
- MD5 hash-based deduplication
- Company-name-based filenames (fallback to hash)
- Collision handling with counter suffix

#### Async Processing
- ThreadPoolExecutor with configurable workers (default: 1)
- asyncio.gather for concurrent processing
- Per-batch delays (2 seconds)

## Troubleshooting

### Authentication Errors

**Error**: `IMAP authentication failed`

**Solutions**:
1. Verify email and password in `.env`
2. For Gmail: Use [app password](https://myaccount.google.com/apppasswords), not regular password
3. For Office 365: Check 2FA, may need app password
4. For Turbify: Ensure IMAP is enabled in account settings

### SSL/TLS Errors

**Error**: `SSL: BAD_LENGTH` or `TLS/SSL connection closed`

**Solutions**:
1. These are automatically handled with retries
2. Reduce `BATCH_SIZE` if too many occur
3. Increase timeout (in `IMAPClient.fetch_email`)

### Connection Timeouts

**Error**: `socket error: timed out`

**Solutions**:
1. Reduce `BATCH_SIZE`
2. Increase delay between batches (modify `email_extractor_v2.py` line ~560)
3. Check network connectivity
4. Reduce IMAP_PORT request count (wait longer between batches)

### No Emails Extracted

1. Check logs in `email_extraction_v2.log`
2. Verify INBOX is selected and has emails
3. Test connection with `test_imap_connection.py`
4. Try with `--limit 1` to test single email

## Performance

Typical performance on Yahoo/Turbify:

- **Connection**: 2-3 seconds
- **Per Email**: 1-2 seconds (includes retries for some)
- **100 emails**: ~2 minutes
- **Batch size**: 20 (recommended for Yahoo)

To improve performance:
- Increase `--batch-size` (may cause SSL errors)
- Process fewer emails with `--limit`
- Run multiple instances on different folders

## Development

### Adding Custom Fields

Edit `Contact` dataclass in `email_extractor_v2.py`:

```python
@dataclass
class Contact:
    email: str
    name: str = ''
    # Add new field:
    custom_field: str = ''
```

### Changing Logo Priority

Edit `EmailParser.parse_signature()` method, logo selection section.

### Adding Email Provider Support

Update `.env` with provider-specific IMAP settings.

## Logging

All logs go to both:
1. Console (real-time monitoring)
2. `email_extraction_v2.log` (persistent record)

Log levels:
- `INFO` - General information, extraction progress
- `WARNING` - Non-fatal issues (failed logo download, decode errors)
- `DEBUG` - Detailed diagnostic information
- `ERROR` - Fatal errors

## Limitations

- IMAP-only (no direct Outlook API support)
- Single connection per run (not concurrent connections)
- Batch processing is sequential (configurable via batch_size)
- Logo extraction only from HTML signatures
- Phone regex may not match all international formats

## Future Improvements

- Multi-connection concurrency
- GUI interface
- Database export (SQLite, PostgreSQL)
- Microsoft Graph API integration
- Advanced phone number parsing
- OCR for images containing contact info

## License

[Your License Here]

## Support

For issues or questions:
1. Check logs: `email_extraction_v2.log`
2. Verify `.env` configuration
3. Test IMAP connection: `python test_imap_connection.py`
4. Check email provider documentation for IMAP settings

## Version History

### v2.0 (Current)
- Restructured with OOP design
- Better error handling and logging
- Comprehensive documentation
- Configurable CLI options

### v1.0 (Original)
- Basic email extraction
- Turbify-specific implementation

