# 📑 PythonPractice Project Index

Quick navigation guide for all projects and their files.

---

## 🎯 Main Projects

### 1️⃣ Email Contact Extractor v2 ⭐ **NEW**

**Status:** ✅ Production Ready

**Location:** `email_extraction/`

**Getting Started:**
```bash
cd email_extraction
python configure_provider.py      # Setup wizard
python test_imap_connection_v2.py # Test connection
python email_extractor_v2.py --limit 10  # Try first 10 emails
```

**Key Files:**
| File | Purpose |
|------|---------|
| `email_extractor_v2.py` | Main application (570 lines) |
| `configure_provider.py` | Interactive setup wizard |
| `test_imap_connection_v2.py` | Test IMAP connection |
| `README_V2.md` | 📖 Full documentation |
| `SETUP_GUIDE_V2.md` | 📖 Provider-specific setup |
| `QUICK_REFERENCE_V2.md` | 📖 Command reference |
| `PROJECT_SUMMARY.md` | 📖 Architecture & features |

**Features:**
- ✅ Async batch processing
- ✅ 7+ email providers (Gmail, Outlook, Yahoo, Turbify, etc.)
- ✅ Logo extraction and deduplication
- ✅ Excel export with formatting
- ✅ Comprehensive error handling
- ✅ Connection resilience

**Output:**
```
output/
├── contacts_YYYYMMDD_HHMMSS.xlsx
└── logos/                      (company logos, deduplicated)
```

---

### 2️⃣ SharePoint Folder Importer

**Status:** ✅ Fixed (TENANT_ID error resolved)

**Location:** `sharepoint/`

**Files:**
| File | Purpose |
|------|---------|
| `import_folder_structure.py` | ⭐ Fixed script |
| `FIX_SUMMARY.md` | 📖 Details of fix |

**Usage:**
```bash
cd sharepoint
python import_folder_structure.py
```

**Output:**
```
sharepoint_tree.json            (JSON folder structure)
```

---

## 📚 Documentation

### Main Directory

| File | Purpose |
|------|---------|
| `PROJECT_STATUS.md` | 📋 Overall project status (THIS IS YOUR START POINT) |
| `INDEX.md` | 📑 This file - navigation guide |
| `.env` | ⚙️ Configuration (email & SharePoint settings) |
| `requirements.txt` | 📦 Python dependencies |

### Email Extraction (`email_extraction/`)

| File | Purpose | Audience |
|------|---------|----------|
| `README_V2.md` | Full documentation | Developers |
| `SETUP_GUIDE_V2.md` | Provider-specific setup | End users |
| `QUICK_REFERENCE_V2.md` | Command cheat sheet | Power users |
| `PROJECT_SUMMARY.md` | Architecture & classes | Developers |
| `email_extraction_v2.log` | Extraction logs | Debugging |

### SharePoint (`sharepoint/`)

| File | Purpose |
|------|---------|
| `FIX_SUMMARY.md` | Fix documentation |

---

## 🚀 Quick Start Guides

### For Email Extraction (First Time)

**Step 1: Configure**
```bash
cd email_extraction
python configure_provider.py
# Follow the wizard:
# 1. Select email provider
# 2. Enter email and password
# 3. Test connection
# 4. Save configuration
```

**Step 2: Verify**
```bash
python test_imap_connection_v2.py
# Should show:
# ✓ SSL connection established
# ✓ Authentication successful
# ✓ INBOX has XXX emails
```

**Step 3: Extract**
```bash
# Test with 10 emails
python email_extractor_v2.py --limit 10

# Or extract all
python email_extractor_v2.py
```

**Step 4: View Results**
```bash
# Windows
start output\contacts_*.xlsx
dir output\logos\

# PowerShell
explorer.exe output\
```

### For Windows Users

```powershell
# PowerShell launcher
cd email_extraction
./run_extractor_v2.ps1

# Or use batch file
run_extractor_v2.bat
```

### For SharePoint

```bash
cd sharepoint
python import_folder_structure.py
# Output: sharepoint_tree.json
```

---

## 📖 Reading Order

### For End Users

1. **`PROJECT_STATUS.md`** - Overview of what's available
2. **`email_extraction/SETUP_GUIDE_V2.md`** - Setup for your email provider
3. **`email_extraction/QUICK_REFERENCE_V2.md`** - Command reference

### For Developers

1. **`PROJECT_STATUS.md`** - Overview
2. **`email_extraction/PROJECT_SUMMARY.md`** - Architecture
3. **`email_extraction/README_V2.md`** - Features & implementation
4. **`email_extraction/email_extractor_v2.py`** - Source code

### For Troubleshooting

1. **`email_extraction_v2.log`** - Check for errors
2. **`email_extraction/SETUP_GUIDE_V2.md`** - Provider-specific help (Troubleshooting section)
3. **`email_extraction/README_V2.md`** - Performance/scaling issues

---

## 🔧 Configuration

### .env File Location
`D:\Projects\PythonPractice\.env`

### Minimal Configuration
```env
EMAIL=your-email@domain.com
PASSWORD=dummy-password
IMAP_SERVER=imap.mail.yahoo.com
IMAP_PORT=993
```

### Provider Servers
- **Gmail:** imap.gmail.com
- **Outlook:** outlook.office365.com
- **Yahoo:** imap.mail.yahoo.com
- **Turbify:** imap.mail.yahoo.com
- **AOL:** imap.aol.com

### Important Notes
- Use **app password** if 2FA is enabled
- For Gmail: Generate at https://myaccount.google.com/apppasswords
- For Outlook: Generate at https://account.microsoft.com/security-settings/app-passwords
- For Yahoo/Turbify: Generate at https://account.yahoo.com/account/security

---

## 📂 File Structure Reference

```
D:\Projects\PythonPractice/
│
├── .env                             ⚙️ Configuration
├── requirements.txt                 📦 Dependencies
├── PROJECT_STATUS.md                📋 Main status (START HERE)
├── INDEX.md                         📑 This file
│
├── email_extraction/                ⭐ Email Extractor v2
│   ├── email_extractor_v2.py       Main app
│   ├── configure_provider.py        Setup wizard
│   ├── test_imap_connection_v2.py  Connection tester
│   ├── run_extractor_v2.ps1        PowerShell launcher
│   ├── run_extractor_v2.bat        Batch launcher
│   ├── README_V2.md                Full docs
│   ├── SETUP_GUIDE_V2.md           Setup guide
│   ├── QUICK_REFERENCE_V2.md       Quick reference
│   ├── PROJECT_SUMMARY.md          Architecture
│   ├── email_extraction_v2.log     Logs
│   ├── output/                     Generated files
│   │   ├── contacts_*.xlsx         Extracted data
│   │   └── logos/                  Downloaded logos
│   │
│   └── [older v1 files]
│
├── sharepoint/                      ✅ SharePoint Importer (Fixed)
│   ├── import_folder_structure.py  Fixed script
│   ├── FIX_SUMMARY.md              Fix docs
│   └── sharepoint_tree.json        Generated output
│
└── [other folders - numpy, pdfExtraction, etc.]
```

---

## ❓ Common Questions

### Q: Where do I start?
**A:** Read `PROJECT_STATUS.md` first, then choose your task (Email or SharePoint)

### Q: How do I extract emails?
**A:** See `email_extraction/SETUP_GUIDE_V2.md` for provider-specific instructions

### Q: Why isn't my configuration working?
**A:** Check `email_extraction/QUICK_REFERENCE_V2.md` Troubleshooting section

### Q: What email providers are supported?
**A:** Gmail, Outlook, Yahoo, Turbify, AOL, Office 365, and any custom IMAP server

### Q: Can I extract from multiple accounts?
**A:** Yes - either run multiple times with different .env values, or modify the code

### Q: Where are the extracted files?
**A:** In `email_extraction/output/` directory

### Q: How do I customize what data is extracted?
**A:** See `email_extraction/PROJECT_SUMMARY.md` - "Adding Custom Fields" section

### Q: Why am I getting SSL errors?
**A:** See `email_extraction/README_V2.md` - "SSL/TLS Errors" troubleshooting section

### Q: The SharePoint script isn't working?
**A:** Check `sharepoint/FIX_SUMMARY.md` - all environment variables should now be loaded correctly

---

## 🎓 Advanced Usage

### Extract Large Inboxes (1000+ emails)
```bash
python email_extractor_v2.py --limit 500 --batch-size 10
# Process 500 emails with batch size of 10
```

### Process Only Recent Emails
```bash
# Use --limit with first N emails from inbox
python email_extractor_v2.py --limit 100
```

### Reduce Batch Size for Slow Connections
```bash
python email_extractor_v2.py --batch-size 5
```

### Use Python API Directly
```python
import asyncio
from email_extractor_v2 import IMAPClient, ContactExtractor, ExcelExporter

async def main():
    imap = IMAPClient('imap.mail.yahoo.com', 993, 'email@yahoo.com', 'password')
    imap.connect()
    email_ids = imap.get_email_ids('INBOX')
    extractor = ContactExtractor(imap)
    contacts = await extractor.extract_all(email_ids, batch_size=20)
    ExcelExporter.export(contacts, 'output.xlsx')
    imap.disconnect()

asyncio.run(main())
```

---

## 🔗 Quick Links

| Need | File | Purpose |
|------|------|---------|
| Overview | PROJECT_STATUS.md | See what's available |
| Setup Help | SETUP_GUIDE_V2.md | Provider-specific setup |
| Commands | QUICK_REFERENCE_V2.md | Command cheat sheet |
| Architecture | PROJECT_SUMMARY.md | Code structure & design |
| Full Docs | README_V2.md | Complete documentation |
| Debugging | email_extraction_v2.log | Check for errors |
| Error Help | Troubleshooting sections in SETUP_GUIDE_V2.md | Common issues |

---

## 📞 Support Resources

### For Email Issues
1. Check `email_extraction_v2.log` for error messages
2. Run `python test_imap_connection_v2.py`
3. Review `SETUP_GUIDE_V2.md` for your provider
4. Check `README_V2.md` - Troubleshooting section

### For SharePoint Issues
1. Verify all `.env` variables are set
2. Check script output for specific errors
3. Review `sharepoint/FIX_SUMMARY.md`

### Email Provider Help
- **Gmail:** https://support.google.com/mail/answer/7126229
- **Outlook:** https://support.microsoft.com/outlook
- **Yahoo:** https://help.yahoo.com/kb
- **Office 365:** https://support.microsoft.com/office365

---

## ✅ Checklist

Before using the projects:

- [ ] Read `PROJECT_STATUS.md`
- [ ] Install dependencies: `pip install -r requirements.txt`
- [ ] Update `.env` file with credentials
- [ ] Test connection: `python test_imap_connection_v2.py`
- [ ] Run configuration wizard: `python configure_provider.py` (optional)
- [ ] Extract first 10 emails: `python email_extractor_v2.py --limit 10`
- [ ] Check `output/` for results

---

**Version:** 1.0  
**Date:** May 4, 2026  
**Status:** ✅ Complete & Ready to Use

For the main status, see: **PROJECT_STATUS.md**

