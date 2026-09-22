"""
Turbify Email Extraction Module
================================

A complete asynchronous email extraction system for Turbify business emails.

Features:
- Asynchronous batch processing (50 emails per batch)
- Multi-threaded execution (5 threads)
- Contact information extraction (From, CC, BCC, Name, Mobile, Logo)
- Signature parsing
- Excel export with professional formatting
- Real-time logging and error handling

Version: 2.0
Status: Production Ready
Date: April 16, 2026

Main Components:
    - extractContacts_Turbify: Enhanced email extractor (recommended)
    - test_imap_connection: Connection and configuration tester
    - turbify_diagnostic: Server finder and diagnostic tool
    - run_extractor: Windows automation scripts (PowerShell & Batch)

Quick Start:
    1. Edit .env and add PASSWORD
    2. Run: python -m email_extraction.test_imap_connection
    3. Run: python -m email_extraction.extractContacts_Turbify

Documentation:
    - docs/TURBIFY_QUICKSTART.md: 5-minute quick start
    - docs/COMPLETE_SETUP_GUIDE.md: Full setup guide
    - docs/README_TURBIFY.md: Technical reference
    - docs/INDEX.md: Documentation index
    - docs/CHECKLIST.md: Completion checklist
"""

__version__ = "2.0"
__author__ = "Turbify Email Extraction System"
__status__ = "Production Ready"

# Import main modules for easy access
try:
    from .extractContacts_Turbify import TurbifyEmailExtractor, main
    from .test_imap_connection import test_imap_connection
    from .turbify_diagnostic import test_connection
except ImportError as e:
    import warnings
    warnings.warn(f"Could not import modules: {e}")

__all__ = [
    'TurbifyEmailExtractor',
    'main',
    'test_imap_connection',
    'test_connection',
]

if __name__ == "__main__":
    print(f"Turbify Email Extraction Module v{__version__}")
    print("Status:", __status__)
    print("\nFor quick start, see: docs/TURBIFY_QUICKSTART.md")
    print("For help, see: docs/INDEX.md")

