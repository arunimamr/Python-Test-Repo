"""
Turbify Email Server Diagnostic Tool
This script helps identify the correct IMAP server settings for Turbify emails
"""

import socket
import imaplib
import sys

# Common Turbify and custom domain patterns
SERVERS_TO_TEST = [
    # Turbify patterns
    ('mail.turbify.com', 993, 'Turbify mail server'),
    ('imap.turbify.com', 993, 'Turbify IMAP server'),
    ('turbify.com', 993, 'Turbify root domain'),
    
    # Your domain based on email arunima@abluae.com
    ('mail.abluae.com', 993, 'ABLUAE mail server'),
    ('imap.abluae.com', 993, 'ABLUAE IMAP server'),
    ('abluae.com', 993, 'ABLUAE root domain'),
    
    # Office 365 / Outlook (backup)
    ('outlook.office365.com', 993, 'Office 365'),
    
    # Generic patterns
    ('smtp.abluae.com', 25, 'ABLUAE SMTP (port 25)'),
]

def test_connection(server, port, description):
    """Test if a server is reachable"""
    print(f"\n{'='*60}")
    print(f"Testing: {description}")
    print(f"Server: {server}:{port}")
    print('='*60)
    
    # Test DNS resolution
    try:
        ip = socket.gethostbyname(server)
        print(f"✓ DNS Resolution: {server} -> {ip}")
    except socket.gaierror as e:
        print(f"✗ DNS Resolution Failed: {e}")
        return False
    
    # Test socket connection
    try:
        sock = socket.create_connection((server, port), timeout=5)
        print(f"✓ Socket Connection: Connected to {server}:{port}")
        sock.close()
    except (socket.timeout, ConnectionRefusedError, ConnectionResetError) as e:
        print(f"✗ Socket Connection Failed: {e}")
        return False
    except Exception as e:
        print(f"✗ Socket Connection Error: {e}")
        return False
    
    # Test IMAP connection (if port 993)
    if port == 993:
        try:
            mail = imaplib.IMAP4_SSL(server, port, timeout=5)
            # Don't call capabilities() as it returns a tuple, just close
            print(f"✓ IMAP Connection: Server is IMAP-capable")
            mail.close()
            return True
        except imaplib.IMAP4.error as e:
            print(f"✗ IMAP Error: {e}")
            return False
        except Exception as e:
            print(f"✗ IMAP Connection Error: {e}")
            return False
    
    return True

def main():
    print("\n" + "="*60)
    print("TURBIFY EMAIL SERVER DIAGNOSTIC TOOL")
    print("="*60)
    print("Testing potential Turbify email servers...")
    print("This script will test connectivity to identify the correct server.\n")
    
    working_servers = []
    
    for server, port, description in SERVERS_TO_TEST:
        result = test_connection(server, port, description)
        if result:
            working_servers.append((server, port, description))
    
    print("\n" + "="*60)
    print("RESULTS SUMMARY")
    print("="*60)
    
    if working_servers:
        print(f"\n✓ Found {len(working_servers)} working server(s):\n")
        for server, port, description in working_servers:
            print(f"  • {server}:{port} ({description})")
        
        print("\n" + "-"*60)
        print("RECOMMENDED CONFIGURATION:")
        print("-"*60)
        # Use first working server
        server, port, description = working_servers[0]
        print(f"""
Update your .env file with:

    IMAP_SERVER={server}
    IMAP_PORT={port}

Configuration: {description}
        """)
    else:
        print("\n✗ No working servers found.")
        print("\nPossible issues:")
        print("  1. Network/firewall blocking outbound connections")
        print("  2. Turbify server uses non-standard hostname")
        print("  3. Your ISP blocking IMAP ports")
        print("\nNext steps:")
        print("  • Check with your IT administrator for the correct IMAP server")
        print("  • Verify if Turbify uses a custom domain name")
        print("  • Ask for the IMAP server address and port number")

if __name__ == "__main__":
    main()

