import os
import json
import sys
from pathlib import Path
import requests
import msal
from dotenv import load_dotenv

# Load .env file from parent directory
env_path = Path(__file__).parent.parent / '.env'
if env_path.exists():
    load_dotenv(env_path)
else:
    # Try current directory
    load_dotenv()

# ----------------------------
# ENV CONFIG
# ----------------------------
SITE_PATH = os.getenv("DEMO_CLIENT_SHAREPOINT_SITE_PATH", "/sites/ABLDashboardDocs")
CLIENT_ID = os.getenv("DEMO_CLIENT_SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = os.getenv("DEMO_CLIENT_SHAREPOINT_CLIENT_SECRET")
TENANT_ID = os.getenv("DEMO_CLIENT_SHAREPOINT_TENANT_ID")
SITE_HOST = os.getenv("DEMO_CLIENT_SHAREPOINT_SITE_HOST")
LIBRARY_NAME = os.getenv("DEMO_CLIENT_SHAREPOINT_LIBRARY_NAME", "Documents")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# ----------------------------
# AUTH (App-only token)
# ----------------------------
def get_token():
    try:
        authority = f"https://login.microsoftonline.com/{TENANT_ID}"
        app = msal.ConfidentialClientApplication(
            CLIENT_ID,
            authority=authority,
            client_credential=CLIENT_SECRET
        )

        result = app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )

        if "access_token" not in result:
            error_msg = result.get("error_description", "Unknown error")
            raise Exception(f"Token acquisition failed: {error_msg}")

        return result["access_token"]
    except Exception as e:
        print(f"ERROR during authentication: {e}")
        raise

# ----------------------------
# GRAPH REQUEST HELPER
# ----------------------------
def graph_get(url, token):
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(url, headers=headers)

    if r.status_code != 200:
        raise Exception(f"Graph error {r.status_code}: {r.text}")

    return r.json()

# ----------------------------
# GET SITE ID
# ----------------------------
def get_site_id(token, site_path=None):
    """Get site id for a given site path. If site_path is not provided, use SITE_PATH."""
    site_path = site_path or SITE_PATH
    url = f"{GRAPH_BASE}/sites/{SITE_HOST}:{site_path}"
    data = graph_get(url, token)
    return data["id"]

# ----------------------------
# GET DRIVE ID (document library)
# ----------------------------
def get_drive_id(token, site_id):
    url = f"{GRAPH_BASE}/sites/{site_id}/drives"
    data = graph_get(url, token)

    for drive in data["value"]:
        if drive["name"] == LIBRARY_NAME:
            return drive["id"]

    raise Exception("Document library not found")

# ----------------------------
# RECURSIVE TREE BUILDER (with error handling)
# ----------------------------
def fetch_children(token, drive_id, item_id=None, depth=0, max_depth=10):
    """Recursively fetch folder structure with error handling and depth limit"""
    import time
    
    # Limit recursion depth to prevent infinite loops
    if depth > max_depth:
        print(f"  [Depth limit reached at level {depth}]")
        return []
    
    try:
        if item_id:
            url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/children"
        else:
            url = f"{GRAPH_BASE}/drives/{drive_id}/root/children"

        data = graph_get(url, token)
        tree = []

        for item in data.get("value", []):
            node = {
                "name": item["name"],
                "id": item["id"],
                "type": "folder" if "folder" in item else "file"
            }

            # Recurse only if folder, with rate limiting
            if "folder" in item:
                time.sleep(0.1)  # Small delay to avoid rate limiting
                try:
                    node["children"] = fetch_children(token, drive_id, item["id"], depth + 1, max_depth)
                except Exception as e:
                    print(f"  [Warning] Failed to fetch children of '{item['name']}': {str(e)[:100]}")
                    node["children"] = []

            tree.append(node)

        return tree
    
    except Exception as e:
        print(f"  [Error] Failed to fetch children at depth {depth}: {str(e)[:100]}")
        return []

# ----------------------------
# MAIN EXECUTION
# ----------------------------
def main():
    # Validate required environment variables
    required_vars = {
        'CLIENT_ID': CLIENT_ID,
        'CLIENT_SECRET': CLIENT_SECRET,
        'TENANT_ID': TENANT_ID,
        'SITE_HOST': SITE_HOST
    }
    
    missing_vars = [k for k, v in required_vars.items() if not v]
    if missing_vars:
        print(f"ERROR: Missing required environment variables: {', '.join(missing_vars)}")
        print(f"\nPlease ensure these are set in .env file at: {env_path}")
        print("\nRequired .env configuration:")
        print("DEMO_CLIENT_SHAREPOINT_CLIENT_ID=xxxx")
        print("DEMO_CLIENT_SHAREPOINT_CLIENT_SECRET=xxxx")
        print("DEMO_CLIENT_SHAREPOINT_TENANT_ID=xxxx")
        print("DEMO_CLIENT_SHAREPOINT_SITE_HOST=xxxx")
        print("DEMO_CLIENT_SHAREPOINT_SITE_PATH=/sites/xxxx")
        print("DEMO_CLIENT_SHAREPOINT_LIBRARY_NAME=Documents")
        sys.exit(1)
    
    print(f"Configuration loaded from: {env_path if env_path.exists() else 'environment'}")
    print(f"Tenant ID: {TENANT_ID}")
    print(f"Site Host: {SITE_HOST}")
    print(f"Site Path: {SITE_PATH}")
    print(f"Library: {LIBRARY_NAME}")
    print()

    token = get_token()
    print("✓ Token acquired")

    # SITE_PATH may include a subfolder (e.g. /sites/ABLDashboardDocs/MB01)
    # Extract the site root (/sites/ABLDashboardDocs) and optional sub_path (MB01/...)
    site_parts = [p for p in SITE_PATH.split('/') if p]
    if len(site_parts) >= 2:
        # site_parts like ['sites', 'ABLDashboardDocs', 'MB01', ...]
        site_root = '/' + '/'.join(site_parts[:2])
        sub_path = '/'.join(site_parts[2:]) if len(site_parts) > 2 else ''
    else:
        site_root = SITE_PATH
        sub_path = ''

    print(f"Resolved site root: {site_root}")
    if sub_path:
        print(f"Resolved sub-path inside site: {sub_path}")

    # Get site id for the site root only
    site_id = get_site_id(token, site_path=site_root)
    print(f"✓ Site ID: {site_id}")

    drive_id = get_drive_id(token, site_id)
    print(f"✓ Drive ID: {drive_id}")

    # If sub_path present, resolve the folder item id in the drive
    start_item_id = None
    if sub_path:
        try:
            # Use the drive root path API to get the item for the sub_path
            url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{sub_path}"
            data = graph_get(url, token)
            start_item_id = data.get('id')
            print(f"✓ Resolved start item id for sub-path: {start_item_id}")
        except Exception as e:
            print(f"\n✗ Error resolving sub-path '{sub_path}': {e}", file=sys.stderr)
            sys.exit(1)

    print("\nBuilding folder structure...")
    tree = []
    if start_item_id:
        # Fetch children of the resolved folder
        tree = fetch_children(token, drive_id, item_id=start_item_id)
        print(f"✓ Found {len(tree)} items in '{sub_path}'")
    else:
        tree = fetch_children(token, drive_id)
        print(f"✓ Found {len(tree)} items in root")

    output = {
        "site_path": SITE_PATH,
        "library": LIBRARY_NAME,
        "structure": tree
    }

    output_file = "sharepoint_tree.json"
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2)

    print(f"\n✓ Tree exported to {output_file}")
    print(f"  Total items: {len(tree)}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n✗ Error: {e}", file=sys.stderr)
        sys.exit(1)
