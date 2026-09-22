import os
import sys
import argparse
import asyncio
import concurrent.futures
from datetime import datetime
from threading import Lock
import requests
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from dotenv import load_dotenv, find_dotenv
import psycopg2
from psycopg2 import pool

# Load environment variables
load_dotenv(find_dotenv())

try:
    from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError
    PLAYWRIGHT_AVAILABLE = True
except ImportError:
    PLAYWRIGHT_AVAILABLE = False

# Import DOM inspection helpers from fg_web_automation if available
try:
    from fg_web_automation import inspect_issue_and_hierarchy, auto_login_facilitygrid, resolve_full_url
    WEB_INSPECT_AVAILABLE = True
except ImportError:
    WEB_INSPECT_AVAILABLE = False


# =====================================================================
# Constants & Configuration
# =====================================================================

# Responsibilities that indicate "Light Blue" when no asset info is provided
RESPONSIBILITY_LIST = [
    "SCH - Schneider Electric",
    "TRN - Trane",
    "SCH - Schneider MV",
    "AER - Aermec",
    "MTU - Rolls Royce",
    "EIG - E and I Engineering Ltd",
    "ARM -Armstrong",
    "SCH - Schneider DCOS",
    "Hazemeyer",
    "SMT - SMARDT",
    "REH - Rehlko Gens",
    "ABEC - DCOS",
    "JAG - Jaeggi AG",
    "ITEC - ITEC Groupe",
    "SIE- Siemens MV",
    "Eppy - EppyGroup",
    "VER - Vertiv Electrical",
    "TES - Tesar Gulf",
    "HAP - Hitachi Energy",
    "SYC - Systecon",
    "SYC – Systecon",  # Include common dash variants
    "HAV - Havtech Mech",
    "AMG - Anord Mardix Group",
    "SCH - Schneider UPS",
    "HAN - Hanley Energy Elec",
    "DEL - Delta Elec",
    "JCI - Johnson Control",
    "SCH - Schneider Mech",
    "SGB-SMIT",
    "SCH - Schneider CA",
    "VER - Vertiv",
    "Ver - Vertiv HVAC",
    "SOC - Socomec",
    "CMM - Cummins"
]

NORMALIZED_RESPONSIBILITIES = [resp.strip().lower() for resp in RESPONSIBILITY_LIST]

# Openpyxl Styles
regular_font = Font(name="Calibri", size=11)
bold_font = Font(name="Calibri", size=11, bold=True)
border_thin = Border(
    left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin")
)

# Colors
header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
cyan_fill = PatternFill(start_color="00FFFF", end_color="00FFFF", fill_type="solid")     # API asset provided + Web Mfg found!
orange_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")  # Web Fallback Match
red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
light_blue_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
purple_fill = PatternFill(start_color="800080", end_color="800080", fill_type="solid")


def get_auth_token() -> str:
    auth_url = os.getenv("EXTERNAL_API_AUTH_ENDPOINT")
    client_id = os.getenv("EXTERNAL_API_CLIENT_ID")
    client_secret = os.getenv("EXTERNAL_API_CLIENT_SECRET")
    grant_type = os.getenv("EXTERNAL_API_GRANT_TYPE", "client_credentials")

    if not all([auth_url, client_id, client_secret]):
        print("[ERROR] Missing FG API Client credentials in .env file.")
        print("Required: EXTERNAL_API_AUTH_ENDPOINT, EXTERNAL_API_CLIENT_ID, EXTERNAL_API_CLIENT_SECRET")
        sys.exit(1)

    print(f"Authenticating with {auth_url}...")
    try:
        response = requests.post(
            auth_url,
            data={
                "grant_type": grant_type,
                "client_id": client_id,
                "client_secret": client_secret
            },
            timeout=10
        )
        response.raise_for_status()
        token = response.json().get("access_token")
        if not token:
            print("[ERROR] Failed to extract access_token from response.")
            sys.exit(1)
        print("[SUCCESS] API Token retrieved successfully!\n")
        return token
    except Exception as e:
        print(f"[ERROR] Authentication failed: {e}")
        if 'response' in locals() and response is not None:
            print(f"Response Body: {response.text}")
        sys.exit(1)


def get_row_val(row, *candidate_keys) -> str:
    """Helper for case-insensitive column lookup from pandas row dict."""
    for key in candidate_keys:
        if key in row and pd.notna(row[key]):
            return str(row[key]).replace(".0", "").strip()
        for r_k in row.keys():
            if str(r_k).strip().lower() == str(key).strip().lower() and pd.notna(row[r_k]):
                return str(row[r_k]).replace(".0", "").strip()
    return ""


def save_styled_excel(records: list[dict], output_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "API Scan Results"

    headers = [
        "Excel Row", "Issue Code", "Project Name", "FG Project ID", "FG Issue ID", "Issue URL",
        "Equipment UUID", "Asset UUID", "Subtask UUID", 
        "Asset Info Provided", "Responsibility Name", "DB Status",
        "Web Cross-Check", "Web Equipment / Asset", "Web Subtask", "Web Manufacturer", "Web Model Number", "Web Source",
        "Status"
    ]
    
    # Write Headers
    for col_idx, header_val in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header_val)
        cell.font = bold_font
        cell.fill = header_fill
        cell.border = border_thin
        cell.alignment = Alignment(horizontal="center", vertical="center")

    sorted_records = sorted(records, key=lambda x: x.get("excel_row", 0))
        
    # Write Data
    for row_idx, r in enumerate(sorted_records, start=2):
        row_data = [
            r.get("excel_row", ""),
            r.get("issue_code", ""),
            r.get("project_name", ""),
            r.get("fg_project_id", ""),
            r.get("fg_issue_uuid", ""),
            r.get("issue_url", ""),
            r.get("equipment_uuid", ""),
            r.get("asset_uuid", ""),
            r.get("subtask_uuid", ""),
            r.get("asset_info_provided", ""),
            r.get("responsibility_name", ""),
            r.get("db_status", ""),
            r.get("web_cross_check", "Not Checked"),
            r.get("web_equipment_name", ""),
            r.get("web_subtask_name", ""),
            r.get("web_manufacturer", ""),
            r.get("web_model_number", ""),
            r.get("web_found_source", ""),
            r.get("status", "")
        ]

        color_assigned = r.get("color_assigned", "")

        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = regular_font
            cell.border = border_thin
            cell.alignment = Alignment(vertical="center")

            if col_idx in [1, 2, 4, 9, 11, 12]:
                cell.alignment = Alignment(horizontal="center", vertical="center")

            # Apply full row color highlights
            if color_assigned == "CYAN":
                cell.fill = cyan_fill
            elif color_assigned == "YELLOW":
                cell.fill = yellow_fill
            elif color_assigned == "ORANGE":
                cell.fill = orange_fill
            elif color_assigned == "RED":
                cell.fill = red_fill
            elif color_assigned == "LIGHT BLUE":
                cell.fill = light_blue_fill
            elif color_assigned == "PURPLE":
                cell.fill = purple_fill

    # Auto-adjust column widths
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 45)

    # Add Legend at the bottom
    legend_start_row = len(sorted_records) + 4
    
    ws.cell(row=legend_start_row, column=1, value="Color Legend").font = bold_font
    
    ws.cell(row=legend_start_row + 1, column=1, value="CYAN").fill = cyan_fill
    ws.cell(row=legend_start_row + 1, column=2, value="API Provided Asset UUIDs AND Manufacturer Info WAS FOUND on Web UI")

    ws.cell(row=legend_start_row + 2, column=1, value="YELLOW").fill = yellow_fill
    ws.cell(row=legend_start_row + 2, column=2, value="API Provided Asset UUIDs, BUT Manufacturer Info Missing on Web UI")

    ws.cell(row=legend_start_row + 3, column=1, value="PURPLE").fill = purple_fill
    ws.cell(row=legend_start_row + 3, column=2, value="DB Mismatch (API has asset info, DB is missing it) -> OUR SIDE CODE FIX NEEDED!")

    ws.cell(row=legend_start_row + 4, column=1, value="ORANGE").fill = orange_fill
    ws.cell(row=legend_start_row + 4, column=2, value="API Asset UUIDs Null, but Info Found on Web UI -> FACILITYGRID (FG) API ISSUE!")

    ws.cell(row=legend_start_row + 5, column=1, value="LIGHT BLUE").fill = light_blue_fill
    ws.cell(row=legend_start_row + 5, column=2, value="No Asset Info (API/Web), but Valid Responsibility Match")
    
    ws.cell(row=legend_start_row + 6, column=1, value="RED").fill = red_fill
    ws.cell(row=legend_start_row + 6, column=2, value="No Asset Info (API/Web) & No Match Responsibility")

    def write_filtered_sheet(sheet_title: str, fill_style, filter_fn, empty_message: str):
        ws_sub = wb.create_sheet(title=sheet_title)
        
        # Write Headers
        for col_idx, header_val in enumerate(headers, start=1):
            cell = ws_sub.cell(row=1, column=col_idx, value=header_val)
            cell.font = bold_font
            cell.fill = header_fill
            cell.border = border_thin
            cell.alignment = Alignment(horizontal="center", vertical="center")

        filtered_rows = [r for r in sorted_records if filter_fn(r)]
        
        if not filtered_rows:
            cell = ws_sub.cell(row=2, column=1, value=empty_message)
            cell.font = bold_font
        else:
            for row_idx, r in enumerate(filtered_rows, start=2):
                row_data = [
                    r.get("excel_row", ""),
                    r.get("issue_code", ""),
                    r.get("project_name", ""),
                    r.get("fg_project_id", ""),
                    r.get("fg_issue_uuid", ""),
                    r.get("issue_url", ""),
                    r.get("equipment_uuid", ""),
                    r.get("asset_uuid", ""),
                    r.get("subtask_uuid", ""),
                    r.get("asset_info_provided", ""),
                    r.get("responsibility_name", ""),
                    r.get("db_status", ""),
                    r.get("web_cross_check", "Not Checked"),
                    r.get("web_equipment_name", ""),
                    r.get("web_subtask_name", ""),
                    r.get("web_manufacturer", ""),
                    r.get("web_model_number", ""),
                    r.get("web_found_source", ""),
                    r.get("status", "")
                ]

                for col_idx, val in enumerate(row_data, start=1):
                    cell = ws_sub.cell(row=row_idx, column=col_idx, value=val)
                    cell.font = regular_font
                    cell.border = border_thin
                    cell.fill = fill_style
                    cell.alignment = Alignment(vertical="center")

                    if col_idx in [1, 2, 4, 9, 11, 12]:
                        cell.alignment = Alignment(horizontal="center", vertical="center")

        # Auto-adjust column widths
        for col in ws_sub.columns:
            col_letter = get_column_letter(col[0].column)
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws_sub.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 45)

        ws_sub.freeze_panes = "A2"

    # Write additional filtered sheets
    write_filtered_sheet("Our Code Issues", purple_fill, lambda r: r.get("color_assigned") == "PURPLE", "No PURPLE (Our Code DB Mismatch) issues found.")
    write_filtered_sheet("FG API Issues", orange_fill, lambda r: r.get("color_assigned") == "ORANGE", "No ORANGE (FG API Web Fallback) issues found.")
    write_filtered_sheet("Cyan - Web Mfg Found", cyan_fill, lambda r: r.get("color_assigned") == "CYAN", "No CYAN issues (API Provided Asset & Web Manufacturer Found).")
    write_filtered_sheet("Yellow - Web Mfg Missing", yellow_fill, lambda r: r.get("color_assigned") == "YELLOW", "No YELLOW issues (API Provided Asset, but Web Manufacturer Missing).")

    try:
        wb.save(output_path)
    except PermissionError:
        base, ext = os.path.splitext(output_path)
        alt_path = f"{base}_updated{ext}"
        print(f"\n[WARNING] Could not save to '{output_path}' because it is open in Excel.", flush=True)
        print(f"[INFO] Saving results to secondary output file instead: {alt_path}", flush=True)
        wb.save(alt_path)


# =====================================================================
# Phase 1: API & DB Inspection
# =====================================================================

def process_row(idx, original_idx, row, base_api_url, token, db_pool, check_all_web: bool = False):
    issue_code = get_row_val(row, "issue_code", "Issue Code")
    project_name = get_row_val(row, "project_name", "Project Name")
    
    project_id = get_row_val(row, "fg_project_id", "FG Project ID")
    issue_id = get_row_val(row, "fg_issue_uuid", "FG Issue ID")
    issue_url = get_row_val(row, "issue_url", "Issue URL")

    if not issue_url and project_id and issue_id:
        issue_url = f"{base_api_url}/inquire?action=eiss&prj={project_id}&rqi={issue_id}"

    api_url = f"{base_api_url}/api/v2_2/project/{project_id}/issue/{issue_id}"
    
    eq_uuid = ""
    asset_uuid = ""
    subtask_uuid = ""
    resp_name = ""
    status_msg = "Success"
    asset_info_provided = "No"
    color_assigned = ""

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }

    try:
        res = requests.get(api_url, headers=headers, timeout=15)
        if res.status_code == 200:
            data = res.json()
            
            eq_val = data.get("equipment_uuid") or ""
            eq_uuid = eq_val.strip() if isinstance(eq_val, str) else str(eq_val).strip()

            sub_val = data.get("subtask_uuid") or ""
            subtask_uuid = sub_val.strip() if isinstance(sub_val, str) else str(sub_val).strip()
            
            asset_obj = data.get("asset_uuid")
            if isinstance(asset_obj, dict):
                asset_uuid = str(asset_obj.get("uuid") or "").strip()
            elif isinstance(asset_obj, str):
                asset_uuid = asset_obj.strip()
            
            resp_name = str(data.get("responsibility_name") or "").strip()
            
            # Check if non-empty UUIDs are present
            has_asset = bool(eq_uuid) or bool(asset_uuid) or bool(subtask_uuid)
            
            if has_asset:
                asset_info_provided = "Yes"
                color_assigned = "YELLOW"
            else:
                asset_info_provided = "No"
                resp_norm = resp_name.strip().lower()
                if any(resp_norm == r for r in NORMALIZED_RESPONSIBILITIES):
                    color_assigned = "LIGHT BLUE"
                else:
                    color_assigned = "RED"
        else:
            status_msg = f"HTTP {res.status_code}: {res.text[:50]}"
            color_assigned = ""
    except Exception as e:
        status_msg = f"Error: {str(e)[:50]}"
        color_assigned = ""

    # Database Mismatch Verification
    db_status = "Not Checked"
    if color_assigned in ["YELLOW", "LIGHT BLUE"] and db_pool:
        conn = None
        try:
            conn = db_pool.getconn()
            if conn:
                cur = conn.cursor()
                cur.execute("SELECT equipment_uuid, asset_uuid, subtask_uuid, responsibility_name FROM issues WHERE fg_issue_uuid = %s", (issue_id,))
                db_row = cur.fetchone()
                cur.close()
                
                if db_row:
                    db_eq, db_asset, db_subtask, db_resp = db_row
                    db_has_asset = bool(db_eq and str(db_eq).strip()) or bool(db_asset and str(db_asset).strip()) or bool(db_subtask and str(db_subtask).strip())
                    db_resp_norm = (db_resp or "").strip().lower()
                    db_has_resp = any(db_resp_norm == r for r in NORMALIZED_RESPONSIBILITIES)
                    
                    if color_assigned == "YELLOW" and not db_has_asset:
                        color_assigned = "PURPLE"
                        db_status = "Mismatch: Missing Asset Info"
                    elif color_assigned == "LIGHT BLUE" and not db_has_resp:
                        color_assigned = "PURPLE"
                        db_status = "Mismatch: Missing Responsibility"
                    else:
                        db_status = "Checked - OK"
                else:
                    color_assigned = "PURPLE"
                    db_status = "Mismatch: Issue missing from DB"
        except Exception as e:
            db_status = f"DB Error: {str(e)[:40]}"
        finally:
            if conn:
                db_pool.putconn(conn)

    return {
        "excel_row": original_idx + 2,
        "issue_code": issue_code,
        "project_name": project_name,
        "fg_project_id": project_id,
        "fg_issue_uuid": issue_id,
        "issue_url": issue_url,
        "equipment_uuid": eq_uuid,
        "asset_uuid": asset_uuid,
        "subtask_uuid": subtask_uuid,
        "asset_info_provided": asset_info_provided,
        "responsibility_name": resp_name,
        "color_assigned": color_assigned,
        "db_status": db_status,
        "web_cross_check": "Pending Web Check" if (check_all_web or color_assigned in ["RED", "LIGHT BLUE"]) else "Skipped (API Has Data)",
        "web_equipment_name": "",
        "web_subtask_name": "",
        "web_manufacturer": "",
        "web_model_number": "",
        "web_found_source": "",
        "status": status_msg
    }


# =====================================================================
# Phase 2: Async Playwright Web UI Cross-Check
# =====================================================================

async def web_worker_task(
    worker_id: int,
    context,
    queue: asyncio.Queue,
    records_dict: dict,
    records_lock: asyncio.Lock,
    total: int,
    completed_counter: list,
    output_path: str,
    timeout_ms: int,
    batch_size: int = 10,
):
    page = await context.new_page()
    try:
        # Block image, media, font network requests to speed up page loading by 5x-10x
        await page.route(
            "**/*", 
            lambda route: route.abort() if route.request.resource_type in ["image", "media", "font"] else route.continue_()
        )
        while not queue.empty():
            try:
                rec = queue.get_nowait()
            except asyncio.QueueEmpty:
                break

            issue_code = rec.get("issue_code", "")
            issue_url = rec.get("issue_url", "")
            excel_row = rec.get("excel_row", 0)

            async with records_lock:
                started_done = completed_counter[0]
                progress_pct = (started_done / total) * 100 if total > 0 else 0.0

            worker_prefix = f"[Phase2-WebWorker-{worker_id}][{started_done}/{total} ({progress_pct:.1f}%)]"
            print(f"{worker_prefix} Cross-checking Web UI for {issue_code}...", flush=True)

            if not issue_url:
                rec["web_cross_check"] = "No Web URL"
                async with records_lock:
                    records_dict[excel_row] = rec
                    completed_counter[0] += 1
                queue.task_done()
                continue

            scan_res = await inspect_issue_and_hierarchy(page, issue_url, timeout_ms=timeout_ms, worker_prefix=worker_prefix)

            mfg_val = scan_res.get("manufacturer", "")
            model_val = scan_res.get("model_number", "")
            found_src = scan_res.get("found_source", "")
            eq_name = scan_res.get("equipment_name", "")
            sub_name = scan_res.get("subtask_name", "")

            found_anything = bool(mfg_val) or bool(model_val) or bool(eq_name) or bool(sub_name)
            mfg_found = bool(mfg_val and str(mfg_val).strip())

            if found_anything:
                rec["web_cross_check"] = "FOUND ON WEB"
                rec["web_equipment_name"] = eq_name
                rec["web_subtask_name"] = sub_name
                rec["web_manufacturer"] = mfg_val
                rec["web_model_number"] = model_val
                rec["web_found_source"] = found_src

                cur_color = rec.get("color_assigned", "")
                if cur_color == "YELLOW":
                    if mfg_found:
                        rec["color_assigned"] = "CYAN"  # API asset provided AND Manufacturer info WAS FOUND on Web UI -> CYAN
                    else:
                        rec["color_assigned"] = "YELLOW" # API asset provided BUT Manufacturer info MISSING on Web UI -> YELLOW
                elif cur_color not in ["CYAN", "PURPLE"]:
                    rec["color_assigned"] = "ORANGE"  # Highlight ORANGE for Web Fallback Match!
            else:
                rec["web_cross_check"] = "Not Found on Web"
                cur_color = rec.get("color_assigned", "")
                if cur_color not in ["YELLOW", "PURPLE", "CYAN"]:
                    resp_norm = str(rec.get("responsibility_name") or "").strip().lower()
                    if any(resp_norm == r for r in NORMALIZED_RESPONSIBILITIES):
                        rec["color_assigned"] = "LIGHT BLUE"
                    else:
                        rec["color_assigned"] = "RED"

            async with records_lock:
                records_dict[excel_row] = rec
                completed_counter[0] += 1
                curr_done = completed_counter[0]
                pct = (curr_done / total) * 100 if total > 0 else 0.0

                if found_anything:
                    print(f"[Phase2-WebWorker-{worker_id}][Done {curr_done}/{total} ({pct:.1f}%)] >>> [WEB MATCH!] {issue_code} -> Eq: '{eq_name}' | Sub: '{sub_name}' | Mfg: '{mfg_val}' (Marked ORANGE)", flush=True)
                else:
                    print(f"[Phase2-WebWorker-{worker_id}][Done {curr_done}/{total} ({pct:.1f}%)] {issue_code} -> No asset info on Web UI.", flush=True)

                if curr_done % batch_size == 0 or curr_done == total:
                    save_styled_excel(list(records_dict.values()), output_path)
                    print(f"      [WEB CHECKPOINT] Saved progress ({curr_done}/{total} completed - {pct:.1f}%) to: {output_path}", flush=True)

            queue.task_done()
    finally:
        await page.close()


async def run_web_cross_check_phase(
    records: list[dict],
    output_path: str,
    web_workers: int = 5,
    timeout_ms: int = 15000,
    headless: bool = False,
):
    target_records = [r for r in records if r.get("web_cross_check") == "Pending Web Check"]
    if not target_records:
        print("\n[PHASE 2 WEB CROSS-CHECK] No rows require web UI cross-checking! All rows have API asset info.")
        return

    print(f"\n=======================================================")
    print(f"[PHASE 2 WEB CROSS-CHECK] Starting Playwright Web UI Inspection...")
    print(f"Rows requiring Web Cross-Check: {len(target_records)}")
    print(f"Web Worker Pages              : {web_workers}")
    print(f"=======================================================\n")

    fg_username = (os.getenv("FG_USERNAME") or os.getenv("FG_EMAIL") or "").strip()
    fg_password = (os.getenv("FG_PASSWORD") or "").strip()

    profile_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".fg_playwright_profile")
    os.makedirs(profile_dir, exist_ok=True)

    records_dict = {r["excel_row"]: r for r in records}

    async with async_playwright() as p:
        context = None
        candidate_profiles = [
            os.path.join(os.path.dirname(os.path.abspath(__file__)), ".fg_playwright_profile"),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), ".fg_api_playwright_profile"),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), ".fg_api_playwright_profile_fallback"),
        ]
        for p_dir in candidate_profiles:
            os.makedirs(p_dir, exist_ok=True)
            try:
                context = await p.chromium.launch_persistent_context(
                    user_data_dir=p_dir,
                    headless=headless,
                    args=["--start-maximized", "--disable-blink-features=AutomationControlled"],
                    no_viewport=True,
                    user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                )
                print(f"[INIT WEB] Successfully initialized browser session with profile: {p_dir}", flush=True)
                break
            except Exception as launch_err:
                print(f"[INIT WEB] Profile {p_dir} locked or unavailable ({str(launch_err)[:50]}...). Trying fallback profile...", flush=True)

        if not context:
            import tempfile
            temp_profile = tempfile.mkdtemp(prefix="fg_playwright_")
            print(f"[INIT WEB] All static profiles locked. Initializing temporary browser profile: {temp_profile}", flush=True)
            context = await p.chromium.launch_persistent_context(
                user_data_dir=temp_profile,
                headless=headless,
                args=["--start-maximized", "--disable-blink-features=AutomationControlled"],
                no_viewport=True,
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            )

        init_page = context.pages[0] if context.pages else await context.new_page()

        # One-time login check on first target URL
        first_url = target_records[0].get("issue_url", "")
        if first_url:
            print(f"[INIT WEB] Checking authentication at: {first_url} ...", flush=True)
            try:
                await init_page.goto(first_url, timeout=timeout_ms, wait_until="domcontentloaded")
                await init_page.wait_for_timeout(2000)
                if "login" in init_page.url.lower() or "signin" in init_page.url.lower():
                    logged_in = False
                    if fg_username and fg_password:
                        print(f"[INIT WEB] Attempting automated login...", flush=True)
                        logged_in = await auto_login_facilitygrid(init_page, fg_username, fg_password, timeout_ms=timeout_ms)

                    if not logged_in:
                        print("\n>>> [ACTION REQUIRED] Please complete login in the opened Chromium browser.")
                        await asyncio.to_thread(input, ">>> Press ENTER here after logging in... ")
                else:
                    print("[INIT WEB] Active login session confirmed! Proceeding with web cross-check...\n", flush=True)
            except Exception as e:
                print(f"[INIT WEB] Note: {e}", flush=True)

        queue = asyncio.Queue()
        for rec in target_records:
            queue.put_nowait(rec)

        records_lock = asyncio.Lock()
        completed_counter = [0]
        total_targets = len(target_records)

        num_workers = min(web_workers, total_targets)
        worker_tasks = [
            asyncio.create_task(
                web_worker_task(
                    worker_id=w_idx + 1,
                    context=context,
                    queue=queue,
                    records_dict=records_dict,
                    records_lock=records_lock,
                    total=total_targets,
                    completed_counter=completed_counter,
                    output_path=output_path,
                    timeout_ms=timeout_ms,
                )
            )
            for w_idx in range(num_workers)
        ]

        await asyncio.gather(*worker_tasks)
        await context.close()

    # Final update
    save_styled_excel(list(records_dict.values()), output_path)


import glob

def find_latest_api_result_file(folder: str) -> str:
    """Finds the most recently modified main api_scan_results_*.xlsx in target folder (excluding temporary _updated.xlsx)."""
    if not folder or not os.path.exists(folder):
        return None
    files = [f for f in glob.glob(os.path.join(folder, "api_scan_results_*.xlsx")) if not f.endswith("_updated.xlsx")]
    if not files:
        files = glob.glob(os.path.join(folder, "api_scan_results_*.xlsx"))
    if not files:
        return None
    return max(files, key=os.path.getmtime)


# =====================================================================
# Main Execution Runner
# =====================================================================

def run_api_scanner(
    excel_path: str = None,
    output_path: str = None,
    limit: int = None,
    max_workers: int = 15,
    enable_web_check: bool = True,
    web_workers: int = 5,
    headless: bool = False,
    phase_2_only: bool = False,
    check_all_web: bool = False,
):
    env_folder = (os.getenv("FG_FOLDER_NAME") or "").strip()
    env_phase1_file = (os.getenv("PHASE_1_FILE_FOR_API_WEB_TEST") or "").strip()
    env_file = (os.getenv("FG_API_TEST_FILE_NAME") or os.getenv("FG_FILE_NAME") or "").strip()
    base_api_url = (os.getenv("EXTERNAL_API_BASE_URL") or "https://equinix.facilitygrid.co.uk").rstrip('/')
    env_check_all = os.getenv("FG_CHECK_ALL_WEB")
    if env_check_all is not None:
        check_all_web = env_check_all.strip().lower() == "true"

    # If Phase 2 only requested and no file passed via CLI, check .env setting first, then latest file
    if phase_2_only and not excel_path:
        if env_phase1_file:
            candidate = os.path.join(env_folder, env_phase1_file) if env_folder and not os.path.isabs(env_phase1_file) else env_phase1_file
            if os.path.exists(candidate):
                excel_path = candidate
                print(f"[INFO] Phase 2 standalone mode: Using PHASE_1_FILE_FOR_API_WEB_TEST from .env: {excel_path}")
        if not excel_path and env_folder:
            latest_file = find_latest_api_result_file(env_folder)
            if latest_file:
                excel_path = latest_file
                print(f"[INFO] Phase 2 standalone mode: Auto-selected latest result file: {excel_path}")

    # Resolve input Excel file path
    if not excel_path:
        if env_folder and env_file:
            excel_path = os.path.join(env_folder, env_file)
        elif env_file:
            excel_path = env_file
        else:
            print("[ERROR] Please provide --file or configure FG_FOLDER_NAME & FG_API_TEST_FILE_NAME in .env")
            return
    elif env_folder and not os.path.isabs(excel_path) and not os.path.exists(excel_path):
        candidate = os.path.join(env_folder, excel_path)
        if os.path.exists(candidate):
            excel_path = candidate

    if not os.path.exists(excel_path):
        print(f"\n[ERROR] Excel input file not found: {excel_path}")
        return

    target_folder = env_folder if env_folder and os.path.isdir(env_folder) else os.path.dirname(os.path.abspath(excel_path))
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    if not output_path:
        if phase_2_only and "api_scan_results_" in os.path.basename(excel_path):
            # Overwrite/update existing result file in-place if running Phase 2 standalone
            output_path = excel_path
        else:
            default_name = f"api_scan_results_{timestamp}.xlsx"
            output_path = os.path.join(target_folder, default_name) if target_folder else default_name
    else:
        if not os.path.isabs(output_path) and not os.path.dirname(output_path) and target_folder:
            output_path = os.path.join(target_folder, output_path)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)

    print(f"\n=======================================================")
    print(f"Loading Excel file        : {os.path.abspath(excel_path)}")
    print(f"Output destination        : {os.path.abspath(output_path)}")
    if phase_2_only:
        print(f"Execution Mode            : STANDALONE PHASE 2 (Playwright Web UI Cross-Check)")
        print(f"Web Worker Pages          : {web_workers}")
    else:
        print(f"API Parallel Threads      : {max_workers}")
        print(f"Web Cross-Check Enabled   : {enable_web_check}")
        if enable_web_check:
            print(f"Web Worker Pages          : {web_workers}")

    df = pd.read_excel(excel_path)
    print(f"Total rows in Excel       : {len(df)}")

    # -------------------------------------------------------------
    # STANDALONE PHASE 2 EXECUTION MODE
    # -------------------------------------------------------------
    if phase_2_only:
        if not (PLAYWRIGHT_AVAILABLE and WEB_INSPECT_AVAILABLE):
            print("[ERROR] Playwright or web inspection modules not available. Cannot run Phase 2.")
            return

        records = []
        is_previous_results = "color_assigned" in df.columns or "asset_info_provided" in df.columns

        for idx, (original_idx, row) in enumerate(df.iterrows(), start=1):
            issue_code = get_row_val(row, "issue_code", "Issue Code")
            project_name = get_row_val(row, "project_name", "Project Name")
            project_id = get_row_val(row, "fg_project_id", "FG Project ID")
            issue_id = get_row_val(row, "fg_issue_uuid", "FG Issue ID")
            
            issue_url = get_row_val(row, "issue_url", "Issue URL")
            if not issue_url and project_id and issue_id:
                issue_url = f"{base_api_url}/inquire?action=eiss&prj={project_id}&rqi={issue_id}"

            # Skip legend rows or empty rows that have no issue_id and no issue_url
            if not issue_url and not issue_id:
                continue

            color_val = get_row_val(row, "color_assigned", "Color Assigned")
            asset_provided = get_row_val(row, "asset_info_provided", "Asset Info Provided") or "No"
            web_status = get_row_val(row, "web_cross_check", "Web Cross-Check")

            # Determine if this row needs web cross-check
            needs_check = False
            if is_previous_results:
                if check_all_web or color_val in ["RED", "LIGHT BLUE"] or asset_provided == "No" or web_status == "Pending Web Check":
                    needs_check = True
            else:
                needs_check = True

            records.append({
                "excel_row": original_idx + 2,
                "issue_code": issue_code,
                "project_name": project_name,
                "fg_project_id": project_id,
                "fg_issue_uuid": issue_id,
                "issue_url": issue_url,
                "equipment_uuid": get_row_val(row, "equipment_uuid", "Equipment UUID"),
                "asset_uuid": get_row_val(row, "asset_uuid", "Asset UUID"),
                "subtask_uuid": get_row_val(row, "subtask_uuid", "Subtask UUID"),
                "asset_info_provided": asset_provided,
                "responsibility_name": get_row_val(row, "responsibility_name", "Responsibility Name"),
                "color_assigned": color_val or "RED",
                "db_status": get_row_val(row, "db_status", "DB Status") or "Not Checked",
                "web_cross_check": "Pending Web Check" if needs_check else (web_status or "Skipped"),
                "web_equipment_name": get_row_val(row, "web_equipment_name", "Web Equipment / Asset"),
                "web_subtask_name": get_row_val(row, "web_subtask_name", "Web Subtask"),
                "web_manufacturer": get_row_val(row, "web_manufacturer", "Web Manufacturer"),
                "web_model_number": get_row_val(row, "web_model_number", "Web Model Number"),
                "web_found_source": get_row_val(row, "web_found_source", "Web Source"),
                "status": get_row_val(row, "status", "Status") or "Success",
            })

        if limit and limit > 0:
            records = records[:limit]
            print(f"Applying limit           : Processing top {limit} rows")

        asyncio.run(
            run_web_cross_check_phase(
                records=records,
                output_path=output_path,
                web_workers=web_workers,
                headless=headless,
            )
        )

        out_df = pd.DataFrame(records)
        cyan_count = (out_df["color_assigned"] == "CYAN").sum() if "color_assigned" in out_df else 0
        yellow_count = (out_df["color_assigned"] == "YELLOW").sum() if "color_assigned" in out_df else 0
        orange_count = (out_df["color_assigned"] == "ORANGE").sum() if "color_assigned" in out_df else 0
        red_count = (out_df["color_assigned"] == "RED").sum() if "color_assigned" in out_df else 0
        blue_count = (out_df["color_assigned"] == "LIGHT BLUE").sum() if "color_assigned" in out_df else 0
        purple_count = (out_df["color_assigned"] == "PURPLE").sum() if "color_assigned" in out_df else 0

        print(f"\n==================== STANDALONE PHASE 2 COMPLETE ====================")
        print(f"Total Rows Checked       : {len(out_df)}")
        print(f"CYAN                     : {cyan_count} (API Asset Provided AND Manufacturer Info Found on Web UI)")
        print(f"YELLOW                   : {yellow_count} (API Asset Provided, BUT Manufacturer Info Missing on Web UI)")
        print(f"PURPLE                   : {purple_count} (DB Mismatch - OUR SIDE CODE FIX NEEDED!)")
        print(f"ORANGE                   : {orange_count} (API Asset UUIDs Null, but Info Found on Web UI -> FACILITYGRID API ISSUE!)")
        print(f"LIGHT BLUE               : {blue_count} (No Asset Info API/Web, but Valid Responsibility Match)")
        print(f"RED                      : {red_count} (No Asset Info API/Web & No Match Responsibility)")
        print(f"Results saved to         : {os.path.abspath(output_path)}")
        print(f"====================================================================\n")
        return

    # -------------------------------------------------------------
    # FULL EXECUTION MODE (PHASE 1 + PHASE 2)
    # -------------------------------------------------------------

    # Initialize DB Pool
    db_pool = None
    try:
        db_url = os.getenv("DATABASE_URL")
        if not db_url:
            db_host = os.getenv("DB_HOST", "localhost")
            db_port = os.getenv("DB_PORT", "5432")
            db_name = os.getenv("DB_NAME", "cx_hub_dashboard_db")
            db_user = os.getenv("DB_USER", "postgres")
            db_pass = os.getenv("DB_PASSWORD", "dummy-password")
            db_url = f"postgresql://{db_user}:{db_pass}@{db_host}:{db_port}/{db_name}"
            
        print(f"Connecting to database pool at {db_url.split('@')[-1]}...")
        db_pool = psycopg2.pool.ThreadedConnectionPool(1, max_workers, db_url)
    except Exception as e:
        print(f"[WARNING] Could not connect to DB pool: {e}. DB verification will be skipped.")

    # Ensure fg_project_id and fg_issue_uuid exist
    required_cols = ["fg_project_id", "fg_issue_uuid"]
    missing_cols = [c for c in required_cols if c not in df.columns]
    if missing_cols:
        print(f"[ERROR] The Excel file is missing required columns: {missing_cols}")
        print(f"Columns found: {list(df.columns)}")
        return

    valid_df = df[df["fg_issue_uuid"].notna() & (df["fg_issue_uuid"].astype(str).str.strip() != "")].copy()
    print(f"Rows with valid fg_issue_uuid: {len(valid_df)}")

    if limit and limit > 0:
        valid_df = valid_df.head(limit)
        print(f"Applying limit           : Processing top {limit} rows")

    if len(valid_df) == 0:
        print("[INFO] No rows to process matching the criteria.")
        return

    token = get_auth_token()
    records = []
    total = len(valid_df)

    print(f"\n--- PHASE 1: Running API & DB Inspection ({max_workers} threads) ---")

    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_idx = {
            executor.submit(process_row, idx, original_idx, row, base_api_url, token, db_pool, check_all_web): idx 
            for idx, (original_idx, row) in enumerate(valid_df.iterrows(), start=1)
        }
        
        completed = 0
        for future in concurrent.futures.as_completed(future_to_idx):
            idx = future_to_idx[future]
            try:
                record = future.result()
                records.append(record)
                completed += 1
                
                print(f"[{completed}/{total}] Checked {record['issue_code']} -> Asset Info: {record['asset_info_provided']} | Resp: '{record['responsibility_name']}' | Color: {record['color_assigned']}", flush=True)
                
                if completed % 50 == 0 or completed == total:
                    sorted_records = sorted(records, key=lambda x: x['excel_row'])
                    save_styled_excel(sorted_records, output_path)
                    
            except Exception as exc:
                print(f"Task {idx} generated an exception: {exc}")

    # Sort Phase 1 records
    records = sorted(records, key=lambda x: x['excel_row'])
    save_styled_excel(records, output_path)

    # --- PHASE 2: Web Cross-Check ---
    if enable_web_check and PLAYWRIGHT_AVAILABLE and WEB_INSPECT_AVAILABLE:
        asyncio.run(
            run_web_cross_check_phase(
                records=records,
                output_path=output_path,
                web_workers=web_workers,
                headless=headless,
            )
        )
    elif enable_web_check and not (PLAYWRIGHT_AVAILABLE and WEB_INSPECT_AVAILABLE):
        print("\n[WARNING] Playwright or web inspect modules not available. Skipping Phase 2 Web Cross-Check.")

    out_df = pd.DataFrame(records)
    cyan_count = (out_df["color_assigned"] == "CYAN").sum() if "color_assigned" in out_df else 0
    yellow_count = (out_df["color_assigned"] == "YELLOW").sum() if "color_assigned" in out_df else 0
    orange_count = (out_df["color_assigned"] == "ORANGE").sum() if "color_assigned" in out_df else 0
    red_count = (out_df["color_assigned"] == "RED").sum() if "color_assigned" in out_df else 0
    blue_count = (out_df["color_assigned"] == "LIGHT BLUE").sum() if "color_assigned" in out_df else 0
    purple_count = (out_df["color_assigned"] == "PURPLE").sum() if "color_assigned" in out_df else 0

    print(f"\n==================== SCAN COMPLETE ====================")
    print(f"Total Rows Checked       : {len(out_df)}")
    print(f"CYAN                     : {cyan_count} (API Asset Provided AND Manufacturer Info Found on Web UI)")
    print(f"YELLOW                   : {yellow_count} (API Asset Provided, BUT Manufacturer Info Missing on Web UI)")
    print(f"PURPLE                   : {purple_count} (DB Mismatch - OUR SIDE CODE FIX NEEDED!)")
    print(f"ORANGE                   : {orange_count} (API Asset UUIDs Null, but Info Found on Web UI -> FACILITYGRID API ISSUE!)")
    print(f"LIGHT BLUE               : {blue_count} (No Asset Info API/Web, but Valid Responsibility Match)")
    print(f"RED                      : {red_count} (No Asset Info API/Web & No Match Responsibility)")
    print(f"Results saved to         : {os.path.abspath(output_path)}")
    print(f"=======================================================\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Scan FG Issues using FacilityGrid API with Web UI Cross-Check")
    parser.add_argument("--file", "-f", default=None, help="Path to the Excel file")
    parser.add_argument("--output", "-o", default=None, help="Output Excel file path")
    parser.add_argument("--limit", "-l", type=int, default=0, help="Number of rows to test (0 for all)")
    parser.add_argument("--workers", "-w", type=int, default=15, help="Number of API concurrent threads (default: 15)")
    parser.add_argument("--no-web-check", action="store_true", help="Disable Phase 2 Playwright Web UI cross-checking")
    parser.add_argument("--web-workers", type=int, default=5, help="Number of parallel Playwright web worker pages (default: 5)")
    parser.add_argument("--headless", action="store_true", help="Run Playwright browser in headless mode")
    parser.add_argument("--phase-2", "--phase-2-only", action="store_true", help="Run Phase 2 (Playwright Web UI Cross-Check) standalone on an input or result Excel file")
    parser.add_argument("--check-all-web", "--check-yellow", action="store_true", default=None, help="Force cross-checking YELLOW rows on Web UI")
    parser.add_argument("--no-check-yellow", action="store_true", help="Skip Web UI cross-checking for YELLOW rows")

    args = parser.parse_args()

    check_yellow = False
    if args.check_all_web:
        check_yellow = True
    elif args.no_check_yellow:
        check_yellow = False

    run_api_scanner(
        excel_path=args.file,
        output_path=args.output,
        limit=args.limit if args.limit > 0 else None,
        max_workers=args.workers,
        enable_web_check=not args.no_web_check,
        web_workers=args.web_workers,
        headless=args.headless,
        phase_2_only=args.phase_2,
        check_all_web=check_yellow,
    )
