"""
Quick verification script -- tests the core MUFAP pipeline logic
against the live MUFAP website WITHOUT sending any emails.
"""
import sys
import io
import os
import logging
from datetime import datetime
# Fix Windows console encoding for Unicode
if sys.platform.startswith("win"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from main import (
    fetch_mufap_page,
    extract_report_date,
    parse_mufap_table,
    build_excel_file,
    build_email_html,
    MUFAP_URL,
)
from bs4 import BeautifulSoup
print("=" * 60)
print("MUFAP Pipeline -- Verification Test (Local Dry-Run)")
print("=" * 60)
print(f"\nCurrent System Time : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
print(f"Target MUFAP URL    : {MUFAP_URL}")
# -- Test 1: Fetch Live Page & Extract Report Date --
print(f"\n{'-' * 40}")
print("TEST 1: Fetching live MUFAP page and checking Report Date...")
try:
    html_content = fetch_mufap_page(MUFAP_URL)
    soup = BeautifulSoup(html_content, "html.parser")
    report_date = extract_report_date(soup)

    print("  [OK] Successfully fetched MUFAP HTML content!")
    print(f"  Page Size          : {len(html_content) / 1024:.1f} KB")
    print(f"  Extracted Date     : {report_date}")
    # -- Test 2: Table Parsing & Filtering AL Habib Funds --
    print(f"\n{'-' * 40}")
    print("TEST 2: Parsing table and filtering 'AL Habib' funds...")

    data_rows, headers = parse_mufap_table(soup, keyword="AL Habib")
    print(f"  [OK] Table parsed successfully!")
    print(f"  Headers Found      : {headers[:5]}...")
    print(f"  AL Habib Funds Count: {len(data_rows)}")

    if data_rows:
        print("\n  Sample Scraped Fund Data (First 5 entries):")
        for i, row in enumerate(data_rows[:5], 1):
            sector = row.get("Sector", "N/A")
            category = row.get("Category", "N/A")
            fund_name = row.get("Fund Name", "N/A")
            print(f"   {i}. [{sector} / {category}] -> {fund_name}")

        # Sanity check: make sure categories are NOT all identical
        # (this was the original bug -- everything collapsing into
        # "Money Market" regardless of the fund's real category)
        distinct_categories = {row.get("Category", "N/A") for row in data_rows}
        print(f"\n  Distinct categories detected: {len(distinct_categories)}")
        for cat in sorted(distinct_categories):
            print(f"   - {cat}")
        if len(distinct_categories) <= 1 and len(data_rows) > 1:
            print("  [WARN] All rows share the same category -- parsing may still be wrong!")
        else:
            print("  [OK] Categories vary correctly across funds.")
    else:
        print("  [WARN] No funds containing 'AL Habib' were detected on the live table.")
    # -- Test 3: Test Excel Generation --
    print(f"\n{'-' * 40}")
    print("TEST 3: Generating test Excel file...")

    test_excel_name = "TEST_AL_Habib_MUFAP_Report.xlsx"
    build_excel_file(data_rows, headers, filename=test_excel_name, report_date=report_date)

    if os.path.exists(test_excel_name):
        file_size = os.path.getsize(test_excel_name) / 1024
        print(f"  [OK] Excel file generated successfully: {test_excel_name} ({file_size:.1f} KB)")
        # Clean up temporary test file
        os.remove(test_excel_name)
        print("  [INFO] Temporary test Excel file cleaned up.")
    # -- Test 4: Verify Email HTML Template Render --
    print(f"\n{'-' * 40}")
    print("TEST 4: Rendering HTML Email Body template...")

    email_html = build_email_html(data_rows, headers, report_date)
    print(f"  [OK] HTML Email rendered successfully! Length: {len(email_html)} characters.")
except Exception as exc:
    print(f"  [FAIL] Test encountered error: {exc}")
    import traceback
    traceback.print_exc()
print(f"\n{'-' * 40}")
print("TEST 5: Email Send Execution (SKIPPED -- Run main.py with valid credentials to send emails)")
print(f"\n{'=' * 60}")
print("VERIFICATION COMPLETED")
print("=" * 60)