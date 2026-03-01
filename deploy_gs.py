#!/usr/bin/env python
"""
deploy_gs.py - Deploy commented .gs formula files to Google Sheets

Reads formula source files (.gs), strips comments, updates deployment metadata,
and writes the cleaned formula to the appropriate cell(s) in a Google Sheet.

Usage:
    python deploy_gs.py "Spreadsheet Name" "Sheet!Column" Column2 "Sheet2!$A$1"
    python deploy_gs.py                          (interactive prompts)
    python deploy_gs.py < deployment.txt         (redirected input)

See README or docstring below for full details.
"""

import os
import re
import sys
import time
import signal
import datetime

import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv

try:
    import pytz
    PYTZ_AVAILABLE = True
except ImportError:
    PYTZ_AVAILABLE = False

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

RETRYABLE_STATUS_CODES = {408, 429, 500, 502, 503, 504}
MAX_RETRIES = 59
RETRY_DELAY = 10  # seconds

# ---------------------------------------------------------------------------
# Graceful Ctrl-C handling
# ---------------------------------------------------------------------------

_shutdown_requested = False

def _sigint_handler(sig, frame):
    global _shutdown_requested
    if _shutdown_requested:
        print("\nForced exit.")
        sys.exit(1)
    _shutdown_requested = True
    print("\nCtrl-C received. Finishing current operation then exiting gracefully...")

signal.signal(signal.SIGINT, _sigint_handler)


def check_shutdown():
    if _shutdown_requested:
        print("Shutting down as requested.")
        sys.exit(0)


# ---------------------------------------------------------------------------
# Retry wrapper
# ---------------------------------------------------------------------------

def with_retry(fn, *args, **kwargs):
    """Call fn(*args, **kwargs), retrying on retryable API errors."""
    for attempt in range(MAX_RETRIES + 1):
        check_shutdown()
        try:
            return fn(*args, **kwargs)
        except gspread.exceptions.APIError as e:
            status = e.response.status_code if hasattr(e, 'response') and e.response else None
            if status in RETRYABLE_STATUS_CODES and attempt < MAX_RETRIES:
                print(f"  [Retryable error {status}] Waiting {RETRY_DELAY}s before retry "
                      f"(attempt {attempt + 1}/{MAX_RETRIES})...")
                for _ in range(RETRY_DELAY):
                    check_shutdown()
                    time.sleep(1)
            else:
                raise
        except Exception:
            raise


# ---------------------------------------------------------------------------
# Comment stripping
# ---------------------------------------------------------------------------

def strip_comments(formula: str) -> str:
    """
    Remove /* ... */ and // ... comments from a formula string.

    Also collapse multiple consecutive blank lines (after comment removal)
    down to a single blank line, and strip leading/trailing blank lines.

    KNOWN LIMITATION: Comment delimiters that appear inside string literals
    will be treated as real comment delimiters. If your formula needs to
    contain the literal strings /* or */ or //, construct them using
    concatenation to avoid accidental stripping:

        LeftCommentDelim,"/"&"*",
        RightCommentDelim,"*"&"/",
        DoubleSlash,"/"&"/",

    The // stripper requires // to appear at the start of a line or be
    preceded by whitespace, which protects URLs like "http://google.com"
    from being clobbered. However, whitespace-preceded // inside a string
    literal would still be stripped.

    Strategy:
      1. Remove block comments /* ... */ (including multi-line).
      2. Remove line comments // ... to end of line.
      3. Collapse runs of blank lines to a single blank line.
      4. Strip leading/trailing whitespace from the result.
    """

    # Step 1: Remove /* ... */ block comments (non-greedy, DOTALL)
    # Also consume any horizontal whitespace immediately preceding the /*
    formula = re.sub(r'[ \t]*/\*.*?\*/', '', formula, flags=re.DOTALL)

    # Step 2: Remove // line comments (to end of line)
    # Also consume any horizontal whitespace immediately preceding the //
    # // must be at start of line or preceded by whitespace to avoid stripping URLs.
    formula = re.sub(r'[ \t]*(?:^|(?<=\s))//[^\n]*', '', formula, flags=re.MULTILINE)

    # Step 3: Collapse runs of blank/whitespace-only lines to at most one blank line
    # A "blank line" here means a line with only whitespace.
    formula = re.sub(r'\n(\s*\n)+', '\n\n', formula)

    # Step 4: Strip leading/trailing whitespace
    formula = formula.strip()

    return formula


# ---------------------------------------------------------------------------
# Metadata substitution
# ---------------------------------------------------------------------------

def update_date_deployed(formula: str, local_tz) -> str:
    """
    Replace the value inside _Date_deployed,N("...") with the current
    datetime formatted as m/d/yyyy hh:mm in the given timezone.
    """
    if local_tz and PYTZ_AVAILABLE:
        now = datetime.datetime.now(tz=local_tz)
    else:
        now = datetime.datetime.now(tz=datetime.timezone.utc)

    month = str(now.month)
    day = str(now.day)
    year = now.strftime('%Y')
    hhmm = now.strftime('%H:%M')
    formatted = f"{month}/{day}/{year} {hhmm}"

    # Replace _Date_deployed,N("anything"),
    # Pattern allows optional whitespace around the value
    pattern = r'(_Date_deployed\s*,\s*N\s*\(\s*")([^"]*)("\s*\)\s*,)'
    replacement = r'\g<1>' + formatted + r'\g<3>'
    new_formula, count = re.subn(pattern, replacement, formula)
    if count == 0:
        print("  [Warning] _Date_deployed marker not found in formula; date not updated.")
    return new_formula


def extract_source_filename(formula: str) -> str | None:
    """
    Extract the filename from _Source,N("<filename>"),
    Returns None if not found.
    """
    pattern = r'_Source\s*,\s*N\s*\(\s*"([^"]+)"\s*\)\s*,'
    m = re.search(pattern, formula)
    return m.group(1) if m else None


# ---------------------------------------------------------------------------
# Input parsing
# ---------------------------------------------------------------------------

def parse_column_refs(raw_refs: list[str]) -> list[dict]:
    """
    Parse a list of column reference strings into structured dicts.

    Formats accepted:
        Sheet!ColumnName        -> sheet="Sheet", col="ColumnName", abs_ref=False
        Sheet!$A$1              -> sheet="Sheet", col="$A$1", abs_ref=True
        ColumnName              -> sheet=None (inherit from previous), col="ColumnName"
        $A$1                    -> sheet=None (inherit), col="$A$1", abs_ref=True

    Returns list of dicts with keys: sheet (str|None), col (str), abs_ref (bool)
    """
    parsed = []
    for ref in raw_refs:
        if '!' in ref:
            sheet_part, col_part = ref.split('!', 1)
        else:
            sheet_part = None
            col_part = ref
        abs_ref = bool(re.match(r'^\$[A-Za-z]+\$\d+$', col_part))
        parsed.append({'sheet': sheet_part, 'col': col_part, 'abs_ref': abs_ref})

    # Fill in inherited sheet names
    current_sheet = None
    for item in parsed:
        if item['sheet'] is not None:
            current_sheet = item['sheet']
        else:
            item['sheet'] = current_sheet

    return parsed


def get_inputs_interactive(spreadsheet_name=None) -> tuple[str, list[str]]:
    """Prompt user for spreadsheet name (if not given) and cell refs interactively."""
    if sys.stdin.isatty():
        if not spreadsheet_name:
            spreadsheet_name = input("Specify Google Sheets spreadsheet name: ").strip()
        print("Specify Cell to update as <sheet name>!<ref> where <ref> is a column name")
        print("or absolute cell reference. You may use just <column name> if it's on the")
        print("same sheet. Press enter after each one.")
        print("Hit Ctrl-Z (Windows) or Ctrl-D (Linux) after the last one:")
    else:
        if not spreadsheet_name:
            spreadsheet_name = sys.stdin.readline().strip()

    col_refs = []
    try:
        for line in sys.stdin:
            line = line.strip()
            if line:
                col_refs.append(line)
    except EOFError:
        pass

    return spreadsheet_name, col_refs


# ---------------------------------------------------------------------------
# Google Sheets helpers
# ---------------------------------------------------------------------------

def open_spreadsheet(gc: gspread.Client, name: str) -> gspread.Spreadsheet:
    return with_retry(gc.open, name)


def batch_get_ranges(spreadsheet: gspread.Spreadsheet, ranges: list[str]) -> list:
    """
    Perform a single batch get for multiple A1 ranges.
    Returns the raw valueRanges list from the API response.
    """
    def _do():
        return spreadsheet.values_batch_get(
            ranges=ranges,
            params={'valueRenderOption': 'FORMULA', 'majorDimension': 'ROWS'}
        )
    return with_retry(_do)


def col_letter_to_index(col_str: str) -> int:
    """Convert a column letter (A, B, ..., Z, AA, ...) to 0-based index."""
    col_str = col_str.upper().strip('$')
    result = 0
    for ch in col_str:
        result = result * 26 + (ord(ch) - ord('A') + 1)
    return result - 1


def col_index_from_name(headers: list[str], col_name: str) -> int | None:
    """Find the 0-based column index of a header name. Case-insensitive."""
    col_name_lower = col_name.lower()
    for i, h in enumerate(headers):
        if h.lower() == col_name_lower:
            return i
    return None


def a1_for_cell(sheet_name: str, row: int, col_index: int) -> str:
    """Build an A1 notation range string for a single cell."""
    col_letter = col_index_to_letter(col_index)
    return f"'{sheet_name}'!{col_letter}{row}"


def col_index_to_letter(index: int) -> str:
    """Convert 0-based column index to column letter(s)."""
    result = ''
    index += 1
    while index > 0:
        index, rem = divmod(index - 1, 26)
        result = chr(ord('A') + rem) + result
    return result


def parse_abs_ref(ref: str) -> tuple[int, int]:
    """
    Parse an absolute cell reference like $A$1 or $AB$23.
    Returns (0-based col index, 1-based row number).
    """
    m = re.match(r'^\$([A-Za-z]+)\$(\d+)$', ref)
    if not m:
        raise ValueError(f"Cannot parse absolute reference: {ref}")
    col_idx = col_letter_to_index(m.group(1))
    row = int(m.group(2))
    return col_idx, row


# ---------------------------------------------------------------------------
# Main logic
# ---------------------------------------------------------------------------

def main():
    load_dotenv()

    # --- Credentials and timezone from .env ---
    creds_file = os.getenv('GOOGLE_CREDENTIALS', 'google_credentials.json')
    tz_name = os.getenv('LOCAL_TIMEZONE', '').strip().strip('"')

    local_tz = None
    if tz_name:
        if PYTZ_AVAILABLE:
            try:
                local_tz = pytz.timezone(tz_name)
            except pytz.UnknownTimeZoneError:
                print(f"[Warning] Unknown timezone '{tz_name}' in .env; using UTC.")
        else:
            print("[Warning] pytz not available; using UTC.")
    else:
        print("[Warning] LOCAL_TIMEZONE not set in .env; using UTC.")

    # --- Parse command-line arguments ---
    if len(sys.argv) > 1:
        spreadsheet_name = sys.argv[1]
        raw_refs = sys.argv[2:]
        if not raw_refs:
            # Spreadsheet given but no columns - prompt or read from stdin
            _, raw_refs = get_inputs_interactive(spreadsheet_name)
    else:
        spreadsheet_name, raw_refs = get_inputs_interactive()

    if not spreadsheet_name:
        print("Error: No spreadsheet name provided.")
        sys.exit(1)

    if not raw_refs:
        print("Error: No column references provided.")
        sys.exit(1)

    col_refs = parse_column_refs(raw_refs)

    # --- Connect to Google Sheets ---
    try:
        creds = Credentials.from_service_account_file(creds_file, scopes=SCOPES)
        gc = gspread.authorize(creds)
    except FileNotFoundError:
        print(f"Error: Credentials file '{creds_file}' not found.")
        sys.exit(1)
    except Exception as e:
        print(f"Error loading credentials: {e}")
        sys.exit(1)

    print(f"\nOpening spreadsheet: {spreadsheet_name}")
    try:
        spreadsheet = open_spreadsheet(gc, spreadsheet_name)
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"Error: Spreadsheet '{spreadsheet_name}' not found or not accessible.")
        sys.exit(1)
    except gspread.exceptions.APIError as e:
        print(f"Error opening spreadsheet: {e}")
        sys.exit(1)

    check_shutdown()

    # --- Group refs by sheet ---
    # For header lookups we need row 1 of each sheet that has non-abs-ref cols.
    # For abs refs we need the specific cell.

    # Collect unique sheets with non-abs columns -> need their row 1
    sheets_needing_headers = {}   # sheet_name -> list of col item dicts (non-abs)
    abs_ref_items = []            # list of col item dicts (abs_ref=True)

    for item in col_refs:
        if item['sheet'] is None:
            print(f"[Warning] Could not determine sheet for '{item['col']}'; skipping.")
            continue
        if item['abs_ref']:
            abs_ref_items.append(item)
        else:
            sheets_needing_headers.setdefault(item['sheet'], []).append(item)

    # --- BATCH GET 1: All header rows for sheets with named columns ---
    header_ranges = [f"'{sname}'!1:1" for sname in sheets_needing_headers]

    header_data = {}   # sheet_name -> list of header strings

    if header_ranges:
        print(f"\nFetching {len(header_ranges)} header row(s) in one batch...")
        try:
            result = batch_get_ranges(spreadsheet, header_ranges)
            value_ranges = result.get('valueRanges', [])
            for i, sheet_name in enumerate(sheets_needing_headers.keys()):
                if i < len(value_ranges):
                    vr = value_ranges[i]
                    rows = vr.get('values', [])
                    header_data[sheet_name] = rows[0] if rows else []
                else:
                    header_data[sheet_name] = []
        except gspread.exceptions.APIError as e:
            print(f"Error fetching header rows: {e}")
            sys.exit(1)

    # --- Build cell ranges for BATCH GET 2: All formula cells ---
    # For named columns: row 2 of the matching column.
    # For abs refs: the specified cell.

    formula_ranges = []   # A1 notation strings
    formula_meta = []     # parallel list of (sheet_name, display_name, is_abs, col_idx, row)

    for sheet_name, items in sheets_needing_headers.items():
        headers = header_data.get(sheet_name, [])
        for item in items:
            col_name = item['col']
            col_idx = col_index_from_name(headers, col_name)
            if col_idx is None:
                print(f"[Warning] Column '{col_name}' not found in sheet '{sheet_name}'; skipping.")
                continue
            a1 = a1_for_cell(sheet_name, 2, col_idx)
            formula_ranges.append(a1)
            formula_meta.append({
                'sheet': sheet_name,
                'display': f"{sheet_name}!{col_name}",
                'col_idx': col_idx,
                'row': 2,
                'a1': a1,
            })

    for item in abs_ref_items:
        sheet_name = item['sheet']
        ref = item['col']
        try:
            col_idx, row = parse_abs_ref(ref)
        except ValueError as e:
            print(f"[Warning] {e}; skipping.")
            continue
        # Build A1 with the literal reference (it's already absolute)
        a1 = f"'{sheet_name}'!{ref}"
        formula_ranges.append(a1)
        formula_meta.append({
            'sheet': sheet_name,
            'display': f"{sheet_name}!{ref}",
            'col_idx': col_idx,
            'row': row,
            'a1': a1,
        })

    if not formula_ranges:
        print("\nNo valid cells to process. Exiting.")
        sys.exit(0)

    # --- BATCH GET 2: All formula cells ---
    print(f"Fetching {len(formula_ranges)} formula cell(s) in one batch...")
    try:
        result = batch_get_ranges(spreadsheet, formula_ranges)
        value_ranges = result.get('valueRanges', [])
    except gspread.exceptions.APIError as e:
        print(f"Error fetching formula cells: {e}")
        sys.exit(1)

    check_shutdown()

    # --- Process each formula ---
    write_data = []   # list of (a1_notation, new_formula)

    for i, meta in enumerate(formula_meta):
        display = meta['display']
        print(f"\nProcessing: {display}")

        vr = value_ranges[i] if i < len(value_ranges) else {}
        rows = vr.get('values', [])
        if not rows or not rows[0]:
            print(f"  [Warning] Cell {display} is empty; skipping.")
            continue

        cell_formula = rows[0][0]

        # Find _Source filename
        source_file = extract_source_filename(cell_formula)
        if not source_file:
            print(f"  [Warning] No _Source,N(\"...\"), marker found in formula; skipping.")
            continue

        print(f"  Source file: {source_file}")

        # Read the source .gs file
        if not os.path.exists(source_file):
            print(f"  [Error] Source file '{source_file}' not found; skipping.")
            continue

        try:
            with open(source_file, 'r', encoding='utf-8') as f:
                gs_content = f.read()
        except IOError as e:
            print(f"  [Error] Cannot read '{source_file}': {e}; skipping.")
            continue

        # Strip comments
        cleaned = strip_comments(gs_content)
        print(f"  Comments stripped. Formula length: {len(gs_content)} -> {len(cleaned)} chars.")

        # Update _Date_deployed
        cleaned = update_date_deployed(cleaned, local_tz)

        write_data.append((meta['a1'], cleaned))

    if not write_data:
        print("\nNo formulas to write. Exiting.")
        sys.exit(0)

    # --- BATCH WRITE: All updated formulas at once ---
    check_shutdown()
    print(f"\nWriting {len(write_data)} formula(s) in one batch...")

    # gspread batch_update takes a list of dicts with 'range' and 'values'
    data_payload = [
        {'range': a1, 'values': [[formula]]}
        for a1, formula in write_data
    ]

    def _do_batch_write():
        spreadsheet.values_batch_update(
            {
                'valueInputOption': 'USER_ENTERED',
                'data': data_payload,
            }
        )

    try:
        with_retry(_do_batch_write)
        print("  Done.")
    except gspread.exceptions.APIError as e:
        print(f"  [Error] Batch write failed after {MAX_RETRIES} retries: {e}")
        sys.exit(1)

    print("\nDeployment complete.")


if __name__ == '__main__':
    main()
