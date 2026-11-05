import os
import logging
import base64
import json
import re
import time
from datetime import datetime
from zoneinfo import ZoneInfo

VANCOUVER_TZ = ZoneInfo("America/Vancouver")

def now_local():
    return datetime.now(VANCOUVER_TZ)

import anthropic
import gspread
from google.oauth2.service_account import Credentials
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler, CallbackQueryHandler,
    filters, ContextTypes
)

# ─── Logging ────────────────────────────────────────────────────────────────
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# ─── Config from env ────────────────────────────────────────────────────────
TELEGRAM_TOKEN    = os.environ["TELEGRAM_TOKEN"]
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]
SPREADSHEET_ID    = os.environ["SPREADSHEET_ID"]
ALLOWED_USER_ID   = int(os.environ["ALLOWED_USER_ID"])

# ─── Google Sheets setup ────────────────────────────────────────────────────
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

MONTH_NAMES_EN = {
    1: "January", 2: "February", 3: "March", 4: "April",
    5: "May", 6: "June", 7: "July", 8: "August",
    9: "September", 10: "October", 11: "November", 12: "December"
}

# Income rows in the "Шкатулки" block of each monthly sheet (column A = label, column B = amount)
SALARY_ROWS = {
    "1": 6,   # "Зарплата  1"
    "2": 7,   # "Зарплата  2"
    "3": 10,  # "Зарплата  3"
}
BONUS_ROW = 8  # "Бонусы"

# Daily expense table columns on the monthly sheet (date is pre-filled per row)
EXPENSE_DATE_COL = "AE"
EXPENSE_DESC_COL = "AF"
EXPENSE_AMOUNT_COL = "AG"

# Expense categories that accumulate row-by-row on the month sheet (e.g. 'June  2026'),
# in the same style as the wallet blocks: a header cell with the category name,
# then date+description / amount rows below, with a 'Total' row further down.
#
# The 6 "fixed monthly payment" categories (Rent, Phone, BC Hydro, Netflix,
# YouTube Music, Car Insurance) used to live only on the FINANCE sheet, but
# the June 2026 sheet ALSO has real row-based blocks for them (column N/O,
# same Total-row-above / header / accumulating-data-rows-below pattern as
# every other category here), confirmed via /debugcategory. As of June 2026
# they're written here instead, to stop the FINANCE-vs-month-sheet confusion.
ROW_BASED_CATEGORIES = [
    "Groceries", "Alcohol", "Fuel", "FHSA", "TFSA", "Restaurant", "Bowling",
    "GIM", "Skiing", "Medication", "Hair Cut", "Dining Out", "Clothing",
    "Sauna", "Audiobooks", "Car loan", "Car Wash", "Charity", "Gifts",
    "Tim Hortons", "Intertainment", "Saving", "FEE", "Housing", "Travel",
    "Rent", "Phone", "BC Hydro", "Netflix", "YouTube Music", "Car Insurance",
    "Parking",
]

# Formerly: fixed monthly-payment categories written as a single cumulative
# cell (column D, "Actual cost") on the 'FINANCE - <Month> <Year>' sheet.
# As of June 2026 this list is empty and the bot no longer writes to the
# FINANCE sheet at all — see ROW_BASED_CATEGORIES above. write_fixed_category_entry()
# is kept below, unused, in case this ever needs to be reverted.
FIXED_CATEGORIES = []

# Income categories that now accumulate row-by-row in the 'INCOM' block on the
# month sheet (e.g. 'June  2026'), same style as expense categories: a header
# cell with the category name, then date+description / amount rows below,
# with a 'Total' row further down (which is then pulled by formula elsewhere).
INCOME_CATEGORIES = ["Зарплата", "Points RBC", "Бонус", "Инвойс"]

# English display names for income categories, used only in Telegram messages.
# The sheet itself still uses the Russian labels above for lookups/writes.
INCOME_CATEGORY_DISPLAY_EN = {
    "Зарплата": "Salary",
    "Бонус": "Bonus",
    "Инвойс": "Invoice",
    "Points RBC": "Points RBC",
}


def display_category(category: str) -> str:
    """Return the English label to show in Telegram for a given internal
    (sheet-matching) category name. Falls back to the raw name if there's
    no mapping (e.g. expense categories, which are already English)."""
    return INCOME_CATEGORY_DISPLAY_EN.get(category, category)

ALL_EXPENSE_CATEGORIES = ROW_BASED_CATEGORIES + FIXED_CATEGORIES

# Column layout for the fixed-payment table on the FINANCE sheet
FIXED_COL_CATEGORY = 2       # column B
FIXED_COL_PROJECTED = 3      # column C ("Projected cost") — never written by the bot
FIXED_COL_ACTUAL = 4         # column D ("Actual cost") — bot adds to this

# ─── Wallets ────────────────────────────────────────────────────────────────
# Column A = wallet name (header row) / date+description (data rows)
# Column B = income, Column C = outcome, Column D = Transfer.
# Header row numbers are looked up dynamically (see find_wallet_header_row)
# rather than hardcoded, since the sheet can be edited elsewhere and rows shift.
WALLET_COL_DATE_DESC = 1  # column A
WALLET_COL_INCOME    = 2  # column B
WALLET_COL_OUTCOME   = 3  # column C
WALLET_COL_TRANSFER  = 4  # column D

WALLET_NAMES = ["RBC Credit", "RBC Checking", "Walmart", "CIBC", "Savings", "Cash"]

# Cache of {sheet_title: {wallet_name: header_row}} to avoid re-scanning the whole
# column on every single write within the same process lifetime. Cleared per-process;
# Railway restarts will naturally refresh it.
_wallet_row_cache: dict = {}


def _sheets_call(fn, *args, **kwargs):
    """Thin wrapper around gspread calls — kept so all API calls go through
    one place (easy to add logging or swap transport later). No blocking
    sleep here: the reduced call count (get_all_values + batch_update) is
    the main guard against 429 rate-limit errors."""
    return fn(*args, **kwargs)


def find_wallet_header_row_from_col(col_a_values: list, wallet: str, sheet_title: str) -> int:
    """Find the header row of a wallet inside a pre-fetched column-A list
    (1-indexed). Updates _wallet_row_cache as a side-effect."""
    for idx, val in enumerate(col_a_values, start=1):
        if (val or "").strip().lower() == wallet.lower():
            _wallet_row_cache.setdefault(sheet_title, {})[wallet] = idx
            return idx
    raise ValueError(
        f"Couldn't find wallet '{wallet}' on sheet '{sheet_title}'. "
        f"Check that the name in column A is spelled exactly the same way."
    )


def find_wallet_header_row(ws, wallet: str) -> int:
    """Find the row where column A exactly equals the wallet's name.
    Searches dynamically instead of relying on a fixed row number, since rows
    can shift if the sheet is edited elsewhere (e.g. budget section above)."""
    cache_key = ws.title
    cached = _wallet_row_cache.get(cache_key, {})
    if wallet in cached:
        # Verify the cached row is still correct before trusting it
        cell_val = (_sheets_call(ws.cell, cached[wallet], WALLET_COL_DATE_DESC).value or "").strip()
        if cell_val.lower() == wallet.lower():
            return cached[wallet]

    col_values = _sheets_call(ws.col_values, WALLET_COL_DATE_DESC)
    return find_wallet_header_row_from_col(col_values, wallet, ws.title)


def get_sheet():
    """Connect to Google Sheets and return the spreadsheet object."""
    google_creds_json = os.environ["GOOGLE_CREDENTIALS_JSON"]
    creds_dict = json.loads(google_creds_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    gc = gspread.authorize(creds)
    return gc.open_by_key(SPREADSHEET_ID)


def get_month_worksheet(spreadsheet):
    """Find the daily-tracker worksheet for the current month, tolerant of spacing
    differences (e.g. 'June 2026' vs 'June  2026'). Skips 'BUDGET - ...' sheets."""
    now = now_local()
    month_name = MONTH_NAMES_EN[now.month]
    year = str(now.year)

    for ws in spreadsheet.worksheets():
        title = ws.title
        if "budget" in title.lower() or "finance" in title.lower():
            continue
        if month_name.lower() in title.lower() and year in title:
            return ws

    raise ValueError(
        f"Couldn't find a sheet for {month_name} {year}. "
        f"Create a tab based on the BUDGET template."
    )


def get_budget_worksheet(spreadsheet):
    """Find the 'BUDGET - <Month> <Year>' (now possibly renamed to
    'FINANCE - <Month> <Year>') worksheet for the current month.
    Currently unused by the bot's write path — kept for future reference/reads."""
    now = now_local()
    month_name = MONTH_NAMES_EN[now.month]
    year = str(now.year)

    for ws in spreadsheet.worksheets():
        title = ws.title
        title_lower = title.lower()
        if ("budget" in title_lower or "finance" in title_lower) and \
           month_name.lower() in title_lower and year in title:
            return ws

    raise ValueError(
        f"Couldn't find sheet 'BUDGET/FINANCE - {month_name} {year}'. "
        f"Create it based on last month's template."
    )


def col_letter_to_index(col_letter: str) -> int:
    """Convert column letter (e.g. 'AG') to 1-based column index."""
    idx = 0
    for ch in col_letter:
        idx = idx * 26 + (ord(ch.upper()) - ord('A') + 1)
    return idx


def find_next_wallet_row(ws, header_row: int) -> int:
    """Find the first empty data row under a wallet's header row.
    A row is considered empty if column A (date+description) is blank.
    Stops scanning after a generous range to avoid runaway loops."""
    row = header_row + 1
    max_row = header_row + 300  # generous safety bound
    while row <= max_row:
        cell_value = _sheets_call(ws.cell, row, WALLET_COL_DATE_DESC).value
        if not cell_value or not str(cell_value).strip():
            return row
        row += 1
    raise ValueError("Couldn't find a free row in the wallet block (block is full).")


def find_next_wallet_row_from_snapshot(all_values: list, header_row: int) -> int:
    """Same as find_next_wallet_row but uses a pre-fetched get_all_values()
    snapshot, avoiding extra API read calls."""
    row = header_row + 1
    max_row = header_row + 300
    while row <= max_row:
        if row - 1 < len(all_values):
            cell_value = all_values[row - 1][WALLET_COL_DATE_DESC - 1] if len(all_values[row - 1]) >= WALLET_COL_DATE_DESC else ""
        else:
            cell_value = ""
        if not cell_value or not str(cell_value).strip():
            return row
        row += 1
    raise ValueError("Couldn't find a free row in the wallet block (block is full).")


def verify_wallet_labels(ws, header_row: int):
    """Safety check: confirm the wallet's header row itself has the
    income/outcome/Transfer labels in columns B/C/D (same row as the wallet
    name in column A), as a sanity check on sheet structure."""
    income_label = (_sheets_call(ws.cell, header_row, WALLET_COL_INCOME).value or "").strip().lower()
    outcome_label = (_sheets_call(ws.cell, header_row, WALLET_COL_OUTCOME).value or "").strip().lower()
    if "income" not in income_label or "outcome" not in outcome_label:
        raise ValueError(
            f"The sheet structure has changed: row {header_row} doesn't have "
            f"income/outcome next to the wallet name. Write cancelled, nothing was changed."
        )


def verify_wallet_labels_from_snapshot(all_values: list, header_row: int):
    """Same as verify_wallet_labels but from a pre-fetched snapshot."""
    row_vals = all_values[header_row - 1] if header_row - 1 < len(all_values) else []
    income_label = (row_vals[WALLET_COL_INCOME - 1] if len(row_vals) >= WALLET_COL_INCOME else "").strip().lower()
    outcome_label = (row_vals[WALLET_COL_OUTCOME - 1] if len(row_vals) >= WALLET_COL_OUTCOME else "").strip().lower()
    if "income" not in income_label or "outcome" not in outcome_label:
        raise ValueError(
            f"The sheet structure has changed: row {header_row} doesn't have "
            f"income/outcome next to the wallet name. Write cancelled, nothing was changed."
        )


def write_wallet_entry(wallet: str, kind: str, amount: float, description: str):
    """Write an income or outcome entry into the wallet's block on the current
    month's worksheet. Uses a single get_all_values() snapshot for all reads,
    then writes with batch_update to minimize API calls."""
    if wallet not in WALLET_NAMES:
        raise ValueError(f"Unknown wallet: {wallet}")

    spreadsheet = _sheets_call(get_sheet)
    ws = get_month_worksheet(spreadsheet)
    all_values = _sheets_call(ws.get_all_values)

    header_row = find_wallet_header_row_from_col(
        [row[WALLET_COL_DATE_DESC - 1] if row else "" for row in all_values],
        wallet, ws.title
    )
    verify_wallet_labels_from_snapshot(all_values, header_row)
    target_row = find_next_wallet_row_from_snapshot(all_values, header_row)

    date_str = now_local().strftime("%d.%m")
    label = f"{date_str}  {description}"
    col = WALLET_COL_INCOME if kind == "income" else WALLET_COL_OUTCOME

    _sheets_call(ws.batch_update, [
        {"range": gspread.utils.rowcol_to_a1(target_row, WALLET_COL_DATE_DESC),
         "values": [[label]]},
        {"range": gspread.utils.rowcol_to_a1(target_row, col),
         "values": [[amount]]},
    ])

    return ws.title, target_row


def write_transfer_entry(from_wallet: str, to_wallet: str, amount: float):
    """Record a transfer between two wallets on the current month's worksheet:
    - from_wallet gets a row with the amount in column D (Transfer).
    - to_wallet gets a row with the amount in column B (Income).
    Uses a single get_all_values() snapshot + batch_update to stay within
    Google Sheets API rate limits. Returns (sheet_title, from_row, to_row)."""
    if from_wallet not in WALLET_NAMES:
        raise ValueError(f"Unknown wallet: {from_wallet}")
    if to_wallet not in WALLET_NAMES:
        raise ValueError(f"Unknown wallet: {to_wallet}")
    if from_wallet == to_wallet:
        raise ValueError("Source wallet and destination wallet are the same.")

    spreadsheet = _sheets_call(get_sheet)
    ws = get_month_worksheet(spreadsheet)
    all_values = _sheets_call(ws.get_all_values)   # ← ONE read for entire sheet

    col_a = [row[WALLET_COL_DATE_DESC - 1] if row else "" for row in all_values]

    from_header = find_wallet_header_row_from_col(col_a, from_wallet, ws.title)
    to_header   = find_wallet_header_row_from_col(col_a, to_wallet,   ws.title)
    verify_wallet_labels_from_snapshot(all_values, from_header)
    verify_wallet_labels_from_snapshot(all_values, to_header)

    from_row = find_next_wallet_row_from_snapshot(all_values, from_header)

    # Mark from_row as occupied so to_wallet lookup doesn't land on the same row
    while len(all_values) < from_row:
        all_values.append([])
    if len(all_values[from_row - 1]) < WALLET_COL_DATE_DESC:
        all_values[from_row - 1] += [""] * (WALLET_COL_DATE_DESC - len(all_values[from_row - 1]))
    all_values[from_row - 1][WALLET_COL_DATE_DESC - 1] = "_reserved_"

    to_row = find_next_wallet_row_from_snapshot(all_values, to_header)

    date_str = now_local().strftime("%d.%m")

    _sheets_call(ws.batch_update, [       # ← ONE write for all 4 cells
        {"range": gspread.utils.rowcol_to_a1(from_row, WALLET_COL_DATE_DESC),
         "values": [[f"{date_str}  Transfer"]]},
        {"range": gspread.utils.rowcol_to_a1(from_row, WALLET_COL_TRANSFER),
         "values": [[amount]]},
        {"range": gspread.utils.rowcol_to_a1(to_row, WALLET_COL_DATE_DESC),
         "values": [[f"{date_str}  Transfer"]]},
        {"range": gspread.utils.rowcol_to_a1(to_row, WALLET_COL_INCOME),
         "values": [[amount]]},
    ])

    return ws.title, from_row, to_row


# ─── Expense category handling ───────────────────────────────────────────────

# Cache of {sheet_title: {category_name: (header_row, header_col)}} to avoid
# re-scanning a wide column range on every write within the same process.
_category_cell_cache: dict = {}


def find_category_header_row(all_values, sheet_title: str, category: str):
    """Find the (row, col) of a row-based category header (expense or income,
    e.g. 'Groceries' or 'Зарплата') within an already-fetched get_all_values()
    snapshot of the sheet (list of rows, each a list of cell strings).

    Some sheets also have a SUMMARY table listing category names next to
    their totals (e.g. an 'INCOM' overview block) — those rows have a value
    already in the adjacent cell. We only want the category's OWN block
    header, which is followed by an empty cell (just like 'Groceries' has
    nothing in the cell next to it). So among all text matches, we pick the
    first one whose adjacent cell (same row, next column) is empty.

    Returns (row, col), 1-indexed."""
    cache_key = sheet_title
    cached = _category_cell_cache.get(cache_key, {})
    if category in cached:
        row, col = cached[category]
        if row - 1 < len(all_values):
            row_vals = all_values[row - 1]
            cell_val = (row_vals[col - 1] if col - 1 < len(row_vals) else "") or ""
            adjacent_val = row_vals[col] if col < len(row_vals) else ""
            if cell_val.strip().lower() == category.lower() and not str(adjacent_val).strip():
                return row, col

    candidates = []
    for r_idx, row_vals in enumerate(all_values, start=1):
        for c_idx, val in enumerate(row_vals[:26], start=1):
            if (val or "").strip().lower() == category.lower():
                candidates.append((r_idx, c_idx))

    if not candidates:
        raise ValueError(
            f"Couldn't find category '{category}' on sheet '{sheet_title}' (columns A-Z)."
        )

    # Prefer a match whose adjacent cell is empty (own block header), not a
    # summary-table row that already has a value/total next to it.
    for row, col in candidates:
        row_vals = all_values[row - 1]
        adjacent_val = row_vals[col] if col < len(row_vals) else ""
        if not str(adjacent_val).strip():
            _category_cell_cache.setdefault(cache_key, {})[category] = (row, col)
            return row, col

    # No candidate had an empty adjacent cell — fall back to the first match
    # rather than failing outright, but this is unexpected.
    row, col = candidates[0]
    _category_cell_cache.setdefault(cache_key, {})[category] = (row, col)
    return row, col


def find_next_category_row(all_values, header_row: int, col: int) -> int:
    """Find the first empty data row under a category header, using the same
    pre-fetched get_all_values() snapshot. Scans the description column
    (same column as the header)."""
    row = header_row + 1
    max_row = header_row + 100  # generous safety bound
    while row <= max_row:
        if row - 1 < len(all_values):
            row_vals = all_values[row - 1]
            cell_value = row_vals[col - 1] if col - 1 < len(row_vals) else ""
        else:
            cell_value = ""
        if not cell_value or not str(cell_value).strip():
            return row
        row += 1
    raise ValueError("Couldn't find a free row in the category block (block is full).")


def write_row_category_entry(category: str, amount: float, description: str):
    """Write an entry into a row-based category block (expense or income) on
    the current month's worksheet (e.g. 'June 2026'). Mirrors the wallet
    entry format: description column gets 'DD.MM. description', amount
    column (next column over) gets the amount."""
    if category not in ROW_BASED_CATEGORIES and category not in INCOME_CATEGORIES:
        raise ValueError(f"Unknown row-based category: {category}")

    spreadsheet = _sheets_call(get_sheet)
    ws = get_month_worksheet(spreadsheet)
    all_values = _sheets_call(ws.get_all_values)  # single API call for the whole sheet

    header_row, header_col = find_category_header_row(all_values, ws.title, category)
    target_row = find_next_category_row(all_values, header_row, header_col)

    date_str = now_local().strftime("%d.%m.")
    label = f"{date_str} {description}"

    _sheets_call(ws.update_cell, target_row, header_col, label)
    _sheets_call(ws.update_cell, target_row, header_col + 1, amount)

    return ws.title, target_row, header_col


def append_comment_to_cell(sheet_title: str, row: int, col: int, comment: str):
    """Append a user comment to an already-written description cell, e.g.
    turning 'DD.MM. Groceries' into 'DD.MM. Groceries [comment]'."""
    spreadsheet = _sheets_call(get_sheet)
    ws = spreadsheet.worksheet(sheet_title)
    existing = _sheets_call(ws.cell, row, col).value or ""
    new_value = f"{existing} [{comment}]" if existing else f"[{comment}]"
    _sheets_call(ws.update_cell, row, col, new_value)


def write_fixed_category_entry(category: str, amount: float):
    """Add `amount` to the 'Actual cost' cell (column D) of a fixed monthly
    payment category (e.g. 'Rent') on the 'FINANCE - <Month> <Year>' sheet.
    Adds to whatever is already there rather than overwriting."""
    if category not in FIXED_CATEGORIES:
        raise ValueError(f"Unknown fixed category: {category}")

    spreadsheet = _sheets_call(get_sheet)
    ws = get_budget_worksheet(spreadsheet)

    found_row = None
    col_values = _sheets_call(ws.col_values, FIXED_COL_CATEGORY)
    for idx, val in enumerate(col_values, start=1):
        if (val or "").strip().lower() == category.lower():
            found_row = idx
            break

    if found_row is None:
        raise ValueError(
            f"Couldn't find category '{category}' in column B on sheet '{ws.title}'."
        )

    existing_raw = _sheets_call(ws.cell, found_row, FIXED_COL_ACTUAL).value or "0"
    existing_clean = str(existing_raw).replace("$", "").strip().replace(",", ".")
    try:
        existing_amount = float(existing_clean) if existing_clean else 0.0
    except ValueError:
        existing_amount = 0.0

    new_amount = existing_amount + amount
    _sheets_call(ws.update_cell, found_row, FIXED_COL_ACTUAL, new_amount)

    return ws.title, found_row


# ─── Income handling ─────────────────────────────────────────────────────────

def write_salary(salary_number: str, amount: float):
    """Write amount into the correct 'Зарплата N' cell of the current month's sheet."""
    spreadsheet = _sheets_call(get_sheet)
    ws = get_month_worksheet(spreadsheet)
    row = SALARY_ROWS.get(salary_number)
    if row is None:
        raise ValueError(f"Unknown salary number: {salary_number}")
    _sheets_call(ws.update_cell, row, 2, amount)  # column B
    return ws.title, row


def write_bonus(amount: float):
    spreadsheet = _sheets_call(get_sheet)
    ws = get_month_worksheet(spreadsheet)
    _sheets_call(ws.update_cell, BONUS_ROW, 2, amount)
    return ws.title


# ─── Expense handling ────────────────────────────────────────────────────────

def write_expense(description: str, amount: float):
    """Write an expense into today's row of the daily expense table."""
    spreadsheet = _sheets_call(get_sheet)
    ws = get_month_worksheet(spreadsheet)

    today = now_local().day
    desc_col_idx = col_letter_to_index(EXPENSE_DESC_COL)
    amount_col_idx = col_letter_to_index(EXPENSE_AMOUNT_COL)

    # Dates are pre-filled sequentially starting at row 2 = day 1
    target_row = today + 1

    existing_desc = _sheets_call(ws.cell, target_row, desc_col_idx).value or ""
    existing_amount = _sheets_call(ws.cell, target_row, amount_col_idx).value or ""

    new_desc = f"{existing_desc}, {description}" if existing_desc else description
    try:
        new_amount = float(str(existing_amount).replace(",", ".") or 0) + amount
    except ValueError:
        new_amount = amount

    _sheets_call(ws.update_cell, target_row, desc_col_idx, new_desc)
    _sheets_call(ws.update_cell, target_row, amount_col_idx, new_amount)
    return ws.title, target_row


# ─── Anthropic AI helpers ────────────────────────────────────────────────────

def analyze_receipt_with_claude(image_bytes: bytes) -> dict:
    """Send receipt image to Claude and get structured expense data."""
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    image_b64 = base64.standard_b64encode(image_bytes).decode("utf-8")
    categories_str = "\n".join(f"- {c}" for c in ALL_EXPENSE_CATEGORIES)

    prompt = f"""You are a financial assistant analyzing a receipt photo.

Extract the total amount spent and a short description in English (store name + main items).
Pick the single best matching category from this list:
{categories_str}

Return ONLY a JSON object, no markdown, no explanation:
{{"description": "...", "amount": 0.00, "category": "..."}}"""

    response = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=500,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": image_b64}},
                {"type": "text", "text": prompt}
            ]
        }]
    )
    text = response.content[0].text.strip()
    text = re.sub(r"^```json|```$", "", text, flags=re.MULTILINE).strip()
    return json.loads(text)


# ─── Wallet keyboard helper ──────────────────────────────────────────────────

def wallet_keyboard(prefix: str) -> InlineKeyboardMarkup:
    """Build an inline keyboard with one button per wallet.
    prefix identifies which pending operation this selection belongs to."""
    buttons = [
        [InlineKeyboardButton(name, callback_data=f"{prefix}|{name}")]
        for name in WALLET_NAMES
    ]
    return InlineKeyboardMarkup(buttons)


def wallet_keyboard_excluding(prefix: str, exclude: str) -> InlineKeyboardMarkup:
    """Same as wallet_keyboard, but leaves out one wallet (used for the
    'transfer to' step, so you can't pick the same wallet as source)."""
    buttons = [
        [InlineKeyboardButton(name, callback_data=f"{prefix}|{name}")]
        for name in WALLET_NAMES if name != exclude
    ]
    return InlineKeyboardMarkup(buttons)


def comment_skip_keyboard() -> InlineKeyboardMarkup:
    """Single button to skip adding a comment to the just-recorded entry."""
    return InlineKeyboardMarkup([[InlineKeyboardButton("No comment", callback_data="skipcomment")]])


# ─── Telegram handlers ────────────────────────────────────────────────────────

def match_category(text: str):
    """Match free-text category input to a known category name (case-insensitive,
    exact match only — no fuzzy matching to avoid silently picking the wrong
    category). Returns the canonical category name, or None if no match."""
    text_lower = text.strip().lower()
    for cat in ALL_EXPENSE_CATEGORIES:
        if cat.lower() == text_lower:
            return cat
    return None


# Common aliases so the person doesn't have to type the exact sheet label.
INCOME_CATEGORY_ALIASES = {
    "зарплата": "Зарплата",
    "salary": "Зарплата",
    "бонус": "Бонус",
    "бонусы": "Бонус",
    "bonus": "Бонус",
    "bonuses": "Бонус",
    "инвойс": "Инвойс",
    "invoice": "Инвойс",
    "self employed": "Инвойс",
    "самозанятость": "Инвойс",
    "points rbc": "Points RBC",
    "поинты": "Points RBC",
    "points": "Points RBC",
}


def match_income_category(text: str):
    """Match free-text input to a known income category name, checking exact
    names first, then common aliases. Returns the canonical category name
    (as it appears on the sheet), or None if no match."""
    text_lower = text.strip().lower()
    for cat in INCOME_CATEGORIES:
        if cat.lower() == text_lower:
            return cat
    return INCOME_CATEGORY_ALIASES.get(text_lower)


def write_expense_to_category(category: str, amount: float, description: str):
    """Route an expense to the correct sheet/structure based on whether the
    category is row-based (June sheet) or fixed (FINANCE sheet). Returns
    (sheet_title, location_label, row, col) — row/col are None for the
    FINANCE-sheet path since comments aren't supported there."""
    if category in ROW_BASED_CATEGORIES:
        sheet_name, row, col = write_row_category_entry(category, amount, description)
        return sheet_name, f"row {row}", row, col
    elif category in FIXED_CATEGORIES:
        sheet_name, row = write_fixed_category_entry(category, amount)
        return sheet_name, f"row {row} (Actual cost)", None, None
    else:
        raise ValueError(f"Unknown category: {category}")


def category_keyboard(prefix: str) -> InlineKeyboardMarkup:
    """Build an inline keyboard with one button per expense category."""
    buttons = [
        [InlineKeyboardButton(name, callback_data=f"{prefix}|{name}")]
        for name in ALL_EXPENSE_CATEGORIES
    ]
    return InlineKeyboardMarkup(buttons)


def income_category_keyboard(prefix: str) -> InlineKeyboardMarkup:
    """Build an inline keyboard with one button per income category."""
    buttons = [
        [InlineKeyboardButton(display_category(name), callback_data=f"{prefix}|{name}")]
        for name in INCOME_CATEGORIES
    ]
    return InlineKeyboardMarkup(buttons)


INCOME_PATTERN = re.compile(r"^(.+?)\s+([\d.,]+)\s*$")
NUMBER_PATTERN = re.compile(r"^\s*([\d.,]+)\s*$")
EXPENSE_PATTERN = re.compile(r"^(.+?)\s+([\d.,]+)\s*$")
TRANSFER_PATTERN = re.compile(r"^(?:transfer|перевод)\s+([\d.,]+)\s*$", re.IGNORECASE)


def check_access(update: Update) -> bool:
    return update.effective_user.id == ALLOWED_USER_ID


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not check_access(update):
        return
    await update.message.reply_text(
        "Hi! Here's what I can do:\n\n"
        "• *Salary 4500* — logs income\n"
        "• *Bonus 300* / *Self employed 200* / *Points RBC 50* — other income types\n"
        "• *Groceries 45.50* — logs an expense\n"
        "• 📸 Photo of a receipt — reads it and logs the expense\n"
        "• *Transfer 100* (or */transfer*) — moves money between wallets\n\n"
        "After logging income or an expense, I'll ask which wallet to use.",
        parse_mode="Markdown"
    )


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not check_access(update):
        return

    text = update.message.text.strip()

    # 0a) If we're waiting for the amount of a transfer (after picking
    # from/to wallets via buttons), treat this message as that amount.
    pending_transfer = context.user_data.get("pending_transfer")
    if pending_transfer is not None and pending_transfer.get("awaiting_amount"):
        m_amount = NUMBER_PATTERN.match(text)
        if not m_amount:
            await update.message.reply_text("❌ Didn't understand the amount. Enter a number, e.g. 100 or 100.50")
            return
        try:
            amount = float(m_amount.group(1).replace(",", "."))
        except ValueError:
            await update.message.reply_text("❌ Didn't understand the amount. Enter a number, e.g. 100 or 100.50")
            return

        context.user_data.pop("pending_transfer", None)
        from_wallet = pending_transfer["from"]
        to_wallet = pending_transfer["to"]

        try:
            sheet_name, from_row, to_row = write_transfer_entry(from_wallet, to_wallet, amount)
        except Exception as e:
            await update.message.reply_text(f"❌ Error recording transfer: {e}")
            return

        context.user_data["pending_comment"] = {
            "targets": [
                {"sheet": sheet_name, "row": from_row, "col": WALLET_COL_DATE_DESC},
                {"sheet": sheet_name, "row": to_row, "col": WALLET_COL_DATE_DESC},
            ]
        }
        await update.message.reply_text(
            f"✅ Transfer ${amount:,.2f}: *{from_wallet}* → *{to_wallet}*\n"
            f"📄 {sheet_name}, rows {from_row} and {to_row}",
            parse_mode="Markdown"
        )
        await update.message.reply_text(
            "Add a comment to this transfer?",
            reply_markup=comment_skip_keyboard()
        )
        return

    # 0) If we're waiting for a comment on the last recorded entry, treat
    # this message as that comment rather than trying to parse a new operation.
    pending_comment = context.user_data.pop("pending_comment", None)
    if pending_comment is not None:
        try:
            if "targets" in pending_comment:
                for t in pending_comment["targets"]:
                    append_comment_to_cell(t["sheet"], t["row"], t["col"], text)
            else:
                append_comment_to_cell(
                    pending_comment["sheet"], pending_comment["row"], pending_comment["col"], text
                )
            await update.message.reply_text("✅ Comment added.")
        except Exception as e:
            await update.message.reply_text(f"❌ Couldn't save the comment: {e}")
        return

    # 0b) Transfer trigger: "Transfer 100" / "Перевод 100" — start the
    # from/to wallet selection flow with the amount already known.
    m_transfer = TRANSFER_PATTERN.match(text)
    if m_transfer:
        try:
            amount = float(m_transfer.group(1).replace(",", "."))
        except ValueError:
            await update.message.reply_text("❌ Didn't understand the amount. Format: Transfer 100")
            return
        context.user_data["pending_transfer"] = {"amount": amount}
        await update.message.reply_text(
            "Transfer between wallets.\nWhich wallet to take the money from?",
            reply_markup=wallet_keyboard("transferfrom")
        )
        return

    # 1) Income: "Зарплата 4500", "Бонус 300", "Self employed 200", "Points RBC 50"
    m = INCOME_PATTERN.match(text)
    if m:
        income_text, amount_str = m.group(1).strip(), m.group(2)
        income_category = match_income_category(income_text)
        if income_category is not None:
            try:
                amount = float(amount_str.replace(",", "."))
            except ValueError:
                await update.message.reply_text("❌ Didn't understand the amount. Format: Salary 4500")
                return

            try:
                sheet_name, row, col = write_row_category_entry(income_category, amount, income_category)
            except Exception as e:
                await update.message.reply_text(f"❌ Error recording entry: {e}")
                return

            context.user_data["pending_wallet_op"] = {
                "kind": "income",
                "amount": amount,
                "description": income_category,
                "entry_sheet": sheet_name,
                "entry_row": row,
                "entry_col": col,
            }
            await update.message.reply_text(
                f"✅ {display_category(income_category)}: ${amount:,.2f}\n📄 {sheet_name}, row {row}\n\n"
                f"Which wallet did the money come into?",
                reply_markup=wallet_keyboard("walletop")
            )
            return
        # else: not a recognized income category — fall through to expense matching below

    # 2) Expense: "Groceries 45.50" — match against known categories
    m = EXPENSE_PATTERN.match(text)
    if m:
        description_raw, amount_str = m.group(1).strip(), m.group(2)
        try:
            amount = float(amount_str.replace(",", "."))
        except ValueError:
            await update.message.reply_text("❌ Didn't understand the amount. Format: Category 45.50")
            return

        category = match_category(description_raw)
        if category is None:
            # Unknown category — ask the user to pick one via buttons,
            # remembering the amount/description for after the pick.
            context.user_data["pending_expense_category_pick"] = {
                "amount": amount,
                "description": description_raw,
            }
            await update.message.reply_text(
                f"Didn't recognize category '{description_raw}'. Pick from the list:",
                reply_markup=category_keyboard("expensecat")
            )
            return

        try:
            sheet_name, location, row, col = write_expense_to_category(category, amount, description_raw)
        except Exception as e:
            await update.message.reply_text(f"❌ Error recording entry: {e}")
            return

        context.user_data["pending_wallet_op"] = {
            "kind": "outcome",
            "amount": amount,
            "description": category,
            "entry_sheet": sheet_name,
            "entry_row": row,
            "entry_col": col,
        }
        await update.message.reply_text(
            f"✅ Expense: {category} — ${amount:,.2f}\n📄 {sheet_name}, {location}\n\n"
            f"Which wallet to take it from?",
            reply_markup=wallet_keyboard("walletop")
        )
        return

    await update.message.reply_text(
        "Didn't understand. Examples:\n*Salary 4500*, *Groceries 45.50*, or a photo of a receipt 📸",
        parse_mode="Markdown"
    )


async def handle_wallet_choice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle the inline-keyboard wallet selection for a pending income/outcome op."""
    query = update.callback_query
    await query.answer()

    if query.from_user.id != ALLOWED_USER_ID:
        return

    try:
        prefix, wallet = query.data.split("|", 1)
    except ValueError:
        return

    if prefix != "walletop":
        return

    op = context.user_data.pop("pending_wallet_op", None)
    if op is None:
        await query.edit_message_text("⚠️ Couldn't find the operation to record (the bot may have restarted). Please try again.")
        return

    try:
        sheet_name, row = write_wallet_entry(
            wallet=wallet,
            kind=op["kind"],
            amount=op["amount"],
            description=op["description"],
        )
        kind_label = "Income" if op["kind"] == "income" else "Expense"
        await query.edit_message_text(
            f"✅ {kind_label} ${op['amount']:,.2f} recorded to wallet *{wallet}*\n"
            f"📄 {sheet_name}, row {row}",
            parse_mode="Markdown"
        )

        # If we know exactly which cell holds the category-block entry,
        # offer to attach a free-text comment to it.
        if op.get("entry_sheet") and op.get("entry_row") and op.get("entry_col"):
            context.user_data["pending_comment"] = {
                "targets": [
                    {"sheet": op["entry_sheet"], "row": op["entry_row"], "col": op["entry_col"]}
                ]
            }
            await query.message.reply_text(
                "Add a comment to this entry?",
                reply_markup=comment_skip_keyboard()
            )
    except Exception as e:
        await query.edit_message_text(f"❌ Error recording to wallet: {e}")


async def handle_transfer_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/transfer — start the wallet-to-wallet transfer flow by asking for
    the source wallet first."""
    if not check_access(update):
        return
    context.user_data.pop("pending_transfer", None)
    await update.message.reply_text(
        "Transfer between wallets.\nWhich wallet to take the money from?",
        reply_markup=wallet_keyboard("transferfrom")
    )


async def handle_transfer_from_choice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle picking the source wallet for a transfer."""
    query = update.callback_query
    await query.answer()

    if query.from_user.id != ALLOWED_USER_ID:
        return

    try:
        _, wallet = query.data.split("|", 1)
    except ValueError:
        return

    pending = context.user_data.setdefault("pending_transfer", {})
    pending["from"] = wallet
    await query.edit_message_text(f"Which wallet to take it from? *{wallet}*", parse_mode="Markdown")
    await query.message.reply_text(
        "Which wallet to transfer to?",
        reply_markup=wallet_keyboard_excluding("transferto", exclude=wallet)
    )


async def handle_transfer_to_choice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle picking the destination wallet for a transfer, then ask for the amount."""
    query = update.callback_query
    await query.answer()

    if query.from_user.id != ALLOWED_USER_ID:
        return

    try:
        _, wallet = query.data.split("|", 1)
    except ValueError:
        return

    pending = context.user_data.get("pending_transfer")
    if pending is None or "from" not in pending:
        await query.edit_message_text("⚠️ Couldn't find the start of the transfer operation (the bot may have restarted). Please start over: Transfer 100 or /transfer")
        return

    pending["to"] = wallet
    from_wallet = pending["from"]

    if "amount" in pending:
        amount = pending["amount"]
        context.user_data.pop("pending_transfer", None)
        await query.edit_message_text(f"Which wallet to transfer to? *{wallet}*", parse_mode="Markdown")

        try:
            sheet_name, from_row, to_row = write_transfer_entry(from_wallet, wallet, amount)
        except Exception as e:
            await query.message.reply_text(f"❌ Error recording transfer: {e}")
            return

        context.user_data["pending_comment"] = {
            "targets": [
                {"sheet": sheet_name, "row": from_row, "col": WALLET_COL_DATE_DESC},
                {"sheet": sheet_name, "row": to_row, "col": WALLET_COL_DATE_DESC},
            ]
        }
        await query.message.reply_text(
            f"✅ Transfer ${amount:,.2f}: *{from_wallet}* → *{wallet}*\n"
            f"📄 {sheet_name}, rows {from_row} and {to_row}",
            parse_mode="Markdown"
        )
        await query.message.reply_text(
            "Add a comment to this transfer?",
            reply_markup=comment_skip_keyboard()
        )
        return

    pending["awaiting_amount"] = True
    await query.edit_message_text(f"Which wallet to transfer to? *{wallet}*", parse_mode="Markdown")
    await query.message.reply_text(f"Transfer amount ({from_wallet} → {wallet})?")


async def handle_skip_comment(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle the 'No comment' button — just clears the pending comment slot."""
    query = update.callback_query
    await query.answer()

    if query.from_user.id != ALLOWED_USER_ID:
        return

    context.user_data.pop("pending_comment", None)
    await query.edit_message_text("No comment.")


async def handle_category_choice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle the inline-keyboard category selection for an expense whose
    typed category didn't match a known category name."""
    query = update.callback_query
    await query.answer()

    if query.from_user.id != ALLOWED_USER_ID:
        return

    try:
        prefix, category = query.data.split("|", 1)
    except ValueError:
        return

    if prefix != "expensecat":
        return

    pending = context.user_data.pop("pending_expense_category_pick", None)
    if pending is None:
        await query.edit_message_text("⚠️ Couldn't find the operation to record (the bot may have restarted). Please try again.")
        return

    try:
        sheet_name, location, row, col = write_expense_to_category(
            category, pending["amount"], pending["description"]
        )
    except Exception as e:
        await query.edit_message_text(f"❌ Error recording entry: {e}")
        return

    context.user_data["pending_wallet_op"] = {
        "kind": "outcome",
        "amount": pending["amount"],
        "description": category,
        "entry_sheet": sheet_name,
        "entry_row": row,
        "entry_col": col,
    }
    await query.edit_message_text(
        f"✅ Expense: {category} — ${pending['amount']:,.2f}\n📄 {sheet_name}, {location}",
    )
    await query.message.reply_text(
        "Which wallet to take it from?",
        reply_markup=wallet_keyboard("walletop")
    )


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not check_access(update):
        return

    await update.message.reply_text("📸 Reading the receipt...")
    photo = update.message.photo[-1]
    file = await context.bot.get_file(photo.file_id)
    image_bytes = await file.download_as_bytearray()

    try:
        data = analyze_receipt_with_claude(bytes(image_bytes))
        description = data["description"]
        amount = float(data["amount"])
        raw_category = data.get("category", "")
        category = match_category(raw_category)

        if category is None:
            # Claude returned something outside our known list — ask the user.
            context.user_data["pending_expense_category_pick"] = {
                "amount": amount,
                "description": description,
            }
            await update.message.reply_text(
                f"Receipt read: {description} — ${amount:,.2f}\n"
                f"Didn't recognize category '{raw_category}'. Pick from the list:",
                reply_markup=category_keyboard("expensecat")
            )
            return

        sheet_name, location, row, col = write_expense_to_category(category, amount, description)

        context.user_data["pending_wallet_op"] = {
            "kind": "outcome",
            "amount": amount,
            "description": category,
            "entry_sheet": sheet_name,
            "entry_row": row,
            "entry_col": col,
        }
        await update.message.reply_text(
            f"✅ Expense: {category} ({description}) — ${amount:,.2f}\n"
            f"📄 {sheet_name}, {location}\n\n"
            f"Which wallet to take it from?",
            reply_markup=wallet_keyboard("walletop")
        )
    except Exception as e:
        await update.message.reply_text(f"❌ Couldn't process the receipt: {e}")


async def handle_debug_income(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Diagnostic: find ALL occurrences of an income category name (e.g.
    'Зарплата') on the month sheet, since the sheet has both a summary block
    listing category names AND each category's own block with the same name.
    Uses a single get_all_values() call to avoid hitting API rate limits.
    Usage: /debugincome [category name, default 'Зарплата']"""
    if not check_access(update):
        return

    category = update.message.text.partition(" ")[2].strip() or "Зарплата"

    try:
        spreadsheet = get_sheet()
        ws = get_month_worksheet(spreadsheet)

        def col_idx_to_letter(idx):
            letters = ""
            while idx > 0:
                idx, rem = divmod(idx - 1, 26)
                letters = chr(65 + rem) + letters
            return letters

        all_values = ws.get_all_values()  # single API call for the whole sheet

        matches = []  # list of (row, col) — 1-indexed
        for r_idx, row_vals in enumerate(all_values, start=1):
            for c_idx, val in enumerate(row_vals[:26], start=1):
                if (val or "").strip().lower() == category.lower():
                    matches.append((r_idx, c_idx))

        if not matches:
            await update.message.reply_text(
                f"❌ Couldn't find '{category}' in columns A-Z on sheet '{ws.title}'."
            )
            return

        def safe_get(r, c):
            if 1 <= r <= len(all_values):
                row_vals = all_values[r - 1]
                if 1 <= c <= len(row_vals):
                    return row_vals[c - 1]
            return None

        lines = [f"📄 Sheet: {ws.title}", f"Occurrences of '{category}' found: {len(matches)}", ""]
        for row, col in matches:
            cl = col_idx_to_letter(col)
            lines.append(f"--- Match at {cl}{row} ---")
            for r in range(row - 1, row + 6):
                vals = []
                for c in range(col, col + 2):
                    ccl = col_idx_to_letter(c)
                    v = safe_get(r, c)
                    vals.append(f"{ccl}{r}={v!r}")
                lines.append(" ".join(vals))
            lines.append("")


        await update.message.reply_text("\n".join(lines))
    except Exception as e:
        await update.message.reply_text(f"❌ Diagnostic error: {e}")


async def handle_debug_fixed(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Diagnostic: scan the FINANCE - <Month> <Year> sheet for a named fixed
    payment category (e.g. 'Rent') and dump raw cell contents around it.
    Usage: /debugfixed Rent"""
    if not check_access(update):
        return

    category = update.message.text.partition(" ")[2].strip()
    if not category:
        await update.message.reply_text("Usage: /debugfixed Rent")
        return

    try:
        spreadsheet = get_sheet()
        ws = get_budget_worksheet(spreadsheet)

        def col_idx_to_letter(idx):
            letters = ""
            while idx > 0:
                idx, rem = divmod(idx - 1, 26)
                letters = chr(65 + rem) + letters
            return letters

        # Scan columns A (1) through AB (28) for a cell matching the category name
        found_row = None
        found_col = None
        for col in range(1, 29):
            col_values = ws.col_values(col)
            for idx, val in enumerate(col_values, start=1):
                if (val or "").strip().lower() == category.lower():
                    found_row = idx
                    found_col = col
                    break
            if found_row:
                break

        if not found_row:
            await update.message.reply_text(
                f"❌ Couldn't find category '{category}' in columns A-AB on sheet '{ws.title}'."
            )
            return

        col_letter = col_idx_to_letter(found_col)
        lines = [
            f"📄 Sheet: {ws.title}",
            f"Category: {category}, found at {col_letter}{found_row}",
            "",
        ]
        for r in range(found_row - 2, found_row + 3):
            vals = []
            for c in range(max(1, found_col - 1), found_col + 3):
                cl = col_idx_to_letter(c)
                v = ws.cell(r, c).value
                vals.append(f"{cl}{r}={v!r}")
            lines.append(" ".join(vals))

        await update.message.reply_text("\n".join(lines))
    except Exception as e:
        await update.message.reply_text(f"❌ Diagnostic error: {e}")


async def handle_debug_category(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Diagnostic: scan columns K onward on the month sheet for a named expense
    category block (e.g. 'Groceries') and dump raw cell contents around it.
    Usage: /debugcategory Groceries"""
    if not check_access(update):
        return

    category = update.message.text.partition(" ")[2].strip()
    if not category:
        await update.message.reply_text("Usage: /debugcategory Groceries")
        return

    try:
        spreadsheet = get_sheet()
        ws = get_month_worksheet(spreadsheet)

        # Scan columns K (11) through Z (26) for a cell matching the category name
        found_row = None
        found_col = None
        for col in range(11, 27):
            col_values = ws.col_values(col)
            for idx, val in enumerate(col_values, start=1):
                if (val or "").strip().lower() == category.lower():
                    found_row = idx
                    found_col = col
                    break
            if found_row:
                break

        if not found_row:
            await update.message.reply_text(
                f"❌ Couldn't find category '{category}' in columns K-Z on sheet '{ws.title}'."
            )
            return

        def col_idx_to_letter(idx):
            letters = ""
            while idx > 0:
                idx, rem = divmod(idx - 1, 26)
                letters = chr(65 + rem) + letters
            return letters

        col_letter = col_idx_to_letter(found_col)
        lines = [
            f"📄 Sheet: {ws.title}",
            f"Category: {category}, found at {col_letter}{found_row}",
            "",
        ]
        for r in range(found_row - 1, found_row + 25):
            vals = []
            for c in range(found_col, found_col + 2):
                cl = col_idx_to_letter(c)
                v = ws.cell(r, c).value
                vals.append(f"{cl}{r}={v!r}")
            lines.append(" ".join(vals))

        await update.message.reply_text("\n".join(lines))
    except Exception as e:
        await update.message.reply_text(f"❌ Diagnostic error: {e}")


async def handle_debug_wallet(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Diagnostic: dump raw cell contents around a wallet block.
    Usage: /debugwallet RBC Checking"""
    if not check_access(update):
        return

    args_text = update.message.text.partition(" ")[2].strip()
    wallet = args_text if args_text in WALLET_NAMES else "RBC Checking"

    try:
        spreadsheet = get_sheet()
        ws = get_month_worksheet(spreadsheet)
        header_row = find_wallet_header_row(ws, wallet)

        lines = [f"📄 Sheet: {ws.title}", f"Wallet: {wallet}, found at row {header_row}", ""]
        for r in range(header_row - 1, header_row + 8):
            a = ws.cell(r, 1).value
            b = ws.cell(r, 2).value
            c = ws.cell(r, 3).value
            d = ws.cell(r, 4).value
            lines.append(f"Row {r}: A={a!r} B={b!r} C={c!r} D={d!r}")

        next_row = find_next_wallet_row(ws, header_row)
        lines.append("")
        lines.append(f"➡️ find_next_wallet_row returned: {next_row}")

        await update.message.reply_text("\n".join(lines))
    except Exception as e:
        await update.message.reply_text(f"❌ Diagnostic error: {e}")


async def handle_status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not check_access(update):
        return
    try:
        spreadsheet = get_sheet()
        ws = get_month_worksheet(spreadsheet)
        all_vals = ws.get_all_values()

        total_income, total_expenses = 0, 0
        for row in all_vals:
            if row and "total income" in row[0].lower():
                for cell in row[1:]:
                    try:
                        total_income = float(cell.replace("$", "").replace(",", "."))
                        break
                    except ValueError:
                        pass
            if row and "total expenses" in row[0].lower():
                for cell in row[1:]:
                    try:
                        total_expenses = float(cell.replace("$", "").replace(",", "."))
                        break
                    except ValueError:
                        pass

        balance = total_income - total_expenses
        await update.message.reply_text(
            f"📊 {ws.title}\n\n"
            f"💚 Income: ${total_income:,.2f}\n"
            f"❤️ Expenses: ${total_expenses:,.2f}\n"
            f"💛 Balance: ${balance:,.2f}"
        )
    except Exception as e:
        await update.message.reply_text(f"❌ Error getting status: {e}")


def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("status", handle_status))
    app.add_handler(CommandHandler("debugwallet", handle_debug_wallet))
    app.add_handler(CommandHandler("debugcategory", handle_debug_category))
    app.add_handler(CommandHandler("debugfixed", handle_debug_fixed))
    app.add_handler(CommandHandler("debugincome", handle_debug_income))
    app.add_handler(CommandHandler("transfer", handle_transfer_command))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(CallbackQueryHandler(handle_wallet_choice, pattern=r"^walletop\|"))
    app.add_handler(CallbackQueryHandler(handle_transfer_from_choice, pattern=r"^transferfrom\|"))
    app.add_handler(CallbackQueryHandler(handle_transfer_to_choice, pattern=r"^transferto\|"))
    app.add_handler(CallbackQueryHandler(handle_category_choice, pattern=r"^expensecat\|"))
    app.add_handler(CallbackQueryHandler(handle_skip_comment, pattern=r"^skipcomment$"))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    logger.info("Bot started!")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
