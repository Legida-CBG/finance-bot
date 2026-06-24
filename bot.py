import os
import logging
import base64
import json
import re
from datetime import datetime

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
# YouTube Music, Auto Insurance) used to live only on the FINANCE sheet, but
# the June 2026 sheet ALSO has real row-based blocks for them (column N/O,
# same Total-row-above / header / accumulating-data-rows-below pattern as
# every other category here), confirmed via /debugcategory. As of June 2026
# they're written here instead, to stop the FINANCE-vs-month-sheet confusion.
ROW_BASED_CATEGORIES = [
    "Groceries", "Alcohol", "Fuel", "FHSA", "TFSA", "Restaurants", "Bowling",
    "GIM", "Skiing", "Medication", "Hair Cut", "Dining Out", "Clothing",
    "Sauna", "Audiobooks", "Vehicle loan", "Car Wash", "Charity", "Gifts",
    "Tim Hortons", "Tickets", "Hotels", "Food in travel",
    "Rent", "Phone", "BC Hydro", "Netflix", "YouTube Music", "Auto Insurance",
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

WALLET_NAMES = ["RBC Credit", "RBC Checking", "Costco", "Walmart", "Cash"]

# Cache of {sheet_title: {wallet_name: header_row}} to avoid re-scanning the whole
# column on every single write within the same process lifetime. Cleared per-process;
# Railway restarts will naturally refresh it.
_wallet_row_cache: dict = {}


def find_wallet_header_row(ws, wallet: str) -> int:
    """Find the row where column A exactly equals the wallet's name.
    Searches dynamically instead of relying on a fixed row number, since rows
    can shift if the sheet is edited elsewhere (e.g. budget section above)."""
    cache_key = ws.title
    cached = _wallet_row_cache.get(cache_key, {})
    if wallet in cached:
        # Verify the cached row is still correct before trusting it
        cell_val = (ws.cell(cached[wallet], WALLET_COL_DATE_DESC).value or "").strip()
        if cell_val.lower() == wallet.lower():
            return cached[wallet]

    col_values = ws.col_values(WALLET_COL_DATE_DESC)  # 1-indexed list, column A
    for idx, val in enumerate(col_values, start=1):
        if (val or "").strip().lower() == wallet.lower():
            _wallet_row_cache.setdefault(cache_key, {})[wallet] = idx
            return idx

    raise ValueError(
        f"Не нашёл кошелёк '{wallet}' на листе '{ws.title}'. "
        f"Проверь, что название в колонке A написано точно так же."
    )


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
    now = datetime.now()
    month_name = MONTH_NAMES_EN[now.month]
    year = str(now.year)

    for ws in spreadsheet.worksheets():
        title = ws.title
        if "budget" in title.lower() or "finance" in title.lower():
            continue
        if month_name.lower() in title.lower() and year in title:
            return ws

    raise ValueError(
        f"Не нашёл лист для {month_name} {year}. "
        f"Создай вкладку на основе шаблона BUDGET."
    )


def get_budget_worksheet(spreadsheet):
    """Find the 'BUDGET - <Month> <Year>' (now possibly renamed to
    'FINANCE - <Month> <Year>') worksheet for the current month.
    Currently unused by the bot's write path — kept for future reference/reads."""
    now = datetime.now()
    month_name = MONTH_NAMES_EN[now.month]
    year = str(now.year)

    for ws in spreadsheet.worksheets():
        title = ws.title
        title_lower = title.lower()
        if ("budget" in title_lower or "finance" in title_lower) and \
           month_name.lower() in title_lower and year in title:
            return ws

    raise ValueError(
        f"Не нашёл лист 'BUDGET/FINANCE - {month_name} {year}'. "
        f"Создай его на основе шаблона предыдущего месяца."
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
        cell_value = ws.cell(row, WALLET_COL_DATE_DESC).value
        if not cell_value or not str(cell_value).strip():
            return row
        row += 1
    raise ValueError("Не нашёл свободную строку в блоке кошелька (блок переполнен).")


def verify_wallet_labels(ws, header_row: int):
    """Safety check: confirm the wallet's header row itself has the
    income/outcome/Transfer labels in columns B/C/D (same row as the wallet
    name in column A), as a sanity check on sheet structure."""
    income_label = (ws.cell(header_row, WALLET_COL_INCOME).value or "").strip().lower()
    outcome_label = (ws.cell(header_row, WALLET_COL_OUTCOME).value or "").strip().lower()
    if "income" not in income_label or "outcome" not in outcome_label:
        raise ValueError(
            f"Структура листа изменилась: в строке {header_row} не нашёл "
            f"income/outcome рядом с именем кошелька. Запись отменена, ничего не испорчено."
        )


def write_wallet_entry(wallet: str, kind: str, amount: float, description: str):
    """Write an income or outcome entry into the wallet's block on the current
    month's worksheet (e.g. 'June 2026'). The wallet name and the
    income/outcome/Transfer column labels live on the SAME row (e.g. row 144:
    A='RBC Checking', B='income', C='outcome', D='Transfer'); data rows start
    immediately below that.
    kind: 'income' or 'outcome'."""
    if wallet not in WALLET_NAMES:
        raise ValueError(f"Неизвестный кошелёк: {wallet}")

    spreadsheet = get_sheet()
    ws = get_month_worksheet(spreadsheet)
    header_row = find_wallet_header_row(ws, wallet)

    verify_wallet_labels(ws, header_row)

    target_row = find_next_wallet_row(ws, header_row)

    date_str = datetime.now().strftime("%d.%m")
    label = f"{date_str}  {description}"

    col = WALLET_COL_INCOME if kind == "income" else WALLET_COL_OUTCOME

    ws.update_cell(target_row, WALLET_COL_DATE_DESC, label)
    ws.update_cell(target_row, col, amount)

    return ws.title, target_row


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
            f"Не нашёл категорию '{category}' на листе '{sheet_title}' (колонки A-Z)."
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
    raise ValueError("Не нашёл свободную строку в блоке категории (блок переполнен).")


def write_row_category_entry(category: str, amount: float, description: str):
    """Write an entry into a row-based category block (expense or income) on
    the current month's worksheet (e.g. 'June 2026'). Mirrors the wallet
    entry format: description column gets 'DD.MM. description', amount
    column (next column over) gets the amount."""
    if category not in ROW_BASED_CATEGORIES and category not in INCOME_CATEGORIES:
        raise ValueError(f"Неизвестная построчная категория: {category}")

    spreadsheet = get_sheet()
    ws = get_month_worksheet(spreadsheet)
    all_values = ws.get_all_values()  # single API call for the whole sheet

    header_row, header_col = find_category_header_row(all_values, ws.title, category)
    target_row = find_next_category_row(all_values, header_row, header_col)

    date_str = datetime.now().strftime("%d.%m.")
    label = f"{date_str} {description}"

    ws.update_cell(target_row, header_col, label)
    ws.update_cell(target_row, header_col + 1, amount)

    return ws.title, target_row, header_col


def append_comment_to_cell(sheet_title: str, row: int, col: int, comment: str):
    """Append a user comment to an already-written description cell, e.g.
    turning 'DD.MM. Groceries' into 'DD.MM. Groceries [comment]'."""
    spreadsheet = get_sheet()
    ws = spreadsheet.worksheet(sheet_title)
    existing = ws.cell(row, col).value or ""
    new_value = f"{existing} [{comment}]" if existing else f"[{comment}]"
    ws.update_cell(row, col, new_value)


def write_fixed_category_entry(category: str, amount: float):
    """Add `amount` to the 'Actual cost' cell (column D) of a fixed monthly
    payment category (e.g. 'Rent') on the 'FINANCE - <Month> <Year>' sheet.
    Adds to whatever is already there rather than overwriting."""
    if category not in FIXED_CATEGORIES:
        raise ValueError(f"Неизвестная фиксированная категория: {category}")

    spreadsheet = get_sheet()
    ws = get_budget_worksheet(spreadsheet)

    found_row = None
    col_values = ws.col_values(FIXED_COL_CATEGORY)
    for idx, val in enumerate(col_values, start=1):
        if (val or "").strip().lower() == category.lower():
            found_row = idx
            break

    if found_row is None:
        raise ValueError(
            f"Не нашёл категорию '{category}' в колонке B на листе '{ws.title}'."
        )

    existing_raw = ws.cell(found_row, FIXED_COL_ACTUAL).value or "0"
    existing_clean = str(existing_raw).replace("$", "").strip().replace(",", ".")
    try:
        existing_amount = float(existing_clean) if existing_clean else 0.0
    except ValueError:
        existing_amount = 0.0

    new_amount = existing_amount + amount
    ws.update_cell(found_row, FIXED_COL_ACTUAL, new_amount)

    return ws.title, found_row


# ─── Income handling ─────────────────────────────────────────────────────────

def write_salary(salary_number: str, amount: float):
    """Write amount into the correct 'Зарплата N' cell of the current month's sheet."""
    spreadsheet = get_sheet()
    ws = get_month_worksheet(spreadsheet)
    row = SALARY_ROWS.get(salary_number)
    if row is None:
        raise ValueError(f"Неизвестный номер зарплаты: {salary_number}")
    ws.update_cell(row, 2, amount)  # column B
    return ws.title, row


def write_bonus(amount: float):
    spreadsheet = get_sheet()
    ws = get_month_worksheet(spreadsheet)
    ws.update_cell(BONUS_ROW, 2, amount)
    return ws.title


# ─── Expense handling ────────────────────────────────────────────────────────

def write_expense(description: str, amount: float):
    """Write an expense into today's row of the daily expense table."""
    spreadsheet = get_sheet()
    ws = get_month_worksheet(spreadsheet)

    today = datetime.now().day
    desc_col_idx = col_letter_to_index(EXPENSE_DESC_COL)
    amount_col_idx = col_letter_to_index(EXPENSE_AMOUNT_COL)

    # Dates are pre-filled sequentially starting at row 2 = day 1
    target_row = today + 1

    existing_desc = ws.cell(target_row, desc_col_idx).value or ""
    existing_amount = ws.cell(target_row, amount_col_idx).value or ""

    new_desc = f"{existing_desc}, {description}" if existing_desc else description
    try:
        new_amount = float(str(existing_amount).replace(",", ".") or 0) + amount
    except ValueError:
        new_amount = amount

    ws.update_cell(target_row, desc_col_idx, new_desc)
    ws.update_cell(target_row, amount_col_idx, new_amount)
    return ws.title, target_row


# ─── Anthropic AI helpers ────────────────────────────────────────────────────

def analyze_receipt_with_claude(image_bytes: bytes) -> dict:
    """Send receipt image to Claude and get structured expense data."""
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    image_b64 = base64.standard_b64encode(image_bytes).decode("utf-8")
    categories_str = "\n".join(f"- {c}" for c in ALL_EXPENSE_CATEGORIES)

    prompt = f"""You are a financial assistant analyzing a receipt photo.

Extract the total amount spent and a short description (store name + main items).
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


def comment_skip_keyboard() -> InlineKeyboardMarkup:
    """Single button to skip adding a comment to the just-recorded entry."""
    return InlineKeyboardMarkup([[InlineKeyboardButton("Без комментария", callback_data="skipcomment")]])


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
        return sheet_name, f"строка {row}", row, col
    elif category in FIXED_CATEGORIES:
        sheet_name, row = write_fixed_category_entry(category, amount)
        return sheet_name, f"строка {row} (Actual cost)", None, None
    else:
        raise ValueError(f"Неизвестная категория: {category}")


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
        [InlineKeyboardButton(name, callback_data=f"{prefix}|{name}")]
        for name in INCOME_CATEGORIES
    ]
    return InlineKeyboardMarkup(buttons)


INCOME_PATTERN = re.compile(r"^(.+?)\s+([\d.,]+)\s*$")
NUMBER_PATTERN = re.compile(r"^\s*([\d.,]+)\s*$")
EXPENSE_PATTERN = re.compile(r"^(.+?)\s+([\d.,]+)\s*$")


def check_access(update: Update) -> bool:
    return update.effective_user.id == ALLOWED_USER_ID


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not check_access(update):
        return
    await update.message.reply_text(
        "Привет! Что я умею:\n\n"
        "• *Зарплата 4500* — запишет доход\n"
        "• *Бонус 300* / *Self employed 200* / *Points RBC 50* — другие виды дохода\n"
        "• *Groceries 45.50* — запишет расход\n"
        "• 📸 Фото чека — распознает и запишет расход\n\n"
        "После записи дохода или расхода я спрошу, какой кошелёк использовать.",
        parse_mode="Markdown"
    )


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not check_access(update):
        return

    text = update.message.text.strip()

    # 0) If we're waiting for a comment on the last recorded entry, treat
    # this message as that comment rather than trying to parse a new operation.
    pending_comment = context.user_data.pop("pending_comment", None)
    if pending_comment is not None:
        try:
            append_comment_to_cell(
                pending_comment["sheet"], pending_comment["row"], pending_comment["col"], text
            )
            await update.message.reply_text("✅ Комментарий добавлен.")
        except Exception as e:
            await update.message.reply_text(f"❌ Не смог записать комментарий: {e}")
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
                await update.message.reply_text("❌ Не понял сумму. Формат: Зарплата 4500")
                return

            try:
                sheet_name, row, col = write_row_category_entry(income_category, amount, income_category)
            except Exception as e:
                await update.message.reply_text(f"❌ Ошибка записи: {e}")
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
                f"✅ {income_category}: ${amount:,.2f}\n📄 {sheet_name}, строка {row}\n\n"
                f"На какой кошелёк пришли деньги?",
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
            await update.message.reply_text("❌ Не понял сумму. Формат: Категория 45.50")
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
                f"Не узнал категорию '{description_raw}'. Выбери из списка:",
                reply_markup=category_keyboard("expensecat")
            )
            return

        try:
            sheet_name, location, row, col = write_expense_to_category(category, amount, description_raw)
        except Exception as e:
            await update.message.reply_text(f"❌ Ошибка записи: {e}")
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
            f"✅ Расход: {category} — ${amount:,.2f}\n📄 {sheet_name}, {location}\n\n"
            f"С какого кошелька списать?",
            reply_markup=wallet_keyboard("walletop")
        )
        return

    await update.message.reply_text(
        "Не понял. Примеры:\n*Зарплата 4500*, *Groceries 45.50*, или фото чека 📸",
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
        await query.edit_message_text("⚠️ Не нашёл операцию для записи (возможно, бот перезапускался). Повтори ввод.")
        return

    try:
        sheet_name, row = write_wallet_entry(
            wallet=wallet,
            kind=op["kind"],
            amount=op["amount"],
            description=op["description"],
        )
        kind_label = "Доход" if op["kind"] == "income" else "Расход"
        await query.edit_message_text(
            f"✅ {kind_label} ${op['amount']:,.2f} записан в кошелёк *{wallet}*\n"
            f"📄 {sheet_name}, строка {row}",
            parse_mode="Markdown"
        )

        # If we know exactly which cell holds the category-block entry,
        # offer to attach a free-text comment to it.
        if op.get("entry_sheet") and op.get("entry_row") and op.get("entry_col"):
            context.user_data["pending_comment"] = {
                "sheet": op["entry_sheet"],
                "row": op["entry_row"],
                "col": op["entry_col"],
            }
            await query.message.reply_text(
                "Добавить комментарий к этой записи?",
                reply_markup=comment_skip_keyboard()
            )
    except Exception as e:
        await query.edit_message_text(f"❌ Ошибка записи в кошелёк: {e}")


async def handle_skip_comment(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle the 'Без комментария' button — just clears the pending comment slot."""
    query = update.callback_query
    await query.answer()

    if query.from_user.id != ALLOWED_USER_ID:
        return

    context.user_data.pop("pending_comment", None)
    await query.edit_message_text("Без комментария.")


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
        await query.edit_message_text("⚠️ Не нашёл операцию для записи (возможно, бот перезапускался). Повтори ввод.")
        return

    try:
        sheet_name, location, row, col = write_expense_to_category(
            category, pending["amount"], pending["description"]
        )
    except Exception as e:
        await query.edit_message_text(f"❌ Ошибка записи: {e}")
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
        f"✅ Расход: {category} — ${pending['amount']:,.2f}\n📄 {sheet_name}, {location}",
    )
    await query.message.reply_text(
        "С какого кошелька списать?",
        reply_markup=wallet_keyboard("walletop")
    )


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not check_access(update):
        return

    await update.message.reply_text("📸 Читаю чек...")
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
                f"Чек прочитан: {description} — ${amount:,.2f}\n"
                f"Не узнал категорию '{raw_category}'. Выбери из списка:",
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
            f"✅ Расход: {category} ({description}) — ${amount:,.2f}\n"
            f"📄 {sheet_name}, {location}\n\n"
            f"С какого кошелька списать?",
            reply_markup=wallet_keyboard("walletop")
        )
    except Exception as e:
        await update.message.reply_text(f"❌ Не смог разобрать чек: {e}")


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
                f"❌ Не нашёл '{category}' в колонках A-Z на листе '{ws.title}'."
            )
            return

        def safe_get(r, c):
            if 1 <= r <= len(all_values):
                row_vals = all_values[r - 1]
                if 1 <= c <= len(row_vals):
                    return row_vals[c - 1]
            return None

        lines = [f"📄 Лист: {ws.title}", f"Найдено вхождений '{category}': {len(matches)}", ""]
        for row, col in matches:
            cl = col_idx_to_letter(col)
            lines.append(f"--- Совпадение в {cl}{row} ---")
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
        await update.message.reply_text(f"❌ Ошибка диагностики: {e}")


async def handle_debug_fixed(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Diagnostic: scan the FINANCE - <Month> <Year> sheet for a named fixed
    payment category (e.g. 'Rent') and dump raw cell contents around it.
    Usage: /debugfixed Rent"""
    if not check_access(update):
        return

    category = update.message.text.partition(" ")[2].strip()
    if not category:
        await update.message.reply_text("Использование: /debugfixed Rent")
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
                f"❌ Не нашёл категорию '{category}' в колонках A-AB на листе '{ws.title}'."
            )
            return

        col_letter = col_idx_to_letter(found_col)
        lines = [
            f"📄 Лист: {ws.title}",
            f"Категория: {category}, найдена в {col_letter}{found_row}",
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
        await update.message.reply_text(f"❌ Ошибка диагностики: {e}")


async def handle_debug_category(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Diagnostic: scan columns K onward on the month sheet for a named expense
    category block (e.g. 'Groceries') and dump raw cell contents around it.
    Usage: /debugcategory Groceries"""
    if not check_access(update):
        return

    category = update.message.text.partition(" ")[2].strip()
    if not category:
        await update.message.reply_text("Использование: /debugcategory Groceries")
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
                f"❌ Не нашёл категорию '{category}' в колонках K-Z на листе '{ws.title}'."
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
            f"📄 Лист: {ws.title}",
            f"Категория: {category}, найдена в {col_letter}{found_row}",
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
        await update.message.reply_text(f"❌ Ошибка диагностики: {e}")


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

        lines = [f"📄 Лист: {ws.title}", f"Кошелёк: {wallet}, найден в строке {header_row}", ""]
        for r in range(header_row - 1, header_row + 8):
            a = ws.cell(r, 1).value
            b = ws.cell(r, 2).value
            c = ws.cell(r, 3).value
            d = ws.cell(r, 4).value
            lines.append(f"Row {r}: A={a!r} B={b!r} C={c!r} D={d!r}")

        next_row = find_next_wallet_row(ws, header_row)
        lines.append("")
        lines.append(f"➡️ find_next_wallet_row вернула: {next_row}")

        await update.message.reply_text("\n".join(lines))
    except Exception as e:
        await update.message.reply_text(f"❌ Ошибка диагностики: {e}")


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
            f"💚 Доходы: ${total_income:,.2f}\n"
            f"❤️ Расходы: ${total_expenses:,.2f}\n"
            f"💛 Баланс: ${balance:,.2f}"
        )
    except Exception as e:
        await update.message.reply_text(f"❌ Ошибка при получении статуса: {e}")


def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("status", handle_status))
    app.add_handler(CommandHandler("debugwallet", handle_debug_wallet))
    app.add_handler(CommandHandler("debugcategory", handle_debug_category))
    app.add_handler(CommandHandler("debugfixed", handle_debug_fixed))
    app.add_handler(CommandHandler("debugincome", handle_debug_income))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(CallbackQueryHandler(handle_wallet_choice, pattern=r"^walletop\|"))
    app.add_handler(CallbackQueryHandler(handle_category_choice, pattern=r"^expensecat\|"))
    app.add_handler(CallbackQueryHandler(handle_skip_comment, pattern=r"^skipcomment$"))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    logger.info("Bot started!")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
