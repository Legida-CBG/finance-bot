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

CATEGORIES = [
    "Groceries", "GAS / Бензин", "Health", "Sport", "Hair Cut",
    "Vehicle", "Tim Horton's", "Sauna", "Alcohol", "Phone",
    "Helping people / Благотворительность", "Clothes", "Travels",
    "Taxes & fees", "Restaurant / Dining out", "Netflix",
    "Audiobook", "Car loan", "Rent / Аренда", "Car insurance",
    "BC Hydro / Комуналка", "YouTube Music", "Other"
]

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
    categories_str = "\n".join(f"- {c}" for c in CATEGORIES)

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


# ─── Telegram handlers ────────────────────────────────────────────────────────

SALARY_PATTERN = re.compile(r"(?i)^\s*зарплата\s*([123])\s*$")
SALARY_WITH_AMOUNT_PATTERN = re.compile(r"(?i)^\s*зарплата\s*([123])\s+([\d.,]+)\s*$")
BONUS_WITH_AMOUNT_PATTERN = re.compile(r"(?i)^\s*бонус(?:ы)?\s+([\d.,]+)\s*$")
NUMBER_PATTERN = re.compile(r"^\s*([\d.,]+)\s*$")
EXPENSE_PATTERN = re.compile(r"^(.+?)\s+([\d.,]+)\s*$")


def check_access(update: Update) -> bool:
    return update.effective_user.id == ALLOWED_USER_ID


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not check_access(update):
        return
    await update.message.reply_text(
        "Привет! Что я умею:\n\n"
        "• *Зарплата 1* (затем отдельным сообщением сумма) — запишет доход\n"
        "• *Зарплата 1 4500* — то же самое одним сообщением\n"
        "• *Бонус 300* — запишет бонус\n"
        "• *Groceries 45.50* — запишет расход\n"
        "• 📸 Фото чека — распознает и запишет расход\n\n"
        "После записи дохода я спрошу, на какой кошелёк он пришёл.",
        parse_mode="Markdown"
    )


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not check_access(update):
        return

    text = update.message.text.strip()

    # 1) "Зарплата 1 4500" — salary with amount in one message
    m = SALARY_WITH_AMOUNT_PATTERN.match(text)
    if m:
        salary_number, amount_str = m.group(1), m.group(2)
        amount = float(amount_str.replace(",", "."))
        sheet_name, row = write_salary(salary_number, amount)
        context.user_data.pop("pending_salary", None)

        # Stash the pending wallet-income operation and ask which wallet
        context.user_data["pending_wallet_op"] = {
            "kind": "income",
            "amount": amount,
            "description": f"Зарплата {salary_number}",
        }
        await update.message.reply_text(
            f"✅ Зарплата {salary_number}: ${amount:,.2f}\n📄 {sheet_name}, ячейка B{row}\n\n"
            f"На какой кошелёк пришли деньги?",
            reply_markup=wallet_keyboard("walletop")
        )
        return

    # 2) "Зарплата 1" alone — wait for the amount in the next message
    m = SALARY_PATTERN.match(text)
    if m:
        salary_number = m.group(1)
        context.user_data["pending_salary"] = salary_number
        await update.message.reply_text(f"Ок, жду сумму для Зарплаты {salary_number} 💰")
        return

    # 3) Pending salary + this message is just a number
    pending = context.user_data.get("pending_salary")
    if pending:
        m = NUMBER_PATTERN.match(text)
        if m:
            amount = float(m.group(1).replace(",", "."))
            sheet_name, row = write_salary(pending, amount)
            context.user_data.pop("pending_salary", None)

            context.user_data["pending_wallet_op"] = {
                "kind": "income",
                "amount": amount,
                "description": f"Зарплата {pending}",
            }
            await update.message.reply_text(
                f"✅ Зарплата {pending}: ${amount:,.2f}\n📄 {sheet_name}, ячейка B{row}\n\n"
                f"На какой кошелёк пришли деньги?",
                reply_markup=wallet_keyboard("walletop")
            )
            return
        else:
            # Message wasn't a number — drop the pending state and fall through
            context.user_data.pop("pending_salary", None)

    # 4) "Бонус 300"
    m = BONUS_WITH_AMOUNT_PATTERN.match(text)
    if m:
        amount = float(m.group(1).replace(",", "."))
        sheet_name = write_bonus(amount)

        context.user_data["pending_wallet_op"] = {
            "kind": "income",
            "amount": amount,
            "description": "Бонус",
        }
        await update.message.reply_text(
            f"✅ Бонус: ${amount:,.2f}\n📄 {sheet_name}\n\n"
            f"На какой кошелёк пришли деньги?",
            reply_markup=wallet_keyboard("walletop")
        )
        return

    # 5) Generic expense: "Groceries 45.50"
    m = EXPENSE_PATTERN.match(text)
    if m:
        description, amount_str = m.group(1).strip(), m.group(2)
        try:
            amount = float(amount_str.replace(",", "."))
        except ValueError:
            await update.message.reply_text("❌ Не понял сумму. Формат: Категория 45.50")
            return
        try:
            sheet_name, row = write_expense(description, amount)
            await update.message.reply_text(
                f"✅ Расход: {description} — ${amount:,.2f}\n📄 {sheet_name}, строка {row}"
            )
        except Exception as e:
            await update.message.reply_text(f"❌ Ошибка записи: {e}")
        return

    await update.message.reply_text(
        "Не понял. Примеры:\n*Зарплата 1*, *Groceries 45.50*, или фото чека 📸",
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
    except Exception as e:
        await query.edit_message_text(f"❌ Ошибка записи в кошелёк: {e}")


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
        category = data.get("category", "Other")
        full_desc = f"{category}: {description}"
        sheet_name, row = write_expense(full_desc, amount)
        await update.message.reply_text(
            f"✅ Записал: {full_desc} — ${amount:,.2f}\n📄 {sheet_name}, строка {row}"
        )
    except Exception as e:
        await update.message.reply_text(f"❌ Не смог разобрать чек: {e}")


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
        for r in range(found_row - 1, found_row + 10):
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
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(CallbackQueryHandler(handle_wallet_choice, pattern=r"^walletop\|"))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    logger.info("Bot started!")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
