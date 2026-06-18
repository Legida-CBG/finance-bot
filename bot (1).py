import os
import logging
import base64
import json
import re
from datetime import datetime

import anthropic
import gspread
from google.oauth2.service_account import Credentials
from telegram import Update
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
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


def get_sheet():
    """Connect to Google Sheets and return the spreadsheet object."""
    google_creds_json = os.environ["GOOGLE_CREDENTIALS_JSON"]
    creds_dict = json.loads(google_creds_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    gc = gspread.authorize(creds)
    return gc.open_by_key(SPREADSHEET_ID)


def get_month_worksheet(spreadsheet):
    """Find the worksheet for the current month, tolerant of spacing differences
    (e.g. 'June 2026' vs 'June  2026')."""
    now = datetime.now()
    month_name = MONTH_NAMES_EN[now.month]
    year = str(now.year)

    for ws in spreadsheet.worksheets():
        title = ws.title
        if month_name.lower() in title.lower() and year in title:
            return ws

    raise ValueError(
        f"Не нашёл лист для {month_name} {year}. "
        f"Создай вкладку на основе шаблона BUDGET."
    )


def col_letter_to_index(col_letter: str) -> int:
    """Convert column letter (e.g. 'AG') to 1-based column index."""
    idx = 0
    for ch in col_letter:
        idx = idx * 26 + (ord(ch.upper()) - ord('A') + 1)
    return idx


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
        "• 📸 Фото чека — распознает и запишет расход",
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
        await update.message.reply_text(
            f"✅ Зарплата {salary_number}: ${amount:,.2f}\n📄 {sheet_name}, ячейка B{row}"
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
            await update.message.reply_text(
                f"✅ Зарплата {pending}: ${amount:,.2f}\n📄 {sheet_name}, ячейка B{row}"
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
        await update.message.reply_text(f"✅ Бонус: ${amount:,.2f}\n📄 {sheet_name}")
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
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    logger.info("Bot started!")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
