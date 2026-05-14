
import os
import re
import shutil
from pathlib import Path
from datetime import datetime

import pdfplumber
from dotenv import load_dotenv
from openpyxl import load_workbook
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes, filters

load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN")
TEMPLATE_FILE = os.getenv("TEMPLATE_FILE", "KP_Client_FireProtect.xlsx")
WORK_DIR = Path(os.getenv("WORK_DIR", "work"))
WORK_DIR.mkdir(exist_ok=True)

# Внутренний лист, куда бот вводит код + количество
INTERNAL_SHEET = "Introducere_interna"
BASE_SHEET = "Baza_preturi"
SYN_SHEET = "Sinonime"

# Внутренние секции в твоем шаблоне
SECTION_ROWS = {
    "Sistem sprinkler": (6, 20),
    "Nod de control și comandă": (21, 35),
    "Stație de pompare": (36, 50),
    "Automatizare": (51, 65),
    "Elemente de fixare": (66, 80),
    "Lucrări și servicii": (81, 95),
}

def normalize_text(text: str) -> str:
    if not text:
        return ""
    t = text.lower()
    t = t.replace("×", "x").replace("*", "x")
    t = t.replace("ţ", "ț").replace("ş", "ș")
    t = re.sub(r"\s+", " ", t)
    return t.strip()

def load_database(template_path: str):
    wb = load_workbook(template_path, data_only=False)
    base = wb[BASE_SHEET]
    syn = wb[SYN_SHEET]

    products = {}
    for row in base.iter_rows(min_row=2, values_only=True):
        code = row[0]
        if not code:
            continue
        products[str(code)] = {
            "category": row[1],
            "name": row[2],
            "um": row[5],
        }

    synonyms = {}
    for row in syn.iter_rows(min_row=2, values_only=True):
        text, code = row[0], row[1]
        if text and code:
            synonyms[normalize_text(str(text))] = str(code).strip()

    return products, synonyms

def extract_pdf_text(pdf_path: Path) -> str:
    chunks = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            chunks.append(txt)
    return "\n".join(chunks)

def guess_qty(line: str):
    # Ищем количество рядом с единицами: buc, m, set, шт, м
    patterns = [
        r"(\d+(?:[.,]\d+)?)\s*(buc|buc\.|m|ml|set|шт|м)\b",
        r"\b(buc|buc\.|m|ml|set|шт|м)\s*(\d+(?:[.,]\d+)?)",
    ]
    for p in patterns:
        m = re.search(p, line, flags=re.IGNORECASE)
        if m:
            nums = [g for g in m.groups() if re.match(r"^\d", str(g))]
            if nums:
                return float(nums[0].replace(",", "."))
    return 1

def recognize_positions(text: str, synonyms: dict):
    found = []
    not_found_lines = []

    lines = [l.strip() for l in text.splitlines() if l.strip()]

    # Ищем синоним как подстроку в строке PDF
    for line in lines:
        nline = normalize_text(line)
        matched_code = None
        matched_syn = None

        # Сначала длинные синонимы, чтобы точнее находить
        for syn_text, code in sorted(synonyms.items(), key=lambda x: len(x[0]), reverse=True):
            if syn_text and syn_text in nline:
                matched_code = code
                matched_syn = syn_text
                break

        if matched_code:
            qty = guess_qty(line)
            found.append({
                "line": line,
                "code": matched_code,
                "qty": qty,
                "matched_synonym": matched_syn,
            })
        else:
            # Не каждую строку надо считать ошибкой, но для MVP сохраняем подозрительные строки
            if any(word in nline for word in ["dn", "teava", "țeav", "труба", "cot", "отвод", "sprinkler", "орос", "vana", "кран", "манометр"]):
                not_found_lines.append(line)

    # Суммируем одинаковые коды
    merged = {}
    for item in found:
        code = item["code"]
        if code not in merged:
            merged[code] = item.copy()
        else:
            merged[code]["qty"] += item["qty"]

    return list(merged.values()), not_found_lines

def clear_internal_sheet(ws):
    # очищаем код и количество
    for start, end in SECTION_ROWS.values():
        for r in range(start, end + 1):
            ws[f"B{r}"] = None
            ws[f"E{r}"] = None

def put_positions_into_template(template_path: str, positions: list, products: dict, output_path: Path):
    shutil.copy(template_path, output_path)
    wb = load_workbook(output_path)
    ws = wb[INTERNAL_SHEET]
    clear_internal_sheet(ws)

    # текущая свободная строка по каждой секции
    pointers = {sec: start for sec, (start, end) in SECTION_ROWS.items()}

    for pos in positions:
        code = pos["code"]
        product = products.get(code, {})
        category = product.get("category", "Sistem sprinkler")
        qty = pos["qty"]

        # Если категория не совпадает ровно с секцией, кладем в Sistem sprinkler
        section = category if category in SECTION_ROWS else "Sistem sprinkler"
        start, end = SECTION_ROWS[section]
        row = pointers[section]

        if row <= end:
            ws[f"B{row}"] = code
            ws[f"E{row}"] = qty
            pointers[section] += 1

    wb.save(output_path)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Привет! Отправь мне PDF спецификацию. Я попробую распознать позиции и собрать КП Excel."
    )

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc:
        return

    if not doc.file_name.lower().endswith(".pdf"):
        await update.message.reply_text("Пока принимаю только PDF. Excel подключим следующим шагом.")
        return

    await update.message.reply_text("Получил PDF. Распознаю спецификацию...")

    file = await doc.get_file()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    pdf_path = WORK_DIR / f"spec_{timestamp}.pdf"
    await file.download_to_drive(str(pdf_path))

    try:
        products, synonyms = load_database(TEMPLATE_FILE)
        text = extract_pdf_text(pdf_path)
        positions, unknown = recognize_positions(text, synonyms)

        if not positions:
            await update.message.reply_text(
                "Я не нашёл ни одной позиции. Нужно добавить больше вариантов названий в лист Sinonime."
            )
            return

        output_path = WORK_DIR / f"KP_generated_{timestamp}.xlsx"
        put_positions_into_template(TEMPLATE_FILE, positions, products, output_path)

        msg = f"Готово. Распознано позиций: {len(positions)}."
        if unknown:
            msg += f"\nНе распознано подозрительных строк: {len(unknown)}.\n\nПервые строки:\n"
            msg += "\n".join(f"- {x[:80]}" for x in unknown[:10])

        await update.message.reply_text(msg)

        with open(output_path, "rb") as f:
            await update.message.reply_document(
                document=f,
                filename=output_path.name,
                caption="КП Excel сформировано."
            )

    except Exception as e:
        await update.message.reply_text(f"Ошибка: {e}")

def main():
    if not BOT_TOKEN:
        raise RuntimeError("Не найден BOT_TOKEN. Заполни .env файл.")

    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.run_polling()

if __name__ == "__main__":
    main()
