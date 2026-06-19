import os
import json
import base64
import logging
from io import BytesIO

import requests
import openpyxl
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes

TELEGRAM_TOKEN = os.environ["TELEGRAM_TOKEN"]
GEMINI_API_KEY = os.environ["GEMINI_API_KEY"]

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def extract_invoice_data(image_bytes: bytes) -> dict:
    image_b64 = base64.b64encode(image_bytes).decode("utf-8")

    url = f"https://generativelanguage.googleapis.com/v1/models/gemini-pro-vision:generateContent?key={GEMINI_API_KEY}"

    payload = {
        "contents": [
            {
                "parts": [
                    {
                        "inline_data": {
                            "mime_type": "image/jpeg",
                            "data": image_b64
                        }
                    },
                    {
                        "text": """Bu faturadaki bilgileri çıkar ve SADECE asagidaki JSON formatinda döndür, baska hicbir sey yazma, markdown kullanma:

{
  "fatura_no": "...",
  "tarih": "...",
  "satici": "...",
  "alici": "...",
  "kalemler": [
    {"aciklama": "...", "miktar": "...", "birim_fiyat": "...", "toplam": "..."}
  ],
  "ara_toplam": "...",
  "kdv": "...",
  "genel_toplam": "...",
  "para_birimi": "..."
}

Eger bir alan görünmüyorsa bos string yaz."""
                    }
                ]
            }
        ]
    }

    response = requests.post(url, json=payload, timeout=30)
    response.raise_for_status()

    content = response.json()["candidates"][0]["content"]["parts"][0]["text"]
    content = content.strip().replace("```json", "").replace("```", "").strip()
    return json.loads(content)


def create_excel(data: dict) -> BytesIO:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Fatura"

    from openpyxl.styles import Font, PatternFill, Alignment
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="2E75B6")

    info_rows = [
        ("Fatura No", data.get("fatura_no", "")),
        ("Tarih", data.get("tarih", "")),
        ("Satici", data.get("satici", "")),
        ("Alici", data.get("alici", "")),
    ]

    for i, (label, value) in enumerate(info_rows, start=1):
        ws.cell(row=i, column=1, value=label).font = Font(bold=True)
        ws.cell(row=i, column=2, value=value)

    headers = ["Aciklama", "Miktar", "Birim Fiyat", "Toplam"]
    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=7, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    items = data.get("kalemler", [])
    for row_idx, item in enumerate(items, start=8):
        ws.cell(row=row_idx, column=1, value=item.get("aciklama", ""))
        ws.cell(row=row_idx, column=2, value=item.get("miktar", ""))
        ws.cell(row=row_idx, column=3, value=item.get("birim_fiyat", ""))
        ws.cell(row=row_idx, column=4, value=item.get("toplam", ""))

    summary_start = 8 + len(items) + 1
    para = data.get("para_birimi", "")
    ws.cell(row=summary_start, column=3, value="Ara Toplam").font = Font(bold=True)
    ws.cell(row=summary_start, column=4, value=f"{data.get('ara_toplam', '')} {para}")
    ws.cell(row=summary_start+1, column=3, value="KDV").font = Font(bold=True)
    ws.cell(row=summary_start+1, column=4, value=f"{data.get('kdv', '')} {para}")
    ws.cell(row=summary_start+2, column=3, value="Genel Toplam").font = Font(bold=True)
    ws.cell(row=summary_start+2, column=4, value=f"{data.get('genel_toplam', '')} {para}")

    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 15

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Merhaba! Ben Fatura Bot.\n\n"
        "📸 Fatura fotografı gonderin → 📊 Excel dosyası alın!\n\n"
        "Hemen bir fatura fotografı gondererek baslayın."
    )

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("⏳ Fatura isleniyor, lutfen bekleyin...")

    try:
        photo = update.message.photo[-1]
        file = await context.bot.get_file(photo.file_id)
        image_bytes = await file.download_as_bytearray()

        data = extract_invoice_data(bytes(image_bytes))
        excel_file = create_excel(data)

        await update.message.reply_document(
            document=excel_file,
            filename="fatura.xlsx",
            caption="✅ Faturanız Excel'e donusturuldu!"
        )

    except json.JSONDecodeError as e:
        logger.error(f"JSON hatası: {e}")
        await update.message.reply_text("❌ Fatura okunamadı. Lutfen daha net bir fotograf gonderin.")
    except Exception as e:
        logger.error(f"Hata: {e}")
        await update.message.reply_text(f"❌ Hata: {str(e)}")

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("📎 Lutfen faturayı fotograf olarak gonderin.")


if __name__ == "__main__":
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    print("✅ Bot calisiyor...")
    app.run_polling()
