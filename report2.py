
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes

from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING

user_data = {}

# -----------------------------
# Start command
# -----------------------------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Type /report to start")

# -----------------------------
# Start report
# -----------------------------
async def report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.message.from_user.id
    user_data[user_id] = {"step": "title"}
    await update.message.reply_text("សូមបំពេញចំណងជើងរបាយការណ៍។ ឧទាហរណ៍ របាយការណ៍ស្តីពី... នៅថ្ងៃទី....")

# -----------------------------
# Handle message
# -----------------------------
async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.message.from_user.id

    if user_id not in user_data:
        return

    text = update.message.text
    step = user_data[user_id]["step"]

    if step == "title":
        user_data[user_id]["title"] = text
        user_data[user_id]["step"] = "date"
        await update.message.reply_text("កាលបរិច្ឆេទធ្វើរបាយការណ៍ ចន្ទគតិ និងសុរិយាគតិ")

    elif step == "date":
        user_data[user_id]["date"] = text
        user_data[user_id]["step"] = "intro"
        await update.message.reply_text("សូមបំពេញព័ត៌មានទូទៅដូចជា ទីកន្លែងប្រារព្ធពិធី កាលបរិច្ឆេទ អធិបតីភាព អ្នកចូលរួម ។ល។")
# ✅ ONE LINE SPACE
        p = doc.add_paragraph("")
        header_style(p)

    elif step == "intro":
        user_data[user_id]["intro"] = text
        user_data[user_id]["step"] = "agenda"
        await update.message.reply_text("សូមបំពេញរបៀបវារៈ")

# ✅ ONE LINE SPACE
        p = doc.add_paragraph("")
        header_style(p)

    elif step == "agenda":
        user_data[user_id]["agenda"] = text
        user_data[user_id]["step"] = "result"
        await update.message.reply_text("ក្រោយពីពិភាក្សា តើទទួលបានលទ្ធផលអ្វីខ្លះ?")

# ✅ ONE LINE SPACE
        p = doc.add_paragraph("")
        header_style(p)

    elif step == "result":
        user_data[user_id]["result"] = text
        user_data[user_id]["step"] = "note"
        await update.message.reply_text("សូមផ្តល់ការកត់សម្គាល់ ឬមតិរបស់អ្នក")

# ✅ ONE LINE SPACE
        p = doc.add_paragraph("")
        header_style(p)

    elif step == "note":
        user_data[user_id]["note"] = text
        user_data[user_id]["step"] = "summary"
        await update.message.reply_text("សូមផ្តល់ការសន្និដ្ឋាន")

# ✅ ONE LINE SPACE
        p = doc.add_paragraph("")
        header_style(p)

    elif step == "summary":
        user_data[user_id]["summary"] = text
        user_data[user_id]["step"] = "name"
        await update.message.reply_text("ឈ្មោះអ្នកធ្វើរបាយការណ៍")

    elif step == "name":
        user_data[user_id]["name"] = text

        # =============================
        # CREATE WORD DOCUMENT
        # =============================
        doc = Document()
        section = doc.sections[0]

        # ✅ A4 + margins
        section.page_height = Cm(29.7)
        section.page_width = Cm(21)
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(2.2)
        section.right_margin = Cm(1.5)

        # =============================
        # STYLE FUNCTIONS
        # =============================
        def style(p):
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

            for run in p.runs:
                run.font.name = "Khmer OS Siemreap"

        def header_style(p):
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

            for run in p.runs:
                run.font.name = "Khmer OS Muol Light"

        # =============================
        # HEADER
        # =============================
        p = doc.add_paragraph("ព្រះរាជាណាចក្រកម្ពុជា")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].bold = True
        header_style(p)

        p = doc.add_paragraph("ជាតិ សាសនា ព្រះមហាក្សត្រ")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].bold = True
        header_style(p)

        # ✅ ONE LINE SPACE
        p = doc.add_paragraph("")
        header_style(p)

        p = doc.add_paragraph("ក្រសួងទេសចរណ៍")
        header_style(p)

        p = doc.add_paragraph("អគ្គនាយកដ្ឋានរដ្ឋបាល និងហិរញ្ញវត្ថុ")
        header_style(p)

        p = doc.add_paragraph("នាយកដ្ឋានសន្តិសុខ និងសុវត្ថិភាពទេសចរ")
        header_style(p)

        # =============================
        # DATE (RIGHT)
        # =============================
        p = doc.add_paragraph(f"កាលបរិច្ឆេទ៖ {user_data[user_id]['date']}")
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        style(p)

        # =============================
        # TITLE (CENTER)
        # =============================
        p = doc.add_paragraph(f"របាយការណ៍ស្តីពី {user_data[user_id]['title']}")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].bold = True
        header_style(p)

        # =============================
        # CONTENT
        # =============================
        for key in ["intro", "agenda", "result", "note", "summary"]:
            p = doc.add_paragraph(user_data[user_id][key])
            style(p)

        # =============================
        # REPORTER
        # =============================
        p = doc.add_paragraph("អ្នកធ្វើរបាយការណ៍")
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        style(p)

        p = doc.add_paragraph(user_data[user_id]["name"])
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        style(p)

        # =============================
        # SAVE FILE
        # =============================
        filename = f"{user_data[user_id]['name']}_report.docx"
        doc.save(filename)

        # =============================
        # SEND TO TELEGRAM
        # =============================
        with open(filename, "rb") as f:
            await update.message.reply_document(document=f, filename=filename)

        user_data.pop(user_id)

# -----------------------------
# RUN BOT
# -----------------------------
app = ApplicationBuilder().token("8503397726:AAHHL2QsPPSmo3cevdNsSies1M80wBDRChw").build()

app.add_handler(CommandHandler("start", start))
app.add_handler(CommandHandler("report", report))
app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

print("✅ Bot is running...")
app.run_polling()