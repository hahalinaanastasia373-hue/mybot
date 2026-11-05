import os
import re
import logging
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer
from datetime import datetime

import openpyxl
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, ContextTypes,
    CallbackQueryHandler, MessageHandler, filters
)

# ======================= CONFIG =======================
TOKEN = os.environ.get("TELEGRAM_TOKEN", "9238378464:AAGcNR6KjMOvzXh_ZmGxS8Mq7ILcivBN24A")
EXCEL_PATH = "Мониторинг.xlsx"
SHEET_NAME = "мониторинг"

# Telegram ID исполнителей (ЗАМЕНИТЕ НА РЕАЛЬНЫЕ)
USERS = {
    "Малиновский Д.С.": 8893744433,
    "Шейко А.Г.": 8923409204,
    "Маскаленчик А.Э.": 731881402,
    "Хахалина А.С.": 8744295897, # Ваш ID
}

CHIEF_ID = 8656941302  # ЗАМЕНИТЕ НА ID НАЧАЛЬНИКА
MY_ID = 8744295897     # ВАШ ID
# ======================================================

logging.basicConfig(format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO)
log = logging.getLogger(__name__)

# ---------- Чтение Excel ----------
def read_tasks():
    if not os.path.exists(EXCEL_PATH):
        return []
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    ws = wb[SHEET_NAME]
    tasks = []
    for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if len(row) < 11: continue
        project_raw, task_text, deadline, assignee, status = row[0], row[3], row[5], row[8], row[10]
        if not project_raw or not task_text: continue

        project = re.sub(r"\.xlsx$", "", str(project_raw)).strip()
        deadline_str = deadline.strftime("%d.%m.%Y") if isinstance(deadline, datetime) else str(deadline)
        
        tasks.append({
            "row": idx,
            "project": project,
            "task": str(task_text).strip(),
            "deadline": deadline_str,
            "assignee": str(assignee).strip() if assignee else "",
            "status": str(status).strip() if status else "",
        })
    return tasks

# ---------- Обновление статуса в Excel ----------
def update_excel_status(row_num, new_status):
    try:
        wb = openpyxl.load_workbook(EXCEL_PATH)
        ws = wb[SHEET_NAME]
        ws.cell(row=row_num, column=11).value = new_status
        wb.save(EXCEL_PATH)
        return True
    except Exception as e:
        log.error(f"Ошибка записи в Excel: {e}")
        return False

# ---------- Форматирование сообщения с задачей ----------
def format_task(t):
    return (
        f"🔔 *Напоминание о задаче!*\n\n"
        f"📁 Проект: {t['project']}\n"
        f"👤 Исполнитель: {t['assignee']}\n"
        f"📋 Задача: {t['task']}\n"
        f"📅 Срок: {t['deadline']}"
    )

# ---------- Приём файла от вас ----------
async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc.file_name.endswith(".xlsx"):
        await update.message.reply_text("Пожалуйста, отправьте файл Excel (.xlsx)")
        return

    await update.message.reply_text("Файл получен, сохраняю...")
    file = await context.bot.get_file(doc.file_id)
    await file.download_to_drive(EXCEL_PATH)
    await update.message.reply_text("✅ Файл обновлён! Теперь напишите /tasks, чтобы разослать задачи.")

# ---------- Рассылка задач (только вы) ----------
async def cmd_tasks(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != MY_ID:
        await update.message.reply_text("⛔ У вас нет прав на рассылку задач.")
        return

    tasks = read_tasks()
    if not tasks:
        await update.message.reply_text("Нет данных. Сначала отправьте файл Excel.")
        return
    sent = 0
    for t in tasks:
        uid = USERS.get(t["assignee"])
        if not uid: continue
        
        keyboard = [[
            InlineKeyboardButton("✅ Выполнено", callback_data=f"done_{t['row']}"),
            InlineKeyboardButton("⏳ Продлить", callback_data=f"extend_{t['row']}")
        ]]
        try:
            await context.bot.send_message(
                chat_id=uid, text=format_task(t),
                parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(keyboard)
            )
            sent += 1
        except Exception as e:
            log.warning(f"Не отправлено {t['assignee']}: {e}")
            
    await update.message.reply_text(f"Отправлено задач: {sent}")

# ---------- Сводка для начальника (тоже только вы) ----------
async def cmd_all(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != MY_ID:
        await update.message.reply_text("⛔ У вас нет прав на просмотр сводки.")
        return

    tasks = read_tasks()
    if not tasks:
        await update.message.reply_text("Нет данных. Сначала отправьте файл Excel.")
        return

    header = f"📋 *Сводка по всем задачам* ({len(tasks)} шт.)\n\n"
    blocks = []
    for i, t in enumerate(tasks, 1):
        block = (
            f"{i}. *{t['project']}*\n"
            f"   👤 {t['assignee']}\n"
            f"   📋 {t['task']}\n"
            f"   📅 Срок: {t['deadline']}\n"
            f"   📊 Статус: {t['status']}\n\n"
        )
        blocks.append(block)

    current_message = header
    for block in blocks:
        if len(current_message) + len(block) > 4000:
            await context.bot.send_message(chat_id=CHIEF_ID, text=current_message, parse_mode="Markdown")
            current_message = ""
        current_message += block

    if current_message.strip():
        await context.bot.send_message(chat_id=CHIEF_ID, text=current_message, parse_mode="Markdown")

    await update.message.reply_text("✅ Сводка отправлена начальнику.")

# ---------- Обработка нажатий на кнопки ----------
async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    action, row_str = query.data.split('_')
    row_num = int(row_str)

    if action == "done":
        new_text = "✅ СУПЕР!"
        action_type = "done"
    else:
        new_text = "⏳ Это печально"
        action_type = "extend"

    new_keyboard = InlineKeyboardMarkup([[InlineKeyboardButton(new_text, callback_data="ignore")]])
    await query.edit_message_reply_markup(reply_markup=new_keyboard)

    context.user_data['pending_action'] = {
        'action': action_type,
        'row': row_num,
        'user': query.from_user.username or query.from_user.first_name
    }
    
    await query.message.reply_text(
        f"Вы выбрали *{new_text}*.\nПожалуйста, напишите комментарий к этой задаче (одним сообщением):",
        parse_mode="Markdown"
    )

# ---------- Обработка текста (комментария) и ОТЧЕТ ----------
async def handle_comment(update: Update, context: ContextTypes.DEFAULT_TYPE):
    pending = context.user_data.get('pending_action')
    if not pending:
        return
        
    comment_text = update.message.text
    row_num = pending['row']
    action_type = pending['action']

    if action_type == "done":
        new_status = "Завершено"
        display_status = "✅ Выполнено"
    else:
        new_status = "В процессе (Продлено)"
        display_status = "⏳ Продлен срок"

    update_excel_status(row_num, new_status)
    
    tasks = read_tasks()
    task_info = next((t for t in tasks if t['row'] == row_num), None)
    
    if task_info:
        real_assignee = task_info['assignee']
        project_name = task_info['project']
        task_name = task_info['task']  # <-- ДОБАВИЛИ САМУ ЗАДАЧУ
    else:
        real_assignee = "—"
        project_name = "—"
        task_name = "—"

    # ОТЧЕТ (ДОБАВЛЕНА СТРОКА С ЗАДАЧЕЙ)
    report = (
        f"🔔 *Отчёт по задаче*\n\n"
        f"📁 *Проект:* {project_name}\n"
        f"📋 *Задача:* {task_name}\n"
        f"👤 *Исполнитель:* {real_assignee}\n"
        f"📊 *Статус:* {display_status}\n"
        f"💬 *Комментарий:* {comment_text}"
    )

    for admin_id in [MY_ID, CHIEF_ID]:
        try:
            await context.bot.send_message(chat_id=admin_id, text=report, parse_mode="Markdown")
        except Exception as e:
            log.warning(f"Не удалось отправить отчет {admin_id}: {e}")

    await update.message.reply_text("Спасибо! Ваш комментарий и статус отправлены руководству.")
    context.user_data.pop('pending_action', None)

# ---------- Служебные команды ----------
async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        f"Ваш ID: `{update.effective_user.id}`\n\n"
        "Отправьте мне файл Excel (.xlsx), затем напишите /tasks для рассылки.",
        parse_mode="Markdown"
    )

# ---------- Веб-сервер для Render ----------
class HealthHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"Bot is alive")

def run_health_server():
    port = int(os.environ.get("PORT", 10000))
    server = HTTPServer(("0.0.0.0", port), HealthHandler)
    server.serve_forever()

# ---------- Запуск ----------
def main():
    app = Application.builder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("tasks", cmd_tasks))
    app.add_handler(CommandHandler("all", cmd_all))
    app.add_handler(CallbackQueryHandler(button_handler))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_comment))

    log.info("Бот запущен и готов к работе")
    app.run_polling()

if __name__ == "__main__":
    threading.Thread(target=run_health_server, daemon=True).start()
    main()
