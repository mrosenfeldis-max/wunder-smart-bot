import os
import logging
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ContextTypes, filters
)
import anthropic
from pypdf import PdfReader
from docx import Document
import io

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")

anthropic_client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

SYSTEM_PROMPT = """Ты — старший стратег digital-агентства Wunder Digital. Проведи де-брифинг клиентского брифа по строгой методологии агентства.

Проанализируй бриф по 7 блокам:
1. Бизнес-задача клиента — какую бизнес-задачу решает клиент (рост продаж, вывод продукта, переключение с конкурентов и т.д.)
2. Маркетинговая задача — рост знания, формирование спроса, лидогенерация, переключение с конкурентов
3. Как будет измеряться успех — KPI: продажи, Brand Lift, лиды, доля рынка
4. Бюджет и рамки — бюджетный диапазон, ограничения. ВАЖНО: отсутствие бюджета = риск
5. Архитектура запуска и медиа-микс — полный медиа-микс, роль digital
6. Методология выбора победителя — как отбирается победитель (для тендеров)
7. Критерии победы — по каким критериям сильное решение

По каждому блоку определи статус:
- ✅ Закрыт — информация есть и достаточна
- ⚠️ Риск — информация частичная или вызывает вопросы
- ❌ Отсутствует — информации нет

Формат ответа — строго следующий (используй эти эмодзи и заголовки):

*ДЕ-БРИФ — WUNDER DIGITAL*

*1. Бизнес-задача* [✅/⚠️/❌]
[краткое резюме что выяснено]
_Открыто:_ [что осталось неизвестным, если есть]

*2. Маркетинговая задача* [✅/⚠️/❌]
[краткое резюме]
_Открыто:_ [...]

*3. Измерение успеха* [✅/⚠️/❌]
[краткое резюме]
_Открыто:_ [...]

*4. Бюджет и рамки* [✅/⚠️/❌]
[краткое резюме]
_Открыто:_ [...]

*5. Архитектура запуска* [✅/⚠️/❌]
[краткое резюме]
_Открыто:_ [...]

*6. Методология выбора* [✅/⚠️/❌]
[краткое резюме]
_Открыто:_ [...]

*7. Критерии победы* [✅/⚠️/❌]
[краткое резюме]
_Открыто:_ [...]

---
*🚨 Риски проекта*
• [риск 1]
• [риск 2]

*❓ Вопросы для доуточнения*
• [вопрос 1]
• [вопрос 2]

*📋 Готовность команды*
[2-3 предложения об общей готовности к старту]

Используй Markdown (*, _, —). Будь конкретным и лаконичным."""

# User states
user_states = {}

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (
        "👋 *Привет! Я бот для де-брифинга Wunder Digital.*\n\n"
        "Я анализирую клиентские брифы по методологии агентства — 7 блоков, риски, вопросы для доуточнения.\n\n"
        "Отправь /debrief чтобы начать."
    )
    await update.message.reply_text(text, parse_mode='Markdown')

async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (
        "*Как пользоваться ботом:*\n\n"
        "1. Напиши /debrief\n"
        "2. Выбери тип проекта\n"
        "3. Отправь бриф — текстом или файлом (PDF, DOCX, TXT)\n"
        "4. Получи анализ по 7 блокам\n\n"
        "*Поддерживаемые форматы:*\n"
        "• Текст — просто напиши или вставь\n"
        "• PDF — прикрепи файл\n"
        "• DOCX / DOC — прикрепи файл\n"
        "• TXT — прикрепи файл\n\n"
        "/debrief — начать анализ\n"
        "/help — эта справка"
    )
    await update.message.reply_text(text, parse_mode='Markdown')

async def debrief_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    keyboard = [
        [
            InlineKeyboardButton("📋 Проект", callback_data="type_project"),
            InlineKeyboardButton("🏆 Тендер", callback_data="type_tender"),
        ]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text(
        "*Де-бриф запущен.*\n\nВыбери тип проекта:",
        reply_markup=reply_markup,
        parse_mode='Markdown'
    )
    user_states[user_id] = {'step': 'choose_type'}

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id

    if query.data in ('type_project', 'type_tender'):
        project_type = 'Проект' if query.data == 'type_project' else 'Тендер'
        user_states[user_id] = {'step': 'waiting_brief', 'type': project_type}
        await query.edit_message_text(
            f"*Тип:* {project_type}\n\n"
            "Теперь отправь бриф — текстом или файлом (PDF, DOCX, TXT).",
            parse_mode='Markdown'
        )

async def extract_text_from_file(file_bytes: bytes, filename: str) -> str:
    name = filename.lower()
    if name.endswith('.pdf'):
        reader = PdfReader(io.BytesIO(file_bytes))
        text = ""
        for page in reader.pages:
            text += page.extract_text() or ""
        return text
    elif name.endswith('.docx') or name.endswith('.doc'):
        doc = Document(io.BytesIO(file_bytes))
        return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    elif name.endswith('.txt'):
        return file_bytes.decode('utf-8', errors='ignore')
    else:
        raise ValueError(f"Формат не поддерживается: {filename}")

async def analyze_brief(brief_text: str, project_type: str) -> str:
    response = anthropic_client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=2000,
        system=SYSTEM_PROMPT,
        messages=[{
            "role": "user",
            "content": f"Тип проекта: {project_type}\n\nБриф клиента:\n\n{brief_text}"
        }]
    )
    return response.content[0].text

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    state = user_states.get(user_id, {})

    if state.get('step') != 'waiting_brief':
        await update.message.reply_text(
            "Напиши /debrief чтобы начать анализ брифа."
        )
        return

    brief_text = update.message.text
    if len(brief_text) < 50:
        await update.message.reply_text(
            "⚠️ Текст слишком короткий. Отправь полный текст брифа."
        )
        return

    project_type = state.get('type', 'Проект')
    thinking_msg = await update.message.reply_text(
        "⏳ Анализирую бриф по 7 блокам..."
    )

    try:
        result = await analyze_brief(brief_text, project_type)
        await thinking_msg.delete()
        await update.message.reply_text(result, parse_mode='Markdown')
        user_states.pop(user_id, None)
    except Exception as e:
        logger.error(f"Error analyzing brief: {e}")
        await thinking_msg.edit_text(
            "❌ Ошибка при анализе. Попробуй снова — /debrief"
        )

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    state = user_states.get(user_id, {})

    if state.get('step') != 'waiting_brief':
        await update.message.reply_text(
            "Напиши /debrief чтобы начать анализ брифа."
        )
        return

    doc = update.message.document
    filename = doc.file_name or ''
    name = filename.lower()

    if not any(name.endswith(ext) for ext in ['.pdf', '.docx', '.doc', '.txt']):
        await update.message.reply_text(
            "⚠️ Формат не поддерживается. Отправь PDF, DOCX или TXT."
        )
        return

    if doc.file_size > 10 * 1024 * 1024:
        await update.message.reply_text("⚠️ Файл слишком большой. Максимум 10 МБ.")
        return

    thinking_msg = await update.message.reply_text(
        f"📄 Читаю файл *{filename}*...", parse_mode='Markdown'
    )

    try:
        file = await context.bot.get_file(doc.file_id)
        file_bytes = await file.download_as_bytearray()
        brief_text = await extract_text_from_file(bytes(file_bytes), filename)

        if not brief_text.strip():
            await thinking_msg.edit_text("❌ Не удалось извлечь текст из файла.")
            return

        project_type = state.get('type', 'Проект')
        await thinking_msg.edit_text("⏳ Анализирую бриф по 7 блокам...")

        result = await analyze_brief(brief_text, project_type)
        await thinking_msg.delete()
        await update.message.reply_text(result, parse_mode='Markdown')
        user_states.pop(user_id, None)

    except Exception as e:
        logger.error(f"Error processing file: {e}")
        await thinking_msg.edit_text(
            "❌ Ошибка при обработке файла. Попробуй снова — /debrief"
        )

def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(CommandHandler("debrief", debrief_start))
    app.add_handler(CallbackQueryHandler(button_handler))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    logger.info("Bot started")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
