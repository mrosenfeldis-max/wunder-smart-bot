import os
import logging
import asyncio
from openpyxl import load_workbook
import xlrd
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ContextTypes, filters
)
from telegram.error import BadRequest
import anthropic
from pypdf import PdfReader
from docx import Document
import io

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")

anthropic_client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

# ───────────────────────────────────────────
# ПРОМПТЫ ДЕ-БРИФИНГА
# ───────────────────────────────────────────

SYSTEM_PROMPT_PROJECT = """Ты - старший стратег digital-агентства Wunder Digital. Проведи де-брифинг клиентского брифа по методологии агентства.

ТИП: ПРОЕКТ. Фокус - на бизнес-задаче, бюджете и медиа-миксе. Блоки 6 и 7 менее критичны.

Проанализируй бриф по 7 блокам:
1. Бизнес-задача - какую бизнес-задачу решает клиент. ВАЖНО: если задача сформулирована только как "нужен охват" - это отсутствует, необходимо докопаться до бизнес-причины.
2. Маркетинговая задача - рост знания, формирование спроса, лидогенерация, переключение с конкурентов. Приоритетная задача и этап воронки.
3. Измерение успеха - KPI и метод оценки: продажи, Brand Lift, лиды, доля рынка.
4. Бюджет и рамки - диапазон бюджета, ограничения. Отсутствие бюджета = критический риск.
5. Архитектура запуска и медиа-микс - полный медиа-микс, роль digital, барьеры восприятия бренда.
6. Методология выбора - для проекта менее критично.
7. Критерии победы - для проекта менее критично.

Используй следующий формат ответа СТРОГО (Markdown + эмодзи):

*📋 ДЕ-БРИФ — WUNDER DIGITAL*
_Тип: Проект_

*1️⃣ Бизнес-задача* — [✅ OK / ⚠️ РИСК / ❌ ОТСУТСТВУЕТ]
[резюме что выяснено]
_❔ Необходимо уточнить:_ [что осталось неизвестным, или "—" если всё закрыто]

*2️⃣ Маркетинговая задача* — [✅ OK / ⚠️ РИСК / ❌ ОТСУТСТВУЕТ]
[резюме]
_❔ Необходимо уточнить:_ [...]

*3️⃣ Измерение успеха* — [✅ OK / ⚠️ РИСК / ❌ ОТСУТСТВУЕТ]
[резюме]
_❔ Необходимо уточнить:_ [...]

*4️⃣ Бюджет и рамки* — [✅ OK / ⚠️ РИСК / ❌ ОТСУТСТВУЕТ]
[резюме]
_❔ Необходимо уточнить:_ [...]

*5️⃣ Архитектура запуска* — [✅ OK / ⚠️ РИСК / ❌ ОТСУТСТВУЕТ]
[резюме]
_❔ Необходимо уточнить:_ [...]

*6️⃣ Методология выбора* — [✅ OK / ⚠️ РИСК / 🔵 НЕ ПРИМЕНИМО]
[резюме]

*7️⃣ Критерии победы* — [✅ OK / ⚠️ РИСК / 🔵 НЕ ПРИМЕНИМО]
[резюме]

➖➖➖➖➖➖➖➖➖➖

*🚨 Риски*
• [риск 1]
• [риск 2]

*❓ Вопросы для доуточнения*
• [вопрос 1]
• [вопрос 2]

*🚀 Готовность команды*
[2-3 предложения. Акцент: достаточно ли вводных по бизнес-задаче, бюджету и медиа-миксу для старта]"""

SYSTEM_PROMPT_TENDER = """Ты - старший стратег digital-агентства Wunder Digital. Проведи де-брифинг тендерного брифа по методологии агентства.

ТИП: ТЕНДЕР. Все 7 блоков критичны. Особый фокус на блоках 6 и 7 - без понимания методологии выбора и критериев победы участие в тендере = слепой риск.

Проанализируй бриф по 7 блокам:
1. Бизнес-задача - какую бизнес-задачу решает клиент. Дополнительно: почему тендер сейчас? Есть ли действующий подрядчик? Чем доволен/недоволен?
2. Маркетинговая задача - рост знания, формирование спроса, лидогенерация, переключение с конкурентов. Приоритетная задача и этап воронки.
3. Измерение успеха - KPI и метод оценки. Отсутствие = риск.
4. Бюджет и рамки - диапазон бюджета, ограничения. Отсутствие = критический риск.
5. Архитектура запуска и медиа-микс - полный медиа-микс, роль digital, барьеры, ограничения по каналам.
6. Методология выбора победителя - КРИТИЧНО. Балльная система? Защита? Кто решает? Кто влияет? Отсутствие = критический риск, сигнал FO.
7. Критерии победы - КРИТИЧНО. Что делает предложение сильным? Отсутствие = критический риск, сигнал FO.

Используй следующий формат ответа СТРОГО (Markdown + эмодзи):

*🏆 ДЕ-БРИФ — WUNDER DIGITAL*
_Тип: Тендер_

*1️⃣ Бизнес-задача* — [✅ OK / ⚠️ РИСК / ❌ ОТСУТСТВУЕТ]
[резюме что выяснено]
_❔ Необходимо уточнить:_ [что осталось неизвестным, или "—" если всё закрыто]

*2️⃣ Маркетинговая задача* — [✅ OK / ⚠️ РИСК / ❌ ОТСУТСТВУЕТ]
[резюме]
_❔ Необходимо уточнить:_ [...]

*3️⃣ Измерение успеха* — [✅ OK / ⚠️ РИСК / ❌ ОТСУТСТВУЕТ]
[резюме]
_❔ Необходимо уточнить:_ [...]

*4️⃣ Бюджет и рамки* — [✅ OK / ⚠️ РИСК / ❌ ОТСУТСТВУЕТ]
[резюме]
_❔ Необходимо уточнить:_ [...]

*5️⃣ Архитектура запуска* — [✅ OK / ⚠️ РИСК / ❌ ОТСУТСТВУЕТ]
[резюме]
_❔ Необходимо уточнить:_ [...]

*6️⃣ Методология выбора* — [✅ OK / ⚠️ РИСК / ❌ ОТСУТСТВУЕТ]
[резюме — если ОТСУТСТВУЕТ добавь: ⚠️ Требует доуточнения через FO до старта работы]
_❔ Необходимо уточнить:_ [...]

*7️⃣ Критерии победы* — [✅ OK / ⚠️ РИСК / ❌ ОТСУТСТВУЕТ]
[резюме — если ОТСУТСТВУЕТ добавь: ⚠️ Требует доуточнения через FO до старта работы]
_❔ Необходимо уточнить:_ [...]

➖➖➖➖➖➖➖➖➖➖

*🚨 Риски тендера*
• [риск 1 — особое внимание на блоки 6, 7 и бюджет]
• [риск 2]

*❓ Вопросы для доуточнения*
• [вопрос 1 — приоритет: причина тендера, подрядчик, методология выбора, критерии]
• [вопрос 2]

*🚀 Готовность команды*
[2-3 предложения. Акцент: можно ли стартовать или блоки 6/7 требуют доуточнения через FO]"""

# ───────────────────────────────────────────
# ПРОМПТ АУДИТА МЕДИАПЛАНА
# ───────────────────────────────────────────

SYSTEM_PROMPT_AUDIT = """Ты — старший медиапланер digital-агентства Wunder Digital. Проведи аудит медиаплана по методологии агентства.

ВАЖНО: Скрытые строки, столбцы и вкладки не проверяются. Аудит охватывает только видимое содержимое файла.

Проверь медиаплан по 5 блокам:

БЛОК 1 — СТРУКТУРА И ФОРМАТ
- КРИТИЧНО: Полнота обязательных полей (площадка, формат, период, ЦА, ГЕО, показы/клики/охват, CPM/CPC/CPV, бюджет по строке, итоговый бюджет). Ни одна ячейка не должна быть пустой.
- ВНИМАНИЕ: Единообразие наименований площадок, форматов, таргетингов.
- ЗАМЕЧАНИЕ: Единый формат дат (ДД.ММ.ГГГГ) и чисел.

БЛОК 2 — БЮДЖЕТ И ФИНАНСЫ
- КРИТИЧНО: Корректность формул — ссылки на правильные ячейки, нет #REF!, #DIV/0!
- КРИТИЧНО: Итоговые суммы — итог суммирует все строки, диапазоны СУММ корректны.
- КРИТИЧНО: НДС и агентская комиссия — ставка единая, явно обозначена. Для Казахстана НДС = 16% без исключений.
- ВНИМАНИЕ: Валюта и курс конвертации — при нескольких валютах курс зафиксирован.
- ВНИМАНИЕ: Модели закупки (CPM/CPC/CPL) соответствуют задаче клиента.

БЛОК 3 — МЕДИАПОКАЗАТЕЛИ
- КРИТИЧНО: Логическая связность — показы × CTR = клики, бюджет / показы × 1000 = CPM.
- ВНИМАНИЕ: Реалистичность CTR, CPM, CPC — соответствие отраслевым бенчмаркам.
- ВНИМАНИЕ: Ставки — соответствие рекомендациям РРС, сравнение с предыдущим флайтом.
- ВНИМАНИЕ: Охват не превышает размер ЦА. Частота выше 10 — требует обоснования.

БЛОК 4 — ТАРГЕТИНГИ И АУДИТОРИЯ
- КРИТИЧНО: Соответствие брифу — возраст, пол, гео, интересы точно по брифу.
- КРИТИЧНО: Геотаргетинг указан корректно в каждой строке.
- ВНИМАНИЕ: Пересечение аудиторий на нескольких площадках учтено или обозначено.

БЛОК 5 — СРОКИ И ПЕРИОДЫ
- КРИТИЧНО: Даты совпадают с календарём кампании, нет "дыр" при непрерывной кампании.
- ВНИМАНИЕ: Учтены выходные, праздники и дедлайны площадок (материалы за 2-5 рабочих дней).
- ВНИМАНИЕ: При поэтапном размещении этапы идут в правильном порядке.

КЛАССИФИКАЦИЯ ОШИБОК:
- КРИТИЧНО — влияет на бюджет, корректность данных или запуск. Исправить до утверждения МП.
- ВНИМАНИЕ — существенное замечание, влияет на качество или риски. Рекомендуется исправить.
- ЗАМЕЧАНИЕ — рекомендация по улучшению, не блокирует утверждение.

Используй следующий формат ответа СТРОГО (Markdown + эмодзи):

*🔍 АУДИТ МЕДИАПЛАНА — WUNDER DIGITAL*

*📊 Сводка*
🔴 Критично: [N]
🟡 Внимание: [N]
🔵 Замечание: [N]
_Вывод:_ [можно утверждать / требуются исправления]

➖➖➖➖➖➖➖➖➖➖

*1️⃣ Структура и формат*
[если нет замечаний: ✅ Замечаний нет]
[если есть:]
🔴 *КРИТИЧНО — [заголовок]*
Описание: [что не так и где]
Решение: [конкретное действие]

🟡 *ВНИМАНИЕ — [заголовок]*
Описание: [...]
Решение: [...]

*2️⃣ Бюджет и финансы*
[аналогично]

*3️⃣ Медиапоказатели*
[аналогично]

*4️⃣ Таргетинги и аудитория*
[аналогично]

*5️⃣ Сроки и периоды*
[аналогично]

➖➖➖➖➖➖➖➖➖➖

*✅ Финальный чеклист*
Структура: [OK / требует правок]
Бюджет: [OK / требует правок]
Метрики: [OK / требует правок]
Аудитория: [OK / требует правок]
Сроки: [OK / требует правок]

*🔁 Второй проход завершён.*"""

# ───────────────────────────────────────────
# СОСТОЯНИЯ ПОЛЬЗОВАТЕЛЕЙ
# ───────────────────────────────────────────

user_states = {}

# ───────────────────────────────────────────
# УТИЛИТЫ
# ───────────────────────────────────────────

async def safe_edit(msg, text):
    try:
        await msg.edit_text(text, parse_mode='Markdown')
    except BadRequest:
        try:
            await msg.edit_text(text)
        except Exception:
            pass

async def safe_delete(msg):
    try:
        await msg.delete()
    except Exception:
        pass

async def send_long_message(update, text):
    """Разбивает длинные сообщения на части по 4000 символов."""
    max_len = 4000
    if len(text) <= max_len:
        await update.message.reply_text(text, parse_mode='Markdown')
        return
    parts = []
    while text:
        if len(text) <= max_len:
            parts.append(text)
            break
        split_at = text.rfind('\n', 0, max_len)
        if split_at == -1:
            split_at = max_len
        parts.append(text[:split_at])
        text = text[split_at:].lstrip('\n')
    for i, part in enumerate(parts):
        try:
            await update.message.reply_text(part, parse_mode='Markdown')
        except Exception:
            await update.message.reply_text(part)

def extract_excel_content(file_bytes: bytes, filename: str) -> str:
    """Извлекает содержимое Excel файла в текстовый формат."""
    name = filename.lower()
    result = []
    try:
        if name.endswith('.xls'):
            wb = xlrd.open_workbook(file_contents=file_bytes)
            for sheet in wb.sheets():
                rows = []
                for rx in range(sheet.nrows):
                    row = [str(sheet.cell_value(rx, cx)) for cx in range(sheet.ncols)]
                    if any(v.strip() for v in row):
                        rows.append("\t".join(row))
                if rows:
                    result.append(f"=== Вкладка: {sheet.name} ===")
                    result.append("\n".join(rows))
        else:
            wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                rows = []
                for row in ws.iter_rows(values_only=True):
                    cells = [str(c) if c is not None else "" for c in row]
                    if any(c.strip() for c in cells):
                        rows.append("\t".join(cells))
                if rows:
                    result.append(f"=== Вкладка: {sheet_name} ===")
                    result.append("\n".join(rows))
            wb.close()
        return "\n\n".join(result) if result else ""
    except Exception as e:
        raise ValueError(f"Не удалось прочитать Excel файл: {e}")

# ───────────────────────────────────────────
# КОМАНДЫ БОТА
# ───────────────────────────────────────────

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 *Привет! Я бот Wunder Digital.*\n\n"
        "Что умею:\n"
        "📋 /debrief — анализ клиентского брифа по 7 блокам\n"
        "🔍 /audit — аудит медиаплана на ошибки\n"
        "❓ /help — справка",
        parse_mode='Markdown'
    )

async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "*Команды бота:*\n\n"
        "📋 */debrief* — де-брифинг клиентского брифа\n"
        "Выбери тип (Проект / Тендер), отправь бриф текстом или файлом (PDF, DOCX, TXT)\n\n"
        "🔍 */audit* — аудит медиаплана\n"
        "Отправь файл медиаплана (XLSX, XLS) после команды\n\n"
        "*Чем отличается де-бриф:*\n"
        "Проект — фокус на бизнес-задаче, бюджете, медиа-миксе\n"
        "Тендер — все 7 блоков критичны, особый акцент на методологии выбора и критериях победы",
        parse_mode='Markdown'
    )

async def debrief_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    keyboard = [[
        InlineKeyboardButton("📋 Проект", callback_data="type_project"),
        InlineKeyboardButton("🏆 Тендер", callback_data="type_tender"),
    ]]
    await update.message.reply_text(
        "*Де-бриф запущен.* Выбери тип проекта:",
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode='Markdown'
    )
    user_states[user_id] = {'step': 'choose_type'}

async def audit_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_states[user_id] = {'step': 'waiting_mediaplan'}
    await update.message.reply_text(
        "🔍 *Аудит медиаплана запущен.*\n\n"
        "📎 Отправь файл медиаплана в формате XLSX или XLS.\n\n"
        "⚠️ _Скрытые строки, столбцы и вкладки не проверяются._",
        parse_mode='Markdown'
    )

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    if query.data in ('type_project', 'type_tender'):
        project_type = 'Проект' if query.data == 'type_project' else 'Тендер'
        user_states[user_id] = {'step': 'waiting_brief', 'type': project_type}
        hint = (
            "📋 Фокус: бизнес-задача, бюджет, медиа-микс."
            if project_type == 'Проект'
            else "🏆 Фокус: все 7 блоков, особенно методология выбора и критерии победы."
        )
        await query.edit_message_text(
            f"*Тип:* {project_type}\n_{hint}_\n\n📎 Отправь бриф — текстом или файлом (PDF, DOCX, TXT).",
            parse_mode='Markdown'
        )

# ───────────────────────────────────────────
# ОБРАБОТКА ФАЙЛОВ
# ───────────────────────────────────────────

async def extract_text_from_file(file_bytes: bytes, filename: str) -> str:
    name = filename.lower()
    if name.endswith('.pdf'):
        reader = PdfReader(io.BytesIO(file_bytes))
        return "".join(page.extract_text() or "" for page in reader.pages)
    elif name.endswith('.docx') or name.endswith('.doc'):
        doc = Document(io.BytesIO(file_bytes))
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    elif name.endswith('.txt'):
        return file_bytes.decode('utf-8', errors='ignore')
    else:
        raise ValueError(f"Формат не поддерживается: {filename}")

async def analyze_brief(brief_text: str, project_type: str) -> str:
    system = SYSTEM_PROMPT_TENDER if project_type == 'Тендер' else SYSTEM_PROMPT_PROJECT
    response = anthropic_client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=2000,
        system=system,
        messages=[{"role": "user", "content": f"Бриф клиента:\n\n{brief_text}"}]
    )
    return response.content[0].text

async def analyze_mediaplan(mp_text: str) -> str:
    response = anthropic_client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=3000,
        system=SYSTEM_PROMPT_AUDIT,
        messages=[{"role": "user", "content": f"Медиаплан для аудита:\n\n{mp_text[:15000]}"}]
    )
    return response.content[0].text

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    state = user_states.get(user_id, {})
    if state.get('step') != 'waiting_brief':
        await update.message.reply_text(
            "Напиши /debrief чтобы начать анализ брифа или /audit для аудита медиаплана."
        )
        return
    brief_text = update.message.text
    if len(brief_text) < 50:
        await update.message.reply_text("⚠️ Текст слишком короткий. Отправь полный текст брифа.")
        return
    project_type = state.get('type', 'Проект')
    thinking_msg = await update.message.reply_text("⏳ Анализирую бриф по 7 блокам...")
    try:
        result = await analyze_brief(brief_text, project_type)
        await safe_delete(thinking_msg)
        await send_long_message(update, result)
        user_states.pop(user_id, None)
    except Exception as e:
        logger.error(f"Error analyzing text: {e}")
        await safe_edit(thinking_msg, "❌ Ошибка при анализе. Попробуй снова — /debrief")

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    state = user_states.get(user_id, {})
    step = state.get('step')

    if step not in ('waiting_brief', 'waiting_mediaplan'):
        await update.message.reply_text(
            "Напиши /debrief для анализа брифа или /audit для аудита медиаплана."
        )
        return

    doc = update.message.document
    filename = doc.file_name or ''
    name = filename.lower()

    if doc.file_size > 10 * 1024 * 1024:
        await update.message.reply_text("⚠️ Файл слишком большой. Максимум 10 МБ.")
        return

    # Аудит медиаплана
    if step == 'waiting_mediaplan':
        if not any(name.endswith(ext) for ext in ['.xlsx', '.xls']):
            await update.message.reply_text("⚠️ Для аудита нужен Excel файл (XLSX или XLS).")
            return
        thinking_msg = await update.message.reply_text(f"📊 Читаю медиаплан {filename}...")
        try:
            file = await context.bot.get_file(doc.file_id)
            file_bytes = await file.download_as_bytearray()
            mp_text = extract_excel_content(bytes(file_bytes), filename)
            if not mp_text.strip():
                await safe_edit(thinking_msg, "❌ Не удалось прочитать данные из файла.")
                return
            await safe_edit(thinking_msg, "🔍 Провожу аудит по 5 блокам...")
            result = await analyze_mediaplan(mp_text)
            await safe_delete(thinking_msg)
            await send_long_message(update, result)
            user_states.pop(user_id, None)
        except Exception as e:
            logger.error(f"Error processing mediaplan: {e}")
            await safe_edit(thinking_msg, f"❌ Ошибка: {str(e)[:100]}. Попробуй снова — /audit")
        return

    # Де-брифинг
    if not any(name.endswith(ext) for ext in ['.pdf', '.docx', '.doc', '.txt']):
        await update.message.reply_text("⚠️ Формат не поддерживается. Отправь PDF, DOCX или TXT.")
        return
    thinking_msg = await update.message.reply_text(f"📄 Читаю файл {filename}...")
    try:
        file = await context.bot.get_file(doc.file_id)
        file_bytes = await file.download_as_bytearray()
        brief_text = await extract_text_from_file(bytes(file_bytes), filename)
        if not brief_text.strip():
            await safe_edit(thinking_msg, "❌ Не удалось извлечь текст из файла.")
            return
        project_type = state.get('type', 'Проект')
        await safe_edit(thinking_msg, "⏳ Анализирую бриф по 7 блокам...")
        result = await analyze_brief(brief_text, project_type)
        await safe_delete(thinking_msg)
        await send_long_message(update, result)
        user_states.pop(user_id, None)
    except Exception as e:
        logger.error(f"Error processing document: {e}")
        await safe_edit(thinking_msg, f"❌ Ошибка: {str(e)[:100]}. Попробуй снова — /debrief")

# ───────────────────────────────────────────
# ЗАПУСК
# ───────────────────────────────────────────

async def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(CommandHandler("debrief", debrief_start))
    app.add_handler(CommandHandler("audit", audit_start))
    app.add_handler(CallbackQueryHandler(button_handler))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    logger.info("Bot started")
    async with app:
        await app.start()
        await app.updater.start_polling(drop_pending_updates=True)
        await asyncio.Event().wait()

if __name__ == "__main__":
    asyncio.run(main())
