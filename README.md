# Wunder Debrief Bot

Telegram-бот для де-брифинга клиентских брифов по методологии Wunder Digital.

## Деплой на Railway

1. Зайди на [railway.app](https://railway.app) и войди через GitHub
2. New Project → Deploy from GitHub repo → выбери этот репозиторий
3. В разделе Variables добавь две переменные:
   - `TELEGRAM_TOKEN` — токен от @BotFather
   - `ANTHROPIC_API_KEY` — ключ от console.anthropic.com
4. Railway сам запустит бота

## Использование

- `/start` — начало работы
- `/debrief` — запустить анализ брифа
- `/help` — справка

Бот принимает брифы в виде текста или файлов: PDF, DOCX, TXT.
