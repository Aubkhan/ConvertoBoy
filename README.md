# ConvertoBoy Telegram Bot

Telegram bot for converting numbered multiple-choice questions into the TXT format used by the «Ассистент» app.

## Start locally (Python 3.10+)

1. Create your bot using https://t.me/BotFather (already created: @ConvertoBoy_bot).
2. Install dependencies: `python -m pip install -r requirements.txt`
3. Set the secret token as an environment variable (never commit or share it):
   - macOS/Linux: `export BOT_TOKEN='YOUR_TOKEN'`
   - Windows PowerShell: `$env:BOT_TOKEN='YOUR_TOKEN'`
4. Run: `python bot.py`
5. Send a test file to https://t.me/ConvertoBoy_bot

## Supported input
TXT, MD, CSV, TSV, DOCX, text PDFs and XLSX. Questions should start with `[1] ...`, `1. ...` or `# ...` and have at least two options. Options can be plain lines, `A) ...` / `1) ...`, or `+ ...` / `- ...`.

Choose whether marked options (`+` or `*`) are correct or assume the first option is correct. The bot does **not** verify the medical accuracy of answers.

## Limits
- 15 MB per document.
- Scanned PDFs require OCR and are not supported.
- Complex tables, old DOC/XLS, PPTX and mixed layouts may not parse correctly.
- XLSX is read as sequential non-empty cells; this is only reliable for spreadsheets laid out like numbered questions and option lines.
- Local polling requires your computer to stay running. For continuous availability, deploy to a host that supports a continuously running worker and environment secrets; check current free-tier limits.
- No files are intentionally persisted to disk. Telegram still processes messages sent to the bot.
