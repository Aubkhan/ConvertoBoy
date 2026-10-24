import asyncio
import io
import logging
import os
import re
from pathlib import Path

from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import Application, CallbackQueryHandler, CommandHandler, ContextTypes, MessageHandler, filters

logging.basicConfig(level=logging.INFO)
MAX_BYTES = 15 * 1024 * 1024
HEADER = re.compile(r'^\s*(?:\[(\d+)\]|(\d+)[.)])\s*(.+)$')
CHOICE = re.compile(r'^\s*(?:([A-Ha-h])[).:]|([1-8])[).])\s+(.+)$')

def extract_text(data: bytes, suffix: str) -> str:
    if suffix in {'.txt', '.md', '.csv', '.tsv'}:
        for enc in ('utf-8-sig', 'utf-16', 'cp1251'):
            try:
                return data.decode(enc)
            except UnicodeError:
                continue
        raise ValueError('Не удалось определить кодировку файла')
    if suffix == '.docx':
        from docx import Document
        doc = Document(io.BytesIO(data))
        lines = [p.text for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                lines.extend(cell.text for cell in row.cells if cell.text.strip())
        return '\n'.join(lines)
    if suffix == '.pdf':
        import pymupdf
        doc = pymupdf.open(stream=data, filetype='pdf')
        try:
            result = '\n'.join(page.get_text(sort=True) for page in doc)
        finally:
            doc.close()
        if not result.strip():
            raise ValueError('Это PDF-скан без текстового слоя. OCR пока не поддерживается.')
        return result
    if suffix == '.xlsx':
        from openpyxl import load_workbook
        wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        lines = []
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                values = [str(v).strip() for v in row if v is not None and str(v).strip()]
                if values:
                    lines.extend(values)
        wb.close()
        return '\n'.join(lines)
    raise ValueError('Поддерживаются TXT, MD, CSV, TSV, DOCX, PDF (текстовый) и XLSX')


def parse_excel_table(data: bytes):
    """Recognize question/answer columns, preserving answer keys from header labels."""
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    all_questions = []
    try:
        for ws in wb.worksheets:
            rows = ws.iter_rows(values_only=True)
            try:
                header = next(rows)
            except StopIteration:
                continue
            def norm(v):
                return re.sub(r'\s+', ' ', str(v or '').strip().lower()).replace('’', "'").replace('`', "'")
            heads = [norm(v) for v in header]
            def is_question(h):
                return ('вопрос' in h or 'savol' in h or 'question' in h)
            def is_correct(h):
                return ('правильн' in h and 'неправильн' not in h) or ('to\'g\'ri' in h and 'noto' not in h) or 'correct answer' in h
            def is_wrong(h):
                return 'неправильн' in h or 'noto' in h or 'incorrect' in h or 'wrong answer' in h
            qidx = next((i for i,h in enumerate(heads) if is_question(h)), None)
            good = [i for i,h in enumerate(heads) if is_correct(h)]
            bad = [i for i,h in enumerate(heads) if is_wrong(h)]
            if qidx is None or not good or not bad:
                continue
            for row in rows:
                def val(i):
                    return str(row[i]).strip() if i < len(row) and row[i] is not None else ''
                question = val(qidx)
                answers = [('+', val(i)) for i in good if val(i)] + [('-', val(i)) for i in bad if val(i)]
                if question and len(answers) >= 2:
                    all_questions.append({'q': question, 'a': answers})
    finally:
        wb.close()
    return all_questions

def parse_questions(text: str):
    questions = []
    current = None
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue

        if current is not None:
            choice = CHOICE.match(line)
            if choice:
                current['a'].append(('', choice.group(3).strip()))
                continue

        header = HEADER.match(line)
        if line.startswith('# '):
            header_text = line[2:].strip()
        elif header:
            header_text = header.group(3).strip()
        else:
            header_text = None

        if header_text is not None:
            if current:
                questions.append(current)
            current = {'q': header_text, 'a': []}
            continue

        if current is None:
            continue

        if line.startswith(('+ ', '- ', '* ')):
            mark, answer = line[0], line[2:].strip()
        else:
            choice = CHOICE.match(line)
            if choice:
                mark, answer = '', choice.group(3).strip()
            else:
                mark, answer = '', line
        if answer:
            current['a'].append((mark, answer))

    if current:
        questions.append(current)
    return [q for q in questions if len(q['a']) >= 2]

def convert(questions, mode: str) -> str:
    out = []
    for q in questions:
        marked = [i for i, (mark, _) in enumerate(q['a']) if mark in ('+', '*')]
        if mode == 'marks' and not marked:
            raise ValueError('В некоторых вопросах нет отметки правильного ответа. Выбери режим «Первый вариант» или добавь ключ.')
        out.append('# ' + q['q'])
        for i, (_, answer) in enumerate(q['a']):
            correct = i in marked if mode == 'marks' else i == 0
            out.append(('+ ' if correct else '- ') + answer)
        out.append('')
    return '\n'.join(out).strip() + '\n'

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text('Привет! Я ConvertoBoy 🤖\nПришли файл с тестами (TXT, DOCX, PDF, XLSX, CSV, TSV или MD). Я преобразую вопросы в TXT для «Ассистента».\n\nВажно: если правильные ответы не отмечены, я предложу выбрать правило, а не буду их угадывать.')

async def document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    suffix = Path(doc.file_name or '').suffix.lower()
    if doc.file_size and doc.file_size > MAX_BYTES:
        await update.message.reply_text('Максимальный размер файла: 15 МБ.')
        return
    if suffix not in {'.txt','.md','.csv','.tsv','.docx','.pdf','.xlsx'}:
        await update.message.reply_text('Этот формат пока не поддерживается. Пришли TXT, DOCX, PDF, XLSX, CSV, TSV или MD.')
        return
    msg = await update.message.reply_text('Читаю файл…')
    try:
        tgfile = await doc.get_file()
        data = bytes(await tgfile.download_as_bytearray())
        questions = await asyncio.to_thread(parse_excel_table, data) if suffix == '.xlsx' else []
        if not questions:
            text = await asyncio.to_thread(extract_text, data, suffix)
            questions = parse_questions(text)
        if not questions:
            raise ValueError('Не нашёл вопросы с вариантами. Для Excel нужны заголовки «Вопрос», «Правильный ответ», «Неправильный ответ» (или их аналоги). Для текстовых файлов — нумерация или #.')
        context.user_data['questions'] = questions
        context.user_data['filename'] = Path(doc.file_name or 'tests').stem
        await msg.edit_text(f'Нашёл вопросов: {len(questions)}.\nКак определить правильные ответы?', reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton('По знаку + или *', callback_data='marks')],
            [InlineKeyboardButton('Первый вариант (предположение)', callback_data='first')]
        ]))
    except Exception as exc:
        logging.exception('Conversion error')
        await msg.edit_text(f'Не удалось обработать файл: {exc}')

async def choose(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    questions = context.user_data.get('questions')
    if not questions:
        await query.edit_message_text('Отправь файл ещё раз: предыдущие данные недоступны.')
        return
    try:
        result = convert(questions, query.data)
        filename = context.user_data.get('filename', 'tests') + '_Assistant.txt'
        bio = io.BytesIO(result.encode('utf-8-sig'))
        bio.name = filename
        await query.message.reply_document(document=bio, filename=filename, caption=f'Готово: {len(questions)} вопросов. Проверь правильность ответов перед импортом.')
        await query.edit_message_text('Файл готов ✅ Отправь новый документ, чтобы конвертировать ещё один.')
    except Exception as exc:
        await query.message.reply_text(str(exc))

def main():
    token = os.getenv('BOT_TOKEN')
    if not token:
        raise RuntimeError('Укажи BOT_TOKEN в переменных окружения. Не публикуй токен.')
    app = Application.builder().token(token).build()
    app.add_handler(CommandHandler('start', start))
    app.add_handler(MessageHandler(filters.Document.ALL, document))
    app.add_handler(CallbackQueryHandler(choose, pattern='^(marks|first)$'))
    app.run_polling(drop_pending_updates=True)

if __name__ == '__main__':
    main()
