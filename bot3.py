import os
import json
import logging
import re
import threading
import time
import io
from http.server import HTTPServer, BaseHTTPRequestHandler
from datetime import datetime, date
from collections import defaultdict
from telegram import Update
from telegram.ext import Application, MessageHandler, filters, ContextTypes
import gspread
from google.oauth2.service_account import Credentials
import openpyxl

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

TELEGRAM_TOKEN = os.environ["TELEGRAM_TOKEN"]
SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]
GOOGLE_CREDS_JSON = os.environ["GOOGLE_CREDS_JSON"]
RECONCILE_SHEET = os.environ.get("RECONCILE_SHEET", "Эквайринг")
PORT = int(os.environ.get("PORT", 8080))

# Комиссия эквайринга (процент). Можно переопределить через env.
ACQUIRING_FEE_PCT = float(os.environ.get("ACQUIRING_FEE_PCT", "2.7"))

# ─── HEALTH CHECK ─────────────────────────────────────────────────────────────

class HealthHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"OK")
    def log_message(self, format, *args):
        pass

def run_health_server():
    server = HTTPServer(("0.0.0.0", PORT), HealthHandler)
    server.serve_forever()

# ─── GOOGLE SHEETS ────────────────────────────────────────────────────────────

_gc = None
_spreadsheet = None

def get_client():
    global _gc, _spreadsheet
    try:
        if _gc is None or _spreadsheet is None:
            raise Exception("reconnect")
        _spreadsheet.title
    except Exception:
        logger.info("Переподключаемся к Google Sheets...")
        creds_dict = json.loads(GOOGLE_CREDS_JSON)
        scopes = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        _gc = gspread.authorize(creds)
        _spreadsheet = _gc.open_by_key(SPREADSHEET_ID)
    return _spreadsheet

def sheets_call_with_retry(func, retries=4, wait=20):
    for attempt in range(retries):
        try:
            return func()
        except Exception as e:
            if "429" in str(e) and attempt < retries - 1:
                logger.warning(f"Лимит Sheets, жду {wait} сек")
                time.sleep(wait)
            else:
                raise

def get_or_create_sheet(name: str):
    """Получить лист или создать если не существует."""
    sp = get_client()
    try:
        return sp.worksheet(name)
    except Exception:
        ws = sheets_call_with_retry(lambda: sp.add_worksheet(title=name, rows=2000, cols=20))
        return ws

# ─── РАЗБОР ОФД ───────────────────────────────────────────────────────────────

def parse_ofd(file_bytes: bytes) -> list[dict]:
    """
    Читает XLSX от ОФД.
    Возвращает список: {date, type, cash, electronic, total}
    type: 'Приход' | 'Возврат прихода'
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    headers = rows[0]
    idx = {}
    for i, h in enumerate(headers):
        if h == 'Дата/время':        idx['dt'] = i
        elif h == 'Признак расчета': idx['type'] = i
        elif h == 'Наличными':       idx['cash'] = i
        elif h == 'Электронными':    idx['elec'] = i
        elif h == 'Итого':           idx['total'] = i

    required = ['dt', 'type', 'cash', 'elec', 'total']
    if not all(k in idx for k in required):
        raise ValueError(f"Не найдены нужные колонки в ОФД. Найдено: {list(idx.keys())}")

    result = []
    for row in rows[1:]:
        dt = row[idx['dt']]
        rtype = str(row[idx['type']] or '').strip()
        cash = float(row[idx['cash']] or 0)
        elec = float(row[idx['elec']] or 0)
        total = float(row[idx['total']] or 0)

        if not dt or not rtype:
            continue
        if isinstance(dt, datetime):
            d = dt.date()
        elif isinstance(dt, date):
            d = dt
        else:
            continue

        result.append({
            'date': d,
            'type': rtype,
            'cash': cash,
            'electronic': elec,
            'total': total,
        })
    return result

# ─── РАЗБОР ВЫПИСКИ БАНКА ─────────────────────────────────────────────────────

def parse_bank_statement(file_bytes: bytes) -> list[dict]:
    """
    Читает XLSX выписки Озон Банка.
    Колонки: Дата(B=1), НомерДок(C=2), Дебет(D=3), Кредит(E=4), Контрагент(F=5), Назначение(H=8)
    Возвращает список поступлений: {date, amount, description, source}
    source: 'acquiring' | 'sbp' | 'other'
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    # Ищем строку с заголовком "Дата"
    header_row = None
    for i, row in enumerate(rows):
        if row[1] == 'Дата':
            header_row = i
            break

    if header_row is None:
        raise ValueError("Не найдена строка с заголовком 'Дата' в выписке")

    result = []
    for row in rows[header_row + 2:]:  # +2 пропускаем подзаголовок
        date_val = row[1]
        debit = row[3]
        credit = row[4]
        description = str(row[8] or '')

        if not date_val or credit is None:
            continue

        # Парсим дату
        if isinstance(date_val, (datetime, date)):
            d = date_val.date() if isinstance(date_val, datetime) else date_val
        elif isinstance(date_val, str):
            try:
                d = datetime.strptime(date_val.strip(), '%d.%m.%Y').date()
            except Exception:
                continue
        else:
            continue

        amount = float(credit or 0)
        if amount <= 0:
            continue

        # Определяем источник
        desc_lower = description.lower()
        if 'торгового эквайринга' in desc_lower or 'реестру' in desc_lower:
            source = 'acquiring'
        elif 'через сбп' in desc_lower or 'оплата по заказу' in desc_lower:
            source = 'sbp'
        else:
            source = 'other'

        result.append({
            'date': d,
            'amount': amount,
            'description': description[:200],
            'source': source,
        })

    return result

# ─── ЛОГИКА СВЕРКИ ────────────────────────────────────────────────────────────

def reconcile(ofd_rows: list[dict], bank_rows: list[dict]) -> dict:
    """
    Сверяет ОФД с банком.
    Возвращает словарь с результатами по дням и итогом.
    """
    # Группируем ОФД по датам
    ofd_by_day = defaultdict(lambda: {'cash_in': 0, 'elec_in': 0, 'cash_ret': 0, 'elec_ret': 0})
    for r in ofd_rows:
        d = r['date']
        if r['type'] == 'Приход':
            ofd_by_day[d]['cash_in'] += r['cash']
            ofd_by_day[d]['elec_in'] += r['electronic']
        elif r['type'] == 'Возврат прихода':
            ofd_by_day[d]['cash_ret'] += r['cash']
            ofd_by_day[d]['elec_ret'] += r['electronic']

    # Суммы из банка по дням (эквайринг + СБП)
    bank_by_day = defaultdict(lambda: {'acquiring': 0, 'sbp': 0})
    for r in bank_rows:
        bank_by_day[r['date']][r['source']] += r['amount']

    # Общий период
    all_dates = sorted(set(list(ofd_by_day.keys()) + list(bank_by_day.keys())))

    days = []
    total_ofd_elec = 0
    total_ofd_elec_ret = 0
    total_bank_elec = 0

    for d in all_dates:
        ofd = ofd_by_day[d]
        bank = bank_by_day[d]

        elec_in = ofd['elec_in']
        elec_ret = ofd['elec_ret']
        ofd_net = elec_in - elec_ret  # безнал ОФД за день

        bank_total = bank['acquiring'] + bank['sbp']

        # Ожидаемое зачисление = ОФД безнал - комиссия
        expected = round(ofd_net * (1 - ACQUIRING_FEE_PCT / 100), 2)
        diff = round(bank_total - expected, 2)

        # Статус
        if bank_total == 0 and ofd_net == 0:
            status = '—'
        elif bank_total == 0 and ofd_net > 0:
            status = '⏳ Ожидается'
        elif abs(diff) < 1:  # погрешность до 1 рубля
            status = '✅ Зачислено'
        elif diff < 0:
            status = '⚠️ Недостача'
        else:
            status = '❓ Расхождение'

        total_ofd_elec += elec_in
        total_ofd_elec_ret += elec_ret
        total_bank_elec += bank_total

        days.append({
            'date': d,
            'ofd_cash': ofd['cash_in'],
            'ofd_cash_ret': ofd['cash_ret'],
            'ofd_elec': elec_in,
            'ofd_elec_ret': elec_ret,
            'ofd_net_elec': ofd_net,
            'bank_acquiring': bank['acquiring'],
            'bank_sbp': bank['sbp'],
            'bank_total': bank_total,
            'expected': expected,
            'diff': diff,
            'status': status,
        })

    total_ofd_net = total_ofd_elec - total_ofd_elec_ret
    total_expected = round(total_ofd_net * (1 - ACQUIRING_FEE_PCT / 100), 2)
    total_diff = round(total_bank_elec - total_expected, 2)

    return {
        'days': days,
        'total_ofd_elec': total_ofd_elec,
        'total_ofd_elec_ret': total_ofd_elec_ret,
        'total_ofd_net': total_ofd_net,
        'total_bank': total_bank_elec,
        'total_expected': total_expected,
        'total_diff': total_diff,
    }

# ─── ЗАПИСЬ В ТАБЛИЦУ ─────────────────────────────────────────────────────────

HEADER = [
    'Дата', 'Нал приход ОФД', 'Нал возврат ОФД',
    'Безнал приход ОФД', 'Безнал возврат ОФД', 'Безнал нетто ОФД',
    'Эквайринг банк', 'СБП банк', 'Итого банк',
    f'Ожидаемо (−{ACQUIRING_FEE_PCT}%)', 'Расхождение', 'Статус'
]

def write_to_sheet(result: dict, period_label: str):
    """Дописывает результаты сверки в конец листа 'Эквайринг'."""
    ws = get_or_create_sheet(RECONCILE_SHEET)

    # Узнаём сколько строк уже есть
    existing = sheets_call_with_retry(lambda: ws.col_values(1))
    next_row = len(existing) + 1

    # Если лист пустой — добавляем заголовок колонок
    rows = []
    if next_row == 1:
        rows.append(HEADER)
        next_row += 1

    # Разделитель с периодом
    rows.append([f"── Сверка за {period_label} (комиссия {ACQUIRING_FEE_PCT}%) ──", '', '', '', '', '', '', '', '', '', '', ''])

    for day in result['days']:
        rows.append([
            day['date'].strftime('%d.%m.%Y'),
            day['ofd_cash'],
            day['ofd_cash_ret'],
            day['ofd_elec'],
            day['ofd_elec_ret'],
            day['ofd_net_elec'],
            day['bank_acquiring'],
            day['bank_sbp'],
            day['bank_total'],
            day['expected'],
            day['diff'],
            day['status'],
        ])

    # Итоговая строка периода
    rows.append([
        f'ИТОГО {period_label}',
        '', '',
        result['total_ofd_elec'],
        result['total_ofd_elec_ret'],
        result['total_ofd_net'],
        '', '',
        result['total_bank'],
        result['total_expected'],
        result['total_diff'],
        '✅ OK' if abs(result['total_diff']) < 10 else '⚠️ Есть расхождение',
    ])
    rows.append([])  # пустая строка-разделитель

    start_cell = f'A{next_row}'
    sheets_call_with_retry(lambda: ws.update(start_cell, rows, value_input_option='USER_ENTERED'))
    logger.info(f"Дописано {len(result['days'])} строк на лист '{RECONCILE_SHEET}' начиная с {start_cell}")
    return RECONCILE_SHEET

# ─── СОСТОЯНИЕ СЕССИЙ ─────────────────────────────────────────────────────────

# Хранит загруженные файлы по chat_id до момента запуска сверки
# {'ofd': [bytes, ...], 'bank': bytes | None}
sessions: dict = {}

def get_session(chat_id):
    if chat_id not in sessions:
        sessions[chat_id] = {'ofd': [], 'bank': None}
    return sessions[chat_id]

# ─── КОМАНДЫ И ОБРАБОТЧИКИ ────────────────────────────────────────────────────

HELP_TEXT = """📊 *Бот сверки эквайринга*

Загрузи файлы и я сверю безналичные продажи по ОФД с зачислениями из банка.

*Как пользоваться:*
1️⃣ Отправь файл(ы) выгрузки ОФД (XLSX) — можно несколько за разные периоды
2️⃣ Отправь выписку по счёту (XLSX) из банка
3️⃣ Напиши *сверка* или */run* — и я всё сопоставлю

*Команды:*
/run — запустить сверку
/status — показать что уже загружено
/reset — начать заново (сбросить файлы)
/help — эта справка

_Файлы хранятся только до момента сверки._"""

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    if not msg:
        return
    chat_id = msg.chat_id

    # Текстовые команды
    if msg.text:
        text = msg.text.strip().lower()

        if text in ['/start', '/help']:
            await msg.reply_text(HELP_TEXT, parse_mode='Markdown')
            return

        if text in ['/reset', 'сброс']:
            sessions.pop(chat_id, None)
            await msg.reply_text("🗑 Файлы сброшены. Загружай заново.")
            return

        if text in ['/status', 'статус']:
            sess = get_session(chat_id)
            ofd_count = len(sess['ofd'])
            bank_ok = "✅ загружена" if sess['bank'] else "❌ не загружена"
            await msg.reply_text(
                f"📂 *Текущие файлы:*\n"
                f"ОФД: {ofd_count} шт.\n"
                f"Выписка банка: {bank_ok}",
                parse_mode='Markdown'
            )
            return

        if text in ['/run', 'сверка', 'запуск', 'начать']:
            await run_reconciliation(msg, chat_id)
            return

        await msg.reply_text(
            "Пришли файлы ОФД и выписку банка, потом напиши *сверка*.\n/help — подробнее.",
            parse_mode='Markdown'
        )
        return

    # Файлы
    if msg.document:
        doc = msg.document
        fname = (doc.file_name or '').lower()

        if not fname.endswith('.xlsx'):
            await msg.reply_text("❌ Поддерживаются только XLSX файлы.")
            return

        file = await context.bot.get_file(doc.file_id)
        file_bytes = bytes(await file.download_as_bytearray())
        sess = get_session(chat_id)

        # Определяем тип файла по названию или содержимому
        caption = (msg.caption or '').lower()

        is_bank = any(x in fname for x in ['выписк', 'bank', 'statement']) or \
                  any(x in caption for x in ['выписка', 'банк', 'bank'])
        is_ofd = any(x in fname for x in ['офд', 'ofd', 'чек', 'check', '4650', '4690']) or \
                 any(x in caption for x in ['офд', 'ofd'])

        # Если не понятно — пробуем угадать по содержимому
        if not is_bank and not is_ofd:
            try:
                wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
                ws = wb.active
                first_row = list(next(ws.iter_rows(values_only=True)))
                headers_str = ' '.join(str(h) for h in first_row if h).lower()
                if 'признак расчета' in headers_str or 'электронными' in headers_str:
                    is_ofd = True
                elif 'назначение платежа' in headers_str or 'дебет' in headers_str or 'кредит' in headers_str:
                    is_bank = True
                elif 'озон банк' in headers_str or 'выписка' in headers_str:
                    is_bank = True
                # Проверяем также вторую строку
                if not is_bank and not is_ofd:
                    for i, row in enumerate(ws.iter_rows(values_only=True)):
                        if i > 5: break
                        row_str = ' '.join(str(c) for c in row if c).lower()
                        if 'выписка по счету' in row_str or 'озон банк' in row_str:
                            is_bank = True
                            break
            except Exception as e:
                logger.warning(f"Не смог определить тип файла: {e}")

        if is_ofd:
            sess['ofd'].append(file_bytes)
            count = len(sess['ofd'])
            await msg.reply_text(
                f"✅ ОФД загружен ({count} шт.).\n"
                + ("Теперь жду выписку банка." if not sess['bank'] else "Выписка уже есть. Пиши *сверка*!"),
                parse_mode='Markdown'
            )
        elif is_bank:
            sess['bank'] = file_bytes
            await msg.reply_text(
                f"✅ Выписка банка загружена.\n"
                + (f"ОФД файлов: {len(sess['ofd'])} шт. Пиши *сверка*!" if sess['ofd'] else "Теперь жди файл(ы) ОФД."),
                parse_mode='Markdown'
            )
        else:
            await msg.reply_text(
                "❓ Не могу определить тип файла.\n"
                "Напиши в подписи к файлу *офд* или *выписка*, и отправь снова."
            )

async def run_reconciliation(msg, chat_id):
    sess = get_session(chat_id)

    if not sess['ofd']:
        await msg.reply_text("❌ Нет файлов ОФД. Загрузи сначала.")
        return
    if not sess['bank']:
        await msg.reply_text("❌ Нет выписки банка. Загрузи сначала.")
        return

    await msg.reply_text("⏳ Начинаю сверку, подожди немного...")

    try:
        # Парсим ОФД (объединяем все файлы)
        ofd_rows = []
        for i, ofd_bytes in enumerate(sess['ofd']):
            try:
                rows = parse_ofd(ofd_bytes)
                ofd_rows.extend(rows)
                logger.info(f"ОФД файл {i+1}: {len(rows)} строк")
            except Exception as e:
                await msg.reply_text(f"❌ Ошибка в файле ОФД #{i+1}: {e}")
                return

        # Парсим выписку
        try:
            bank_rows = parse_bank_statement(sess['bank'])
            logger.info(f"Выписка банка: {len(bank_rows)} поступлений")
        except Exception as e:
            await msg.reply_text(f"❌ Ошибка в выписке банка: {e}")
            return

        if not ofd_rows:
            await msg.reply_text("❌ В файлах ОФД не найдено данных.")
            return

        # Сверка
        result = reconcile(ofd_rows, bank_rows)

        # Определяем период для названия листа
        all_dates = [d['date'] for d in result['days']]
        if all_dates:
            dt_from = min(all_dates).strftime('%d.%m')
            dt_to = max(all_dates).strftime('%d.%m.%Y')
            period_label = f"{dt_from}–{dt_to}"
        else:
            period_label = datetime.now().strftime('%d.%m.%Y')

        # Записываем в таблицу
        sheet_name = write_to_sheet(result, period_label)

        # Формируем итоговое сообщение
        days_count = len(result['days'])
        problems = [d for d in result['days'] if '⚠️' in d['status'] or '❓' in d['status']]
        pending = [d for d in result['days'] if '⏳' in d['status']]

        lines = [
            f"✅ *Сверка завершена* за {period_label}",
            f"",
            f"📊 *Итого по ОФД (безнал):* {result['total_ofd_net']:,.2f} ₽",
            f"🏦 *Зачислено в банке:* {result['total_bank']:,.2f} ₽",
            f"📉 *Комиссия ({ACQUIRING_FEE_PCT}%):* {result['total_ofd_net'] - result['total_expected']:,.2f} ₽",
            f"",
        ]

        if abs(result['total_diff']) < 10:
            lines.append("✅ Общий баланс сходится!")
        else:
            sign = "+" if result['total_diff'] > 0 else ""
            lines.append(f"⚠️ *Расхождение итогов:* {sign}{result['total_diff']:,.2f} ₽")

        if problems:
            lines.append(f"\n⚠️ *Дни с расхождениями ({len(problems)}):*")
            for d in problems[:5]:  # Показываем первые 5
                lines.append(f"  • {d['date'].strftime('%d.%m')} — {d['status']} ({d['diff']:+.2f} ₽)")
            if len(problems) > 5:
                lines.append(f"  _...и ещё {len(problems)-5} дней_")

        if pending:
            lines.append(f"\n⏳ *Ещё не зачислено ({len(pending)} дней):*")
            for d in pending[:3]:
                lines.append(f"  • {d['date'].strftime('%d.%m')} — {d['ofd_net_elec']:,.0f} ₽")

        lines.append(f"\n📋 Детализация по дням → лист *\"{sheet_name}\"*")

        await msg.reply_text('\n'.join(lines), parse_mode='Markdown')

        # Сбрасываем сессию после успешной сверки
        sessions.pop(chat_id, None)

    except Exception as e:
        logger.error(f"Ошибка сверки: {e}", exc_info=True)
        await msg.reply_text(f"❌ Ошибка: {str(e)[:300]}")

# ─── ЗАПУСК ───────────────────────────────────────────────────────────────────

async def post_init(application):
    t = threading.Thread(target=run_health_server, daemon=True)
    t.start()
    logger.info(f"Health check на порту {PORT}")

if __name__ == "__main__":
    app = Application.builder().token(TELEGRAM_TOKEN).post_init(post_init).build()
    app.add_handler(MessageHandler(filters.ALL, handle_message))
    logger.info("Бот сверки эквайринга запущен!")
    app.run_polling(drop_pending_updates=True)
