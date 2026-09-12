"""
ربات بله برای ارسال جدول قیمت به‌روز به کاربران
اجرا: python bot.py
تنظیمات لازم: BALE_BOT_TOKEN و BALE_ADMIN_CHAT_ID (به README.md مراجعه کنید)
"""

import os
import re
import time
import json
import sqlite3
import logging
import datetime
import zipfile
import requests
import jdatetime
from zoneinfo import ZoneInfo
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side

BOT_TOKEN = os.environ.get("BALE_BOT_TOKEN", "your-bot-token")
ADMIN_CHAT_ID = int(os.environ.get("BALE_ADMIN_CHAT_ID", "your-chat-id"))

API_URL = f"https://tapi.bale.ai/bot{BOT_TOKEN}"
FILE_URL = f"https://tapi.bale.ai/file/bot{BOT_TOKEN}"
DB_PATH = "users.db"
OUTPUT_DIR = "sent_price_lists"  # فایل‌های نهایی (اکسل) قبل از ارسال اینجا ذخیره می‌شوند
IRAN_TZ = ZoneInfo("Asia/Tehran")

# ---------- تنظیمات سرویس «سفیر» بله (ارسال پیام به شماره، حتی برای کسانی که ربات را استارت نکرده‌اند) ----------
# TODO: به‌محض دریافت این دو مقدار از پنل کسب‌وکار بله، این‌جا یا در متغیرهای محیطی پر کنید.
# تا وقتی خالی باشند، دکمهٔ «تلاش دوباره از طریق سفیر» در ربات غیرفعال می‌ماند و بقیهٔ ربات
# دقیقاً مثل قبل کار می‌کند؛ یعنی همین الان می‌توانید کد را اجرا کنید.
SAFIR_API_ACCESS_KEY = os.environ.get("SAFIR_API_ACCESS_KEY", "")
SAFIR_BOT_ID = os.environ.get("SAFIR_BOT_ID", "")
SAFIR_API_URL = "https://safir.bale.ai/api/v3/send_message"


def safir_configured():
    return bool(SAFIR_API_ACCESS_KEY and SAFIR_BOT_ID)

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("price-bot")


# ---------- دیتابیس کاربران ----------

def db_init():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS users (
            chat_id INTEGER PRIMARY KEY,
            phone TEXT,
            first_name TEXT,
            joined_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
    # پیام‌ها/فایل‌هایی که هنوز ارسال نشده‌اند (چه به‌خاطر خطای شبکه، چه هر دلیل دیگری)
    # اینجا نگه داشته می‌شوند تا چیزی گم نشود؛ حتی اگر ربات ری‌استارت شود،
    # این صف روی دیسک باقی می‌ماند و در اجرای بعدی دوباره تلاش می‌شود.
    conn.execute("""
        CREATE TABLE IF NOT EXISTS outbox (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            chat_id INTEGER NOT NULL,
            text TEXT NOT NULL DEFAULT '',
            file_path TEXT,
            parse_mode TEXT,
            attempts INTEGER DEFAULT 0,
            last_error TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
    # مهاجرت برای دیتابیس‌های قدیمی‌تر که ستون file_path را ندارند
    cols = [row[1] for row in conn.execute("PRAGMA table_info(outbox)").fetchall()]
    if "file_path" not in cols:
        conn.execute("ALTER TABLE outbox ADD COLUMN file_path TEXT")

    # لیست اصلی شماره مشتریان که ادمین از طریق دکمهٔ اکسل آپلود می‌کند (جدا از
    # جدول users که فقط کسانی‌اند که خودشان ربات را استارت کرده‌اند)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS customer_numbers (
            phone TEXT PRIMARY KEY,
            added_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # لاگ دائمی هر ارسال (برخلاف outbox که بعد از ارسال موفق پاک می‌شود)، برای
    # گزارش «به چه شماره‌هایی در چه تاریخی ارسال شد / نشد»
    conn.execute("""
        CREATE TABLE IF NOT EXISTS send_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            batch_date TEXT NOT NULL,
            file_name TEXT,
            phone TEXT NOT NULL,
            status TEXT NOT NULL,
            error TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # وضعیت فعلی ادمین در گفتگو (مثلاً منتظر آپلود لیست شماره‌ها یا فایل قیمت است)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS admin_state (
            admin_chat_id INTEGER PRIMARY KEY,
            state TEXT
        )
    """)

    conn.commit()
    return conn


# ---------- وضعیت ادمین (کدام دکمه را زده و منتظر چه فایلی است) ----------

def set_admin_state(conn, admin_chat_id, state):
    conn.execute(
        "INSERT INTO admin_state (admin_chat_id, state) VALUES (?, ?) "
        "ON CONFLICT(admin_chat_id) DO UPDATE SET state=excluded.state",
        (admin_chat_id, state),
    )
    conn.commit()


def get_admin_state(conn, admin_chat_id):
    row = conn.execute(
        "SELECT state FROM admin_state WHERE admin_chat_id=?", (admin_chat_id,)
    ).fetchone()
    return row[0] if row else None


_PERSIAN_DIGIT_MAP = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")


def normalize_phone(phone: str) -> str:
    """شماره را به فرمت یکسان +98XXXXXXXXXX تبدیل می‌کند"""
    # ارقام فارسی/عربی را به انگلیسی تبدیل می‌کنیم تا \D آن‌ها را حذف نکند
    text = (phone or "").translate(_PERSIAN_DIGIT_MAP)
    digits = re.sub(r"\D", "", text)
    # بعضی فایل‌های اکسل عدد را به صورت 98912... با .0 ذخیره می‌کنند
    if digits.endswith("0") and len(digits) > 12 and "." in text:
        digits = re.sub(r"\D", "", text.split(".", 1)[0])
    if digits.startswith("0098"):
        digits = digits[2:]
    if digits.startswith("98") and len(digits) == 12:
        digits = "+" + digits
    elif digits.startswith("0") and len(digits) == 11:
        digits = "+98" + digits[1:]
    elif digits.startswith("9") and len(digits) == 10:
        digits = "+98" + digits
    return digits


IRAN_PHONE_RE = re.compile(r"^\+989\d{9}$")


def filter_valid_phones(phones):
    """فقط شماره‌های موبایل ایرانی معتبر را برمی‌گرداند (با نرمال‌سازی)."""
    valid = []
    seen = set()
    for p in phones or []:
        n = normalize_phone(str(p))
        if n and IRAN_PHONE_RE.match(n) and n not in seen:
            seen.add(n)
            valid.append(n)
    return valid


def save_customer_numbers(conn, phones):
    """لیست کامل شماره مشتریان را جایگزین لیست قبلی می‌کند (چون فایل جدید آپلودشده
    قرار است «لیست به‌روز» باشد، نه اضافه‌شدن به لیست قدیمی). هر چیزی که بعد از
    نرمال‌سازی به شکل یک شمارهٔ موبایل ایرانی معتبر (+98912...) درنیاید، نادیده
    گرفته می‌شود (مثلاً اگر فایل ستون ردیف یا عدد دیگری هم داشته باشد)."""
    normalized = sorted(filter_valid_phones(phones))
    conn.execute("DELETE FROM customer_numbers")
    conn.executemany(
        "INSERT INTO customer_numbers (phone) VALUES (?)",
        [(p,) for p in normalized],
    )
    conn.commit()
    return normalized


def get_customer_numbers(conn):
    rows = conn.execute("SELECT phone FROM customer_numbers ORDER BY phone").fetchall()
    # فیلتر دوباره برای پاک‌سازی داده‌های قدیمی/نامعتبر احتمالی در دیتابیس
    return filter_valid_phones([r[0] for r in rows])


def save_user(conn, chat_id, phone=None, first_name=None):
    conn.execute(
        "INSERT INTO users (chat_id, phone, first_name) VALUES (?,?,?) "
        "ON CONFLICT(chat_id) DO UPDATE SET "
        "phone=COALESCE(excluded.phone, users.phone), "
        "first_name=COALESCE(excluded.first_name, users.first_name)",
        (chat_id, normalize_phone(phone) if phone else None, first_name),
    )
    conn.commit()


# ---------- توابع کمکی API بله ----------

def api(method, **params):
    """درخواست به API بله را می‌فرستد و در صورت خطای موقت شبکه (مثل قطع SSL)
    چند بار با فاصلهٔ زمانی افزایشی دوباره تلاش می‌کند"""
    max_attempts = 3
    last_error = None
    for attempt in range(1, max_attempts + 1):
        try:
            r = requests.post(f"{API_URL}/{method}", data=params, timeout=30)
            r.raise_for_status()
            return r.json()
        except (requests.exceptions.ConnectionError,
                requests.exceptions.SSLError,
                requests.exceptions.Timeout) as e:
            last_error = e
            if attempt < max_attempts:
                wait = 1.5 * attempt  # ۱.۵، ۳ ثانیه
                log.warning("خطای موقت شبکه در %s (تلاش %s/%s): %s — %s ثانیه صبر و تلاش دوباره",
                            method, attempt, max_attempts, e, wait)
                time.sleep(wait)
    raise last_error


def send_message(chat_id, text, reply_markup=None, parse_mode=None):
    payload = {"chat_id": chat_id, "text": text}
    if reply_markup:
        payload["reply_markup"] = json.dumps(reply_markup)
    if parse_mode:
        payload["parse_mode"] = parse_mode
    return api("sendMessage", **payload)


def send_document(chat_id, file_path, caption=None):
    """فایل را به‌عنوان سند برای chat_id می‌فرستد؛ مثل api() چند بار در صورت
    خطای موقت شبکه دوباره تلاش می‌کند"""
    filename = os.path.basename(file_path)
    max_attempts = 3
    last_error = None
    for attempt in range(1, max_attempts + 1):
        try:
            with open(file_path, "rb") as f:
                data = {"chat_id": chat_id}
                if caption:
                    data["caption"] = caption
                files = {"document": (filename, f)}
                r = requests.post(f"{API_URL}/sendDocument", data=data, files=files, timeout=60)
                r.raise_for_status()
                return r.json()
        except (requests.exceptions.ConnectionError,
                requests.exceptions.SSLError,
                requests.exceptions.Timeout) as e:
            last_error = e
            if attempt < max_attempts:
                wait = 1.5 * attempt
                log.warning("خطای موقت شبکه در sendDocument (تلاش %s/%s): %s — %s ثانیه صبر و تلاش دوباره",
                            attempt, max_attempts, e, wait)
                time.sleep(wait)
    raise last_error


CONTACT_KEYBOARD = {
    "keyboard": [[{"text": "ارسال شماره تماس من", "request_contact": True}]],
    "resize_keyboard": True,
    "one_time_keyboard": True,
}

# منوی دکمه‌ای ادمین؛ با دستور /panel نمایش داده می‌شود
ADMIN_MENU_KEYBOARD = {
    "inline_keyboard": [
        [{"text": "📋 آپلود لیست شماره مشتریان", "callback_data": "await_phone_list"}],
        [{"text": "💰 آپلود فایل قیمت و موجودی", "callback_data": "await_price_file"}],
        [{"text": "📊 گزارش آخرین ارسال", "callback_data": "report_last"}],
        [{"text": "🔁 تلاش دوباره برای جامانده‌ها (سفیر)", "callback_data": "retry_missing"}],
    ]
}


def answer_callback_query(callback_query_id, text=None):
    """به بله اطلاع می‌دهد که کلیک روی دکمه پردازش شد (تا چرخش لودینگ دکمه قطع شود).
    اگر بله این متد را نداشته باشد یا خطا بدهد، فقط لاگ می‌شود و ربات متوقف نمی‌شود."""
    try:
        params = {"callback_query_id": callback_query_id}
        if text:
            params["text"] = text
        api("answerCallbackQuery", **params)
    except Exception as e:
        log.warning("answerCallbackQuery ناموفق بود (مهم نیست، ادامه می‌دهیم): %s", e)


def download_file(file_id, dest_path):
    info = api("getFile", file_id=file_id)
    file_path = info["result"]["file_path"]
    r = requests.get(f"{FILE_URL}/{file_path}", timeout=60)
    r.raise_for_status()
    with open(dest_path, "wb") as f:
        f.write(r.content)
    return dest_path


def read_text_any_encoding(path):
    """فایل را صرف‌نظر از پسوند، به‌عنوان متن می‌خواند و انکودینگ را حدس می‌زند"""
    for enc in ("utf-8-sig", "utf-8", "windows-1256", "cp1252"):
        try:
            with open(path, "r", encoding=enc) as f:
                return f.read()
        except (UnicodeDecodeError, LookupError):
            continue
    with open(path, "rb") as f:
        return f.read().decode("utf-8", errors="replace")


# ---------- تجزیهٔ ورودی‌های مختلف (متن، csv/tsv، اکسل) ----------

# جداکننده‌های قابل قبول بین نام محصول و قیمت در فایل/متن ساده:
# تب، دو فاصله یا بیشتر، ویرگول، سمی‌کالن
LINE_SPLIT_RE = re.compile(r"\t+|(?<=\S)\s{2,}(?=\S)|,|;")

# اگر اولین سطر شبیه هدر جدول باشد (مثلاً «نام محصول / قیمت»)، نادیده گرفته می‌شود
HEADER_KEYWORDS = ("نام", "قیمت", "محصول", "ردیف", "name", "product", "price")


def _looks_like_header(cells):
    joined = " ".join(str(c) for c in cells).strip().lower()
    return any(k in joined for k in HEADER_KEYWORDS)


def parse_price_lines(raw_text):
    """هر خط از متن ساده را به (نام محصول, قیمت) تبدیل می‌کند تا ترتیب هرگز جابه‌جا نشود"""
    rows = []
    first = True
    for line in raw_text.splitlines():
        line = line.strip()
        if not line:
            continue
        parts = [p.strip() for p in LINE_SPLIT_RE.split(line) if p.strip()]
        if first:
            first = False
            if _looks_like_header(parts if parts else [line]):
                continue
        if len(parts) >= 2:
            name = " ".join(parts[:-1])
            price = parts[-1]
        else:
            name, price = line, ""
        rows.append((name, price))
    return rows


def _excel_cell_to_str(value):
    """مقدار سلول اکسل را به رشتهٔ تمیز تبدیل می‌کند (بدون .0 اضافه برای اعداد صحیح)."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if value == int(value):
            return str(int(value))
        return str(value)
    text = str(value).strip()
    # اعدادی که openpyxl گاهی به صورت '1931.0' می‌دهد
    if re.fullmatch(r"\d+\.0+", text):
        return text.split(".", 1)[0]
    return text


def parse_price_xlsx(path):
    """فایل اکسل (xlsx/xlsm) را می‌خواند؛ هر سطر باید نام محصول در یک سلول و
    قیمت در آخرین سلول پر شدهٔ همان سطر باشد"""
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    rows = []
    first = True
    for excel_row in ws.iter_rows(values_only=True):
        cells = []
        for c in excel_row:
            text = _excel_cell_to_str(c)
            if text != "":
                cells.append(text)
        if not cells:
            continue
        if first:
            first = False
            if _looks_like_header(cells):
                continue
        # اگر خود فایل ورودی قبلاً یک ستون ردیف/شماره در ابتدا دارد، آن را نادیده
        # می‌گیریم تا با نام محصول قاطی نشود (مثلاً وقتی فایل خروجی خود ربات
        # دوباره به آن داده شود)
        if len(cells) >= 3 and re.fullmatch(r"\d+", cells[0]):
            cells = cells[1:]
        if len(cells) >= 2:
            name = " ".join(cells[:-1])
            price = cells[-1]
        else:
            name, price = cells[0], ""
        rows.append((name, price))
    return rows


def detect_spreadsheet_kind(local_path, original_filename="", mime_type=""):
    """نوع واقعی فایل را از روی پسوند، mime-type و بایت‌های ابتدای فایل تشخیص می‌دهد.
    بعضی کلاینت‌های بله file_name را بدون پسوند می‌فرستند؛ openpyxl هم بدون پسوند
    .xlsx مسیر را رد می‌کند. بنابراین فقط به پسوند اکتفا نمی‌کنیم."""
    ext = os.path.splitext(original_filename or "")[1].lower()
    if not ext:
        ext = os.path.splitext(local_path or "")[1].lower()

    mime = (mime_type or "").lower()
    if not ext:
        if "spreadsheetml" in mime or "xlsx" in mime:
            ext = ".xlsx"
        elif "ms-excel" in mime or mime.endswith("xls"):
            ext = ".xls"

    try:
        with open(local_path, "rb") as f:
            header = f.read(8)
        if header.startswith(b"PK"):
            if zipfile.is_zipfile(local_path):
                with zipfile.ZipFile(local_path) as zf:
                    if any(name.startswith("xl/") for name in zf.namelist()):
                        return "xlsx"
            return "xlsx"
        if header.startswith(b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"):
            return "xls"
    except OSError:
        pass

    if ext in (".xlsx", ".xlsm"):
        return "xlsx"
    if ext == ".xls":
        return "xls"
    return "text"


def ensure_spreadsheet_extension(local_path, kind):
    """openpyxl مسیر بدون پسوند .xlsx/.xlsm را نمی‌پذیرد؛ در صورت نیاز فایل را
    با پسوند درست کپی می‌کند و مسیر جدید را برمی‌گرداند."""
    ext = os.path.splitext(local_path)[1].lower()
    if kind == "xlsx" and ext not in (".xlsx", ".xlsm"):
        new_path = local_path + ".xlsx"
        if not os.path.exists(new_path):
            with open(local_path, "rb") as src, open(new_path, "wb") as dst:
                dst.write(src.read())
        return new_path
    if kind == "xls" and ext != ".xls":
        new_path = local_path + ".xls"
        if not os.path.exists(new_path):
            with open(local_path, "rb") as src, open(new_path, "wb") as dst:
                dst.write(src.read())
        return new_path
    return local_path


def _load_xls_sheet_rows(path):
    """خواندن فایل قدیمی .xls با xlrd (در صورت نصب بودن)"""
    try:
        import xlrd
    except ImportError as e:
        raise ValueError(
            "این فایل فرمت قدیمی .xls است. یا آن را در اکسل با Save As به صورت "
            ".xlsx ذخیره کنید، یا روی سرور دستور pip install xlrd را بزنید."
        ) from e
    book = xlrd.open_workbook(path)
    return book.sheet_by_index(0)


def parse_price_xls(path):
    sheet = _load_xls_sheet_rows(path)
    rows = []
    first = True
    for r in range(sheet.nrows):
        cells = []
        for c in range(sheet.ncols):
            val = sheet.cell_value(r, c)
            if val is None or val == "":
                continue
            if isinstance(val, float) and val == int(val):
                val = int(val)
            text = str(val).strip()
            if text:
                cells.append(text)
        if not cells:
            continue
        if first:
            first = False
            if _looks_like_header(cells):
                continue
        if len(cells) >= 3 and re.fullmatch(r"\d+", cells[0]):
            cells = cells[1:]
        if len(cells) >= 2:
            name = " ".join(cells[:-1])
            price = cells[-1]
        else:
            name, price = cells[0], ""
        rows.append((name, price))
    return rows


def parse_phone_list_xls(path):
    sheet = _load_xls_sheet_rows(path)
    phones = []
    for r in range(sheet.nrows):
        for c in range(sheet.ncols):
            val = sheet.cell_value(r, c)
            if val is None or val == "":
                continue
            if isinstance(val, float) and val == int(val):
                val = int(val)
            text = str(val).strip()
            if not text or _looks_like_header([text]):
                continue
            phones.append(text)
    return phones


def parse_price_file(local_path, original_filename="", mime_type=""):
    """بر اساس پسوند/محتوای فایل ارسالی، روش مناسب تجزیه را انتخاب می‌کند"""
    kind = detect_spreadsheet_kind(local_path, original_filename, mime_type)
    if kind == "text":
        raw_text = read_text_any_encoding(local_path)
        return parse_price_lines(raw_text)

    path = ensure_spreadsheet_extension(local_path, kind)
    tmp_copy = path if path != local_path else None
    try:
        try:
            if kind == "xlsx":
                return parse_price_xlsx(path)
            if kind == "xls":
                return parse_price_xls(path)
        except ValueError:
            raise
        except Exception as e:
            msg = str(e).lower()
            if "xls" in msg or "not support" in msg:
                try:
                    xls_path = ensure_spreadsheet_extension(local_path, "xls")
                    try:
                        return parse_price_xls(xls_path)
                    finally:
                        if xls_path != local_path and xls_path != path and os.path.exists(xls_path):
                            os.remove(xls_path)
                except ValueError:
                    raise
                except Exception:
                    pass
            raise ValueError(
                f"نتوانستم فایل اکسل را بخوانم: {e}\n"
                "لطفاً فایل را با فرمت .xlsx ذخیره کنید و دوباره بفرستید."
            ) from e
    finally:
        if tmp_copy and os.path.exists(tmp_copy):
            os.remove(tmp_copy)
    return []


def parse_phone_list_file(local_path, original_filename="", mime_type=""):
    """فایل لیست شماره مشتریان (اکسل یا متنی) را می‌خواند و یک لیست ساده از رشته‌های
    خام شماره برمی‌گرداند (نرمال‌سازی نهایی توسط normalize_phone در save_customer_numbers
    انجام می‌شود، پس اینجا لازم نیست فرمت خاصی رعایت شود)"""
    kind = detect_spreadsheet_kind(local_path, original_filename, mime_type)
    phones = []
    if kind == "text":
        raw_text = read_text_any_encoding(local_path)
        for line in raw_text.splitlines():
            for part in re.split(r"[,\t;]+|\s{2,}", line.strip()):
                part = part.strip()
                if part:
                    phones.append(part)
        return phones

    path = ensure_spreadsheet_extension(local_path, kind)
    tmp_copy = path if path != local_path else None
    try:
        try:
            if kind == "xlsx":
                wb = load_workbook(path, data_only=True)
                ws = wb.active
                for excel_row in ws.iter_rows(values_only=True):
                    for cell in excel_row:
                        text = _excel_cell_to_str(cell)
                        if not text or _looks_like_header([text]):
                            continue
                        phones.append(text)
                return phones
            if kind == "xls":
                return parse_phone_list_xls(path)
        except ValueError:
            raise
        except Exception as e:
            msg = str(e).lower()
            if "xls" in msg or "not support" in msg:
                try:
                    xls_path = ensure_spreadsheet_extension(local_path, "xls")
                    try:
                        return parse_phone_list_xls(xls_path)
                    finally:
                        if xls_path != local_path and xls_path != path and os.path.exists(xls_path):
                            os.remove(xls_path)
                except ValueError:
                    raise
                except Exception:
                    pass
            raise ValueError(
                f"نتوانستم فایل اکسل شماره‌ها را بخوانم: {e}\n"
                "لطفاً فایل را با فرمت .xlsx ذخیره کنید و دوباره بفرستید."
            ) from e
    finally:
        if tmp_copy and os.path.exists(tmp_copy):
            os.remove(tmp_copy)
    return phones


# ---------- ساخت فایل خروجی (اکسل) با نام تاریخ‌دار شمسی ----------

def jalali_now():
    """تاریخ و ساعت فعلی به وقت ایران، به تقویم شمسی (به‌عنوان آبجکت jdatetime)"""
    now_ir = datetime.datetime.now(IRAN_TZ)
    return jdatetime.datetime.fromgregorian(datetime=now_ir)


def price_list_filename(jnow):
    return f"لیست-قیمت-{jnow.strftime('%Y-%m-%d_%H-%M')}.xlsx"


WEEKDAY_FA = {
    0: "شنبه", 1: "یکشنبه", 2: "دوشنبه", 3: "سه‌شنبه",
    4: "چهارشنبه", 5: "پنجشنبه", 6: "جمعه",
}


def price_list_caption(jnow):
    """پیامی که همراه فایل برای کاربران فرستاده می‌شود تا مشخص باشد این لیست
    قیمت مربوط به چه تاریخ و ساعتی است"""
    weekday = WEEKDAY_FA.get(jnow.weekday(), "")
    return (
        f"📋 لیست قیمت جدید\n"
        f"🗓 تاریخ: {weekday} {jnow.strftime('%Y/%m/%d')}\n"
        f"⏰ ساعت ارسال: {jnow.strftime('%H:%M')} (به وقت ایران)"
    )


def build_price_file(rows, jnow):
    """یک فایل اکسل جدول‌بندی‌شده (ردیف، نام محصول، قیمت) می‌سازد و مسیر آن را برمی‌گرداند"""
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "لیست قیمت"
    ws.sheet_view.rightToLeft = True

    header = ["ردیف", "نام محصول", "قیمت"]
    ws.append(header)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for i, (name, price) in enumerate(rows, start=1):
        ws.append([i, name, price])

    thin = Side(style="thin", color="AAAAAA")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=3):
        for cell in row:
            cell.border = border
            if cell.column_letter != "B":
                cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.column_dimensions["A"].width = 8
    name_w = max((len(str(n)) for n, _ in rows), default=10)
    ws.column_dimensions["B"].width = min(max(name_w + 4, 15), 50)
    ws.column_dimensions["C"].width = 16
    ws.freeze_panes = "A2"

    path = os.path.join(OUTPUT_DIR, price_list_filename(jnow))
    wb.save(path)
    return path


# ---------- صف پیام‌ها/فایل‌های ارسال‌نشده (outbox) ----------

def enqueue(conn, chat_id, text="", file_path=None, parse_mode=None):
    """پیام یا فایل را قبل از هر تلاشی روی دیسک ذخیره می‌کند تا هرگز گم نشود"""
    cur = conn.execute(
        "INSERT INTO outbox (chat_id, text, file_path, parse_mode) VALUES (?,?,?,?)",
        (chat_id, text or "", file_path, parse_mode),
    )
    conn.commit()
    return cur.lastrowid


def flush_outbox(conn, max_per_call=200):
    """تلاش می‌کند پیام‌ها/فایل‌های موجود در صف را بفرستد؛ هر کدام که موفق ارسال شود
    از صف حذف می‌شود، و هر کدام که هنوز ناموفق باشد در صف می‌ماند تا دفعهٔ بعد.
    این تابع هم بعد از هر broadcast و هم به‌صورت دوره‌ای در حلقهٔ اصلی صدا زده می‌شود،
    پس اگر شبکه قطع باشد، چیزی گم نمی‌شود و به‌محض وصل شدن دوباره شبکه ارسال خواهد شد."""
    rows = conn.execute(
        "SELECT id, chat_id, text, file_path, parse_mode FROM outbox ORDER BY id LIMIT ?",
        (max_per_call,),
    ).fetchall()

    sent, still_pending = 0, 0
    for row_id, chat_id, text, file_path, parse_mode in rows:
        try:
            if file_path:
                if not os.path.exists(file_path):
                    raise FileNotFoundError(f"فایل {file_path} دیگر روی دیسک موجود نیست")
                send_document(chat_id, file_path, caption=text or None)
            else:
                send_message(chat_id, text, parse_mode=parse_mode)
            conn.execute("DELETE FROM outbox WHERE id=?", (row_id,))
            conn.commit()
            sent += 1
        except Exception as e:
            conn.execute(
                "UPDATE outbox SET attempts=attempts+1, last_error=? WHERE id=?",
                (str(e), row_id),
            )
            conn.commit()
            still_pending += 1
            log.warning("ارسال آیتم صف (id=%s) به %s هنوز ناموفق است: %s", row_id, chat_id, e)
        time.sleep(0.4)  # رعایت محدودیت نرخ ارسال بله (حداکثر ۲ پیام در ثانیه)

    if sent or still_pending:
        log.info("outbox: %s مورد ارسال شد، %s مورد هنوز در صف مانده", sent, still_pending)
    return sent, still_pending


# ---------- ارسال گروهی ----------

def broadcast(conn, text="", file_path=None, only_phones=None, parse_mode=None):
    all_users = conn.execute("SELECT chat_id, phone FROM users").fetchall()

    if only_phones:
        # فقط شماره‌های معتبر؛ و مقایسه با نسخهٔ نرمال‌شدهٔ شمارهٔ ذخیره‌شده در users
        wanted = set(filter_valid_phones(only_phones))
        targets = []
        matched = set()
        for cid, ph in all_users:
            norm = normalize_phone(ph) if ph else ""
            if norm and norm in wanted:
                targets.append((cid, norm))
                matched.add(norm)
            elif ph in wanted:
                targets.append((cid, ph))
                matched.add(ph)
        missing = wanted - matched
    else:
        targets = all_users
        missing = set()

    # همهٔ آیتم‌ها را اول در صف روی دیسک ذخیره می‌کنیم؛ این‌طوری حتی اگر
    # درست همین‌جا برق برود یا برنامه کرش کند، چیزی گم نمی‌شود.
    for chat_id, _ in targets:
        enqueue(conn, chat_id, text=text, file_path=file_path, parse_mode=parse_mode)

    sent, queued = flush_outbox(conn)
    return sent, queued, missing


# ---------- لاگ دائمی ارسال‌ها (برای گزارش) ----------

def log_send(conn, batch_date, file_name, phone, status, error=None):
    conn.execute(
        "INSERT INTO send_log (batch_date, file_name, phone, status, error) VALUES (?,?,?,?,?)",
        (batch_date, file_name, phone, status, error),
    )
    conn.commit()


def log_batch_result(conn, batch_date, file_name, conn_users, target_phones, missing_phones):
    """برای هر شماره در لیست هدف، وضعیت نهایی را در send_log ثبت می‌کند:
    'sent_or_queued' یعنی پیام به‌صف فرستادن رسید (چه همان لحظه ارسال شود چه بعداً
    توسط outbox دوباره تلاش شود)، 'not_started' یعنی آن شماره اصلاً ربات را استارت
    نکرده و بله اجازهٔ ارسال مستقیم به آن را نمی‌دهد."""
    users_by_phone = {ph: cid for cid, ph in conn_users}
    for phone in target_phones:
        if phone in missing_phones:
            log_send(conn, batch_date, file_name, phone, "not_started")
        elif phone in users_by_phone:
            log_send(conn, batch_date, file_name, phone, "sent_or_queued")


def last_batch_date(conn):
    row = conn.execute(
        "SELECT batch_date FROM send_log ORDER BY id DESC LIMIT 1"
    ).fetchone()
    return row[0] if row else None


def build_report(conn, batch_date=None):
    """گزارش یک تاریخ مشخص (یا آخرین تاریخ اگر داده نشود) را می‌سازد: تعداد هر
    وضعیت + لیست شماره‌هایی که ارسال نشده‌اند (not_started)، برای این‌که ادمین
    بداند برای چه کسانی باید دوباره تلاش کند (مثلاً از طریق سفیر)."""
    if batch_date is None:
        batch_date = last_batch_date(conn)
    if batch_date is None:
        return "هنوز هیچ ارسالی ثبت نشده است.", []

    rows = conn.execute(
        "SELECT phone, status FROM send_log WHERE batch_date=? ORDER BY phone",
        (batch_date,),
    ).fetchall()
    if not rows:
        return f"برای تاریخ «{batch_date}» هیچ گزارشی پیدا نشد.", []

    sent = [p for p, s in rows if s in ("sent_or_queued", "sent_via_safir")]
    not_started = [p for p, s in rows if s == "not_started"]
    failed_safir = [p for p, s in rows if s == "failed_via_safir"]

    text = (
        f"📊 گزارش ارسال — {batch_date}\n"
        f"تعداد کل شماره‌های هدف: {len(rows)}\n"
        f"✅ ارسال‌شده (از طریق ربات بله): {len(sent)}\n"
        f"❌ استارت‌نکرده (پیام دریافت نکردند): {len(not_started)}"
    )
    if failed_safir:
        text += f"\n⚠️ تلاش ناموفق از طریق سفیر: {len(failed_safir)}"
    if not_started:
        preview = "\n".join(not_started[:30])
        text += f"\n\nشماره‌های استارت‌نکرده:\n{preview}"
        if len(not_started) > 30:
            text += f"\n… و {len(not_started) - 30} شمارهٔ دیگر (فایل کامل ضمیمه می‌شود)."
    return text, not_started


# ---------- سرویس سفیر بله (ارسال مستقیم به شماره، بدون نیاز به استارت قبلی) ----------

def send_via_safir(phone_98_format, text):
    """یک پیام متنی را از طریق سرویس سفیر بله مستقیماً به یک شماره می‌فرستد.
    phone_98_format باید به‌شکل 98XXXXXXXXXX باشد (بدون + و بدون خط تیره).
    این تابع فقط وقتی SAFIR_API_ACCESS_KEY و SAFIR_BOT_ID پر شده باشند کار می‌کند."""
    if not safir_configured():
        raise RuntimeError("اطلاعات سرویس سفیر (SAFIR_API_ACCESS_KEY / SAFIR_BOT_ID) هنوز تنظیم نشده است.")
    payload = {
        "request_id": f"price-{int(time.time() * 1000)}-{phone_98_format}",
        "bot_id": int(SAFIR_BOT_ID),
        "phone_number": phone_98_format,
        "message_data": {"message": {"text": text}},
    }
    headers = {"api-access-key": SAFIR_API_ACCESS_KEY, "Content-Type": "application/json"}
    r = requests.post(SAFIR_API_URL, json=payload, headers=headers, timeout=30)
    r.raise_for_status()
    return r.json()


def to_safir_phone(normalized_phone):
    """+98XXXXXXXXXX (فرمت داخلی ربات) را به 98XXXXXXXXXX (فرمت موردنیاز سفیر) تبدیل می‌کند"""
    return normalized_phone.lstrip("+")


# ---------- پردازش پیام‌های ورودی ----------

def handle_price_rows(conn, admin_chat_id, rows, only_phones=None):
    """از یک لیست (نام, قیمت) فایل اکسل تاریخ‌دار می‌سازد، همراه با پیام تاریخ/ساعت
    برای کاربران broadcast می‌کند، نتیجه را در send_log ثبت می‌کند (برای گزارش)،
    و خلاصه را برای ادمین در صف می‌گذارد"""
    if not rows:
        enqueue(conn, admin_chat_id, "هیچ ردیف قیمتی در ورودی پیدا نشد؛ چیزی ارسال نشد.")
        return

    jnow = jalali_now()  # یک‌بار محاسبه می‌شود تا نام فایل، پیام همراه، و گزارش دقیقاً هم‌تاریخ باشند
    out_path = build_price_file(rows, jnow)
    caption = price_list_caption(jnow)

    # اگر از caption یا جای دیگر لیست شماره آمده، فقط شماره‌های موبایل معتبر را نگه دار.
    # مقادیر نامعتبر مثل کد کالا / قیمت / تاریخ (مثلاً 14050611 یا 1931) نباید
    # باعث شوند ارسال به هیچ‌کس نرسد.
    if only_phones is not None:
        only_phones = filter_valid_phones(only_phones)
        if not only_phones:
            only_phones = None

    # اگر ادمین لیست شماره خاصی نداده باشد و لیست اصلی مشتریان (از دکمهٔ
    # «آپلود لیست شماره مشتریان») پر باشد، هدف پیش‌فرض همان لیست است؛ در غیر این
    # صورت رفتار قبلی حفظ می‌شود (همهٔ کاربران ثبت‌شده در ربات)
    if only_phones is None:
        customer_list = get_customer_numbers(conn)
        if customer_list:
            only_phones = customer_list

    sent, queued, missing = broadcast(conn, text=caption, file_path=out_path, only_phones=only_phones)

    batch_date = jnow.strftime("%Y-%m-%d %H:%M")
    file_name = os.path.basename(out_path)
    all_users = conn.execute("SELECT chat_id, phone FROM users").fetchall()
    target_phones = only_phones if only_phones is not None else [ph for _, ph in all_users if ph]
    log_batch_result(conn, batch_date, file_name, all_users, target_phones, missing)

    report = (
        f"ارسال شد ✅\n"
        f"تاریخ/دستهٔ ارسال: {batch_date}\n"
        f"نام فایل ارسالی: {file_name}\n"
        f"تعداد ردیف: {len(rows)}\n"
        f"تعداد موفق: {sent}"
    )
    if queued:
        report += f"\nتعداد در صف (به‌خاطر مشکل شبکه، به‌محض وصل شدن دوباره تلاش می‌شود): {queued}"
    if missing:
        report += ("\n\nاین شماره‌ها هنوز ربات را استارت نکرده‌اند و پیام دریافت نکردند:\n"
                   + "\n".join(sorted(missing)))
        report += "\n\nبرای گزارش کامل و امکان تلاش دوباره، از دکمهٔ «📊 گزارش آخرین ارسال» استفاده کنید."
    enqueue(conn, admin_chat_id, report)


def find_batch_date_for_day(conn, day_str):
    """اگر ادمین فقط تاریخ (بدون ساعت) بدهد، آخرین دستهٔ ارسال همان روز را پیدا می‌کند"""
    row = conn.execute(
        "SELECT batch_date FROM send_log WHERE batch_date LIKE ? ORDER BY id DESC LIMIT 1",
        (day_str + "%",),
    ).fetchone()
    return row[0] if row else None


def send_report(conn, chat_id, batch_date=None):
    text, not_started = build_report(conn, batch_date)
    enqueue(conn, chat_id, text)
    if len(not_started) > 30:
        # لیست کامل شماره‌های ارسال‌نشده را به‌صورت فایل اکسل هم می‌فرستیم چون در متن جا نمی‌شود
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        wb = Workbook()
        ws = wb.active
        ws.title = "استارت‌نکرده‌ها"
        ws.append(["شماره"])
        for p in not_started:
            ws.append([p])
        path = os.path.join(OUTPUT_DIR, f"شماره‌های-استارت‌نکرده-{jalali_now().strftime('%Y-%m-%d_%H-%M')}.xlsx")
        wb.save(path)
        enqueue(conn, chat_id, "", file_path=path)


def handle_retry_missing(conn, chat_id):
    if not safir_configured():
        enqueue(
            conn, chat_id,
            "سرویس سفیر هنوز متصل نشده. وقتی SAFIR_API_ACCESS_KEY و SAFIR_BOT_ID را از پنل "
            "کسب‌وکار بله گرفتید، آن‌ها را به‌عنوان متغیر محیطی تنظیم کنید و ربات را دوباره اجرا کنید.",
        )
        return

    batch_date = last_batch_date(conn)
    if not batch_date:
        enqueue(conn, chat_id, "هنوز هیچ ارسالی ثبت نشده که بخواهیم برایش دوباره تلاش کنیم.")
        return

    rows = conn.execute(
        "SELECT phone FROM send_log WHERE batch_date=? AND status='not_started'",
        (batch_date,),
    ).fetchall()
    phones = [r[0] for r in rows]
    if not phones:
        enqueue(conn, chat_id, f"برای دستهٔ {batch_date} شماره‌ی جامانده‌ای ثبت نشده.")
        return

    nudge_text = (
        "📋 لیست قیمت جدید آماده است.\n"
        "لطفاً برای دریافت فایل کامل قیمت و موجودی، به ربات ما پیام /start بزنید."
    )
    ok, failed = 0, 0
    for phone in phones:
        try:
            send_via_safir(to_safir_phone(phone), nudge_text)
            log_send(conn, batch_date, "retry-via-safir", phone, "sent_via_safir")
            ok += 1
        except Exception as e:
            log_send(conn, batch_date, "retry-via-safir", phone, "failed_via_safir", str(e))
            failed += 1
        time.sleep(0.3)

    enqueue(
        conn, chat_id,
        f"تلاش از طریق سفیر برای دستهٔ {batch_date} تمام شد.\nموفق: {ok}\nناموفق: {failed}",
    )


def handle_callback_query(conn, callback_query):
    """کلیک روی دکمه‌های منوی ادمین (inline keyboard) را پردازش می‌کند"""
    cq_id = callback_query.get("id")
    from_user = callback_query.get("from", {})
    chat_id = (callback_query.get("message") or {}).get("chat", {}).get("id") or from_user.get("id")
    data = callback_query.get("data", "")

    if chat_id != ADMIN_CHAT_ID:
        answer_callback_query(cq_id)
        return

    if data == "await_phone_list":
        set_admin_state(conn, chat_id, "awaiting_phone_list")
        send_message(chat_id, "لطفاً فایل اکسل (یا متنی) لیست شماره مشتریان را بفرستید؛ یک شماره در هر سطر/سلول کافی است.")
    elif data == "await_price_file":
        set_admin_state(conn, chat_id, "awaiting_price_file")
        send_message(chat_id, "لطفاً فایل قیمت و موجودی (اکسل یا متنی) را بفرستید.")
    elif data == "report_last":
        send_report(conn, chat_id)
    elif data == "retry_missing":
        handle_retry_missing(conn, chat_id)
    else:
        log.warning("callback_data ناشناخته: %s", data)

    answer_callback_query(cq_id)


def handle_update(conn, update):
    if "callback_query" in update:
        handle_callback_query(conn, update["callback_query"])
        return

    message = update.get("message")
    if not message:
        return

    chat_id = message["chat"]["id"]
    from_user = message.get("from", {})
    log.info("پیام دریافت شد از chat_id=%s", chat_id)  # برای پیدا کردن chat_id ادمین در کنسول
    if "document" in message and chat_id != ADMIN_CHAT_ID:
        log.info(
            "فایلی از chat_id=%s رسید ولی با ADMIN_CHAT_ID=%s برابر نیست، پس نادیده گرفته می‌شود",
            chat_id, ADMIN_CHAT_ID,
        )

    if "contact" in message:
        save_user(conn, chat_id, phone=message["contact"].get("phone_number"),
                   first_name=from_user.get("first_name"))
        send_message(chat_id, "شماره شما ثبت شد. قیمت‌های به‌روز از این پس برای شما ارسال می‌شود.")
        return

    text = message.get("text", "")

    if text == "/start":
        save_user(conn, chat_id, first_name=from_user.get("first_name"))
        send_message(
            chat_id,
            "سلام! برای دریافت قیمت‌های به‌روز، لطفاً شماره تماس‌تان را با دکمهٔ زیر ارسال کنید.",
            reply_markup=CONTACT_KEYBOARD,
        )
        return

    # از این‌جا به بعد فقط دستورات مخصوص ادمین
    if chat_id != ADMIN_CHAT_ID:
        return

    if text == "/panel":
        set_admin_state(conn, chat_id, None)
        send_message(chat_id, "پنل مدیریت:", reply_markup=ADMIN_MENU_KEYBOARD)
        return

    if text == "/count":
        n = conn.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        m = len(get_customer_numbers(conn))
        send_message(chat_id, f"تعداد کاربران استارت‌کرده: {n}\nتعداد شماره در لیست اصلی مشتریان: {m}")
        return

    if text == "/queue":
        n = conn.execute("SELECT COUNT(*) FROM outbox").fetchone()[0]
        send_message(chat_id, f"تعداد آیتم‌های در صف انتظار ارسال: {n}")
        return

    if text == "گزارش" or text.startswith("گزارش "):
        parts = text.split(maxsplit=1)
        day = parts[1].strip() if len(parts) > 1 else None
        batch_date = find_batch_date_for_day(conn, day) if day else None
        if day and not batch_date:
            enqueue(conn, chat_id, f"برای تاریخ «{day}» گزارشی پیدا نشد.")
            return
        send_report(conn, chat_id, batch_date)
        return

    admin_state = get_admin_state(conn, chat_id)

    # ---- حالت ۱: ادمین یک فایل فرستاده (txt، csv، xlsx و ...) ----
    if "document" in message:
        doc = message["document"]
        caption = (message.get("caption") or "").strip()
        original_filename = doc.get("file_name") or ""
        mime_type = doc.get("mime_type") or ""
        name_ext = os.path.splitext(original_filename)[1].lower()
        if not name_ext:
            mt = mime_type.lower()
            if "spreadsheetml" in mt or "xlsx" in mt:
                name_ext = ".xlsx"
            elif "ms-excel" in mt or mt.endswith("xls"):
                name_ext = ".xls"
        safe_id = re.sub(r"[^A-Za-z0-9_-]", "_", str(doc["file_id"]))
        local_path = f"incoming_{safe_id}{name_ext}"
        try:
            download_file(doc["file_id"], local_path)

            if admin_state == "awaiting_phone_list":
                try:
                    raw_phones = parse_phone_list_file(local_path, original_filename, mime_type)
                except ValueError as e:
                    enqueue(conn, chat_id, str(e))
                    return
                saved = save_customer_numbers(conn, raw_phones)
                set_admin_state(conn, chat_id, None)
                enqueue(conn, chat_id, f"✅ لیست شماره مشتریان ذخیره شد. تعداد شماره‌های معتبر: {len(saved)}")
                return

            # حالت awaiting_price_file یا حالت قدیمی (بدون دکمه) — هر دو یکسان پردازش می‌شوند
            try:
                rows = parse_price_file(local_path, original_filename, mime_type)
            except ValueError as e:
                enqueue(conn, chat_id, str(e))
                return
            phones = None
            if caption:
                # caption فقط وقتی به‌عنوان فیلتر گیرنده استفاده می‌شود که
                # واقعاً شماره موبایل معتبر داخلش باشد؛ وگرنه نادیده گرفته می‌شود
                # تا کد کالا/قیمت اشتباهاً جای شماره ننشیند.
                candidates = [p.strip() for p in re.split(r"[,\n]+", caption) if p.strip()]
                phones = filter_valid_phones(candidates) or None
            handle_price_rows(conn, chat_id, rows, only_phones=phones)
            if admin_state == "awaiting_price_file":
                set_admin_state(conn, chat_id, None)
        finally:
            if os.path.exists(local_path):
                os.remove(local_path)
            # اگر به‌خاطر نبود پسوند، کپی .xlsx/.xls ساخته شده بود آن را هم پاک کن
            for extra_ext in (".xlsx", ".xls"):
                extra = local_path + extra_ext
                if extra != local_path and os.path.exists(extra):
                    os.remove(extra)
        return

    # ---- حالت ۲: ادمین شماره‌ها یا لیست قیمت را مستقیم و خط‌به‌خط تایپ کرده ----
    if text.strip():
        if admin_state == "awaiting_phone_list":
            raw_phones = re.split(r"[,\n]+", text)
            saved = save_customer_numbers(conn, raw_phones)
            set_admin_state(conn, chat_id, None)
            enqueue(conn, chat_id, f"✅ لیست شماره مشتریان ذخیره شد. تعداد شماره‌های معتبر: {len(saved)}")
            return
        rows = parse_price_lines(text)
        handle_price_rows(conn, chat_id, rows)
        if admin_state == "awaiting_price_file":
            set_admin_state(conn, chat_id, None)
        return


# ---------- حلقهٔ اصلی (polling) ----------

def main():
    if BOT_TOKEN == "PUT-YOUR-TOKEN-HERE":
        raise SystemExit("لطفاً متغیر محیطی BALE_BOT_TOKEN را تنظیم کنید (به README.md نگاه کنید)")

    conn = db_init()
    offset = None
    log.info("ربات اجرا شد و منتظر پیام است...")
    while True:
        params = {"timeout": 30}
        if offset:
            params["offset"] = offset
        try:
            resp = api("getUpdates", **params)
        except Exception as e:
            log.warning("خطا در دریافت پیام‌ها: %s", e)
            time.sleep(3)
            continue

        for update in resp.get("result", []):
            offset = update["update_id"] + 1
            try:
                handle_update(conn, update)
            except Exception as e:
                log.exception("خطا در پردازش پیام: %s", e)

        # هر دور از حلقه (چه پیام جدیدی رسیده باشد چه نه) یک‌بار صف را
        # خالی کردن امتحان می‌کنیم؛ این‌طوری به‌محض وصل شدن دوباره شبکه،
        # آیتم‌های معطل‌مانده بدون نیاز به پیام جدید ارسال می‌شوند.
        try:
            flush_outbox(conn)
        except Exception as e:
            log.warning("خطا هنگام خالی کردن صف: %s", e)


if __name__ == "__main__":
    main()
