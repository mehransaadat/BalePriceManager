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
import requests
import jdatetime
from zoneinfo import ZoneInfo
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side

BOT_TOKEN = os.environ.get("BALE_BOT_TOKEN", "1258122671:KfFt7JNbCDAE2gIvgNSBWPcMT-i-kinpZAg")
ADMIN_CHAT_ID = int(os.environ.get("BALE_ADMIN_CHAT_ID", "1804507729"))

API_URL = f"https://tapi.bale.ai/bot{BOT_TOKEN}"
FILE_URL = f"https://tapi.bale.ai/file/bot{BOT_TOKEN}"
DB_PATH = "users.db"
OUTPUT_DIR = "sent_price_lists"  # فایل‌های نهایی (اکسل) قبل از ارسال اینجا ذخیره می‌شوند
IRAN_TZ = ZoneInfo("Asia/Tehran")

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
    conn.commit()
    return conn


def normalize_phone(phone: str) -> str:
    """شماره را به فرمت یکسان +98XXXXXXXXXX تبدیل می‌کند"""
    digits = re.sub(r"\D", "", phone or "")
    if digits.startswith("0098"):
        digits = digits[2:]
    if digits.startswith("98") and len(digits) == 12:
        digits = "+" + digits
    elif digits.startswith("0") and len(digits) == 11:
        digits = "+98" + digits[1:]
    elif digits.startswith("9") and len(digits) == 10:
        digits = "+98" + digits
    return digits


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


def parse_price_xlsx(path):
    """فایل اکسل (xlsx/xlsm) را می‌خواند؛ هر سطر باید نام محصول در یک سلول و
    قیمت در آخرین سلول پر شدهٔ همان سطر باشد"""
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    rows = []
    first = True
    for excel_row in ws.iter_rows(values_only=True):
        cells = [str(c).strip() for c in excel_row if c is not None and str(c).strip() != ""]
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


def parse_price_file(local_path, original_filename=""):
    """بر اساس پسوند فایل ارسالی، روش مناسب تجزیه را انتخاب می‌کند"""
    ext = os.path.splitext(original_filename or local_path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        return parse_price_xlsx(local_path)
    if ext == ".xls":
        raise ValueError(
            "فرمت قدیمی xls پشتیبانی نمی‌شود. لطفاً فایل را با فرمت xlsx ذخیره کنید "
            "یا آن را به‌صورت فایل متنی (txt / csv) بفرستید."
        )
    # هر فایل دیگری (txt، csv، tsv، بدون پسوند و ...) به‌عنوان متن ساده خوانده می‌شود
    raw_text = read_text_any_encoding(local_path)
    return parse_price_lines(raw_text)


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
        wanted = {normalize_phone(p) for p in only_phones}
        targets = [(cid, ph) for cid, ph in all_users if ph in wanted]
        matched = {ph for _, ph in targets}
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


# ---------- پردازش پیام‌های ورودی ----------

def handle_price_rows(conn, admin_chat_id, rows, only_phones=None):
    """از یک لیست (نام, قیمت) فایل اکسل تاریخ‌دار می‌سازد، همراه با پیام تاریخ/ساعت
    برای کاربران broadcast می‌کند و گزارش را برای ادمین در صف می‌گذارد"""
    if not rows:
        enqueue(conn, admin_chat_id, "هیچ ردیف قیمتی در ورودی پیدا نشد؛ چیزی ارسال نشد.")
        return

    jnow = jalali_now()  # یک‌بار محاسبه می‌شود تا نام فایل و پیام همراه، دقیقاً هم‌تاریخ باشند
    out_path = build_price_file(rows, jnow)
    caption = price_list_caption(jnow)
    sent, queued, missing = broadcast(conn, text=caption, file_path=out_path, only_phones=only_phones)

    report = (
        f"ارسال شد ✅\n"
        f"نام فایل ارسالی: {os.path.basename(out_path)}\n"
        f"تعداد ردیف: {len(rows)}\n"
        f"تعداد موفق: {sent}"
    )
    if queued:
        report += f"\nتعداد در صف (به‌خاطر مشکل شبکه، به‌محض وصل شدن دوباره تلاش می‌شود): {queued}"
    if missing:
        report += ("\n\nاین شماره‌ها هنوز ربات را استارت نکرده‌اند و پیام دریافت نکردند:\n"
                   + "\n".join(sorted(missing)))
    enqueue(conn, admin_chat_id, report)


def handle_update(conn, update):
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

    if text == "/count":
        n = conn.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        send_message(chat_id, f"تعداد کاربران ثبت‌شده: {n}")
        return

    if text == "/queue":
        n = conn.execute("SELECT COUNT(*) FROM outbox").fetchone()[0]
        send_message(chat_id, f"تعداد آیتم‌های در صف انتظار ارسال: {n}")
        return

    # ---- حالت ۱: ادمین یک فایل فرستاده (txt، csv، xlsx و ...) ----
    if "document" in message:
        doc = message["document"]
        caption = (message.get("caption") or "").strip()
        original_filename = doc.get("file_name", "")
        safe_id = re.sub(r"[^A-Za-z0-9_-]", "_", doc["file_id"])
        local_path = f"incoming_{safe_id}"
        try:
            download_file(doc["file_id"], local_path)
            try:
                rows = parse_price_file(local_path, original_filename)
            except ValueError as e:
                enqueue(conn, chat_id, str(e))
                return
            phones = None
            if caption:
                phones = [p.strip() for p in re.split(r"[,\n]+", caption) if p.strip()]
            handle_price_rows(conn, chat_id, rows, only_phones=phones)
        finally:
            if os.path.exists(local_path):
                os.remove(local_path)
        return

    # ---- حالت ۲: ادمین لیست قیمت را مستقیم و خط‌به‌خط تایپ کرده (پیام متنی معمولی) ----
    if text.strip():
        rows = parse_price_lines(text)
        handle_price_rows(conn, chat_id, rows)
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
