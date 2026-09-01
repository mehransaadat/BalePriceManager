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
import requests

BOT_TOKEN = os.environ.get("BALE_BOT_TOKEN", "1258122671:KfFt7JNbCDAE2gIvgNSBWPcMT-i-kinpZAg")
ADMIN_CHAT_ID = int(os.environ.get("BALE_ADMIN_CHAT_ID", "1804507729"))

API_URL = f"https://tapi.bale.ai/bot{BOT_TOKEN}"
FILE_URL = f"https://tapi.bale.ai/file/bot{BOT_TOKEN}"
DB_PATH = "users.db"

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
    # پیام‌هایی که هنوز ارسال نشده‌اند (چه به‌خاطر خطای شبکه، چه هر دلیل دیگری)
    # اینجا نگه داشته می‌شوند تا چیزی گم نشود؛ حتی اگر ربات ری‌استارت شود،
    # این صف روی دیسک باقی می‌ماند و در اجرای بعدی دوباره تلاش می‌شود.
    conn.execute("""
        CREATE TABLE IF NOT EXISTS outbox (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            chat_id INTEGER NOT NULL,
            text TEXT NOT NULL,
            parse_mode TEXT,
            attempts INTEGER DEFAULT 0,
            last_error TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
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


# ---------- تجزیهٔ فایل قیمت و ساخت جدول مرتب ----------

# جداکننده‌های قابل قبول بین نام محصول و قیمت: تب، دو فاصله یا بیشتر، ویرگول، سمی‌کالن
LINE_SPLIT_RE = re.compile(r"\t+|(?<=\S)\s{2,}(?=\S)|,|;")


def parse_price_lines(raw_text):
    """هر خط را به (نام محصول, قیمت) تبدیل می‌کند تا ترتیب هرگز جابه‌جا نشود"""
    rows = []
    for line in raw_text.splitlines():
        line = line.strip()
        if not line:
            continue
        parts = [p.strip() for p in LINE_SPLIT_RE.split(line) if p.strip()]
        if len(parts) >= 2:
            name = " ".join(parts[:-1])
            price = parts[-1]
        else:
            name, price = line, ""
        rows.append((name, price))
    return rows


def build_table(rows):
    if not rows:
        return "```\nلیست قیمتی خالی است.\n```"
    idx_w = len(str(len(rows)))
    name_w = max(len(n) for n, _ in rows)
    price_w = max(len(p) for _, p in rows)
    lines = [
        f"{str(i).rjust(idx_w)}. {name.ljust(name_w)}  {price.rjust(price_w)}"
        for i, (name, price) in enumerate(rows, start=1)
    ]
    table = "\n".join(lines)
    # داخل بلاک کد (```) قرار می‌گیرد تا فونت ثابت (monospace) باشد
    # و فاصله‌ها/ترتیب ستون‌ها در گوشی کاربر به‌هم نریزد
    return "```\n" + table + "\n```"


# ---------- صف پیام‌های ارسال‌نشده (outbox) ----------

def enqueue(conn, chat_id, text, parse_mode=None):
    """پیام را قبل از هر تلاشی روی دیسک ذخیره می‌کند تا هرگز گم نشود"""
    cur = conn.execute(
        "INSERT INTO outbox (chat_id, text, parse_mode) VALUES (?,?,?)",
        (chat_id, text, parse_mode),
    )
    conn.commit()
    return cur.lastrowid


def flush_outbox(conn, max_per_call=200):
    """تلاش می‌کند پیام‌های موجود در صف را بفرستد؛ هر پیامی که موفق ارسال شود
    از صف حذف می‌شود، و هر پیامی که هنوز ناموفق باشد در صف می‌ماند تا دفعهٔ بعد.
    این تابع هم بعد از هر broadcast و هم به‌صورت دوره‌ای در حلقهٔ اصلی صدا زده می‌شود،
    پس اگر شبکه قطع باشد، پیام گم نمی‌شود و به‌محض وصل شدن دوباره شبکه ارسال خواهد شد."""
    rows = conn.execute(
        "SELECT id, chat_id, text, parse_mode FROM outbox ORDER BY id LIMIT ?",
        (max_per_call,),
    ).fetchall()

    sent, still_pending = 0, 0
    for row_id, chat_id, text, parse_mode in rows:
        try:
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
            log.warning("ارسال پیام صف (id=%s) به %s هنوز ناموفق است: %s", row_id, chat_id, e)
        time.sleep(0.4)  # رعایت محدودیت نرخ ارسال بله (حداکثر ۲ پیام در ثانیه)

    if sent or still_pending:
        log.info("outbox: %s پیام ارسال شد، %s پیام هنوز در صف مانده", sent, still_pending)
    return sent, still_pending


# ---------- ارسال گروهی ----------

def broadcast(conn, text, only_phones=None):
    all_users = conn.execute("SELECT chat_id, phone FROM users").fetchall()

    if only_phones:
        wanted = {normalize_phone(p) for p in only_phones}
        targets = [(cid, ph) for cid, ph in all_users if ph in wanted]
        matched = {ph for _, ph in targets}
        missing = wanted - matched
    else:
        targets = all_users
        missing = set()

    # همهٔ پیام‌ها را اول در صف روی دیسک ذخیره می‌کنیم؛ این‌طوری حتی اگر
    # درست همین‌جا برق برود یا برنامه کرش کند، هیچ پیامی گم نمی‌شود.
    for chat_id, _ in targets:
        enqueue(conn, chat_id, text, parse_mode="Markdown")

    sent, queued = flush_outbox(conn)
    return sent, queued, missing


# ---------- پردازش پیام‌های ورودی ----------

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
        send_message(chat_id, f"تعداد پیام‌های در صف انتظار ارسال: {n}")
        return

    if "document" in message:
        doc = message["document"]
        caption = (message.get("caption") or "").strip()
        safe_id = re.sub(r"[^A-Za-z0-9_-]", "_", doc["file_id"])
        local_path = f"incoming_{safe_id}"
        try:
            download_file(doc["file_id"], local_path)
            raw_text = read_text_any_encoding(local_path)
            rows = parse_price_lines(raw_text)
            table_msg = build_table(rows)

            if caption:
                phones = [p.strip() for p in re.split(r"[,\n]+", caption) if p.strip()]
                sent, queued, missing = broadcast(conn, table_msg, only_phones=phones)
            else:
                sent, queued, missing = broadcast(conn, table_msg)

            report = f"ارسال شد ✅\nتعداد موفق: {sent}"
            if queued:
                report += (f"\nتعداد در صف (به‌خاطر مشکل شبکه، به‌محض وصل شدن دوباره تلاش می‌شود): {queued}")
            if missing:
                report += ("\n\nاین شماره‌ها هنوز ربات را استارت نکرده‌اند و پیام دریافت نکردند:\n"
                           + "\n".join(sorted(missing)))
            enqueue(conn, chat_id, report)
        finally:
            if os.path.exists(local_path):
                os.remove(local_path)
        return

    # از این‌جا به بعد: پیام متنی معمولی از ادمین (نه فایل، نه یکی از دستورات بالا)
    # یعنی خود ادمین مستقیماً لیست قیمت را خط‌به‌خط تایپ کرده است
    if text.strip():
        rows = parse_price_lines(text)
        if rows:
            table_msg = build_table(rows)
            sent, queued, missing = broadcast(conn, table_msg)
            report = f"ارسال شد ✅\nتعداد موفق: {sent}"
            if queued:
                report += (f"\nتعداد در صف (به‌خاطر مشکل شبکه، به‌محض وصل شدن دوباره تلاش می‌شود): {queued}")
            enqueue(conn, chat_id, report)
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
        # پیام‌های معطل‌مانده بدون نیاز به پیام جدید ارسال می‌شوند.
        try:
            flush_outbox(conn)
        except Exception as e:
            log.warning("خطا هنگام خالی کردن صف: %s", e)


if __name__ == "__main__":
    main()
