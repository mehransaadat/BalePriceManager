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

BOT_TOKEN = os.environ.get("BALE_BOT_TOKEN", "your bot token")
ADMIN_CHAT_ID = int(os.environ.get("BALE_ADMIN_CHAT_ID", "put your chat_id"))

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
    r = requests.post(f"{API_URL}/{method}", data=params, timeout=30)
    r.raise_for_status()
    return r.json()


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
    name_w = max(len(n) for n, _ in rows)
    price_w = max(len(p) for _, p in rows)
    lines = [f"{name.ljust(name_w)}  {price.rjust(price_w)}" for name, price in rows]
    table = "\n".join(lines)
    # داخل بلاک کد (```) قرار می‌گیرد تا فونت ثابت (monospace) باشد
    # و فاصله‌ها/ترتیب ستون‌ها در گوشی کاربر به‌هم نریزد
    return "```\n" + table + "\n```"


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

    sent, failed = 0, 0
    for chat_id, _ in targets:
        try:
            send_message(chat_id, text, parse_mode="Markdown")
            sent += 1
        except Exception as e:
            log.warning("ارسال به %s ناموفق بود: %s", chat_id, e)
            failed += 1
        time.sleep(0.4)  # رعایت محدودیت نرخ ارسال بله (حداکثر ۲ پیام در ثانیه)

    return sent, failed, missing


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
                sent, failed, missing = broadcast(conn, table_msg, only_phones=phones)
            else:
                sent, failed, missing = broadcast(conn, table_msg)

            report = f"ارسال شد ✅\nتعداد موفق: {sent}\nتعداد ناموفق: {failed}"
            if missing:
                report += ("\n\nاین شماره‌ها هنوز ربات را استارت نکرده‌اند و پیام دریافت نکردند:\n"
                           + "\n".join(sorted(missing)))
            send_message(chat_id, report)
        finally:
            if os.path.exists(local_path):
                os.remove(local_path)
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


if __name__ == "__main__":
    main()
