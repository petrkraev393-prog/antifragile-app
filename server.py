# -*- coding: utf-8 -*-
"""
Сервер для Telegram Mini App дашборда.
Доступ только у участников закрытого канала: при открытии Mini App Telegram
передаёт подписанные данные (initData); сервер проверяет подпись ботом и
спрашивает у Telegram статус пользователя в канале (getChatMember).

Запуск: BOT_TOKEN=... CHANNEL_ID=... python server.py   (или через gunicorn/systemd)
Настройки берутся из переменных окружения:
  BOT_TOKEN   — токен бота от @BotFather
  CHANNEL_ID  — id канала (например -1001234567890) или @username публичного
  PORT        — порт (по умолчанию 8080)
"""
import os, json, hmac, hashlib, time, urllib.request, urllib.parse
from flask import Flask, request, Response, send_file

BOT_TOKEN  = os.environ.get("BOT_TOKEN", "")
CHANNEL_ID = os.environ.get("CHANNEL_ID", "")
PORT       = int(os.environ.get("PORT", "8080"))
HERE       = os.path.dirname(os.path.abspath(__file__))
DATA_FILE  = os.path.join(HERE, "data.json")
PAGE_FILE  = os.path.join(HERE, "miniapp.html")

app = Flask(__name__)
_member_cache = {}   # uid -> (timestamp, bool), кэш на 5 минут

def check_init_data(init_data: str):
    """Проверяет подпись Telegram initData. Возвращает user_id или None."""
    if not init_data:
        return None
    try:
        parsed = dict(urllib.parse.parse_qsl(init_data, keep_blank_values=True))
    except Exception:
        return None
    recv_hash = parsed.pop("hash", None)
    if not recv_hash:
        return None
    data_check = "\n".join(f"{k}={parsed[k]}" for k in sorted(parsed))
    secret = hmac.new(b"WebAppData", BOT_TOKEN.encode(), hashlib.sha256).digest()
    calc = hmac.new(secret, data_check.encode(), hashlib.sha256).hexdigest()
    if not hmac.compare_digest(calc, recv_hash):
        return None
    # данные не старше суток
    try:
        if time.time() - int(parsed.get("auth_date", "0")) > 86400:
            return None
    except ValueError:
        return None
    try:
        return json.loads(parsed.get("user", "{}")).get("id")
    except Exception:
        return None

def is_member(uid: int) -> bool:
    """Состоит ли пользователь в канале (с кэшем 5 мин)."""
    now = time.time()
    cached = _member_cache.get(uid)
    if cached and now - cached[0] < 300:
        return cached[1]
    url = (f"https://api.telegram.org/bot{BOT_TOKEN}/getChatMember"
           f"?chat_id={urllib.parse.quote(str(CHANNEL_ID))}&user_id={uid}")
    ok = False
    try:
        with urllib.request.urlopen(url, timeout=10) as r:
            res = json.loads(r.read().decode("utf-8"))
        if res.get("ok"):
            ok = res["result"]["status"] in ("creator", "administrator", "member", "restricted")
    except Exception:
        ok = False
    _member_cache[uid] = (now, ok)
    return ok

@app.route("/")
def index():
    return send_file(PAGE_FILE)

@app.route("/data")
def data():
    init_data = request.headers.get("X-Init-Data", "") or request.args.get("initData", "")
    uid = check_init_data(init_data)
    if not uid:
        return Response('{"error":"auth"}', status=401, mimetype="application/json")
    if not is_member(uid):
        return Response('{"error":"not_member"}', status=403, mimetype="application/json")
    try:
        with open(DATA_FILE, encoding="utf-8") as f:
            return Response(f.read(), mimetype="application/json")
    except FileNotFoundError:
        return Response('{"error":"no_data"}', status=503, mimetype="application/json")

@app.route("/health")
def health():
    return "ok"

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=PORT)
