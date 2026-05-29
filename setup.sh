#!/usr/bin/env bash
# Установщик дашборда (Telegram Mini App). Запускать на сервере: bash /opt/antifragile/setup.sh
set -e
DIR=/opt/antifragile
cd "$DIR"

echo "===== 1/6 Установка пакетов ====="
apt-get update -y
apt-get install -y python3-venv python3-pip curl jq debian-keyring debian-archive-keyring apt-transport-https
if ! command -v caddy >/dev/null 2>&1; then
  curl -1sLf 'https://dl.cloudsmith.io/public/caddy/stable/gpg.key' | gpg --dearmor -o /usr/share/keyrings/caddy-stable-archive-keyring.gpg
  curl -1sLf 'https://dl.cloudsmith.io/public/caddy/stable/debian.deb.txt' > /etc/apt/sources.list.d/caddy-stable.list
  apt-get update -y && apt-get install -y caddy
fi

echo "===== 2/6 Python-окружение ====="
python3 -m venv venv
venv/bin/pip install -q --upgrade pip
venv/bin/pip install -q -r requirements.txt

echo
echo "===== 3/6 Ввод данных (печатай аккуратно, Enter после каждого) ====="
read -rp "Токен бота от BotFather: " BOT
read -rp "ID таблицы Antifragile:  " SA
read -rp "ID таблицы журнала:      " SZ
read -rp "Домен (например 161563.com): " DOM

# ID канала определяем автоматически (бот должен быть админом канала + туда отправлено сообщение)
CHID=$(curl -s "https://api.telegram.org/bot${BOT}/getUpdates" | jq -r '[.result[] | (.channel_post.chat.id // .my_chat_member.chat.id // empty)] | last // 0' 2>/dev/null || echo 0)
[ -z "$CHID" ] && CHID=0
echo ">> Определён ID канала: $CHID"

cat > "$DIR/secrets.env" <<EOF
BOT_TOKEN=${BOT}
CHANNEL_ID=${CHID}
SHEET_ANTIFRAGILE=${SA}
SHEET_ZHURNAL=${SZ}
EOF
chmod 600 "$DIR/secrets.env"
sed -i "s/dashboard\.example\.com/${DOM}/" "$DIR/Caddyfile"

echo
echo "===== 4/6 Первый расчёт (Google + Мосбиржа) ====="
set -a; . "$DIR/secrets.env"; set +a
venv/bin/python reconstruct.py || true
if [ -f "$DIR/data.json" ]; then echo ">> data.json создан, OK"; else echo ">> ВНИМАНИЕ: data.json НЕ создан"; fi

echo
echo "===== 5/6 Запуск сервера (служба) ====="
cp "$DIR/dashboard.service" /etc/systemd/system/
systemctl daemon-reload
systemctl enable --now dashboard
sleep 1
if systemctl is-active --quiet dashboard; then echo ">> сервер запущен (active)"; else echo ">> сервер НЕ запустился:"; journalctl -u dashboard -n 15 --no-pager; fi

echo
echo "===== 6/6 HTTPS (Caddy) + автообновление (cron) ====="
cp "$DIR/Caddyfile" /etc/caddy/Caddyfile
systemctl reload caddy 2>/dev/null || systemctl restart caddy || true
( crontab -l 2>/dev/null | grep -v 'reconstruct.py' ; echo "17 16 * * * cd $DIR && set -a && . ./secrets.env && set +a && venv/bin/python reconstruct.py >> $DIR/update.log 2>&1" ) | crontab -

echo
echo "############ ГОТОВО ############"
echo "ID канала: $CHID"
echo "  (если 0 — добавь бота АДМИНОМ в канал, напиши там любое сообщение и снова запусти: bash $DIR/setup.sh)"
echo "Домен: $DOM  -> проверь позже https://$DOM/health (после того как DNS заработает; ждём A-запись -> 78.17.3.14)"
echo "Локальная проверка сервера:"
curl -s http://127.0.0.1:8080/health && echo " <- сервер отвечает" || echo "сервер не отвечает локально"
