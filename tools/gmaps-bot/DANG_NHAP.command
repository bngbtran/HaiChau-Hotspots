#!/bin/bash
cd "$(dirname "$0")"
command -v node >/dev/null || { echo "Chưa cài Node.js. Tải bản LTS tại https://nodejs.org rồi chạy lại."; read -p "Nhấn Enter để đóng"; exit 1; }
[ -d node_modules ] || { echo "Đang cài thư viện lần đầu, chờ 1-2 phút..."; npm install --silent; }
node bot.js login
read -p "Xong. Nhấn Enter để đóng"
