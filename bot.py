import os
import requests
import time

TOKEN = "8665708174:AAF9WFb0j0No6XBPXeeZM21vxKgq4Zeux04"
URL = f"https://api.telegram.org/bot{TOKEN}/"

def get_updates(offset=None):
    res = requests.get(URL + "getUpdates", params={"offset": offset, "timeout": 30})
    return res.json().get("result", [])

def send_file(chat_id):
    if os.path.exists("davomat_hisobot.xlsx"):
        with open("davomat_hisobot.xlsx", "rb") as f:
            requests.post(URL + "sendDocument", data={"chat_id": chat_id, "caption": "📊 Bugungi yangilangan davomat hisoboti:"}, files={"document": f})

print("Bot ishga tushdi...")
offset = None
while True:
    updates = get_updates(offset)
    for update in updates:
        offset = update["update_id"] + 1
        message = update.get("message", {})
        text = message.get("text", "")
        chat_id = message.get("chat", {}).get("id")
        
        if text == "/hisobot":
            send_file(chat_id)
    time.sleep(1)