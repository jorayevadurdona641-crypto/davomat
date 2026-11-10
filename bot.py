import os
import asyncio
import csv
from datetime import datetime
from zoneinfo import ZoneInfo
import pandas as pd
import openpyxl
from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command
from aiogram.types import Message, ReplyKeyboardMarkup, KeyboardButton
from aiogram.fsm.state import StatesGroup, State
from aiogram.fsm.context import FSMContext
from flask import Flask
from threading import Thread

# --- SOZLAMALAR ---
TELEGRAM_BOT_TOKEN = "9241880191:AAHfYqQSQC9KOpeAfFdomyft9nhIBm-h6sk"
ADMIN_IDS_INT = [1179846016, 1838498079, 1611805500, 727029645]

EXCEL_HISOBOT = "davomat_hisobot.xlsx"
XODIMLAR_FAYLI = "Umumiy maʼlumot.xlsx"
LOG_FILE = "attendance_log.csv"

UZB_TZ = ZoneInfo("Asia/Tashkent")
TEHNIKUM_LAT = 38.970790
TEHNIKUM_LON = 66.706597
DAVOMAT_MAX_MASOFA = 100

bot = Bot(token=TELEGRAM_BOT_TOKEN)
dp = Dispatcher()

# --- FLASK SERVER (Render port talabini qondirish uchun) ---
app = Flask('')

@app.route('/')
def home():
    return "Bot is active!"

def run():
    app.run(host='0.0.0.0', port=10000)

def keep_alive():
    t = Thread(target=run)
    t.start()


class AttendanceState(StatesGroup):
    waiting_for_jshr = State()
    waiting_for_location = State()
    waiting_for_photo = State()

def hisobla_masofa(lat1, lon1, lat2, lon2):
    import math
    R = 6371000  
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = math.sin(dlat / 2)**2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlon / 2)**2
    return R * (2 * math.atan2(math.sqrt(a), math.sqrt(1 - a)))

def get_teacher_from_excel_bot(jshr: str):
    if not os.path.exists(XODIMLAR_FAYLI):
        return None, None, None
    try:
        df = pd.read_excel(XODIMLAR_FAYLI, sheet_name=0, skiprows=2)
        kerak_emas_sozlar = ["qorovul", "farrosh", "soatbay"]
        
        for index, row in df.iterrows():
            fio = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
            lavozim = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else "O'qituvchi"
            tg_id = str(row.iloc[3]).strip() if len(row) > 3 and pd.notna(row.iloc[3]) else ""
            pnfl = str(row.iloc[4]).strip() if len(row) > 4 and pd.notna(row.iloc[4]) else ""
            
            clean_input_jshr = str(jshr).strip().split('.')[0]
            clean_row_pnfl = str(pnfl).strip().split('.')[0]
            
            if clean_row_pnfl == clean_input_jshr and fio and fio != "nan":
                lavozim_kichik = lavozim.lower()
                for soz in kerak_emas_sozlar:
                    if soz in lavozim_kichik:
                        return "EXCLUDED", "EXCLUDED", None
                return fio, lavozim, tg_id
    except Exception as e:
        print("Excelni botda o'qishda xatolik:", e)
    return None, None, None

def bot_bugun_kelganmi(jshr):
    bugun = datetime.now(UZB_TZ).strftime("%Y-%m-%d")
    if os.path.exists(EXCEL_HISOBOT):
        try:
            df = pd.read_excel(EXCEL_HISOBOT)
            clean_jshr = str(jshr).strip().split('.')[0]
            mavjud = df[(df["ID"].astype(str).str.split('.').str[0] == clean_jshr) & (df["Sana"] == bugun) & (df["Kelish_Vaqti"] != "-")]
            if not mavjud.empty:
                return True
        except Exception as e:
            print("Tekshirishda xatolik:", e)
    return False

def davomatni_saqlash_bot(jshr, teacher_name, position, vaqt_str, holat_turi):
    """
    Kelish va ketish vaqtlarini bitta qatorda (bitta variqda) saqlash funksiyasi.
    """
    bugun = datetime.now(UZB_TZ).strftime("%Y-%m-%d")
    clean_jshr = str(jshr).strip().split('.')[0]
    
    if not os.path.exists(EXCEL_HISOBOT):
        df = pd.DataFrame(columns=["ID", "F.I.O", "Lavozimi", "Sana", "Kelish_Vaqti", "Ketish_Vaqti", "Holat", "Oxirgi_Holat"])
    else:
        df = pd.read_excel(EXCEL_HISOBOT)
    
    if "Oxirgi_Holat" not in df.columns:
        df["Oxirgi_Holat"] = "Hududda"

    mask = (df["ID"].astype(str).str.split('.').str[0] == clean_jshr) & (df["Sana"] == bugun)
    
    if not df[mask].empty:
        # Agar bugungi yozuvi mavjud bo'lsa, demak bu ketish vaqti
        df.loc[mask, "Ketish_Vaqti"] = vaqt_str
        df.loc[mask, "Oxirgi_Holat"] = "Ketdi"
    else:
        # Agar bugun hali ro'yxatga o'tmagan bo'lsa, bu kelish vaqti
        yangi_qator = {
            "ID": jshr, 
            "F.I.O": teacher_name, 
            "Lavozimi": position, 
            "Sana": bugun, 
            "Kelish_Vaqti": vaqt_str, 
            "Ketish_Vaqti": "-", 
            "Holat": holat_turi, 
            "Oxirgi_Holat": "Hududda"
        }
        df = pd.concat([df, pd.DataFrame([yangi_qator])], ignore_index=True)
        
    df.to_excel(EXCEL_HISOBOT, index=False)

@dp.message(Command("start"))
async def cmd_start_bot(message: Message, state: FSMContext):
    await state.set_state(AttendanceState.waiting_for_jshr)
    await message.answer("Assalomu alaykum! Davomat botiga xush kelibsiz. PNFL (JSHR) raqamingizni kiriting:")

@dp.message(AttendanceState.waiting_for_jshr, F.text)
async def process_jshr_bot(message: Message, state: FSMContext):
    jshr = message.text.strip()
    if bot_bugun_kelganmi(jshr):
        await message.answer("⚠️ <b>Siz bugun allaqachon davomatdan o'tgansiz!</b> Qayta o'tish mumkin emas.", parse_mode="HTML")
        await state.clear()
        return

    teacher_name, position, excel_tg_id = get_teacher_from_excel_bot(jshr)
    if teacher_name == "EXCLUDED":
        await message.answer("❌ Kechirasiz, lavozimingiz kiritilmagan.")
        await state.clear()
    elif teacher_name:
        await state.update_data(jshr=jshr, teacher_name=teacher_name, position=position)
        kb = ReplyKeyboardMarkup(keyboard=[[KeyboardButton(text="📍 Geolokatsiyani yuborish", request_location=True)]], resize_keyboard=True, one_time_keyboard=True)
        await state.set_state(AttendanceState.waiting_for_location)
        await message.answer(f"Salom, <b>{teacher_name}</b> ({position})! Lokatsiyangizni yuboring:", reply_markup=kb, parse_mode="HTML")
    else:
        await message.answer("❌ Bunday PNFL (JSHR) bazadan topilmadi. Qaytadan to'g'ri kiriting:")

@dp.message(AttendanceState.waiting_for_location, F.location)
async def process_location_bot(message: Message, state: FSMContext):
    user_lat = message.location.latitude
    user_lon = message.location.longitude
    data = await state.get_data()
    jshr = data.get("jshr")
    
    if bot_bugun_kelganmi(jshr):
        await message.answer("⚠️ <b>Siz bugun allaqachon davomatdan o'tgansiz!</b>", parse_mode="HTML")
        await state.clear()
        return

    distance = hisobla_masofa(user_lat, user_lon, TEHNIKUM_LAT, TEHNIKUM_LON)
    if distance > DAVOMAT_MAX_MASOFA:
        await message.answer(f"❌ Xatolik! Texnikum hududidan uzoqdasiz ({int(distance)}m). Davomat {DAVOMAT_MAX_MASOFA}m ichida bo'lishi kerak.")
        await state.clear()
        return

    await state.update_data(distance=int(distance))
    await state.set_state(AttendanceState.waiting_for_photo)
    await message.answer("Lokatsiya tasdiqlandi ✅ Endi davomat uchun o'z <b>rasmingizni (selfie)</b> yuboring:", parse_mode="HTML", reply_markup=None)

@dp.message(AttendanceState.waiting_for_photo, F.photo)
async def process_photo_bot(message: Message, state: FSMContext):
    data = await state.get_data()
    jshr = data.get("jshr")
    if bot_bugun_kelganmi(jshr):
        await message.answer("⚠️ <b>Siz bugun allaqachon davomatdan o'tgansiz!</b>", parse_mode="HTML")
        await state.clear()
        return

    photo_file_id = message.photo[-1].file_id
    teacher_name = data.get("teacher_name")
    position = data.get("position")
    distance = data.get("distance", 0)
    
    now = datetime.now(UZB_TZ)
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H:%M:%S")
    soat, minut = now.hour, now.minute
    
    is_late = (soat > 8) or (soat == 8 and minut > 30)
    status_text = "Kechikib keldi ⚠️" if is_late else "O'z vaqtida keldi ✅"
    holat_matn = "Kechikdi ⚠️" if is_late else "Hududda"

    davomatni_saqlash_bot(jshr, teacher_name, position, time_str, holat_matn)

    if not os.path.exists(LOG_FILE):
        with open(LOG_FILE, mode="w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["PNFL", "F.I.Sh.", "Lavozim", "Sana", "Vaqt", "Holati", "Oxirgi_Holat"])

    with open(LOG_FILE, mode="a", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow([jshr, teacher_name, position, date_str, time_str, status_text, "Keldi"])

    await state.clear()
    
    caption_text = (
        f"🟢 <b>KIRISH (Kelish) - {date_str}</b>\n\n"
        f"👤 <b>F.I.O:</b> {teacher_name}\n"
        f"📋 <b>Lavozimi:</b> {position}\n"
        f"⏰ <b>Vaqt:</b> {time_str}\n"
        f"Holat: {status_text}\n"
        f"📍 <b>Masofa:</b> {distance}m"
    )
    
    await message.answer(f"Rahmat, {teacher_name}! Kelgan vaqtingiz muvaffaqiyatli qayd etildi ✅")

    target_ids = list(ADMIN_IDS_INT)
    if message.from_user.id not in target_ids:
        target_ids.append(message.from_user.id)

    for admin_id in target_ids:
        try:
            await bot.send_photo(chat_id=admin_id, photo=photo_file_id, caption=caption_text, parse_mode="HTML")
        except Exception as e:
            print("Adminchaga yuborishda xatolik:", e)

async def main():
    await bot.delete_webhook(drop_pending_updates=True)
    print("Bot polling rejimida ishga tushdi...")
    await dp.start_polling(bot)

if __name__ == "__main__":
    keep_alive()  # Flask serverni fon rejimida yoqamiz
    asyncio.run(main())
