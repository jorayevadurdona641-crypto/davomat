import os
import asyncio
import csv
from datetime import datetime
from zoneinfo import ZoneInfo
import pandas as pd
import openpyxl
from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command
from aiogram.types import Message, FSInputFile, ReplyKeyboardMarkup, KeyboardButton
from aiogram.fsm.state import StatesGroup, State
from aiogram.fsm.context import FSMContext

# --- SOZLAMALAR ---
TELEGRAM_BOT_TOKEN = "8418617685:AAEeepVkX_wp3opdVUbW0z99bXxWbJQOfmA"
ADMIN_IDS_INT = [1179846016, 1838498079, 1611805500, 727029645]

EXCEL_HISOBOT = "davomat_hisobot.xlsx"
XODIMLAR_FAYLI = "Umumiy maʼlumot.xlsx"
LOG_FILE = "attendance_log.csv"
EXCEL_REPORT = "xodimlar_hisoboti.xlsx"

UZB_TZ = ZoneInfo("Asia/Tashkent")
TEHNIKUM_LAT = 38.970790
TEHNIKUM_LON = 66.706597
DAVOMAT_MAX_MASOFA = 100

bot = Bot(token=TELEGRAM_BOT_TOKEN)
dp = Dispatcher()

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
        wb = openpyxl.load_workbook(XODIMLAR_FAYLI, data_only=True)
        ws = wb.active
        kerak_emas_sozlar = ["qorovul", "farrosh", "soatbay"]
        for row in ws.iter_rows(min_row=4, values_only=True):
            if row and len(row) >= 5:
                row_name = str(row[1]).strip() if row[1] is not None else ""
                row_position = str(row[2]).strip() if row[2] is not None else ""
                row_tg = str(row[3]).strip() if row[3] is not None else ""
                row_pnfl = str(row[4]).strip() if row[4] is not None else ""
                if row_pnfl == str(jshr).strip():
                    pos_lower = row_position.lower()
                    for soz in kerak_emas_sozlar:
                        if soz in pos_lower:
                            return "EXCLUDED", "EXCLUDED", None
                    return row_name, row_position, row_tg
    except Exception as e:
        print("Excelni botda o'qishda xatolik:", e)
    return None, None, None

def bot_bugun_kelganmi(jshr):
    bugun = datetime.now(UZB_TZ).strftime("%Y-%m-%d")
    if os.path.exists(EXCEL_HISOBOT):
        df = pd.read_excel(EXCEL_HISOBOT)
        for _, row in df.iterrows():
            if str(row.get('Sana')) == bugun and str(row.get('Kelish_Vaqti')) != "-":
                pass
    if os.path.exists(LOG_FILE):
        with open(LOG_FILE, mode="r", encoding="utf-8") as f:
            reader = csv.reader(f)
            next(reader, None)
            for row in reader:
                if len(row) >= 6 and row[0].strip() == str(jshr).strip() and row[3].strip() == bugun and "Keldi" in row[5]:
                    return True
    return False

def davomatni_saqlash_bot(jshr, teacher_name, position, vaqt_str, holat_str):
    bugun = datetime.now(UZB_TZ).strftime("%Y-%m-%d")
    if not os.path.exists(EXCEL_HISOBOT):
        df = pd.DataFrame(columns=["ID", "F.I.O", "Lavozimi", "Sana", "Kelish_Vaqti", "Ketish_Vaqti", "Holat", "Oxirgi_Holat"])
    else:
        df = pd.read_excel(EXCEL_HISOBOT)
    
    if "Oxirgi_Holat" not in df.columns:
        df["Oxirgi_Holat"] = "Hududda"

    yangi_qator = {
        "ID": jshr, "F.I.O": teacher_name, "Lavozimi": position, 
        "Sana": bugun, "Kelish_Vaqti": vaqt_str, "Ketish_Vaqti": "-", 
        "Holat": "Keldi", "Oxirgi_Holat": holat_str
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
        current_user_id = str(message.from_user.id)
        if excel_tg_id and excel_tg_id != "nan" and excel_tg_id != "":
            if str(excel_tg_id) != current_user_id:
                await message.answer("❌ <b>Xatolik!</b> Bu PNFL boshqa Telegram akkauntga biriktirilgan.", parse_mode="HTML")
                await state.clear()
                return
        await state.update_data(jshr=jshr, teacher_name=teacher_name, position=position)
        kb = ReplyKeyboardMarkup(keyboard=[[KeyboardButton(text="📍 Geolokatsiyani yuborish", request_location=True)]], resize_keyboard=True, one_time_keyboard=True)
        await state.set_state(AttendanceState.waiting_for_location)
        await message.answer(f"Salom, <b>{teacher_name}</b> ({position})! Lokatsiyangizni yuboring:", reply_markup=kb, parse_mode="HTML")
    else:
        await message.answer("Bunday PNFL (JSHR) bazadan topilmadi. Qaytadan kiriting:")

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
    asyncio.run(main())
