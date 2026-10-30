from flask import Flask, render_template_string, request, send_from_directory
from datetime import datetime
import math
import requests
import os
import pandas as pd

app = Flask(__name__)

# --- TELEGRAM BOT VA ADMINLAR SOZLAMASI ---
TELEGRAM_BOT_TOKEN = "8665708174:AAF9WFb0j0No6XBPXeeZM21vxKgq4Zeux04"
ADMIN_CHAT_IDS = ["1179846016", "727029645", "1838498079", "1611805500"]

EXCEL_HISOBOT = "davomat_hisobot.xlsx"
XODIMLAR_FAYLI = "Umumiy maʼlumot.xlsx"
UPLOAD_FOLDER = "uploads"

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Xodimlarni Exceldan o'qib olish (Qorovul va farroshlar chiqarib tashlanadi)
def xodimlarni_yuklash():
    if os.path.exists(XODIMLAR_FAYLI):
        try:
            df = pd.read_excel(XODIMLAR_FAYLI, sheet_name=0, skiprows=2)
            xodimlar_royxati = []
            kerak_emas_sozlar = ["qorovul", "farrosh"]
            
            for index, row in df.iterrows():
                fio = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                lavozim = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else "O'qituvchi"
                
                if fio and fio != "nan":
                    lavozim_kichik = lavozim.lower()
                    tashlab_yuborish = False
                    for soz in kerak_emas_sozlar:
                        if soz in lavozim_kichik:
                            tashlab_yuborish = True
                            break
                    
                    if not tashlab_yuborish:
                        xodim_id = len(xodimlar_royxati) + 1
                        bugun = datetime.now().strftime("%Y-%m-%d")
                        bugungi_rasm = f"xodim_{xodim_id}_{bugun}.jpg"
                        rasm_path = os.path.join(UPLOAD_FOLDER, bugungi_rasm)
                        
                        photo_url = f"/uploads/{bugungi_rasm}" if os.path.exists(rasm_path) else ""
                        
                        xodimlar_royxati.append({
                            "id": xodim_id,
                            "fame": fio,
                            "lavozim": lavozim,
                            "photo": photo_url
                        })
            return xodimlar_royxati
        except Exception as e:
            print("Excelni o'qishda xatolik:", e)
    return []

XODIMLAR = xodimlarni_yuklash()

def bazani_tekshir():
    if not os.path.exists(EXCEL_HISOBOT):
        df = pd.DataFrame(columns=["ID", "F.I.O", "Lavozimi", "Sana", "Kelish_Vaqti", "Ketish_Vaqti", "Holat"])
        df.to_excel(EXCEL_HISOBOT, index=False)

bazani_tekshir()

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

def davomatni_saqlash(xodim, amal, vaqt):
    bugun = datetime.now().strftime("%Y-%m-%d")
    df = pd.read_excel(EXCEL_HISOBOT)
    mavjud = df[(df["ID"] == xodim["id"]) & (df["Sana"] == bugun)]
    
    if amal == "KIRISH":
        if mavjud.empty:
            yangi_qator = {"ID": xodim["id"], "F.I.O": xodim["fame"], "Lavozimi": xodim["lavozim"], "Sana": bugun, "Kelish_Vaqti": vaqt, "Ketish_Vaqti": "-", "Holat": "Keldi"}
            df = pd.concat([df, pd.DataFrame([yangi_qator])], ignore_index=True)
    elif amal == "CHIQISH":
        if not mavjud.empty:
            df.loc[(df["ID"] == xodim["id"]) & (df["Sana"] == bugun), "Ketish_Vaqti"] = vaqt
            df.loc[(df["ID"] == xodim["id"]) & (df["Sana"] == bugun), "Holat"] = "Ketdi"
        else:
            yangi_qator = {"ID": xodim["id"], "F.I.O": xodim["fame"], "Lavozimi": xodim["lavozim"], "Sana": bugun, "Kelish_Vaqti": "-", "Ketish_Vaqti": vaqt, "Holat": "Faqat Chiqish"}
            df = pd.concat([df, pd.DataFrame([yangi_qator])], ignore_index=True)
            
    df.to_excel(EXCEL_HISOBOT, index=False)

def telegramga_rasmli_xabar_yuborish(photo_path, matn):
    for chat_id in ADMIN_CHAT_IDS:
        try:
            url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendPhoto"
            if os.path.exists(photo_path):
                with open(photo_path, 'rb') as photo:
                    data = {"chat_id": chat_id, "caption": matn, "parse_mode": "HTML"}
                    requests.post(url, data=data, files={"photo": photo})
            else:
                requests.post(f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage", data={"chat_id": chat_id, "text": matn, "parse_mode": "HTML"})
        except Exception as e:
            print("Xatolik:", e)

# Siz yuborgan xaritadagi aniq koordinata
TEHNIKUM_LAT = 38.970790
TEHNIKUM_LON = 66.706597
MAX_MASOFA = 150  # Aniqlik uchun chegara 150 metr qilindi

def hisobla_masofa(lat1, lon1, lat2, lon2):
    R = 6371000  
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = math.sin(dlat / 2)**2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlon / 2)**2
    return R * (2 * math.atan2(math.sqrt(a), math.sqrt(1 - a)))

@app.route('/')
def index():
    return render_template_string('''
    <!DOCTYPE html>
    <html lang="uz">
    <head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Davomat</title>
    <style>
        body { font-family: 'Segoe UI', sans-serif; background: #0f172a; color: #fff; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; }
        .card { background: #1e293b; padding: 30px; border-radius: 16px; width: 100%; max-width: 400px; text-align: center; box-shadow: 0 10px 25px rgba(0,0,0,0.4); }
        h2 { color: #38bdf8; margin-bottom: 10px; }
        .btn { width: 100%; padding: 15px; border: none; border-radius: 10px; font-size: 16px; font-weight: bold; cursor: pointer; text-decoration: none; color: white; display: block; margin-bottom: 15px; }
        .btn-kirish { background: #22c55e; }
        .btn-chiqish { background: #ef4444; }
    </style></head>
    <body><div class="card">
        <h2>🏢 Davomat Tizimi</h2>
        <p style="color:#94a3b8;margin-bottom:20px;">Amal turini tanlang:</p>
        <a href="/select-employee?action=KIRISH" class="btn btn-kirish">🟢 KELISH (Kirish)</a>
        <a href="/select-employee?action=CHIQISH" class="btn btn-chiqish">🔴 KETISH (Chiqish)</a>
    </div></body></html>
    ''')

@app.route('/select-employee')
def select_employee():
    action = request.args.get('action', 'KIRISH')
    global XODIMLAR
    XODIMLAR = xodimlarni_yuklash()
    
    return render_template_string('''
    <!DOCTYPE html><html lang="uz"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Xodim</title>
    <style>
        body { font-family: 'Segoe UI', sans-serif; background: #0f172a; color: #fff; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; }
        .card { background: #1e293b; padding: 25px; border-radius: 16px; width: 100%; max-width: 400px; text-align: center; box-shadow: 0 10px 25px rgba(0,0,0,0.4); }
        select { width: 100%; padding: 12px; border-radius: 8px; background: #0f172a; color: #fff; border: 1px solid #334155; margin-bottom: 20px; font-size:15px; }
        .profile-img { width: 120px; height: 120px; border-radius: 50%; object-fit: cover; border: 4px solid #38bdf8; display: none; margin: 0 auto 15px; }
        .btn { width: 100%; padding: 12px; border: none; border-radius: 8px; background: #3b82f6; color: white; font-size: 16px; font-weight: bold; cursor: pointer; }
        .upload-box { margin-bottom: 15px; display: none; text-align: left; background: #0f172a; padding: 10px; border-radius: 8px; border: 1px dashed #38bdf8; }
    </style></head>
    <body><div class="card">
        <h3 style="color: #38bdf8;">Ismingizni tanlang</h3>
        <p style="color: #94a3b8; font-size: 13px;">Harakat: <b>{{ action }}</b></p>
        <img id="empPhoto" class="profile-img" alt="Rasmi">
        
        <form action="/verify" method="GET" id="mainForm">
            <input type="hidden" name="action" value="{{ action }}">
            <select name="xodim_id" required onchange="checkEmployee(this)">
                <option value="" disabled selected>-- Ism familiyangizni tanlang --</option>
                {% for x in xodimlar %}
                <option value="{{ x.id }}" data-photo="{{ x.photo }}">{{ x.fame }} ({{ x.lavozim }})</option>
                {% endfor %}
            </select>
            
            <div id="uploadContainer" class="upload-box">
                <label style="font-size: 13px; color: #38bdf8; display:block; margin-bottom:5px;">📸 Bugungi rasmingiz topilmadi. Rasm yuklang (yoki rasmga tushing):</label>
                <input type="file" name="emp_photo" id="empPhotoInput" accept="image/*" capture="user" style="font-size: 12px; color: #fff; width:100%;">
            </div>

            <button type="submit" class="btn" id="submitBtn">Davomatni Tasdiqlash</button>
        </form>
        <a href="/" style="display:block;margin-top:15px;color:#94a3b8;text-decoration:none;font-size:14px;">⟵ Bosh sahifaga</a>
    </div>
    <script>
        function checkEmployee(sel) {
            let opt = sel.options[sel.selectedIndex];
            let url = opt.getAttribute('data-photo');
            let img = document.getElementById('empPhoto');
            let uploadBox = document.getElementById('uploadContainer');
            
            if(url && url !== "") {
                img.src = url; 
                img.style.display = 'block';
                uploadBox.style.display = 'none';
                document.getElementById('empPhotoInput').removeAttribute('required');
            } else {
                img.style.display = 'none';
                uploadBox.style.display = 'block';
                document.getElementById('empPhotoInput').setAttribute('required', 'true');
            }
        }
    </script></body></html>
    ''', action=action, xodimlar=XODIMLAR)

@app.route('/verify')
def verify():
    action = request.args.get('action')
    xodim_id = request.args.get('xodim_id')
    xodim = next((x for x in XODIMLAR if str(x['id']) == str(xodim_id)), None)
    if not xodim: return "Xodim topilmadi!"
    return render_template_string('''
    <!DOCTYPE html><html lang="uz"><head><meta charset="UTF-8"><title>GPS</title></head>
    <body style="background:#0f172a;color:#fff;display:flex;justify-content:center;align-items:center;height:100vh;margin:0;font-family:sans-serif;">
        <div style="background:#1e293b;padding:30px;border-radius:16px;text-align:center;">
            <h3>Lokatsiya aniqlanmoqda...</h3><p style="color:#94a3b8;">Iltimos kuting...</p>
            <form id="geoForm" action="/result" method="POST" enctype="multipart/form-data">
                <input type="hidden" name="action" value="{{ action }}">
                <input type="hidden" name="xodim_id" value="{{ xodim.id }}">
                <input type="hidden" name="lat" id="lat"><input type="hidden" name="lon" id="lon">
            </form>
        </div>
        <script>
            navigator.geolocation.getCurrentPosition(pos => {
                document.getElementById('lat').value = pos.coords.latitude;
                document.getElementById('lon').value = pos.coords.longitude;
                document.getElementById('geoForm').submit();
            }, err => { 
                alert("GPS xatosi! Texnikum hududida ekanligingizni tasdiqlash uchun lokatsiyaga ruxsat bering."); 
                window.location.href='/'; 
            }, {enableHighAccuracy:true, timeout: 10000});
        </script>
    </body></html>
    ''', action=action, xodim=xodim)

@app.route('/result', methods=['POST'])
def result():
    action = request.form.get('action')
    xodim_id = request.form.get('xodim_id')
    lat = float(request.form.get('lat', 0))
    lon = float(request.form.get('lon', 0))
    
    xodim = next((x for x in XODIMLAR if str(x['id']) == str(xodim_id)), None)
    if not xodim: return "Xodim topilmadi!"
    
    # Aniq masofani tekshirish (150 metr)
    masofa = hisobla_masofa(lat, lon, TEHNIKUM_LAT, TEHNIKUM_LON)
    if masofa > MAX_MASOFA:
        return f"<h2 style='color:red;text-align:center;margin-top:50px;'>❌ Masofa xatosi! Siz texnikumdan {int(masofa)} metr uzoqdasiz. Davomat faqat texnikum hududida (150m ichida) amalga oshiriladi. <a href='/'>Qaytish</a></h2>"

    bugun = datetime.now().strftime("%Y-%m-%d")
    fayl_nomi = f"xodim_{xodim['id']}_{bugun}.jpg"
    fayl_path = os.path.join(UPLOAD_FOLDER, fayl_nomi)
    
    if not os.path.exists(fayl_path):
        file = request.files.get('emp_photo')
        if file and file.filename != '':
            file.save(fayl_path)
            xodim['photo'] = f"/uploads/{fayl_nomi}"

    hozir = datetime.now()
    soat, minut = hozir.hour, hozir.minute
    vaqt_str = hozir.strftime("%H:%M:%S")
    
    if action == 'KIRISH':
        if (soat > 8) or (soat == 8 and minut > 30):
            return "<h2 style='color:orange;text-align:center;margin-top:50px;'>⚠️ Kechikdingiz! Afsuski, kelish vaqti tugagan. <a href='/'>Qaytish</a></h2>"
        davomatni_saqlash(xodim, "KIRISH", vaqt_str)
        telegramga_rasmli_xabar_yuborish(fayl_path, f"🟢 <b>Kelish (Kirish) - {bugun}</b>\n👤 <b>{xodim['fame']}</b>\n📋 {xodim['lavozim']}\n⏰ Vaqt: {vaqt_str}\n📍 Hududda (Masofa: {int(masofa)}m)")

    elif action == 'CHIQISH':
        if soat < 16:
            return "<h2 style='color:red;text-align:center;margin-top:50px;'>❌ Hali ketish vaqti emas! 16:00 dan oldin chiqib bo'lmaydi. <a href='/'>Qaytish</a></h2>"
        davomatni_saqlash(xodim, "CHIQISH", vaqt_str)
        telegramga_rasmli_xabar_yuborish(fayl_path, f"🔴 <b>Ketish (Chiqish) - {bugun}</b>\n👤 <b>{xodim['fame']}</b>\n📋 {xodim['lavozim']}\n⏰ Vaqt: {vaqt_str}\n📍 Hududda (Masofa: {int(masofa)}m)")

    return render_template_string('''
    <!DOCTYPE html><html lang="uz"><head><meta charset="UTF-8"><title>Natija</title>
    <style>body{font-family:sans-serif;background:#0f172a;color:#fff;display:flex;justify-content:center;align-items:center;height:100vh;margin:0;}
    .card{background:#1e293b;padding:30px;border-radius:16px;text-align:center;width:400px;box-shadow:0 10px 25px rgba(0,0,0,0.4);}
    .profile-img{width:120px;height:120px;border-radius:50%;object-fit:cover;border:4px solid #38bdf8;margin-bottom:15px;}</style></head>
    <body><div class="card">
        <h2 style="color:#22c55e;">Muvaffaqiyatli!</h2>
        <img src="{{ xodim.photo }}" class="profile-img">
        <h3>{{ xodim.fame }}</h3><p style="color:#94a3b8;">{{ xodim.lavozim }}</p>
        <p style="color:#38bdf8;font-weight:bold;">Vaqt: {{ vaqt_str }}</p>
        <p style="color:#22c55e;font-weight:bold;margin:15px 0;">Bugungi davomat muvaffaqiyatli qayd etildi!</p>
        <a href="/" style="display:inline-block;padding:10px 20px;background:#38bdf8;color:#0f172a;font-weight:bold;border-radius:8px;text-decoration:none;">Bosh sahifa</a>
    </div></body></html>
    ''', xodim=xodim, vaqt_str=vaqt_str)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
