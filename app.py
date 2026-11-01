from flask import Flask, render_template_string, request, send_from_directory, session, redirect, url_for
from datetime import datetime
from zoneinfo import ZoneInfo
import math
import requests
import os
import pandas as pd

app = Flask(__name__)
app.secret_key = "yakkabog_davomat_maxfiy_kalit"

TELEGRAM_BOT_TOKEN = "8665708174:AAF9WFb0j0No6XBPXeeZM21vxKgq4Zeux04"
ADMIN_CHAT_IDS = ["1179846016", "1838498079", "1611805500", "727029645"]

EXCEL_HISOBOT = "davomat_hisobot.xlsx"
XODIMLAR_FAYLI = "Umumiy maʼlumot.xlsx"
UPLOAD_FOLDER = "uploads"

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
UZB_TZ = ZoneInfo("Asia/Tashkent")

def xodimlarni_yuklash():
    if os.path.exists(XODIMLAR_FAYLI):
        try:
            df = pd.read_excel(XODIMLAR_FAYLI, sheet_name=0, skiprows=2)
            xodimlar_royxati = []
            kerak_emas_sozlar = ["qorovul", "farrosh"]
            
            for index, row in df.iterrows():
                fio = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                lavozim = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else "O'qituvchi"
                tg_id = str(row.iloc[3]).strip() if len(row) > 3 and pd.notna(row.iloc[3]) else ""
                
                if fio and fio != "nan":
                    lavozim_kichik = lavozim.lower()
                    tashlab_yuborish = False
                    for soz in kerak_emas_sozlar:
                        if soz in lavozim_kichik:
                            tashlab_yuborish = True
                            break
                    
                    if not tashlab_yuborish:
                        xodim_id = len(xodimlar_royxati) + 1
                        xodimlar_royxati.append({
                            "id": xodim_id,
                            "fame": fio,
                            "lavozim": lavozim,
                            "telegram_id": tg_id
                        })
            return xodimlar_royxati
        except Exception as e:
            print("Excelni o'qishda xatolik:", e)
    return []

XODIMLAR = xodimlarni_yuklash()

def bazani_tekshir():
    if not os.path.exists(EXCEL_HISOBOT):
        df = pd.DataFrame(columns=["ID", "F.I.O", "Lavozimi", "Sana", "Kelish_Vaqti", "Ketish_Vaqti", "Holat", "Oxirgi_Holat"])
        df.to_excel(EXCEL_HISOBOT, index=False)

bazani_tekshir()

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

def davomatni_saqlash(xodim, amal, vaqt, holat_str="Hududda"):
    bugun = datetime.now(UZB_TZ).strftime("%Y-%m-%d")
    df = pd.read_excel(EXCEL_HISOBOT)
    
    if "Oxirgi_Holat" not in df.columns:
        df["Oxirgi_Holat"] = "Hududda"

    mavjud = df[(df["ID"] == xodim["id"]) & (df["Sana"] == bugun)]
    
    if amal == "KIRISH":
        if mavjud.empty:
            yangi_qator = {"ID": xodim["id"], "F.I.O": xodim["fame"], "Lavozimi": xodim["lavozim"], "Sana": bugun, "Kelish_Vaqti": vaqt, "Ketish_Vaqti": "-", "Holat": "Keldi", "Oxirgi_Holat": holat_str}
            df = pd.concat([df, pd.DataFrame([yangi_qator])], ignore_index=True)
        else:
            df.loc[(df["ID"] == xodim["id"]) & (df["Sana"] == bugun), "Oxirgi_Holat"] = holat_str
    elif amal == "CHIQISH":
        if not mavjud.empty:
            df.loc[(df["ID"] == xodim["id"]) & (df["Sana"] == bugun), "Ketish_Vaqti"] = vaqt
            df.loc[(df["ID"] == xodim["id"]) & (df["Sana"] == bugun), "Holat"] = "Ketdi"
            df.loc[(df["ID"] == xodim["id"]) & (df["Sana"] == bugun), "Oxirgi_Holat"] = "Ketib qoldi"
        else:
            yangi_qator = {"ID": xodim["id"], "F.I.O": xodim["fame"], "Lavozimi": xodim["lavozim"], "Sana": bugun, "Kelish_Vaqti": "-", "Ketish_Vaqti": vaqt, "Holat": "Faqat Chiqish", "Oxirgi_Holat": "Ketib qoldi"}
            df = pd.concat([df, pd.DataFrame([yangi_qator])], ignore_index=True)
            
    df.to_excel(EXCEL_HISOBOT, index=False)

def telegramga_rasmli_xabar_yuborish(photo_path, matn, target_chat_ids):
    for chat_id in target_chat_ids:
        if not chat_id or str(chat_id).strip() == "nan":
            continue
        try:
            url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendPhoto"
            if photo_path and os.path.exists(photo_path):
                with open(photo_path, 'rb') as photo:
                    data = {"chat_id": str(chat_id).strip(), "caption": matn, "parse_mode": "HTML"}
                    requests.post(url, data=data, files={"photo": photo})
            else:
                requests.post(f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage", data={"chat_id": str(chat_id).strip(), "text": matn, "parse_mode": "HTML"})
        except Exception as e:
            print("Xatolik:", e)

def telegramga_xabar_yuborish(chat_id, matn):
    if not chat_id or str(chat_id).strip() == "nan":
        return
    try:
        url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
        requests.post(url, data={"chat_id": str(chat_id).strip(), "text": matn, "parse_mode": "HTML"})
    except Exception as e:
        print("Xatolik:", e)

TEHNIKUM_LAT = 38.970790
TEHNIKUM_LON = 66.706597
DAVOMAT_MAX_MASOFA = 150    # Davomat uchun 150 metr
CHIQISH_XABAR_MASOFA = 400  # 400 metrdan chiqsa ogohlantirish

def hisobla_masofa(lat1, lon1, lat2, lon2):
    R = 6371000  
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = math.sin(dlat / 2)**2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlon / 2)**2
    return R * (2 * math.atan2(math.sqrt(a), math.sqrt(1 - a)))

@app.route('/api/check-location', methods=['POST'])
def check_location():
    data = request.get_json()
    if not data:
        return {"status": "error", "message": "Ma'lumot kelmadi"}, 400
        
    xodim_id = data.get('xodim_id')
    lat = float(data.get('lat', 0))
    lon = float(data.get('lon', 0))
    
    xodim = next((x for x in XODIMLAR if str(x['id']) == str(xodim_id)), None)
    if not xodim:
        return {"status": "error", "message": "Xodim topilmadi"}, 404
        
    masofa = hisobla_masofa(lat, lon, TEHNIKUM_LAT, TEHNIKUM_LON)
    
    if masofa > CHIQISH_XABAR_MASOFA:
        bugun = datetime.now(UZB_TZ).strftime("%Y-%m-%d")
        hozir_vaqt = datetime.now(UZB_TZ).strftime("%H:%M:%S")
        
        davomatni_saqlash(xodim, "KIRISH", hozir_vaqt, holat_str=f"⚠️ 400m dan chiqdi ({int(masofa)}m)")
        
        matn_admin = f"🚨 <b>DIQQAT! XODIM 400 METRDAN UZOQLASHDI!</b>\n\n👤 <b>F.I.O:</b> {xodim['fame']}\n📋 <b>Lavozimi:</b> {xodim['lavozim']}\n📏 Masofa: <b>{int(masofa)} metr</b> uzoqda!\n⏰ Vaqt: {hozir_vaqt}"
        fayl_nomi = f"xodim_{xodim['id']}_{bugun}.jpg"
        fayl_path = os.path.join(UPLOAD_FOLDER, fayl_nomi)
        
        target_ids = list(ADMIN_CHAT_IDS)
        telegramga_rasmli_xabar_yuborish(fayl_path, matn_admin, target_chat_ids=target_ids)
        
        if xodim.get('telegram_id') and str(xodim['telegram_id']) != "nan":
            matn_xodim = f"⚠️ <b>DIQQAT! Siz texnikum hududidan uzoqlashyapsiz!</b>\n\n👤 {xodim['fame']}\n📏 Masofa: <b>{int(masofa)} metr</b>\n⏰ Vaqt: {hozir_vaqt}\n\nIltimos, hududga (400m ichiga) qayting!"
            telegramga_xabar_yuborish(str(xodim['telegram_id']), matn_xodim)
        
        return {"status": "out_of_bounds", "masofa": int(masofa)}
        
    return {"status": "in_bounds", "masofa": int(masofa)}

@app.route('/admin-login', methods=['GET', 'POST'])
def admin_login():
    error = None
    if request.method == 'POST':
        parol = request.form.get('parol')
        if parol == "7777":
            session['admin_logged'] = True
            return redirect(url_for('admin_panel'))
        else:
            error = "Parol noto'g'ri! Qaytadan urinib ko'ring."
            
    return render_template_string('''
    <!DOCTYPE html><html lang="uz"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Admin Kirish</title>
    <style>
        body { font-family: 'Segoe UI', sans-serif; background: #0f172a; color: #fff; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; }
        .card { background: #1e293b; padding: 30px; border-radius: 16px; width: 100%; max-width: 350px; text-align: center; box-shadow: 0 10px 25px rgba(0,0,0,0.4); }
        input[type="password"] { width: 100%; padding: 12px; border-radius: 8px; background: #0f172a; color: #fff; border: 1px solid #334155; margin-bottom: 15px; font-size: 16px; box-sizing: border-box; text-align: center; }
        .btn { width: 100%; padding: 12px; border: none; border-radius: 8px; background: #38bdf8; color: #0f172a; font-size: 16px; font-weight: bold; cursor: pointer; }
    </style></head>
    <body><div class="card">
        <h3 style="color: #38bdf8; margin-top:0;">🛡️ Admin Tasdiqlash</h3>
        <p style="color: #94a3b8; font-size: 13px; margin-bottom: 20px;">Admin panelga kirish uchun maxfiy parolni kiriting:</p>
        {% if error %}<p style="color: #ef4444; font-size: 13px; margin-bottom: 10px;">{{ error }}</p>{% endif %}
        <form method="POST">
            <input type="password" name="parol" placeholder="🔑 Parolni kiriting" required autofocus>
            <button type="submit" class="btn">Kirish</button>
        </form>
        <a href="/" style="display:block; margin-top:20px; color:#94a3b8; text-decoration:none; font-size:14px;">⟵ Bosh sahifaga qaytish</a>
    </div></body></html>
    ''', error=error)

@app.route('/admin-panel')
def admin_panel():
    if not session.get('admin_logged'):
        return redirect(url_for('admin_login'))
        
    bugun = datetime.now(UZB_TZ).strftime("%Y-%m-%d")
    df = pd.read_excel(EXCEL_HISOBOT) if os.path.exists(EXCEL_HISOBOT) else pd.DataFrame()
    bugungi_df = df[df["Sana"] == bugun] if not df.empty else pd.DataFrame()
    
    kelgan_ids = set(bugungi_df["ID"].astype(str)) if not bugungi_df.empty else set()
    kelmaganlar = [x for x in XODIMLAR if str(x['id']) not in kelgan_ids]
    
    return render_template_string('''
    <!DOCTYPE html><html lang="uz"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Admin Paneli</title>
    <style>
        body { font-family: 'Segoe UI', sans-serif; background: #0f172a; color: #fff; margin: 0; padding: 20px; }
        .container { max-width: 950px; margin: 0 auto; background: #1e293b; padding: 25px; border-radius: 16px; box-shadow: 0 10px 25px rgba(0,0,0,0.4); }
        h2 { color: #38bdf8; text-align: center; margin-bottom: 5px; }
        h3 { color: #f43f5e; margin-top: 30px; border-bottom: 1px solid #334155; padding-bottom: 8px; }
        table { width: 100%; border-collapse: collapse; margin-top: 15px; }
        th, td { padding: 12px; text-align: left; border-bottom: 1px solid #334155; font-size: 14px; }
        th { background: #0f172a; color: #38bdf8; }
        .badge-keldi { background: #22c55e; padding: 5px 10px; border-radius: 6px; font-weight: bold; font-size: 12px; }
        .badge-kechikdi { background: #f97316; padding: 5px 10px; border-radius: 6px; font-weight: bold; font-size: 12px; color: #fff; }
        .badge-chiqdi { background: #ef4444; padding: 5px 10px; border-radius: 6px; font-weight: bold; font-size: 12px; }
        .badge-ogoh { background: #f59e0b; padding: 5px 10px; border-radius: 6px; font-weight: bold; font-size: 12px; color: #000; }
        .badge-kelmagan { background: #64748b; padding: 5px 10px; border-radius: 6px; font-weight: bold; font-size: 12px; color: #fff; }
        .logout-btn { background: #ef4444; color: white; padding: 8px 15px; border-radius: 6px; text-decoration: none; font-size: 13px; float: right; font-weight: bold; }
    </style></head>
    <body><div class="container">
        <a href="/admin-logout" class="logout-btn">Chiqish 🔓</a>
        <h2>🛡️ Admin Kuzatuv Paneli</h2>
        <p style="color:#94a3b8; text-align:center;">Sana: <b>{{ bugun }}</b></p>
        
        <h3 style="color:#38bdf8; border-color:#38bdf8;">✅ Bugun Kelgan Xodimlar</h3>
        <table>
            <tr>
                <th>F.I.O</th>
                <th>Lavozimi</th>
                <th>Kelish Vaqti</th>
                <th>Ketish Vaqti</th>
                <th>Joriy Holat</th>
            </tr>
            {% for index, row in data.iterrows() %}
            <tr>
                <td><b>{{ row['F.I.O'] }}</b></td>
                <td>{{ row['Lavozimi'] }}</td>
                <td>{{ row['Kelish_Vaqti'] }}</td>
                <td>{{ row['Ketish_Vaqti'] }}</td>
                <td>
                    {% if '400m' in row['Oxirgi_Holat'] or 'uzoqla' in row['Oxirgi_Holat'] %}
                        <span class="badge-ogoh">{{ row['Oxirgi_Holat'] }}</span>
                    {% elif 'Kechikdi' in row['Oxirgi_Holat'] %}
                        <span class="badge-kechikdi">{{ row['Oxirgi_Holat'] }}</span>
                    {% elif row['Holat'] == 'Keldi' %}
                        <span class="badge-keldi">Hududda</span>
                    {% else %}
                        <span class="badge-chiqdi">{{ row['Holat'] }}</span>
                    {% endif %}
                </td>
            </tr>
            {% else %}
            <tr><td colspan="5" style="text-align:center; color:#94a3b8;">Bugun hali kelganlar yo'q.</td></tr>
            {% endfor %}
        </table>

        <h3>❌ Bugun Ishga Kelmaganlar</h3>
        <table>
            <tr>
                <th>F.I.O</th>
                <th>Lavozimi</th>
                <th>Holati</th>
            </tr>
            {% for x in kelmaganlar %}
            <tr>
                <td><b>{{ x.fame }}</b></td>
                <td>{{ x.lavozim }}</td>
                <td><span class="badge-kelmagan">Kelmagan</span></td>
            </tr>
            {% else %}
            <tr><td colspan="3" style="text-align:center; color:#22c55e;">Ajoyib! Hamma xodimlar kelgan.</td></tr>
            {% endfor %}
        </table>

        <a href="/" style="display:block; text-align:center; margin-top:25px; color:#38bdf8; text-decoration:none;">⟵ Bosh sahifaga qaytish</a>
    </div></body></html>
    ''', bugun=bugun, data=bugungi_df, kelmaganlar=kelmaganlar)

@app.route('/admin-logout')
def admin_logout():
    session.pop('admin_logged', None)
    return redirect(url_for('admin_login'))

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
        .btn { width: 100%; padding: 15px; border: none; border-radius: 10px; font-size: 16px; font-weight: bold; cursor: pointer; text-decoration: none; color: white; display: block; margin-bottom: 15px; box-sizing: border-box; }
        .btn-kirish { background: #22c55e; }
        .btn-chiqish { background: #ef4444; }
        .btn-admin { background: #334155; font-size: 14px; padding: 10px; }
    </style></head>
    <body><div class="card">
        <h2>🏢 Davomat Tizimi</h2>
        <p style="color:#94a3b8;margin-bottom:20px;">Amal turini tanlang:</p>
        <a href="/select-employee?action=KIRISH" class="btn btn-kirish">🟢 KELISH (Kirish)</a>
        <a href="/select-employee?action=CHIQISH" class="btn btn-chiqish">🔴 KETISH (Chiqish)</a>
        <a href="/admin-login" class="btn btn-admin">🛡️ Admin Kuzatuv Paneli</a>
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
        .card { background: #1e293b; padding: 25px; border-radius: 16px; width: 100%; max-width: 400px; text-align: center; box-sizing: border-box; box-shadow: 0 10px 25px rgba(0,0,0,0.4); }
        input[type="text"], select { width: 100%; padding: 12px; border-radius: 8px; background: #0f172a; color: #fff; border: 1px solid #334155; margin-bottom: 15px; font-size:15px; box-sizing: border-box; }
        .btn { width: 100%; padding: 12px; border: none; border-radius: 8px; background: #3b82f6; color: white; font-size: 16px; font-weight: bold; cursor: pointer; box-sizing: border-box; }
        .upload-box { margin-bottom: 15px; text-align: left; background: #0f172a; padding: 10px; border-radius: 8px; border: 1px dashed #38bdf8; }
    </style></head>
    <body><div class="card">
        <h3 style="color: #38bdf8; margin-top:0;">Xodimni qidirish va tanlash</h3>
        <p style="color: #94a3b8; font-size: 13px; margin-bottom:15px;">Harakat: <b>{{ action }}</b></p>
        
        <form action="/verify" method="GET" id="mainForm">
            <input type="hidden" name="action" value="{{ action }}">
            
            <input type="text" id="searchInput" placeholder="🔍 Ism yoki familiyani yozib qidiring..." onkeyup="filterEmployees()">
            
            <select name="xodim_id" id="employeeSelect" required>
                <option value="" disabled selected>-- Ism familiyani tanlang --</option>
                {% for x in xodimlar %}
                <option value="{{ x.id }}" data-name="{{ x.fame | lower }}">{{ x.fame }} ({{ x.lavozim }})</option>
                {% endfor %}
            </select>
            
            <div class="upload-box">
                <label style="font-size: 13px; color: #38bdf8; display:block; margin-bottom:5px;">📸 Jonli rasmga tushish (Selfie):</label>
                <input type="file" name="emp_photo" id="empPhotoInput" accept="image/*" capture="user" required style="font-size: 12px; color: #fff; width:100%;">
            </div>

            <button type="submit" class="btn" id="submitBtn">Davomatni Tasdiqlash</button>
        </form>
        <a href="/" style="display:block;margin-top:15px;color:#94a3b8;text-decoration:none;font-size:14px;">⟵ Bosh sahifaga</a>
    </div>
    <script>
        function filterEmployees() {
            let input = document.getElementById('searchInput').value.toLowerCase();
            let select = document.getElementById('employeeSelect');
            let options = select.options;
            
            for (let i = 1; i < options.length; i++) {
                let name = options[i].getAttribute('data-name');
                if (name.includes(input)) {
                    options[i].style.display = "";
                } else {
                    options[i].style.display = "none";
                }
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
    
    masofa = hisobla_masofa(lat, lon, TEHNIKUM_LAT, TEHNIKUM_LON)
    if masofa > DAVOMAT_MAX_MASOFA:
        return f"<h2 style='color:red;text-align:center;margin-top:50px;'>❌ Masofa xatosi! Siz texnikumdan {int(masofa)} metr uzoqdasiz. Davomat faqat texnikum hududida (150m ichida) amalga oshiriladi. <a href='/'>Qaytish</a></h2>"

    bugun = datetime.now(UZB_TZ).strftime("%Y-%m-%d")
    fayl_nomi = f"xodim_{xodim['id']}_{bugun}_{datetime.now(UZB_TZ).strftime('%H%M%S')}.jpg"
    fayl_path = os.path.join(UPLOAD_FOLDER, fayl_nomi)
    
    file = request.files.get('emp_photo')
    if file and file.filename != '':
        file.save(fayl_path)
        xodim['photo_url'] = f"/uploads/{fayl_nomi}"
    else:
        xodim['photo_url'] = ""

    hozir = datetime.now(UZB_TZ)
    soat, minut = hozir.hour, hozir.minute
    hafta_kuni = hozir.weekday() 
    vaqt_str = hozir.strftime("%H:%M:%S")
    
    target_chat_ids = list(ADMIN_CHAT_IDS)
    if xodim.get('telegram_id') and str(xodim['telegram_id']).strip() != "nan" and str(xodim['telegram_id']).strip() != "":
        target_chat_ids.append(str(xodim['telegram_id']).strip())

    if action == 'KIRISH':
        is_late = (soat > 8) or (soat == 8 and minut > 35)
        holat_matn = "Kechikdi ⚠️" if is_late else "Hududda"
        kechikish_statusi = "<b>Holat:</b> Kechikib keldi ⚠️" if is_late else "<b>Holat:</b> O'z vaqtida keldi ✅"
        
        davomatni_saqlash(xodim, "KIRISH", vaqt_str, holat_str=holat_matn)
        
        matn_xabar = f"🟢 <b>KIRISH (Kelish) - {bugun}</b>\n\n👤 <b>F.I.O:</b> {xodim['fame']}\n📋 <b>Lavozimi:</b> {xodim['lavozim']}\n⏰ <b>Vaqt:</b> {vaqt_str}\n{kechikish_statusi}\n📍 <b>Masofa:</b> {int(masofa)}m"
        telegramga_rasmli_xabar_yuborish(fayl_path, matn_xabar, target_chat_ids=target_chat_ids)

    elif action == 'CHIQISH':
        if hafta_kuni == 5:
            if soat < 13:
                return "<h2 style='color:red;text-align:center;margin-top:50px;'>❌ Shanba kuni 13:00 dan oldin chiqib bo'lmaydi! <a href='/'>Qaytish</a></h2>"
        else:
            if (soat < 15) or (soat == 15 and minut < 50):
                return "<h2 style='color:red;text-align:center;margin-top:50px;'>❌ Hali ketish vaqti emas! 15:50 dan oldin chiqib bo'lmaydi. <a href='/'>Qaytish</a></h2>"
                
        davomatni_saqlash(xodim, "CHIQISH", vaqt_str, holat_str="Ketib qoldi")
        matn_xabar = f"🔴 <b>CHIQISH (Ketish) - {bugun}</b>\n\n👤 <b>F.I.O:</b> {xodim['fame']}\n📋 <b>Lavozimi:</b> {xodim['lavozim']}\n⏰ <b>Vaqt:</b> {vaqt_str}\n<b>Holat:</b> Ishdan ketdi 🔴\n📍 <b>Masofa:</b> {int(masofa)}m"
        telegramga_rasmli_xabar_yuborish(fayl_path, matn_xabar, target_chat_ids=target_chat_ids)

    return render_template_string('''
    <!DOCTYPE html><html lang="uz"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Natija</title>
    <style>
        body{font-family:'Segoe UI',sans-serif;background:#0f172a;color:#fff;display:flex;justify-content:center;align-items:center;min-height:100vh;margin:0;}
        .card{background:#1e293b;padding:40px 30px;border-radius:24px;text-align:center;width:100%;max-width:460px;box-sizing:border-box;box-shadow: 0 20px 40px rgba(0,0,0,0.6);}
        .profile-img{width:160px;height:160px;border-radius:50%;object-fit:cover;border:5px solid #38bdf8;margin-bottom:20px;}
        h2{color:#22c55e;font-size:30px;margin-top:0;margin-bottom:15px;}
        h3{font-size:24px;margin:10px 0 5px 0;color:#f8fafc;}
        p{font-size:18px;color:#94a3b8;margin:5px 0;}
        .time-box{background:#0f172a;padding:16px;border-radius:12px;margin:25px 0;font-size:22px;color:#38bdf8;font-weight:bold;border:2px solid #334155;}
        .btn{display:block;width:100%;padding:16px;background:#38bdf8;color:#0f172a;font-weight:bold;font-size:18px;border-radius:12px;text-decoration:none;box-sizing:border-box;}
    </style></head>
    <body><div class="card">
        <h2>✅ Tasdiqlandi!</h2>
        {% if xodim.photo_url %}
        <img src="{{ xodim.photo_url }}" class="profile-img" alt="Xodim rasmi">
        {% endif %}
        <h3>{{ xodim.fame }}</h3>
        <p>{{ xodim.lavozim }}</p>
        
        <div class="time-box">
            ⏰ Vaqt: {{ vaqt_str }}
        </div>
        
        <p style="color:#22c55e;font-weight:bold;font-size:18px;margin-bottom:30px;">Davomatingiz muvaffaqiyatli qayd etildi va botga yuborildi!</p>
        
        <a href="/" class="btn">Bosh sahifaga qaytish</a>
    </div>

    <script>
        const xodimId = "{{ xodim.id }}";
        if (navigator.geolocation) {
            setInterval(() => {
                navigator.geolocation.getCurrentPosition(pos => {
                    fetch('/api/check-location', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            xodim_id: xodimId,
                            lat: pos.coords.latitude,
                            lon: pos.coords.longitude
                        })
                    }).then(res => res.json()).then(data => {
                        if (data.status === 'out_of_bounds') {
                            console.log("Diqqat! 400m dan chiqib ketildi: " + data.masofa + "m");
                        }
                    });
                }, err => { console.log("GPS kuzatuv xatosi"); }, { enableHighAccuracy: true });
            }, 30000);
        }
    </script>
    </body></html>
    ''', xodim=xodim, vaqt_str=vaqt_str)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
