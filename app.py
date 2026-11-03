from flask import Flask, render_template_string, request, send_from_directory, session, redirect, url_for
from datetime import datetime
from zoneinfo import ZoneInfo
import math
import requests
import os
import pandas as pd
import csv
import openpyxl

app = Flask(__name__)
app.secret_key = "yakkabog_davomat_maxfiy_kalit"

ADMIN_CHAT_IDS = ["1179846016", "1838498079", "1611805500", "727029645"]
EXCEL_HISOBOT = "davomat_hisobot.xlsx"
XODIMLAR_FAYLI = "Umumiy maʼlumot.xlsx"
LOG_FILE = "attendance_log.csv"
UPLOAD_FOLDER = "uploads"

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
UZB_TZ = ZoneInfo("Asia/Tashkent")

TEHNIKUM_LAT = 38.970790
TEHNIKUM_LON = 66.706597
DAVOMAT_MAX_MASOFA = 100

def xodimlarni_yuklash():
    if os.path.exists(XODIMLAR_FAYLI):
        try:
            df = pd.read_excel(XODIMLAR_FAYLI, sheet_name=0, skiprows=2)
            xodimlar_royxati = []
            kerak_emas_sozlar = ["qorovul", "farrosh", "soatbay"]
            
            for index, row in df.iterrows():
                fio = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                lavozim = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else "O'qituvchi"
                tg_id = str(row.iloc[3]).strip() if len(row) > 3 and pd.notna(row.iloc[3]) else ""
                pnfl = str(row.iloc[4]).strip() if len(row) > 4 and pd.notna(row.iloc[4]) else ""
                
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
                            "id": xodim_id, "fame": fio, "lavozim": lavozim,
                            "telegram_id": tg_id, "pnfl": pnfl
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

def bugun_kelganmi(xodim_id):
    bugun = datetime.now(UZB_TZ).strftime("%Y-%m-%d")
    if os.path.exists(EXCEL_HISOBOT):
        df = pd.read_excel(EXCEL_HISOBOT)
        mavjud = df[(df["ID"] == xodim_id) & (df["Sana"] == bugun) & (df["Kelish_Vaqti"] != "-")]
        if not mavjud.empty:
            return True
    return False

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
    df.to_excel(EXCEL_HISOBOT, index=False)

def hisobla_masofa(lat1, lon1, lat2, lon2):
    R = 6371000  
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = math.sin(dlat / 2)**2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlon / 2)**2
    return R * (2 * math.atan2(math.sqrt(a), math.sqrt(1 - a)))

@app.route('/')
def index():
    return render_template_string('''
    <!DOCTYPE html><html lang="uz"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Davomat</title>
    <style>
        body { font-family: 'Segoe UI', sans-serif; background: #0f172a; color: #fff; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; }
        .card { background: #1e293b; padding: 30px; border-radius: 16px; width: 100%; max-width: 400px; text-align: center; box-shadow: 0 10px 25px rgba(0,0,0,0.4); }
        h2 { color: #38bdf8; margin-bottom: 10px; }
        .btn { width: 100%; padding: 15px; border: none; border-radius: 10px; font-size: 16px; font-weight: bold; cursor: pointer; text-decoration: none; color: white; display: block; margin-bottom: 15px; box-sizing: border-box; }
        .btn-kirish { background: #22c55e; }
        .btn-admin { background: #334155; font-size: 14px; padding: 10px; }
    </style></head>
    <body><div class="card">
        <h2>🏢 Davomat Tizimi</h2>
        <p style="color:#94a3b8;margin-bottom:20px;">Amal turini tanlang:</p>
        <a href="/select-employee?action=KIRISH" class="btn btn-kirish">🟢 KELISH (Kirish)</a>
        <a href="/admin-login" class="btn btn-admin">🛡️ Admin Kuzatuv Paneli</a>
    </div></body></html>
    ''')

@app.route('/admin-login', methods=['GET', 'POST'])
def admin_login():
    error = None
    if request.method == 'POST':
        if request.form.get('parol') == "7777":
            session['admin_logged'] = True
            return redirect(url_for('admin_panel'))
        else:
            error = "Parol noto'g'ri! (7777)"
    return render_template_string('''
    <!DOCTYPE html><html lang="uz"><head><meta charset="UTF-8"><title>Admin Kirish</title>
    <style>
        body { font-family: 'Segoe UI', sans-serif; background: #0f172a; color: #fff; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; }
        .card { background: #1e293b; padding: 30px; border-radius: 16px; width: 100%; max-width: 360px; text-align: center; }
        input[type="password"] { width: 100%; padding: 12px; border-radius: 8px; background: #0f172a; color: #fff; border: 1px solid #334155; margin-bottom: 15px; text-align: center; }
        .btn { width: 100%; padding: 12px; border: none; border-radius: 8px; background: #38bdf8; color: #0f172a; font-weight: bold; cursor: pointer; }
    </style></head>
    <body><div class="card">
        <h3>🛡️ Admin Paneli</h3>
        {% if error %}<p style="color:red;">{{ error }}</p>{% endif %}
        <form method="POST"><input type="password" name="parol" placeholder="••••" required autofocus><button type="submit" class="btn">Kirish</button></form>
        <a href="/" style="display:block;margin-top:15px;color:#94a3b8;text-decoration:none;">⟵ Bosh sahifaga</a>
    </div></body></html>
    ''', error=error)

@app.route('/admin')
def admin_panel():
    if not session.get('admin_logged'):
        return redirect(url_for('admin_login'))
    bugun = datetime.now(UZB_TZ).strftime("%Y-%m-%d")
    df = pd.read_excel(EXCEL_HISOBOT) if os.path.exists(EXCEL_HISOBOT) else pd.DataFrame()
    bugun_df = df[df['Sana'] == bugun] if not df.empty else pd.DataFrame()
    return render_template_string('''
    <!DOCTYPE html><html lang="uz"><head><meta charset="UTF-8"><title>Admin</title>
    <style>body{font-family:sans-serif;background:#0f172a;color:#fff;padding:20px;} .container{max-width:800px;margin:auto;background:#1e293b;padding:20px;border-radius:12px;}</style>
    </head><body><div class="container">
    <h2>🛡️ Bugungi Kelganlar ({{ bugun }})</h2><a href="/" style="color:#38bdf8;">Bosh sahifa</a>
    <table width="100%" border="1" style="border-collapse:collapse;margin-top:15px;border-color:#334155;text-align:left;padding:8px;">
    <tr><th>F.I.O</th><th>Lavozim</th><th>Vaqt</th><th>Holat</th></tr>
    {% for _, r in data.iterrows() %}
    <tr><td>{{ r['F.I.O'] }}</td><td>{{ r['Lavozimi'] }}</td><td>{{ r['Kelish_Vaqti'] }}</td><td>{{ r['Oxirgi_Holat'] }}</td></tr>
    {% endfor %}
    </table></div></body></html>
    ''', bugun=bugun, data=bugun_df)

@app.route('/select-employee')
def select_employee():
    action = request.args.get('action', 'KIRISH')
    global XODIMLAR
    XODIMLAR = xodimlarni_yuklash()
    return render_template_string('''
    <!DOCTYPE html><html lang="uz"><head><meta charset="UTF-8"><title>Xodim</title>
    <style>body{font-family:sans-serif;background:#0f172a;color:#fff;display:flex;justify-content:center;align-items:center;height:100vh;margin:0;}
    .card{background:#1e293b;padding:25px;border-radius:16px;width:100%;max-width:400px;text-align:center;}
    input, select {width:100%;padding:12px;border-radius:8px;background:#0f172a;color:#fff;border:1px solid #334155;margin-bottom:15px;box-sizing:border-box;}
    .btn{width:100%;padding:12px;border:none;border-radius:8px;background:#3b82f6;color:white;font-weight:bold;cursor:pointer;}</style></head>
    <body><div class="card">
        <h3>Xodimni tanlang</h3>
        <form action="/verify" method="GET">
            <input type="hidden" name="action" value="{{ action }}">
            <select name="xodim_id" required><option value="" disabled selected>-- Ism familiya --</option>
            {% for x in xodimlar %}<option value="{{ x.id }}">{{ x.fame }} ({{ x.lavozim }})</option>{% endfor %}</select>
            <div style="text-align:left;margin-bottom:15px;"><label style="font-size:13px;color:#38bdf8;">📸 Selfie:</label>
            <input type="file" name="emp_photo" accept="image/*" capture="user" required style="font-size:12px;margin-top:5px;"></div>
            <button type="submit" class="btn">Tasdiqlash</button>
        </form><a href="/" style="display:block;margin-top:15px;color:#94a3b8;text-decoration:none;">⟵ Qaytish</a>
    </div></body></html>
    ''', action=action, xodimlar=XODIMLAR)

@app.route('/verify')
def verify():
    action = request.args.get('action')
    xodim_id = request.args.get('xodim_id')
    xodim = next((x for x in XODIMLAR if str(x['id']) == str(xodim_id)), None)
    if not xodim: return "Xodim topilmadi!"
    return render_template_string('''
    <!DOCTYPE html><html lang="uz"><head><title>GPS</title></head>
    <body style="background:#0f172a;color:#fff;display:flex;justify-content:center;align-items:center;height:100vh;margin:0;font-family:sans-serif;">
        <div style="background:#1e293b;padding:30px;border-radius:16px;text-align:center;">
            <h3>Lokatsiya aniqlanmoqda...</h3>
            <form id="geoForm" action="/result" method="POST"><input type="hidden" name="action" value="{{ action }}"><input type="hidden" name="xodim_id" value="{{ xodim.id }}"><input type="hidden" name="lat" id="lat"><input type="hidden" name="lon" id="lon"></form>
        </div>
        <script>
            navigator.geolocation.getCurrentPosition(pos => {
                document.getElementById('lat').value = pos.coords.latitude;
                document.getElementById('lon').value = pos.coords.longitude;
                document.getElementById('geoForm').submit();
            }, err => { alert("GPS xatosi! Lokatsiyaga ruxsat bering."); window.location.href='/'; }, {enableHighAccuracy: true});
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
        return f"<h2 style='color:red;text-align:center;margin-top:50px;'>❌ Masofa xatosi! Texnikumdan {int(masofa)} metr uzoqdasiz ({DAVOMAT_MAX_MASOFA}m ruxsat etilgan). <a href='/'>Qaytish</a></h2>"

    hozir = datetime.now(UZB_TZ)
    vaqt_str = hozir.strftime("%H:%M:%S")
    is_late = (hozir.hour > 8) or (hozir.hour == 8 and hozir.minute > 30)
    holat_matn = "Kechikdi ⚠️" if is_late else "Hududda"

    davomatni_saqlash(xodim, action, vaqt_str, holat_str)
    return f"<h2 style='color:green;text-align:center;margin-top:50px;'>✅ Muvaffaqiyatli saqlandi! Vaqt: {vaqt_str}. <a href='/'>Bosh sahifa</a></h2>"

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)
