from flask import Flask, render_template_string, request, send_from_directory, session, redirect, url_for, send_file
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import math
import requests
import os
import pandas as pd
import csv
import openpyxl

app = Flask(__name__)
app.secret_key = "yakkabog_davomat_maxfiy_kalit"

# --- TELEGRAM BOT TOKENI ---
TELEGRAM_BOT_TOKEN = "9241880191:AAEGdwBHjqdBmqLv4hOrL8bIvJl2bRAV8QE"
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
OGOHLANTIRISH_MASOFA = 500   

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
    if not os.path.exists(LOG_FILE):
        with open(LOG_FILE, mode="w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["PNFL", "F.I.Sh.", "Lavozim", "Sana", "Vaqt", "Holati", "Oxirgi_Holat"])

bazani_tekshir()

def bugun_kelganmi(xodim_id):
    bugun = datetime.now(UZB_TZ).strftime("%Y-%m-%d")
    if os.path.exists(EXCEL_HISOBOT):
        try:
            df = pd.read_excel(EXCEL_HISOBOT)
            mavjud = df[(df["ID"].astype(str) == str(xodim_id)) & (df["Sana"] == bugun) & (df["Kelish_Vaqti"] != "-")]
            if not mavjud.empty:
                return True
        except Exception as e:
            print("Tekshirishda xatolik:", e)
    return False

def bugun_chiqqanmi(xodim_id):
    bugun = datetime.now(UZB_TZ).strftime("%Y-%m-%d")
    if os.path.exists(EXCEL_HISOBOT):
        try:
            df = pd.read_excel(EXCEL_HISOBOT)
            mavjud = df[(df["ID"].astype(str) == str(xodim_id)) & (df["Sana"] == bugun) & (df["Ketish_Vaqti"] != "-")]
            if not mavjud.empty:
                return True
        except Exception as e:
            print("Tekshirishda xatolik:", e)
    return False

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

def davomatni_saqlash(xodim, amal, vaqt, holat_str="Hududda"):
    bugun = datetime.now(UZB_TZ).strftime("%Y-%m-%d")
    df = pd.read_excel(EXCEL_HISOBOT)
    if "Oxirgi_Holat" not in df.columns:
        df["Oxirgi_Holat"] = "Hududda"
    mavjud = df[(df["ID"].astype(str) == str(xodim["id"])) & (df["Sana"] == bugun)]
    
    if amal == "KIRISH":
        if mavjud.empty:
            yangi_qator = {"ID": xodim["id"], "F.I.O": xodim["fame"], "Lavozimi": xodim["lavozim"], "Sana": bugun, "Kelish_Vaqti": vaqt, "Ketish_Vaqti": "-", "Holat": "Keldi", "Oxirgi_Holat": holat_str}
            df = pd.concat([df, pd.DataFrame([yangi_qator])], ignore_index=True)
        else:
            df.loc[(df["ID"].astype(str) == str(xodim["id"])) & (df["Sana"] == bugun), "Kelish_Vaqti"] = vaqt
            df.loc[(df["ID"].astype(str) == str(xodim["id"])) & (df["Sana"] == bugun), "Holat"] = "Keldi"
            df.loc[(df["ID"].astype(str) == str(xodim["id"])) & (df["Sana"] == bugun), "Oxirgi_Holat"] = holat_str
    elif amal == "CHIQISH":
        if not mavjud.empty:
            df.loc[(df["ID"].astype(str) == str(xodim["id"])) & (df["Sana"] == bugun), "Ketish_Vaqti"] = vaqt
            df.loc[(df["ID"].astype(str) == str(xodim["id"])) & (df["Sana"] == bugun), "Holat"] = "Ketdi"
            df.loc[(df["ID"].astype(str) == str(xodim["id"])) & (df["Sana"] == bugun), "Oxirgi_Holat"] = "Chiqib ketdi ❌"
        else:
            yangi_qator = {"ID": xodim["id"], "F.I.O": xodim["fame"], "Lavozimi": xodim["lavozim"], "Sana": bugun, "Kelish_Vaqti": "-", "Ketish_Vaqti": vaqt, "Holat": "Faqat Chiqish", "Oxirgi_Holat": "Chiqib ketdi ❌"}
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
            print("Telegramga yuborishda xatolik:", e)

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
        body { font-family: 'Segoe UI', sans-serif; background: #0f172a; color: #fff; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; }
        .card { background: #1e293b; padding: 40px 30px; border-radius: 20px; width: 100%; max-width: 440px; text-align: center; box-shadow: 0 15px 35px rgba(0,0,0,0.5); box-sizing: border-box; }
        h2 { color: #38bdf8; font-size: 26px; margin-bottom: 10px; }
        p { color: #94a3b8; font-size: 16px; margin-bottom: 25px; }
        .btn { width: 100%; padding: 18px; border: none; border-radius: 14px; font-size: 18px; font-weight: bold; cursor: pointer; text-decoration: none; color: white; display: block; margin-bottom: 18px; box-sizing: border-box; }
        .btn-kirish { background: #22c55e; font-size: 20px; }
        .btn-chiqish { background: #ef4444; font-size: 20px; }
        .btn-admin { background: #334155; font-size: 15px; padding: 14px; color: #cbd5e1; }
    </style></head>
    <body><div class="card">
        <h2>🏢 Davomat Tizimi</h2>
        <p>Amal turini tanlang:</p>
        <a href="/select-employee?action=KIRISH" class="btn btn-kirish">🟢 KELISH (Kirish)</a>
        <a href="/select-employee?action=CHIQISH" class="btn btn-chiqish">🔴 KETISH (Chiqish)</a>
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
    <!DOCTYPE html><html lang="uz"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Admin Kirish</title>
    <style>
        body { font-family: 'Segoe UI', sans-serif; background: #0f172a; color: #fff; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; }
        .card { background: #1e293b; padding: 40px 30px; border-radius: 20px; width: 100%; max-width: 380px; text-align: center; box-sizing: border-box; }
        input[type="password"] { width: 100%; padding: 15px; border-radius: 12px; background: #0f172a; color: #fff; border: 1px solid #334155; margin-bottom: 20px; font-size: 18px; text-align: center; box-sizing: border-box; }
        .btn { width: 100%; padding: 15px; border: none; border-radius: 12px; background: #38bdf8; color: #0f172a; font-size: 18px; font-weight: bold; cursor: pointer; }
    </style></head>
    <body><div class="card">
        <h3 style="color: #38bdf8; font-size: 22px; margin-top:0;">🛡️ Admin Paneli</h3>
        {% if error %}<p style="color:red; font-size:15px;">{{ error }}</p>{% endif %}
        <form method="POST"><input type="password" name="parol" placeholder="••••" required autofocus><button type="submit" class="btn">Kirish</button></form>
        <a href="/" style="display:block;margin-top:20px;color:#94a3b8;text-decoration:none;font-size:15px;">⟵ Bosh sahifaga</a>
    </div></body></html>
    ''', error=error)

@app.route('/admin')
def admin_panel():
    if not session.get('admin_logged'):
        return redirect(url_for('admin_login'))
    
    sana_filter = request.args.get('sana', datetime.now(UZB_TZ).strftime("%Y-%m-%d"))
    df = pd.read_excel(EXCEL_HISOBOT) if os.path.exists(EXCEL_HISOBOT) else pd.DataFrame()
    
    if not df.empty and 'Sana' in df.columns:
        sana_df = df[df['Sana'].astype(str) == str(sana_filter)]
    else:
        sana_df = pd.DataFrame()

    kelgan_ids = set(sana_df["ID"].astype(str)) if not sana_df.empty else set()
    kelmaganlar = [x for x in XODIMLAR if str(x['id']) not in kelgan_ids]

    return render_template_string('''
    <!DOCTYPE html><html lang="uz"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Admin Paneli</title>
    <style>
        body { font-family: 'Segoe UI', sans-serif; background: #0f172a; color: #fff; padding: 20px; }
        .container { max-width: 1000px; margin: auto; background: #1e293b; padding: 30px; border-radius: 20px; box-shadow: 0 15px 35px rgba(0,0,0,0.5); }
        h2 { color: #38bdf8; text-align: center; font-size: 24px; }
        table { width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 15px; }
        th, td { padding: 14px; text-align: left; border-bottom: 1px solid #334155; }
        th { background: #0f172a; color: #38bdf8; }
        .badge-keldi { background: #22c55e; padding: 6px 12px; border-radius: 6px; font-weight: bold; color: #fff; }
        .badge-kechikdi { background: #f97316; padding: 6px 12px; border-radius: 6px; font-weight: bold; color: #fff; }
        .badge-chiqdi { background: #ef4444; padding: 6px 12px; border-radius: 6px; font-weight: bold; color: #fff; }
        .badge-kelmagan { background: #64748b; padding: 6px 12px; border-radius: 6px; font-weight: bold; color: #fff; }
        .filter-box { background: #0f172a; padding: 15px; border-radius: 12px; margin-bottom: 20px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px; }
        input[type="date"] { padding: 10px; border-radius: 8px; background: #1e293b; color: #fff; border: 1px solid #334155; font-size: 16px; }
        .btn-download { background: #3b82f6; color: white; padding: 10px 20px; border-radius: 8px; text-decoration: none; font-weight: bold; font-size: 15px; }
        .btn-filter { background: #22c55e; color: white; padding: 10px 20px; border-radius: 8px; border: none; font-weight: bold; cursor: pointer; font-size: 15px; }
    </style></head>
    <body><div class="container">
        <h2>🛡️ Xodimlar Davomati Kuzatuv Paneli</h2>
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:20px;">
            <a href="/" style="color:#38bdf8; text-decoration:none; font-size:16px; font-weight:bold;">⟵ Bosh sahifa</a>
            <a href="/download-excel" class="btn-download">📥 To'liq Excel Hisobotni Yuklab Olish</a>
        </div>

        <form method="GET" class="filter-box">
            <div>
                <label style="margin-right: 10px; font-weight: bold; color: #38bdf8;">📅 Sanani tanlang:</label>
                <input type="date" name="sana" value="{{ sana_filter }}">
            </div>
            <button type="submit" class="btn-filter">Saralash / Ko'rish</button>
        </form>

        <h3 style="color:#38bdf8; margin-top:25px; font-size:18px;">✅ Kelganlar ({{ sana_filter }})</h3>
        <table>
            <tr><th>F.I.O</th><th>Lavozimi</th><th>Kelish Vaqti</th><th>Ketish Vaqti</th><th>Holati</th></tr>
            {% for _, r in data.iterrows() %}
            <tr>
                <td><b>{{ r['F.I.O'] }}</b></td>
                <td>{{ r['Lavozimi'] }}</td>
                <td>{{ r['Kelish_Vaqti'] }}</td>
                <td>{{ r['Ketish_Vaqti'] }}</td>
                <td>
                    {% if 'Chiqib' in r['Oxirgi_Holat'] %}<span class="badge-chiqdi">{{ r['Oxirgi_Holat'] }}</span>
                    {% elif 'Kechikdi' in r['Oxirgi_Holat'] %}<span class="badge-kechikdi">{{ r['Oxirgi_Holat'] }}</span>
                    {% else %}<span class="badge-keldi">O'z vaqtida ✅</span>{% endif %}
                </td>
            </tr>
            {% else %}
            <tr><td colspan="5" style="text-align:center; color:#94a3b8;">Bu kunda kelganlar yo'q.</td></tr>
            {% endfor %}
        </table>

        <h3 style="color:#f43f5e; margin-top:35px; font-size:18px;">❌ Kelmaganlar ({{ sana_filter }})</h3>
        <table>
            <tr><th>F.I.O</th><th>Lavozimi</th><th>Holati</th></tr>
            {% for x in kelmaganlar %}
            <tr><td><b>{{ x.fame }}</b></td><td>{{ x.lavozim }}</td><td><span class="badge-kelmagan">Kelmagan</span></td></tr>
            {% else %}
            <tr><td colspan="3" style="text-align:center; color:#22c55e;">Bu kunda hamma kelgan!</td></tr>
            {% endfor %}
        </table>
    </div></body></html>
    ''', sana_filter=sana_filter, data=sana_df, kelmaganlar=kelmaganlar)

@app.route('/download-excel')
def download_excel():
    if not session.get('admin_logged'):
        return redirect(url_for('admin_login'))
    if os.path.exists(EXCEL_HISOBOT):
        return send_file(EXCEL_HISOBOT, as_attachment=True)
    return "Hali hisobot fayli yaratilmagan!"

@app.route('/select-employee')
def select_employee():
    action = request.args.get('action', 'KIRISH')
    global XODIMLAR
    XODIMLAR = xodimlarni_yuklash()
    return render_template_string('''
    <!DOCTYPE html><html lang="uz"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Xodim</title>
    <style>
        body { font-family: 'Segoe UI', sans-serif; background: #0f172a; color: #fff; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; }
        .card { background: #1e293b; padding: 30px; border-radius: 20px; width: 100%; max-width: 420px; text-align: center; box-sizing: border-box; }
        input[type="text"], select { width: 100%; padding: 15px; border-radius: 12px; background: #0f172a; color: #fff; border: 1px solid #334155; margin-bottom: 20px; font-size: 16px; box-sizing: border-box; }
        .btn { width: 100%; padding: 16px; border: none; border-radius: 12px; background: #3b82f6; color: white; font-size: 18px; font-weight: bold; cursor: pointer; margin-top: 10px; }
        video, canvas { width: 100%; border-radius: 12px; margin-bottom: 15px; background: #000; }
    </style></head>
    <body><div class="card">
        <h3 style="color:#38bdf8; font-size: 22px; margin-top:0;">Xodimni tanlash va Selfie</h3>
        <p style="color:#94a3b8; font-size: 15px;">Harakat: <b>{{ action }}</b></p>
        
        <input type="text" id="searchInput" placeholder="🔍 Ism yoki familiyani yozib qidiring..." onkeyup="filterEmployees()">
        <select id="employeeSelect" required><option value="" disabled selected>-- Ism familiyani tanlang --</option>
        {% for x in xodimlar %}<option value="{{ x.id }}" data-name="{{ x.fame | lower }}">{{ x.fame }} ({{ x.lavozim }})</option>{% endfor %}</select>

        <div style="text-align:left; margin-bottom:10px;">
            <label style="font-size:15px;color:#38bdf8;display:block;margin-bottom:8px;">📸 Jonli kamera (Selfie):</label>
            <video id="video" autoplay playsinline></video>
            <canvas id="canvas" style="display:none;"></canvas>
            <button type="button" id="snap" class="btn" style="background:#22c55e;">📸 Suratga olish</button>
        </div>

        <button type="button" id="submitBtn" class="btn" style="background:#64748b; cursor:not-allowed;" disabled>Tasdiqlash va yuborish</button>
        <a href="/" style="display:block;margin-top:20px;color:#94a3b8;text-decoration:none;font-size:15px;">⟵ Qaytish</a>
    </div>
    
    <script>
        let video = document.getElementById('video');
        let canvas = document.getElementById('canvas');
        let snap = document.getElementById('snap');
        let submitBtn = document.getElementById('submitBtn');
        let capturedBlob = null;

        navigator.mediaDevices.getUserMedia({ video: { facingMode: "user" }, audio: false })
            .then(stream => { video.srcObject = stream; })
            .catch(err => { alert("Kameraga ruxsat berilmadi yoki kamera topilmadi!"); });

        snap.addEventListener('click', function() {
            canvas.width = video.videoWidth;
            canvas.height = video.videoHeight;
            let ctx = canvas.getContext('2d');
            ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
            
            canvas.toBlob(blob => {
                capturedBlob = blob;
                alert("Surat muvaffaqiyatli olindi! Endi 'Tasdiqlash va yuborish' tugmasini bosing.");
                submitBtn.style.background = "#3b82f6";
                submitBtn.style.cursor = "pointer";
                submitBtn.removeAttribute('disabled');
            }, 'image/jpeg');
        });

        submitBtn.addEventListener('click', function() {
            let select = document.getElementById('employeeSelect');
            if (!select.value) {
                alert("Iltimos, avval ism-familiyangizni tanlang!");
                return;
            }
            if (!capturedBlob) {
                alert("Iltimos, avval kameraga qarab 'Suratga olish' tugmasini bosing!");
                return;
            }

            navigator.geolocation.getCurrentPosition(pos => {
                let lat = pos.coords.latitude;
                let lon = pos.coords.longitude;
                let actionVal = "{{ action }}";
                let xodimIdVal = select.value;

                let formData = new FormData();
                formData.append('action', actionVal);
                formData.append('xodim_id', xodimIdVal);
                formData.append('lat', lat);
                formData.append('lon', lon);
                formData.append('emp_photo', capturedBlob, 'selfie.jpg');

                submitBtn.innerText = "Yuborilmoqda...";
                submitBtn.disabled = true;

                fetch('/result', {
                    method: 'POST',
                    body: formData
                })
                .then(res => res.text())
                .then(html => {
                    document.open();
                    document.write(html);
                    document.close();
                })
                .catch(err => {
                    alert("Xatolik yuz berdi: " + err);
                    submitBtn.innerText = "Tasdiqlash va yuborish";
                    submitBtn.disabled = false;
                });
            }, err => { alert("GPS xatosi! Lokatsiyaga ruxsat bering."); }, {enableHighAccuracy: true});
        });

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
    return redirect(url_for('index'))

@app.route('/result', methods=['POST'])
def result():
    action = request.form.get('action')
    xodim_id = request.form.get('xodim_id')
    lat = float(request.form.get('lat', 0))
    lon = float(request.form.get('lon', 0))
    xodim = next((x for x in XODIMLAR if str(x['id']) == str(xodim_id)), None)
    if not xodim: return "Xodim topilmadi!"

    # Takroriy kirishni tekshirish
    if action == 'KIRISH' and bugun_kelganmi(xodim['id']):
        return f"""
        <!DOCTYPE html><html lang="uz"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Xatolik</title>
        <style>body{{font-family:'Segoe UI',sans-serif;background:#0f172a;color:#fff;display:flex;justify-content:center;align-items:center;min-height:100vh;margin:0;}}
        .card{{background:#1e293b;padding:40px 30px;border-radius:20px;text-align:center;max-width:420px;width:100%;box-shadow:0 15px 35px rgba(0,0,0,0.5);box-sizing:border-box;}}
        h2{{color:#ef4444;font-size:24px;margin-bottom:15px;}}p{{color:#94a3b8;font-size:16px;margin-bottom:25px;}}
        .btn{{display:block;width:100%;padding:16px;background:#38bdf8;color:#0f172a;font-weight:bold;font-size:18px;border-radius:12px;text-decoration:none;box-sizing:border-box;}}</style></head>
        <body><div class="card">
            <h2>⚠️ Allaqachon o'tgansiz!</h2>
            <p><b>{xodim['fame']}</b>, siz bugun allaqachon kelish davomatidan o'tgansiz. Qayta o'tish mumkin emas!</p>
            <a href="/" class="btn">Bosh sahifaga qaytish</a>
        </div></body></html>
        """

    # Takroriy chiqishni tekshirish
    if action == 'CHIQISH' and bugun_chiqqanmi(xodim['id']):
        return f"""
        <!DOCTYPE html><html lang="uz"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Xatolik</title>
        <style>body{{font-family:'Segoe UI',sans-serif;background:#0f172a;color:#fff;display:flex;justify-content:center;align-items:center;min-height:100vh;margin:0;}}
        .card{{background:#1e293b;padding:40px 30px;border-radius:20px;text-align:center;max-width:420px;width:100%;box-shadow:0 15px 35px rgba(0,0,0,0.5);box-sizing:border-box;}}
        h2{{color:#ef4444;font-size:24px;margin-bottom:15px;}}p{{color:#94a3b8;font-size:16px;margin-bottom:25px;}}
        .btn{{display:block;width:100%;padding:16px;background:#38bdf8;color:#0f172a;font-weight:bold;font-size:18px;border-radius:12px;text-decoration:none;box-sizing:border-box;}}</style></head>
        <body><div class="card">
            <h2>⚠️ Allaqachon chiqib ketgansiz!</h2>
            <p><b>{xodim['fame']}</b>, siz bugun allaqachon ishdan ketish davomatidan o'tgansiz. Ikkinchi imkoniyat berilmaydi!</p>
            <a href="/" class="btn">Bosh sahifaga qaytish</a>
        </div></body></html>
        """

    # CHIQISH UCHUN VAQT VA KUN TEKshiruvi (Shanba: 13:00, Boshqa kunlar: 16:00)
    hozir = datetime.now(UZB_TZ)
    if action == 'CHIQISH':
        hafta_kuni = hozir.weekday() # 5 - Shanba, 6 - Yakshanba
        soat = hozir.hour

        if hafta_kuni == 5: # Shanba
            if soat < 13:
                return f"""
                <!DOCTYPE html><html lang="uz"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Vaqt xatosi</title>
                <style>body{{font-family:'Segoe UI',sans-serif;background:#0f172a;color:#fff;display:flex;justify-content:center;align-items:center;min-height:100vh;margin:0;}}
                .card{{background:#1e293b;padding:40px 30px;border-radius:20px;text-align:center;max-width:420px;width:100%;box-shadow:0 15px 35px rgba(0,0,0,0.5);box-sizing:border-box;}}
                h2{{color:#f97316;font-size:24px;margin-bottom:15px;}}p{{color:#94a3b8;font-size:16px;margin-bottom:25px;}}
                .btn{{display:block;width:100%;padding:16px;background:#38bdf8;color:#0f172a;font-weight:bold;font-size:18px;border-radius:12px;text-decoration:none;box-sizing:border-box;}}</style></head>
                <body><div class="card">
                    <h2>⚠️ Vaqt yetmadi!</h2>
                    <p>Shanba kuni ketish davomati faqat <b>soat 13:00 dan keyin</b> ochiladi. (Hozirgi vaqt: {hozir.strftime('%H:%M')})</p>
                    <a href="/" class="btn">Bosh sahifaga qaytish</a>
                </div></body></html>
                """
        else: # Dushanba - Juma
            if soat < 16:
                return f"""
                <!DOCTYPE html><html lang="uz"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Vaqt xatosi</title>
                <style>body{{font-family:'Segoe UI',sans-serif;background:#0f172a;color:#fff;display:flex;justify-content:center;align-items:center;min-height:100vh;margin:0;}}
                .card{{background:#1e293b;padding:40px 30px;border-radius:20px;text-align:center;max-width:420px;width:100%;box-shadow:0 15px 35px rgba(0,0,0,0.5);box-sizing:border-box;}}
                h2{{color:#f97316;font-size:24px;margin-bottom:15px;}}p{{color:#94a3b8;font-size:16px;margin-bottom:25px;}}
                .btn{{display:block;width:100%;padding:16px;background:#38bdf8;color:#0f172a;font-weight:bold;font-size:18px;border-radius:12px;text-decoration:none;box-sizing:border-box;}}</style></head>
                <body><div class="card">
                    <h2>⚠️ Vaqt yetmadi!</h2>
                    <p>Ish kunlari ketish davomati faqat <b>soat 16:00 dan keyin</b> ochiladi. (Hozirgi vaqt: {hozir.strftime('%H:%M')})</p>
                    <a href="/" class="btn">Bosh sahifaga qaytish</a>
                </div></body></html>
                """

    masofa = hisobla_masofa(lat, lon, TEHNIKUM_LAT, TEHNIKUM_LON)

    if masofa > DAVOMAT_MAX_MASOFA:
        return f"""
        <!DOCTYPE html><html lang="uz"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Masofa xatosi</title>
        <style>body{{font-family:'Segoe UI',sans-serif;background:#0f172a;color:#fff;display:flex;justify-content:center;align-items:center;min-height:100vh;margin:0;}}
        .card{{background:#1e293b;padding:40px 30px;border-radius:20px;text-align:center;max-width:420px;width:100%;box-shadow:0 15px 35px rgba(0,0,0,0.5);box-sizing:border-box;}}
        h2{{color:#ef4444;font-size:24px;margin-bottom:15px;}}p{{color:#94a3b8;font-size:16px;margin-bottom:25px;}}
        .btn{{display:block;width:100%;padding:16px;background:#38bdf8;color:#0f172a;font-weight:bold;font-size:18px;border-radius:12px;text-decoration:none;box-sizing:border-box;}}</style></head>
        <body><div class="card">
            <h2>❌ Masofa xatosi!</h2>
            <p>Siz texnikumdan <b>{int(masofa)} metr</b> uzoqdasiz. Davomat faqat {DAVOMAT_MAX_MASOFA} metr ichida qabul qilinadi!</p>
            <a href="/" class="btn">Bosh sahifaga qaytish</a>
        </div></body></html>
        """

    bugun = datetime.now(UZB_TZ).strftime("%Y-%m-%d")
    fayl_nomi = f"xodim_{xodim['id']}_{bugun}_{datetime.now(UZB_TZ).strftime('%H%M%S')}.jpg"
    fayl_path = os.path.join(UPLOAD_FOLDER, fayl_nomi)
    file = request.files.get('emp_photo')
    if file and file.filename != '':
        file.save(fayl_path)
        xodim['photo_url'] = f"/uploads/{fayl_nomi}"
    else:
        xodim['photo_url'] = ""

    soat_val, minut_val = hozir.hour, hozir.minute
    vaqt_str = hozir.strftime("%H:%M:%S")

    target_chat_ids = list(ADMIN_CHAT_IDS)
    if xodim.get('telegram_id') and str(xodim['telegram_id']).strip() not in ["nan", ""]:
        target_chat_ids.append(str(xodim['telegram_id']).strip())

    if action == 'KIRISH':
        is_late = (soat_val > 8) or (soat_val == 8 and minut_val > 30)

        if masofa > OGOHLANTIRISH_MASOFA:
            holat_matn = f"Hududdan chiqib ketgan 🚨 ({int(masofa)}m)"
            tabrik_matn = f"Diqqat! Siz hududdan chiqib ketgansiz ({int(masofa)}m)! ⚠️"
            rang = "#eab308"
            holat_telegram = f"Holat: <b>🚨 DIQQAT! XODIM HUDUDDAN CHIQIB KETGAN! ({int(masofa)}m)</b>"
        elif is_late:
            holat_matn = "Kechikdi ⚠️"
            tabrik_matn = "Afsuski, siz kechikdingiz! ⚠️"
            rang = "#f97316"
            holat_telegram = "Holat: Kechikib keldi ⚠️"
        else:
            holat_matn = "Hududda"
            tabrik_matn = "Tabriklaymiz, siz vaqtida o'tdingiz! 🎉"
            rang = "#22c55e"
            holat_telegram = "Holat: O'z vaqtida keldi ✅"

        davomatni_saqlash(xodim, "KIRISH", vaqt_str, holat_str=holat_matn)
        
        matn_xabar = f"🟢 <b>KIRISH (Kelish) - {bugun}</b>\n\n👤 <b>F.I.O:</b> {xodim['fame']}\n📋 <b>Lavozimi:</b> {xodim['lavozim']}\n⏰ <b>Vaqt:</b> {vaqt_str}\n{holat_telegram}\n📍 <b>Masofa:</b> {int(masofa)}m"
        telegramga_rasmli_xabar_yuborish(fayl_path, matn_xabar, target_chat_ids=target_chat_ids)

        return f"""
        <!DOCTYPE html><html lang="uz"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Natija</title>
        <style>
            body{{font-family:'Segoe UI',sans-serif;background:#0f172a;color:#fff;display:flex;justify-content:center;align-items:center;min-height:100vh;margin:0;}}
            .card{{background:#1e293b;padding:40px 30px;border-radius:20px;text-align:center;max-width:440px;width:100%;box-shadow:0 15px 35px rgba(0,0,0,0.5);box-sizing:border-box;}}
            .profile-img{{width:130px;height:130px;border-radius:50%;object-fit:cover;border:4px solid {rang};margin-bottom:20px;}}
            h2{{color:{rang};font-size:22px;margin-bottom:10px;}}
            p{{color:#94a3b8;font-size:16px;margin:5px 0;}}
            .time-box{{background:#0f172a;padding:14px;border-radius:12px;margin:20px 0;font-size:20px;color:#38bdf8;font-weight:bold;border:1px solid #334155;}}
            .btn{{display:block;width:100%;padding:16px;background:#38bdf8;color:#0f172a;font-weight:bold;font-size:18px;border-radius:12px;text-decoration:none;box-sizing:border-box;}}
        </style></head>
        <body><div class="card">
            <h2>{tabrik_matn}</h2>
            {"<img src='" + xodim['photo_url'] + "' class='profile-img'>" if xodim['photo_url'] else ""}
            <h3 style="font-size:22px;margin:10px 0 5px 0;color:#f8fafc;">{xodim['fame']}</h3>
            <p>{xodim['lavozim']}</p>
            <div class="time-box">⏰ Vaqt: {vaqt_str}</div>
            <p style="color:#22c55e;font-size:15px;margin-bottom:25px;">Davomatingiz saqlandi va botga yuborildi! (Masofa: {int(masofa)}m)</p>
            <a href="/" class="btn">Bosh sahifaga qaytish</a>
        </div></body></html>
        """

    elif action == 'CHIQISH':
        davomatni_saqlash(xodim, "CHIQISH", vaqt_str, holat_str="Chiqib ketdi ❌")
        matn_xabar = f"🔴 <b>CHIQISH (Ketish) - {bugun}</b>\n\n👤 <b>F.I.O:</b> {xodim['fame']}\n📋 <b>Lavozimi:</b> {xodim['lavozim']}\n⏰ <b>Vaqt:</b> {vaqt_str}\n<b>Holat:</b> Ishdan ketdi 🔴\n📍 <b>Masofa:</b> {int(masofa)}m"
        telegramga_rasmli_xabar_yuborish(fayl_path, matn_xabar, target_chat_ids=target_chat_ids)

        return f"""
        <!DOCTYPE html><html lang="uz"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Natija</title>
        <style>
            body{{font-family:'Segoe UI',sans-serif;background:#0f172a;color:#fff;display:flex;justify-content:center;align-items:center;min-height:100vh;margin:0;}}
            .card{{background:#1e293b;padding:40px 30px;border-radius:20px;text-align:center;max-width:440px;width:100%;box-shadow:0 15px 35px rgba(0,0,0,0.5);box-sizing:border-box;}}
            .profile-img{{width:130px;height:130px;border-radius:50%;object-fit:cover;border:4px solid #ef4444;margin-bottom:20px;}}
            h2{{color:#ef4444;font-size:24px;margin-bottom:10px;}}
            p{{color:#94a3b8;font-size:16px;margin:5px 0;}}
            .time-box{{background:#0f172a;padding:14px;border-radius:12px;margin:20px 0;font-size:20px;color:#38bdf8;font-weight:bold;border:1px solid #334155;}}
            .btn{{display:block;width:100%;padding:16px;background:#38bdf8;color:#0f172a;font-weight:bold;font-size:18px;border-radius:12px;text-decoration:none;box-sizing:border-box;}}
        </style></head>
        <body><div class="card">
            <h2>🔴 Ketish qayd etildi</h2>
            {"<img src='" + xodim['photo_url'] + "' class='profile-img'>" if xodim['photo_url'] else ""}
            <h3 style="font-size:22px;margin:10px 0 5px 0;color:#f8fafc;">{xodim['fame']}</h3>
            <p>{xodim['lavozim']}</p>
            <div class="time-box">⏰ Vaqt: {vaqt_str}</div>
            <p style="color:#ef4444;font-size:15px;margin-bottom:25px;">Ishdan ketganingiz muvaffaqiyatli saqlandi va botga yuborildi! (Masofa: {int(masofa)}m)</p>
            <a href="/" class="btn">Bosh sahifaga qaytish</a>
        </div></body></html>
        """

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)
