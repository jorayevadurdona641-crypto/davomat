import os
import math
from flask import Flask, request, render_template_string

app = Flask(__name__)

# Texnikum koordinatalari (misol uchun)
TEXNIKUM_LAT = 38.995  # Kenglik
TEXNIKUM_LON = 66.885  # Uzunlik

# Masofani hisoblash funksiyasi (metrda)
def calculate_distance(lat1, lon1, lat2, lon2):
    R = 6371000  # Yer radiusi metrda
    phi1 = math.radians(lat1)
    phi2 = math.radians(lat2)
    delta_phi = math.radians(lat2 - lat1)
    delta_lambda = math.radians(lon2 - lon1)

    a = math.sin(delta_phi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(delta_lambda / 2) ** 2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))

    return R * c

@app.route('/')
def index():
    return "Davomat boti ishlayapti!"

@app.route('/check_location', methods=['POST'])
def check_location():
    data = request.json
    user_lat = data.get('lat')
    user_lon = data.get('lon')
    
    if not user_lat or not user_lon:
        return {"status": "error", "message": "Koordinatalar topilmadi!"}, 400

    distance = calculate_distance(user_lat, user_lon, TEXNIKUM_LAT, TEXNIKUM_LON)
    
    # 8000 metr (8 km) qilib belgilandi
    if distance > 8000:
        return {
            "status": "error", 
            "message": f"Masofa xatosi! Siz texnikumdan {int(distance)} metr uzoqdasiz. Davomat faqat 8000m ichida amalga oshiriladi."
        }
    
    return {"status": "success", "message": f"Masofa to'g'ri! {int(distance)} metr masofadasiz."}

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
