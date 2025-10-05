from flask import Flask, jsonify, request, send_file, render_template
from pymongo import MongoClient
from flask_cors import CORS
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment
from io import BytesIO
from datetime import datetime, timedelta, timezone
import os, calendar

app = Flask(__name__, template_folder="templates")
CORS(app)

# === MongoDB ===
MONGO_URI = os.getenv("MONGO_URI", "mongodb://localhost:27017/")
client = MongoClient(MONGO_URI)
db = client["gps_attendance"]
admins = db["admins"]
idx = db["idx_collection"]
checkins = db["alt_checkins"]

VN_TZ = timezone(timedelta(hours=7))

@app.route("/")
def home():
    return render_template("index.html")

# ===== ĐĂNG NHẬP =====
@app.route("/api/login", methods=["POST"])
def api_login():
    data = request.get_json()
    email, password = data.get("email"), data.get("password")
    admin = admins.find_one({"email": email})
    if not admin or not check_password_hash(admin["password"], password):
        return jsonify({"error": "Sai email hoặc mật khẩu"}), 401
    return jsonify({"message": "Đăng nhập thành công", "email": email})

# ===== RESET PASSWORD =====
@app.route("/api/reset-password", methods=["POST"])
def api_reset_password():
    data = request.get_json()
    email = data.get("email")
    new_pw = data.get("new_password")
    if not email or not new_pw:
        return jsonify({"error": "Thiếu email hoặc mật khẩu mới"}), 400
    if not admins.find_one({"email": email}):
        return jsonify({"error": "Không tìm thấy tài khoản"}), 404
    hashed_pw = generate_password_hash(new_pw)
    admins.update_one({"email": email}, {"$set": {"password": hashed_pw}})
    return jsonify({"message": "✅ Mật khẩu đã được cập nhật!"})

# ===== LẤY DỮ LIỆU CHẤM CÔNG =====
@app.route("/api/attendances", methods=["GET"])
def api_attendances():
    email = request.args.get("email")
    if not email:
        return jsonify([])
    data = list(checkins.find({}, {"_id": 0}))
    for r in data:
        if "CheckinTime" in r and isinstance(r["CheckinTime"], datetime):
            r["CheckinDate"] = r["CheckinTime"].astimezone(VN_TZ).strftime("%d/%m/%Y")
            r["CheckinTime"] = r["CheckinTime"].astimezone(VN_TZ).strftime("%H:%M:%S")
    return jsonify(data)

# ===== XUẤT EXCEL =====
@app.route("/api/export-excel", methods=["GET"])
def export_excel():
    try:
        email = request.args.get("email")
        if not email:
            return jsonify({"error": "Thiếu email"}), 400

        employees = list(idx.find({}, {"EmployeeId": 1, "EmployeeName": 1, "_id": 0}))
        template_path = os.path.join("templates", "template.xlsx")
        wb = load_workbook(template_path)
        ws = wb.active

        now = datetime.now(VN_TZ)
        month_name = calendar.month_name[now.month]

        thin = Side(border_style="thin", color="000000")
        border = Border(top=thin, left=thin, right=thin, bottom=thin)
        align = Alignment(horizontal="center", vertical="center", wrap_text=True)

        row = 4
        for emp in employees:
            emp_id, emp_name = emp.get("EmployeeId"), emp.get("EmployeeName")
            records = list(checkins.find({"EmployeeId": emp_id}).sort("CheckinTime", 1))
            if not records: continue
            ws.cell(row=row, column=1, value=emp_name)
            ws.cell(row=row, column=2, value=emp_id)
            for j, r in enumerate(records[:10], start=1):
                time_str = r.get("CheckinTime")
                if isinstance(time_str, datetime):
                    time_str = time_str.astimezone(VN_TZ).strftime("%H:%M:%S")
                entry = "; ".join(filter(None, [
                    time_str, str(r.get("Tasks", "")), r.get("Status", ""), r.get("Address", "")
                ]))
                ws.cell(row=row, column=2 + j, value=entry)
            row += 1

        for r in ws.iter_rows(min_row=4, max_row=row - 1, min_col=1, max_col=15):
            for c in r:
                c.border, c.alignment = border, align

        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        filename = f"ChamCong_{month_name}_{now.year}.xlsx"
        return send_file(buf, as_attachment=True, download_name=filename)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True, port=5000)
