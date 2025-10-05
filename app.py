from flask import Flask, jsonify, request, render_template, redirect, url_for, session
from pymongo import MongoClient
from flask_cors import CORS
from werkzeug.security import generate_password_hash, check_password_hash
from io import BytesIO
from datetime import datetime, timedelta, timezone
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment
import calendar, os

# ====== Cấu hình Flask ======
app = Flask(__name__, template_folder="templates")
CORS(app)
app.secret_key = "supersecretkey"  # dùng session lưu đăng nhập

# ====== Kết nối MongoDB ======
MONGO_URI = os.getenv("MONGO_URI", "mongodb://localhost:27017/")
client = MongoClient(MONGO_URI)
db = client["gps_attendance"]
idx_collection = db["idx_collection"]
alt_checkins = db["alt_checkins"]
admins_collection = db["admins"]

# ====== Múi giờ Việt Nam ======
VN_TZ = timezone(timedelta(hours=7))

# ======================================
# 🔹 TRANG CHỦ
# ======================================
@app.route("/")
def home():
    if "email" in session:
        return f"Xin chào {session['email']}! <a href='/logout'>Đăng xuất</a>"
    return redirect(url_for("login_form"))

# ======================================
# 🔹 ĐĂNG KÝ ADMIN
# ======================================
@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        data = request.json if request.is_json else request.form
        email = data.get("email")
        password = data.get("password")

        if not email or not password:
            return jsonify({"error": "Thiếu email hoặc mật khẩu"}), 400

        if admins_collection.find_one({"email": email}):
            return jsonify({"error": "Email đã tồn tại"}), 400

        hashed_pw = generate_password_hash(password)
        admins_collection.insert_one({"email": email, "password": hashed_pw})
        return jsonify({"message": "Đăng ký thành công"}), 200

    return render_template("register.html")

# ======================================
# 🔹 ĐĂNG NHẬP
# ======================================
@app.route("/login", methods=["GET", "POST"])
def login_form():
    if request.method == "POST":
        data = request.json if request.is_json else request.form
        email = data.get("email")
        password = data.get("password")

        admin = admins_collection.find_one({"email": email})
        if not admin or not check_password_hash(admin["password"], password):
            return jsonify({"error": "Sai email hoặc mật khẩu"}), 401

        session["email"] = email
        return jsonify({"message": "Đăng nhập thành công", "email": email})

    return render_template("login.html")

# ======================================
# 🔹 QUÊN MẬT KHẨU – Form tự reset
# ======================================
@app.route("/forgot-password", methods=["GET", "POST"])
def forgot_password():
    if request.method == "POST":
        data = request.json if request.is_json else request.form
        email = data.get("email")
        new_password = data.get("new_password")

        admin = admins_collection.find_one({"email": email})
        if not admin:
            return jsonify({"error": "Không tìm thấy tài khoản"}), 404

        hashed_pw = generate_password_hash(new_password)
        admins_collection.update_one({"email": email}, {"$set": {"password": hashed_pw}})
        return jsonify({"message": "Mật khẩu đã được đặt lại thành công"}), 200

    return render_template("forgot_password.html")

# ======================================
# 🔹 ĐĂNG XUẤT
# ======================================
@app.route("/logout")
def logout():
    session.pop("email", None)
    return redirect(url_for("login_form"))

# ======================================
# 🔹 XUẤT FILE EXCEL (bạn giữ nguyên logic này)
# ======================================
@app.route("/api/export-excel", methods=["GET"])
def export_to_excel():
    try:
        email = request.args.get("email")
        if not email:
            return jsonify({"error": "❌ Thiếu email người dùng"}), 400

        admin = admins_collection.find_one({"email": email})
        if not admin:
            return jsonify({"error": "❌ Không tìm thấy tài khoản admin với email này"}), 404

        employees = list(idx_collection.find({}, {"EmployeeId": 1, "EmployeeName": 1, "_id": 0}))
        template_path = os.path.join("templates", "template.xlsx")
        wb = load_workbook(template_path)
        ws = wb.active

        now = datetime.now(VN_TZ)
        current_month, current_year = now.month, now.year
        month_name = calendar.month_name[current_month]

        thin = Side(border_style="thin", color="000000")
        border = Border(top=thin, left=thin, right=thin, bottom=thin)
        center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

        row = 4
        for emp in employees:
            emp_id = emp.get("EmployeeId")
            emp_name = emp.get("EmployeeName")
            records = list(
                alt_checkins.find(
                    {
                        "EmployeeId": emp_id,
                        "CheckinTime": {
                            "$gte": datetime(current_year, current_month, 1, tzinfo=VN_TZ),
                            "$lt": datetime(
                                current_year,
                                current_month + 1 if current_month < 12 else 1,
                                1 if current_month < 12 else 1,
                                tzinfo=VN_TZ,
                            ),
                        },
                    }
                ).sort("CheckinTime", 1)
            )

            if not records:
                continue

            ws.cell(row=row, column=1, value=emp_name)
            ws.cell(row=row, column=2, value=emp_id)

            for j, rec in enumerate(records[:10], start=1):
                checkin_time = rec.get("CheckinTime")
                if isinstance(checkin_time, datetime):
                    time_str = checkin_time.astimezone(VN_TZ).strftime("%H:%M:%S")
                else:
                    time_str = str(checkin_time)

                parts = []
                if time_str:
                    parts.append(time_str)
                if rec.get("Tasks"):
                    parts.append(str(rec["Tasks"]))
                if rec.get("Status"):
                    parts.append(rec["Status"])
                if rec.get("Address"):
                    parts.append(rec["Address"])

                ws.cell(row=row, column=2 + j, value="; ".join(parts))
            row += 1

        for r in ws.iter_rows(min_row=4, max_row=row - 1, min_col=1, max_col=15):
            for cell in r:
                cell.border = border
                cell.alignment = center_align

        output = BytesIO()
        wb.save(output)
        output.seek(0)
        filename = f"ChamCong_{month_name}_{current_year}.xlsx"
        return send_file(output, as_attachment=True, download_name=filename)

    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ======================================
# 🔹 CHẠY ỨNG DỤNG
# ======================================
if __name__ == "__main__":
    app.run(debug=True, port=5000)
