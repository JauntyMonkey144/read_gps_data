from flask import Flask, render_template, jsonify, request, redirect, url_for, send_file
from pymongo import MongoClient
from flask_cors import CORS
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime, timedelta, timezone
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment
from io import BytesIO
import os, smtplib, uuid, re, calendar
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ---- Flask config ----
app = Flask(__name__, template_folder="templates")
CORS(app, methods=["GET", "POST"])

# ---- Timezone VN ----
VN_TZ = timezone(timedelta(hours=7))

# ---- MongoDB Config ----
MONGO_URI = os.getenv(
    "MONGO_URI",
    "mongodb+srv://banhbaobeo2205:lm2hiCLXp6B0D7hq@cluster0.festnla.mongodb.net/?retryWrites=true&w=majority"
)
DB_NAME = os.getenv("DB_NAME", "Sun_Database_1")

# ---- Kết nối MongoDB ----
client = MongoClient(MONGO_URI)
db = client[DB_NAME]
admins = db["admins"]
collection = db["alt_checkins"]       # ✅ Collection dữ liệu chấm công
reset_tokens = db["reset_tokens"]     # ✅ Token reset mật khẩu

# ---- SMTP Config ----
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))
SMTP_USER = os.getenv("SMTP_USER", "banhbaobeo2205@gmail.com")
SMTP_PASS = os.getenv("SMTP_PASS", "vynqvvvmbcigpdvy")  # App password Gmail

# ==========================================
# 🔹 ROUTES
# ==========================================

@app.route("/")
def index():
    success = request.args.get("success")
    return render_template("index.html", success=success)

# ---- Đăng nhập ----
@app.route("/login", methods=["POST", "GET"])
def login():
    if request.method == "GET":
        return redirect(url_for("index"))

    email = request.form.get("email")
    password = request.form.get("password")

    if not email or not password:
        return jsonify({"success": False, "message": "❌ Vui lòng nhập email và mật khẩu"}), 400

    admin = admins.find_one({"email": email})
    if not admin or not check_password_hash(admin.get("password", ""), password):
        return jsonify({"success": False, "message": "🚫 Email hoặc mật khẩu không đúng!"}), 401

    return jsonify({
        "success": True,
        "message": "✅ Đăng nhập thành công",
        "username": admin["username"],
        "email": admin["email"]
    })

# ==========================================
# 🔹 QUÊN MẬT KHẨU - GỬI LINK RESET
# ==========================================
@app.route("/forgot-password", methods=["GET", "POST"])
def forgot_password():
    if request.method == "GET":
        return """
        <html><head><meta charset="UTF-8"><title>Quên mật khẩu</title></head>
        <body style="font-family:Arial;background:#f4f6f9;padding:40px;text-align:center;">
            <div style="max-width:400px;margin:auto;background:white;padding:30px;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,0.1);">
                <h2>🔒 Quên mật khẩu</h2>
                <form method="POST">
                    <input type="email" name="email" placeholder="Nhập email của bạn" required
                        style="width:100%;padding:10px;margin:10px 0;border-radius:4px;border:1px solid #ccc;">
                    <button type="submit" style="padding:10px;width:100%;background:#007bff;color:white;border:none;border-radius:4px;cursor:pointer;">Gửi link đặt lại</button>
                    <p><a href="/">⬅ Quay lại đăng nhập</a></p>
                </form>
            </div>
        </body></html>
        """

    # POST → gửi mail
    email = request.form.get("email")
    admin = admins.find_one({"email": email})
    if not admin:
        return jsonify({"success": False, "message": "🚫 Email không tồn tại"}), 404

    token = str(uuid.uuid4())
    expire_time = datetime.now(VN_TZ) + timedelta(hours=1)
    reset_tokens.insert_one({"email": email, "token": token, "expire_at": expire_time})

    reset_link = f"http://localhost:5000/reset-password?token={token}"

    # Gửi mail
    try:
        msg = MIMEMultipart()
        msg["From"] = SMTP_USER
        msg["To"] = email
        msg["Subject"] = "🔐 Đặt lại mật khẩu - Hệ thống chấm công GPS"

        body = f"""
        <html>
        <body style="font-family:Arial;">
            <p>Xin chào <b>{admin.get('username', email)}</b>,</p>
            <p>Bạn đã yêu cầu đặt lại mật khẩu. Nhấn vào link dưới đây để tiếp tục:</p>
            <p><a href="{reset_link}" style="background:#007bff;color:white;padding:10px 15px;border-radius:6px;text-decoration:none;">🔑 Đặt lại mật khẩu</a></p>
            <p>Link này có hiệu lực trong 1 giờ. Nếu bạn không yêu cầu, hãy bỏ qua email này.</p>
            <hr><p style="font-size:12px;color:#888;">© 2025 Sun Automation - Hệ thống chấm công GPS</p>
        </body>
        </html>
        """
        msg.attach(MIMEText(body, "html"))

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASS)
            server.send_message(msg)

        return f"""
        <html><body style="font-family:Arial;text-align:center;padding:50px;">
        <h2>📩 Đã gửi email đặt lại mật khẩu!</h2>
        <p>Vui lòng kiểm tra hộp thư đến: <b>{email}</b></p>
        <a href="/">⬅ Quay lại đăng nhập</a>
        </body></html>
        """

    except Exception as e:
        print("❌ Lỗi gửi mail:", e)
        return jsonify({"success": False, "message": f"Lỗi gửi mail: {e}"}), 500

# ==========================================
# 🔹 RESET MẬT KHẨU
# ==========================================
@app.route("/reset-password", methods=["GET", "POST"])
def reset_password():
    token = request.args.get("token")
    if not token:
        return "<h3>❌ Token không hợp lệ</h3>"

    record = reset_tokens.find_one({"token": token})
    if not record or record["expire_at"] < datetime.now(VN_TZ):
        return "<h3>🚫 Token không tồn tại hoặc đã hết hạn</h3>"

    if request.method == "GET":
        return f"""
        <html><body style="font-family:Arial;text-align:center;padding:50px;">
        <h2>🔑 Nhập mật khẩu mới</h2>
        <form method="POST">
            <input type="hidden" name="token" value="{token}">
            <input type="password" name="new_password" placeholder="Mật khẩu mới" required
                   style="padding:10px;width:250px;"><br><br>
            <button type="submit" style="padding:10px 20px;background:#28a745;color:white;border:none;border-radius:4px;cursor:pointer;">Cập nhật</button>
        </form>
        </body></html>
        """

    new_pw = request.form.get("new_password")
    hashed_pw = generate_password_hash(new_pw)
    admins.update_one({"email": record["email"]}, {"$set": {"password": hashed_pw}})
    reset_tokens.delete_one({"token": token})

    return """
    <html><body style="font-family:Arial;text-align:center;padding:50px;">
    <h2>✅ Mật khẩu đã được cập nhật thành công!</h2>
    <a href="/">⬅ Quay lại đăng nhập</a>
    </body></html>
    """

# ==========================================
# 🔹 HÀM BUILD QUERY LỌC
# ==========================================
def build_query(filter_type, start_date, end_date, search):
    query = {}
    today = datetime.now(VN_TZ)

    if filter_type == "custom" and start_date and end_date:
        query["CheckinDate"] = {"$gte": start_date, "$lte": end_date}
    elif filter_type == "hôm nay":
        query["CheckinDate"] = today.strftime("%Y-%m-%d")
    elif filter_type == "tuần":
        start = (today - timedelta(days=today.weekday())).strftime("%Y-%m-%d")
        end = (today + timedelta(days=6 - today.weekday())).strftime("%Y-%m-%d")
        query["CheckinDate"] = {"$gte": start, "$lte": end}
    elif filter_type == "tháng":
        start = today.replace(day=1).strftime("%Y-%m-%d")
        end = today.replace(day=calendar.monthrange(today.year, today.month)[1]).strftime("%Y-%m-%d")
        query["CheckinDate"] = {"$gte": start, "$lte": end}
    elif filter_type == "năm":
        query["CheckinDate"] = {"$regex": f"^{today.year}"}
    elif filter_type == "nghỉ phép":
        regex = re.compile("Nghỉ phép", re.IGNORECASE)
        query["$or"] = [
            {"Tasks": {"$regex": regex}},
            {"Status": {"$regex": regex}},
            {"OtherNote": {"$regex": regex}}
        ]

    if search:
        regex = re.compile(search, re.IGNORECASE)
        query["$or"] = [
            {"EmployeeId": {"$regex": regex}},
            {"EmployeeName": {"$regex": regex}}
        ]
    return query

# ==========================================
# 🔹 API LẤY DỮ LIỆU CHẤM CÔNG
# ==========================================
@app.route("/api/attendances", methods=["GET"])
def get_attendances():
    try:
        email = request.args.get("email")
        if not email:
            return jsonify({"error": "❌ Thiếu email"}), 400

        admin = admins.find_one({"email": email})
        if not admin:
            return jsonify({"error": "🚫 Email không hợp lệ"}), 403

        filter_type = request.args.get("filter", "hôm nay").lower()
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()

        query = build_query(filter_type, start_date, end_date, search)
        data = list(collection.find(query, {"_id": 0}))
        print(f"DEBUG: {len(data)} records fetched for {email}")
        return jsonify(data)
    except Exception as e:
        print("❌ Error in get_attendances:", e)
        return jsonify({"error": str(e)}), 500

# ==========================================
# 🔹 CHẠY ỨNG DỤNG
# ==========================================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
