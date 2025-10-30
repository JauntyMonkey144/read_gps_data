from flask import Flask, render_template, jsonify, request, redirect, url_for, send_file
from pymongo import MongoClient
from flask_cors import CORS
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime, timedelta, timezone
import os
import re
import calendar
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment
import secrets
import requests
from dotenv import load_dotenv

app = Flask(__name__, template_folder="templates")
CORS(app, methods=["GET", "POST"])
# ---- Init App ----
# ---- Cấu hình Domain chính cho ứng dụng ----
# Đặt 2 dòng này để url_for(_external=True) tạo link đúng (https://app.sun-automation.id.vn)
app.config["SERVER_NAME"] = os.environ.get("SERVER_NAME", "system.sun-automation.id.vn")
app.config["PREFERRED_URL_SCHEME"] = "https"

# ---- Timezone VN ----
VN_TZ = timezone(timedelta(hours=7))
load_dotenv() # Tải các biến từ file .env
# ---- MongoDB Config ----
MONGO_URI = os.getenv("MONGO_URI") 
DB_NAME = os.getenv("DB_NAME", "Sun_Database_1") 

# ---- Resend API Config ----
RESEND_API_KEY = os.getenv("RESEND_API_KEY")
# Email gửi đi từ domain bạn đã xác thực
RESEND_FROM_EMAIL = os.getenv("RESEND_FROM_EMAIL", "system@sun-automation.id.vn")

def send_email_resend(to_email, subject, html_body):
    """Gửi email bằng Resend API."""
    if not RESEND_API_KEY:
        print("❌ Lỗi: Biến môi trường RESEND_API_KEY chưa được đặt.")
        return False
    
    url = "https://api.resend.com/emails"
    payload = {
        "from": f"Sun Automation System <{RESEND_FROM_EMAIL}>",
        "to": [to_email],
        "subject": subject,
        "html": html_body
    }
    headers = {
        "Authorization": f"Bearer {RESEND_API_KEY}",
        "Content-Type": "application/json"
    }
    
    try:
        response = requests.post(url, json=payload, headers=headers)
        if response.status_code == 200:
            print(f"✅ Gửi email thành công đến {to_email}")
            return True
        else:
            print(f"❌ Lỗi Resend API ({response.status_code}): {response.text}")
            return False
    except Exception as e:
        print(f"❌ Lỗi ngoại lệ khi gửi email: {e}")
        return False

# ---- Kết nối MongoDB ----
client = MongoClient(MONGO_URI)
db = client[DB_NAME]

# Các collection sử dụng
admins = db["admins"]
users = db["users"]
collection = db["alt_checkins"]
reset_tokens = db["reset_tokens"]  # Collection for password reset tokens

# ---- Trang chủ (đăng nhập chính) ----
@app.route("/")
def index():
    return render_template("index.html")

# ---- Đăng nhập API ----
@app.route("/login", methods=["POST", "GET"])
def login():
    if request.method == "GET":
        return redirect(url_for("index"))
    email = request.form.get("email")
    password = request.form.get("password")
    if not email or not password:
        return jsonify({"success": False, "message": "❌ Vui lòng nhập email và mật khẩu"}), 400
    admin = admins.find_one({"email": email})
    if admin and check_password_hash(admin.get("password", ""), password):
        return jsonify({
            "success": True, "message": "✅ Đăng nhập thành công",
            "username": admin["username"], "email": admin["email"], "role": "admin"
        })
    user = users.find_one({"email": email})
    if user and check_password_hash(user.get("password", ""), password):
        return jsonify({
            "success": True, "message": "✅ Đăng nhập thành công",
            "username": user["username"], "email": user["email"], "role": "user"
        })
    return jsonify({"success": False, "message": "🚫 Email hoặc mật khẩu không đúng!"}), 401

# THAY THẾ TOÀN BỘ HÀM CŨ BẰNG HÀM MỚI NÀY

@app.route("/request-reset-password", methods=["POST"])
def request_reset_password():
    email = request.form.get("email")
    if not email:
        return """
        <!DOCTYPE html><html lang="vi"><head><meta charset="UTF-8"><title>Lỗi</title>
        <style>body{font-family:Arial,sans-serif;background:#f4f6f9;padding:20px}.container{max-width:400px;margin:100px auto;background:white;padding:30px;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,.1)}p{color:#dc3545;text-align:center}</style>
        </head><body><div class="container"><p>❌ Vui lòng nhập email</p>
        <a href="/forgot-password">Thử lại</a></div></body></html>""", 400

    account = admins.find_one({"email": email}) or users.find_one({"email": email})
    if not account:
        return """
        <!DOCTYPE html><html lang="vi"><head><meta charset="UTF-8"><title>Lỗi</title>
        <style>body{font-family:Arial,sans-serif;background:#f4f6f9;padding:20px}.container{max-width:400px;margin:100px auto;background:white;padding:30px;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,.1)}p{color:#dc3545;text-align:center}</style>
        </head><body><div class="container"><p>🚫 Email không tồn tại!</p>
        <a href="/forgot-password">Thử lại</a></div></body></html>""", 404

    # Generate reset token
    token = secrets.token_urlsafe(32)
    expiration = datetime.now(VN_TZ).astimezone(timezone.utc).replace(tzinfo=None) + timedelta(days=1)
    reset_tokens.insert_one({
        "email": email,
        "token": token,
        "expiration": expiration
    })

    # Nội dung email HTML (giữ nguyên)
    reset_link = url_for("reset_password", token=token, _external=True)
    html_body = f"""
    <div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: auto; padding: 20px; border: 1px solid #ddd; border-radius: 10px;">
        <h2 style="color: #007bff; text-align: center;">Yêu cầu đặt lại mật khẩu</h2>
        <p>Xin chào,</p>
        <p>Bạn đã yêu cầu đặt lại mật khẩu cho tài khoản của mình. Vui lòng nhấp vào nút bên dưới để hoàn tất quá trình:</p>
        <p style="text-align: center; margin: 30px 0;">
            <a href="{reset_link}" 
               style="background-color: #28a745; color: white; padding: 12px 25px; text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block;">
               Đặt lại mật khẩu
            </a>
        </p>
        <p>Liên kết này sẽ hết hạn sau 1 ngày.</p>
        <p>Nếu bạn không yêu cầu đặt lại mật khẩu, vui lòng bỏ qua email này một cách an toàn.</p>
        <hr style="border: none; border-top: 1px solid #eee;">
        <p style="font-size: 0.9em; color: #6c757d; text-align: center;">
            Trân trọng,<br>
            Hệ thống tự động Sun Automation
        </p>
    </div>
    """

    # Gửi email bằng Resend
    if send_email_resend(email, "Yêu cầu đặt lại mật khẩu", html_body):
        return """
        <!DOCTYPE html><html lang="vi"><head><meta charset="UTF-8"><title>Gửi liên kết thành công</title>
        <style>body{font-family:Arial,sans-serif;background:#f4f6f9;padding:20px}.container{max-width:400px;margin:100px auto;background:white;padding:30px;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,.1)}.success{color:#28a745;text-align:center;font-size:18px;margin-bottom:20px}button{background:#28a745;color:white;padding:12px;width:100%;border:none;border-radius:4px;cursor:pointer;font-size:16px}</style>
        </head><body><div class="container"><div class="success">✅ Email chứa liên kết đặt lại mật khẩu đã được gửi thành công! Vui lòng kiểm tra hộp thư của bạn.</div>
        <a href="/"><button>Quay về trang chủ</button></a></div></body></html>"""
    else:
        return """
        <!DOCTYPE html><html lang="vi"><head><meta charset="UTF-8"><title>Lỗi</title>
        <style>body{font-family:Arial,sans-serif;background:#f4f6f9;padding:20px}.container{max-width:400px;margin:100px auto;background:white;padding:30px;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,.1)}p{color:#dc3545;text-align:center}</style>
        </head><body><div class="container"><p>❌ Lỗi khi gửi email, vui lòng thử lại sau</p>
        <a href="/forgot-password">Thử lại</a></div></body></html>""", 500

# ---- Trang reset mật khẩu với token (ĐÃ SỬA LỖI) ----
@app.route("/reset-password/<token>", methods=["GET", "POST"])
def reset_password(token):
    # Sửa lỗi: Xử lý cả yêu cầu GET và HEAD
    if request.method in ["GET", "HEAD"]:
        token_data = reset_tokens.find_one({"token": token})
        # Compare expiration as offset-naive (UTC) datetime
        if not token_data or token_data["expiration"] < datetime.now(timezone.utc).replace(tzinfo=None):
            return """
            <!DOCTYPE html><html lang="vi"><head><meta charset="UTF-8"><title>Lỗi</title>
            <style>body{font-family:Arial,sans-serif;background:#f4f6f9;padding:20px}.container{max-width:400px;margin:100px auto;background:white;padding:30px;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,.1)}p{color:#dc3545;text-align:center}</style>
            </head><body><div class="container"><p>🚫 Liên kết đặt lại mật khẩu không hợp lệ hoặc đã hết hạn!</p>
            <a href="/forgot-password">Thử lại</a></div></body></html>""", 400

        return """
        <!DOCTYPE html><html lang="vi"><head><meta charset="UTF-8"><title>Đặt lại mật khẩu</title>
        <style>body{font-family:Arial,sans-serif;background:#f4f6f9;padding:20px}.container{max-width:400px;margin:100px auto;background:white;padding:30px;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,.1)}input{width:100%;padding:10px;margin:10px 0;box-sizing:border-box;border:1px solid #ddd;border-radius:4px}button{background:#28a745;color:white;padding:12px;width:100%;border:none;border-radius:4px;cursor:pointer;font-size:16px}</style>
        </head><body><div class="container"><h2>🔒 Đặt lại mật khẩu</h2><form method="POST">
        <input type="password" name="new_password" placeholder="Mật khẩu mới" required>
        <input type="password" name="confirm_password" placeholder="Xác nhận mật khẩu" required>
        <button type="submit">Cập nhật mật khẩu</button></form></div></body></html>"""

    if request.method == "POST":
        token_data = reset_tokens.find_one({"token": token})
        # Compare expiration as offset-naive (UTC) datetime
        if not token_data or token_data["expiration"] < datetime.now(timezone.utc).replace(tzinfo=None):
            return """
            <!DOCTYPE html><html lang="vi"><head><meta charset="UTF-8"><title>Lỗi</title>
            <style>body{font-family:Arial,sans-serif;background:#f4f6f9;padding:20px}.container{max-width:400px;margin:100px auto;background:white;padding:30px;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,.1)}p{color:#dc3545;text-align:center}</style>
            </head><body><div class="container"><p>❌ Liên kết không hợp lệ hoặc đã hết hạn</p>
            <a href="/forgot-password">Thử lại</a></div></body></html>""", 400

        new_password = request.form.get("new_password")
        confirm_password = request.form.get("confirm_password")
        if not new_password or not confirm_password:
            return """
            <!DOCTYPE html><html lang="vi"><head><meta charset="UTF-8"><title>Lỗi</title>
            <style>body{font-family:Arial,sans-serif;background:#f4f6f9;padding:20px}.container{max-width:400px;margin:100px auto;background:white;padding:30px;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,.1)}p{color:#dc3545;text-align:center}</style>
            </head><body><div class="container"><p>❌ Vui lòng điền đầy đủ thông tin</p>
            <a href="/reset-password/{}">Thử lại</a></div></body></html>""".format(token), 400
        if new_password != confirm_password:
            return """
            <!DOCTYPE html><html lang="vi"><head><meta charset="UTF-8"><title>Lỗi</title>
            <style>body{font-family:Arial,sans-serif;background:#f4f6f9;padding:20px}.container{max-width:400px;margin:100px auto;background:white;padding:30px;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,.1)}p{color:#dc3545;text-align:center}</style>
            </head><body><div class="container"><p>❌ Mật khẩu xác nhận không khớp</p>
            <a href="/reset-password/{}">Thử lại</a></div></body></html>""".format(token), 400

        email = token_data["email"]
        account = admins.find_one({"email": email}) or users.find_one({"email": email})
        if not account:
            return """
            <!DOCTYPE html><html lang="vi"><head><meta charset="UTF-8"><title>Lỗi</title>
            <style>body{font-family:Arial,sans-serif;background:#f4f6f9;padding:20px}.container{max-width:400px;margin:100px auto;background:white;padding:30px;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,.1)}p{color:#dc3545;text-align:center}</style>
            </head><body><div class="container"><p>🚫 Email không tồn tại!</p>
            <a href="/forgot-password">Thử lại</a></div></body></html>""", 404

        hashed_pw = generate_password_hash(new_password)
        collection_to_update = admins if "username" in account else users
        collection_to_update.update_one({"email": email}, {"$set": {"password": hashed_pw}})
        reset_tokens.delete_one({"token": token}) # Remove used token

        return """
        <!DOCTYPE html><html lang="vi"><head><title>Thay đổi mật khẩu thành công</title>
        <style>body{font-family:Arial,sans-serif;background:#f4f6f9;padding:20px}.container{max-width:400px;margin:100px auto;background:white;padding:30px;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,.1)}.success{color:#28a745;text-align:center;font-size:18px;margin-bottom:20px}button{background:#28a745;color:white;padding:12px;width:100%;border:none;border-radius:4px;cursor:pointer;font-size:16px}</style>
        </head><body><div class="container"><div class="success">✅ Thay đổi mật khẩu thành công! Bạn có thể đăng nhập với mật khẩu mới.</div>
        <a href="/"><button>Quay về trang chủ</button></a></div></body></html>"""

# ---- Reset mật khẩu (giữ nguyên chức năng cũ) ----
@app.route("/forgot-password", methods=["GET", "POST"])
def forgot_password():
    if request.method == "GET":
        return """
        <!DOCTYPE html><html lang="vi"><head><meta charset="UTF-8"><title>Đặt lại mật khẩu</title>
        <style>body{font-family:Arial,sans-serif;background:#f4f6f9;padding:20px}.container{max-width:400px;margin:100px auto;background:white;padding:30px;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,.1)}input{width:100%;padding:10px;margin:10px 0;box-sizing:border-box;border:1px solid #ddd;border-radius:4px}button{background:#28a745;color:white;padding:12px;width:100%;border:none;border-radius:4px;cursor:pointer;font-size:16px}</style>
        </head><body><div class="container"><h2>🔒 Đặt lại mật khẩu</h2><form method="POST" action="/request-reset-password">
        <input type="email" name="email" placeholder="Email" required>
        <button type="submit">Gửi liên kết đặt lại</button><a href="/">Quay về trang chủ</a></form></div></body></html>"""
    if request.method == "POST":
        email = request.form.get("email")
        new_password = request.form.get("new_password")
        confirm_password = request.form.get("confirm_password")
        if not all([email, new_password, confirm_password]):
            return """
            <!DOCTYPE html><html lang="vi"><head><meta charset="UTF-8"><title>Lỗi</title>
            <style>body{font-family:Arial,sans-serif;background:#f4f6f9;padding:20px}.container{max-width:400px;margin:100px auto;background:white;padding:30px;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,.1)}p{color:#dc3545;text-align:center}</style>
            </head><body><div class="container"><p>❌ Vui lòng điền đầy đủ thông tin</p>
            <a href="/forgot-password">Thử lại</a></div></body></html>""", 400
        if new_password != confirm_password:
            return """
            <!DOCTYPE html><html lang="vi"><head><meta charset="UTF-8"><title>Lỗi</title>
            <style>body{font-family:Arial,sans-serif;background:#f4f6f9;padding:20px}.container{max-width:400px;margin:100px auto;background:white;padding:30px;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,.1)}p{color:#dc3545;text-align:center}</style>
            </head><body><div class="container"><p>❌ Mật khẩu xác nhận không khớp</p>
            <a href="/forgot-password">Thử lại</a></div></body></html>""", 400
        account = admins.find_one({"email": email}) or users.find_one({"email": email})
        if not account:
            return """
            <!DOCTYPE html><html lang="vi"><head><meta charset="UTF-8"><title>Lỗi</title>
            <style>body{font-family:Arial,sans-serif;background:#f4f6f9;padding:20px}.container{max-width:400px;margin:100px auto;background:white;padding:30px;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,.1)}p{color:#dc3545;text-align:center}</style>
            </head><body><div class="container"><p>🚫 Email không tồn tại!</p>
            <a href="/forgot-password">Thử lại</a></div></body></html>""", 404
        hashed_pw = generate_password_hash(new_password)
        collection_to_update = admins if "username" in account else users
        collection_to_update.update_one({"email": email}, {"$set": {"password": hashed_pw}})
        return """
        <!DOCTYPE html><html lang="vi"><head><title>Thay đổi mật khẩu thành công</title>
        <style>body{font-family:Arial,sans-serif;background:#f4f6f9;padding:20px}.container{max-width:400px;margin:100px auto;background:white;padding:30px;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,.1)}.success{color:#28a745;text-align:center;font-size:18px;margin-bottom:20px}button{background:#28a745;color:white;padding:12px;width:100%;border:none;border-radius:4px;cursor:pointer;font-size:16px}</style>
        </head><body><div class="container"><div class="success">✅ Thay đổi mật khẩu thành công! Bạn có thể đăng nhập với mật khẩu mới.</div>
        <a href="/"><button>Quay về trang chủ</button></a></div></body></html>"""
        
# ---- Build leave query (lọc theo dateType)----
def build_leave_query(filter_type, start_date_str, end_date_str, search, date_type="CheckinTime", username=None):
    today = datetime.now(VN_TZ)
    regex_leave = re.compile("Nghỉ phép", re.IGNORECASE)
    conditions = [{"$or": [{"Tasks": regex_leave}, {"Reason": {"$exists": True}}]}]
    date_filter = {}

    start_dt, end_dt = None, None

    # Xác định khoảng thời gian start_dt và end_dt
    if filter_type == "custom" and start_date_str and end_date_str:
        try:
            start_dt = datetime.strptime(start_date_str, "%Y-%m-%d").replace(tzinfo=VN_TZ)
            end_dt = datetime.strptime(end_date_str, "%Y-%m-%d").replace(hour=23, minute=59, second=59, tzinfo=VN_TZ)
        except ValueError:
            pass # Bỏ qua nếu định dạng ngày không hợp lệ
    elif filter_type != "tất cả":
        if filter_type == "hôm nay":
            start_dt, end_dt = today.replace(hour=0, minute=0, second=0), today.replace(hour=23, minute=59, second=59)
        elif filter_type == "tuần":
            start_dt = (today - timedelta(days=today.weekday())).replace(hour=0, minute=0, second=0)
            end_dt = (start_dt + timedelta(days=6)).replace(hour=23, minute=59, second=59)
        elif filter_type == "tháng":
            start_dt = today.replace(day=1, hour=0, minute=0, second=0)
            _, last_day = calendar.monthrange(today.year, today.month)
            end_dt = today.replace(day=last_day, hour=23, minute=59, second=59)
        elif filter_type == "năm":
            start_dt = today.replace(month=1, day=1, hour=0, minute=0, second=0)
            end_dt = today.replace(month=12, day=31, hour=23, minute=59, second=59)

    # Xây dựng bộ lọc ngày tháng nếu có khoảng thời gian hợp lệ
    if start_dt and end_dt:
        # Quan trọng: Không lọc theo LeaveDate ở đây nữa, sẽ xử lý sau
        if date_type == "CheckinTime":
            date_filter = {"CreationTime": {"$gte": start_dt, "$lte": end_dt}}
        elif date_type == "ApprovalDate1":
            date_filter = {"ApprovalDate1": {"$gte": start_dt, "$lte": end_dt}}
        elif date_type == "ApprovalDate2":
            date_filter = {"ApprovalDate2": {"$gte": start_dt, "$lte": end_dt}}

    if date_filter:
        conditions.append(date_filter)

    if search:
        regex = re.compile(search, re.IGNORECASE)
        conditions.append({"$or": [{"EmployeeId": regex}, {"EmployeeName": regex}]})
    if username:
        conditions.append({"EmployeeName": username})

    return {"$and": conditions}

# ---- Build attendance query ----
def build_attendance_query(filter_type, start_date, end_date, search, username=None):
    today = datetime.now(VN_TZ)
    conditions = [{"CheckType": {"$in": ["checkin", "checkout"]}}]
    date_filter = {}
    if filter_type == "custom" and start_date and end_date:
        start_dt = datetime.strptime(start_date, "%Y-%m-%d").replace(tzinfo=VN_TZ)
        end_dt = datetime.strptime(end_date, "%Y-%m-%d").replace(hour=23, minute=59, second=59, tzinfo=VN_TZ)
        date_filter = {"Timestamp": {"$gte": start_dt, "$lte": end_dt}}
    elif filter_type == "hôm nay":
        date_filter = {"CheckinDate": today.strftime("%d/%m/%Y")}
    elif filter_type == "tuần":
        start_dt = (today - timedelta(days=today.weekday())).replace(hour=0, minute=0, second=0)
        end_dt = (start_dt + timedelta(days=6)).replace(hour=23, minute=59, second=59)
        date_filter = {"Timestamp": {"$gte": start_dt, "$lte": end_dt}}
    elif filter_type == "tháng":
        date_filter = {"CheckinDate": {"$regex": f"/{today.month:02d}/{today.year}$"}}
    elif filter_type == "năm":
        date_filter = {"CheckinDate": {"$regex": f"/{today.year}$"}}
    if date_filter: conditions.append(date_filter)
    if search:
        regex = re.compile(search, re.IGNORECASE)
        conditions.append({"$or": [{"EmployeeId": regex}, {"EmployeeName": regex}]})
    if username:
        conditions.append({"EmployeeName": username})
    return {"$and": conditions}

# ---- Helper functions ----
def calculate_leave_days_for_month(record, export_year, export_month):
    """
    Calculates the number of leave days for a given record within a specific month and year.
    - Returns (0, False) if cannot parse dates.
    - Computes days_in_month ignoring status.
    - If days_in_month == 0, returns (0, False)
    - Else, determines final status and adjusts days_in_month accordingly.
    """
    display_date = record.get("DisplayDate", "").strip().lower()
    start_date, end_date = None, None

    # Parse start and end dates from various possible fields
    try:
        if display_date:
            if "từ" in display_date and "đến" in display_date:
                # Format: "Từ YYYY-MM-DD ... đến YYYY-MM-DD ..."
                date_parts = re.findall(r"\d{4}-\d{2}-\d{2}", display_date)
                if len(date_parts) == 2:
                    start_date = datetime.strptime(date_parts[0], "%Y-%m-%d")
                    end_date = datetime.strptime(date_parts[1], "%Y-%m-%d")
            else:
                 # Format: "YYYY-MM-DD ..." (single day)
                date_part = display_date.split()[0]
                start_date = end_date = datetime.strptime(date_part, "%Y-%m-%d")
        elif record.get('StartDate') and record.get('EndDate'):
             start_date = datetime.strptime(record['StartDate'], "%Y-%m-%d")
             end_date = datetime.strptime(record['EndDate'], "%Y-%m-%d")
        elif record.get('LeaveDate'):
             start_date = end_date = datetime.strptime(record['LeaveDate'], "%Y-%m-%d")
    except (ValueError, TypeError, IndexError):
        return 0.0, False # Cannot parse date

    if not start_date or not end_date:
        return 0.0, False

    # Compute days_in_month
    days_in_month = 0.0
    _, last_day = calendar.monthrange(export_year, export_month)
    month_start = datetime(export_year, export_month, 1)
    month_end = datetime(export_year, export_month, last_day)
    if start_date == end_date:
        if month_start <= start_date <= month_end and start_date.weekday() <= 5:
            if "cả ngày" in display_date or not ("sáng" in display_date or "chiều" in display_date or record.get('Session', '').lower() in ['sáng', 'chiều']):
                days_in_month = 1.0
            else:
                days_in_month = 0.5
    else:
        current_date = max(start_date, month_start)
        end = min(end_date, month_end)
        while current_date <= end:
            if current_date.weekday() <= 5:
                days_in_month += 1.0
            current_date += timedelta(days=1)

    if days_in_month == 0:
        return 0.0, False

    # Determine final status
    status1 = str(record.get("Status1") or "").lower()
    status2 = str(record.get("Status2") or "").lower()

    if "từ chối" in status2:
        return 0.0, True
    elif "duyệt" in status2:
        return days_in_month, True
    elif "duyệt" in status1:
        return days_in_month, True
    else:
        return 0.0, True

def get_formatted_approval_date(approval_date):
    if not approval_date: return ""
    try: return approval_date.astimezone(VN_TZ).strftime("%d/%m/%Y %H:%M:%S") if isinstance(approval_date, datetime) else str(approval_date)
    except: return str(approval_date)

# ---- API lấy dữ liệu chấm công ----
@app.route("/api/attendances", methods=["GET"])
def get_attendances():
    try:
        email = request.args.get("email")
        admin = admins.find_one({"email": email})
        user = users.find_one({"email": email})
        if not admin and not user: return jsonify({"error": "🚫 Email không tồn tại"}), 403
        username = None if admin else user["username"]
        query = build_attendance_query(
            request.args.get("filter", "hôm nay").lower(),
            request.args.get("startDate"), request.args.get("endDate"),
            request.args.get("search", "").strip(), username=username
        )
        
        all_relevant_data = list(collection.find(query, {"_id": 0}))
        daily_hours_map, monthly_hours_map = {}, {}
        emp_data = {}
        for rec in all_relevant_data:
            emp_id = rec.get("EmployeeId")
            if emp_id: emp_data.setdefault(emp_id, []).append(rec)
        
        for emp_id, records in emp_data.items():
            daily_groups = {}
            for rec in records:
                date_str = rec.get("CheckinDate")
                if date_str: daily_groups.setdefault(date_str, []).append(rec)
            
            for date_str, day_records in daily_groups.items():
                checkins = []
                for r in day_records:
                    if r.get('CheckType') == 'checkin' and r.get('Timestamp'):
                        try:
                            if isinstance(r['Timestamp'], str):
                                timestamp = datetime.strptime(r['Timestamp'], "%Y-%m-%d %H:%M:%S")
                            elif isinstance(r['Timestamp'], datetime):
                                timestamp = r['Timestamp']
                            else:
                                continue
                            checkins.append(timestamp)
                        except (ValueError, TypeError):
                            continue
                checkins = sorted(checkins)
                checkouts = []
                for r in day_records:
                    if r.get('CheckType') == 'checkout' and r.get('Timestamp'):
                        try:
                            if isinstance(r['Timestamp'], str):
                                timestamp = datetime.strptime(r['Timestamp'], "%Y-%m-%d %H:%M:%S")
                            elif isinstance(r['Timestamp'], datetime):
                                timestamp = r['Timestamp']
                            else:
                                continue
                            checkouts.append(timestamp)
                        except (ValueError, TypeError):
                            continue
                checkouts = sorted(checkouts)
                daily_seconds = 0
                if checkins and checkouts and checkouts[-1] > checkins[0]:
                    daily_seconds = (checkouts[-1] - checkins[0]).total_seconds()
                daily_hours_map[(emp_id, date_str)] = daily_seconds
                # Update all records for this employee and date with DailyHours and _dailySeconds
                h, rem = divmod(daily_seconds, 3600)
                m, s = divmod(rem, 60)
                daily_hours = f"{int(h)}h {int(m)}m {int(s)}s" if daily_seconds > 0 else ""
                collection.update_many(
                    {"EmployeeId": emp_id, "CheckinDate": date_str, "CheckType": {"$in": ["checkin", "checkout"]}},
                    {"$set": {"DailyHours": daily_hours, "_dailySeconds": daily_seconds}}
                )
            monthly_groups = {}
            for (map_emp_id, map_date_str), daily_seconds in daily_hours_map.items():
                if map_emp_id == emp_id:
                    try: month_key = datetime.strptime(map_date_str, "%d/%m/%Y").strftime("%Y-%m")
                    except: continue
                    monthly_groups.setdefault(month_key, []).append((map_date_str, daily_seconds))
            
            for month, days in monthly_groups.items():
                sorted_days = sorted(days, key=lambda x: datetime.strptime(x[0], "%d/%m/%Y"))
                running_total = 0
                for date_str, daily_seconds in sorted_days:
                    running_total += daily_seconds
                    monthly_hours_map[(emp_id, date_str)] = running_total
                    # Update all records for this employee and date with MonthlyHours and _monthlySeconds
                    h, rem = divmod(running_total, 3600)
                    m, s = divmod(rem, 60)
                    monthly_hours = f"{int(h)}h {int(m)}m {int(s)}s" if running_total > 0 else ""
                    collection.update_many(
                        {"EmployeeId": emp_id, "CheckinDate": date_str, "CheckType": {"$in": ["checkin", "checkout"]}},
                        {"$set": {"MonthlyHours": monthly_hours, "_monthlySeconds": running_total}}
                    )
        
        for item in all_relevant_data:
            emp_id, date_str = item.get("EmployeeId"), item.get("CheckinDate")
            daily_sec = daily_hours_map.get((emp_id, date_str), 0)
            h, rem = divmod(daily_sec, 3600)
            m, s = divmod(rem, 60)
            item['DailyHours'], item['_dailySeconds'] = (f"{int(h)}h {int(m)}m {int(s)}s" if daily_sec > 0 else ""), daily_sec
            monthly_sec = monthly_hours_map.get((emp_id, date_str), 0)
            h, rem = divmod(monthly_sec, 3600)
            m, s = divmod(rem, 60)
            item['MonthlyHours'], item['_monthlySeconds'] = (f"{int(h)}h {int(m)}m {int(s)}s" if monthly_sec > 0 else ""), monthly_sec
            if item.get('Timestamp'):
                try:
                    if isinstance(item['Timestamp'], str):
                        timestamp = datetime.strptime(item['Timestamp'], "%Y-%m-%d %H:%M:%S")
                    elif isinstance(item['Timestamp'], datetime):
                        timestamp = item['Timestamp']
                    else:
                        timestamp = None
                    item['CheckinTime'] = timestamp.astimezone(VN_TZ).strftime('%H:%M:%S') if timestamp else ""
                except (ValueError, TypeError):
                    item['CheckinTime'] = ""
        return jsonify(all_relevant_data)
    except Exception as e:
        print(f"❌ Lỗi tại get_attendances: {e}")
        return jsonify({"error": str(e)}), 500

# ---- API lấy dữ liệu nghỉ phép ----
@app.route("/api/leaves", methods=["GET"])
def get_leaves():
    try:
        email = request.args.get("email")
        admin = admins.find_one({"email": email})
        user = users.find_one({"email": email})
        if not admin and not user: return jsonify({"error": "🚫 Email không tồn tại"}), 403
        
        username = None if admin else user["username"]
        date_type = request.args.get("dateType", "CheckinDate")
        filter_type = request.args.get("filter", "tất cả").lower()
        start_date_str = request.args.get("startDate")
        end_date_str = request.args.get("endDate")
        search = request.args.get("search", "").strip()

        query = build_leave_query(filter_type, start_date_str, end_date_str, search, date_type, username=username)
        data = list(collection.find(query, {"_id": 0}))

        # <<< PHẦN SỬA LỖI BẮT ĐẦU >>>
        # Xử lý lọc theo Ngày nghỉ (LeaveDate) sau khi đã lấy dữ liệu từ DB
        if date_type == "LeaveDate" and filter_type != "tất cả":
            filter_start_dt, filter_end_dt = None, None
            today = datetime.now(VN_TZ)

            # Xác định khoảng thời gian lọc
            if filter_type == "custom" and start_date_str and end_date_str:
                try:
                    filter_start_dt = datetime.strptime(start_date_str, "%Y-%m-%d").date()
                    filter_end_dt = datetime.strptime(end_date_str, "%Y-%m-%d").date()
                except (ValueError, TypeError):
                    pass
            else:
                if filter_type == "hôm nay":
                    filter_start_dt = filter_end_dt = today.date()
                elif filter_type == "tuần":
                    filter_start_dt = (today - timedelta(days=today.weekday())).date()
                    filter_end_dt = (filter_start_dt + timedelta(days=6))
                elif filter_type == "tháng":
                    filter_start_dt = today.replace(day=1).date()
                    _, last_day = calendar.monthrange(today.year, today.month)
                    filter_end_dt = today.replace(day=last_day).date()
                elif filter_type == "năm":
                    filter_start_dt = today.replace(month=1, day=1).date()
                    filter_end_dt = today.replace(month=12, day=31).date()

            # Nếu có khoảng thời gian lọc hợp lệ, tiến hành lọc dữ liệu
            if filter_start_dt and filter_end_dt:
                filtered_data = []
                for item in data:
                    display_date = item.get("DisplayDate", "")
                    if not display_date: continue

                    record_start_dt, record_end_dt = None, None
                    try:
                        if "đến" in display_date: # Dạng "Từ YYYY-MM-DD đến YYYY-MM-DD"
                            dates = re.findall(r"\d{4}-\d{2}-\d{2}", display_date)
                            if len(dates) == 2:
                                record_start_dt = datetime.strptime(dates[0], "%Y-%m-%d").date()
                                record_end_dt = datetime.strptime(dates[1], "%Y-%m-%d").date()
                        else: # Dạng "YYYY-MM-DD ..."
                            date_part = display_date.split()[0]
                            record_start_dt = record_end_dt = datetime.strptime(date_part, "%Y-%m-%d").date()
                        
                        # Kiểm tra sự giao thoa giữa khoảng thời gian của bản ghi và bộ lọc
                        if record_start_dt and record_end_dt:
                            if record_start_dt <= filter_end_dt and record_end_dt >= filter_start_dt:
                                filtered_data.append(item)
                    except (ValueError, TypeError, IndexError):
                        continue # Bỏ qua nếu không thể phân tích ngày tháng
                data = filtered_data # Ghi đè dữ liệu gốc bằng dữ liệu đã lọc
        # <<< PHẦN SỬA LỖI KẾT THÚC >>>

        if not data:
            return jsonify([])

        # Định dạng dữ liệu trước khi trả về (giữ nguyên)
        for item in data:
            item["ApprovalDate1"] = get_formatted_approval_date(item.get("ApprovalDate1"))
            item["ApprovalDate2"] = get_formatted_approval_date(item.get("ApprovalDate2"))
            item["Status1"] = item.get("Status1", "")
            item["Status2"] = item.get("Status2", "")
            item["Note"] = item.get("LeaveNote", "")
            if item.get('CreationTime'):
                try:
                    timestamp = item['CreationTime'] if isinstance(item['CreationTime'], datetime) else datetime.fromisoformat(item['CreationTime'].replace('Z', '+00:00'))
                    item['CheckinTime'] = timestamp.astimezone(VN_TZ).strftime('%d/%m/%Y %H:%M:%S')
                except (ValueError, TypeError):
                    item['CheckinTime'] = ""
            else:
                item['CheckinTime'] = ""
            
            display_date = item.get('DisplayDate', "")
            if display_date:
                # Tìm và định dạng lại tất cả các chuỗi ngày tháng YYYY-MM-DD
                def reformat_date(match):
                    return datetime.strptime(match.group(0), "%Y-%m-%d").strftime("%d/%m/%Y")
                display_date = re.sub(r"\d{4}-\d{2}-\d{2}", reformat_date, display_date)
            
            item['CheckinDate'] = display_date # Gán lại giá trị đã định dạng
            tasks = item.get("Tasks", [])
            tasks_str = (", ".join(tasks) if isinstance(tasks, list) else str(tasks or "")).replace("Nghỉ phép: ", "")
            item['Tasks'] = item.get("Reason") or tasks_str

        return jsonify(data)
    except Exception as e:
        import traceback
        print(f"❌ Lỗi tại get_leaves: {e}")
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

# ---- HÀM HỖ TRỢ XUẤT THEO KHOẢNG ----
def get_export_date_range():
    start_date = request.args.get("startDate")
    end_date = request.args.get("endDate")
    month = request.args.get("month")
    year = request.args.get("year")

    if month and year:
        try:
            export_dt = datetime(int(year), int(month), 1)
            _, last_day = calendar.monthrange(int(year), int(month))
            start_date = export_dt.strftime("%Y-%m-%d")
            end_date = export_dt.replace(day=last_day).strftime("%Y-%m-%d")
        except:
            pass
    if not start_date or not end_date:
        return None, None
    return start_date, end_date

def get_export_filename(prefix, start_date, end_date, export_date_str):
    """
    Tạo tên file đẹp, dễ đọc cho mọi trường hợp:
    - 1 ngày:   Chấm công_30-10-2025_30-10-2025.xlsx
    - Nhiều ngày: Chấm công_01-10-2025_đến_30-10-2025_30-10-2025.xlsx
    - Tháng:    Chấm công_Tháng_10-2025_30-10-2025.xlsx
    """
    try:
        # Chuyển sang định dạng DD-MM-YYYY
        start_dt = datetime.strptime(start_date, "%Y-%m-%d")
        end_dt = datetime.strptime(end_date, "%Y-%m-%d")
        start_str = start_dt.strftime("%d-%m-%Y")
        end_str = end_dt.strftime("%d-%m-%Y")

        if start_dt == end_dt:
            # 1 ngày duy nhất
            file_prefix = start_str
        elif start_dt.replace(day=1) == end_dt.replace(day=1) and \
             start_dt.day == 1 and end_dt.day == (end_dt.replace(day=28) + timedelta(days=4)).day:
            # Cả tháng (từ ngày 1 đến cuối tháng)
            file_prefix = f"Tháng {end_dt.month:02d}-{end_dt.year}"
        else:
            # Khoảng ngày
            file_prefix = f"{start_str} đến {end_str}"

        return f"{prefix} {file_prefix}_{export_date_str}.xlsx"
    except:
        # Dự phòng nếu lỗi định dạng
        return f"{prefix} {start_date} đến {end_date}_{export_date_str}.xlsx"
    
@app.route("/api/export-excel", methods=["GET"])
def export_to_excel():
    try:
        email = request.args.get("email")
        admin = admins.find_one({"email": email})
        user = users.find_one({"email": email})
        if not admin and not user:
            return jsonify({"error": "Email không tồn tại"}), 403
        username = None if admin else user.get("username")

        # === 1. XÁC ĐỊNH KHOẢNG XUẤT ===
        start_date, end_date = get_export_date_range()
        if not start_date or not end_date:
            return jsonify({"error": "Thiếu thông tin ngày xuất"}), 400

        # === 2. XÁC ĐỊNH THÁNG ĐỂ TÍNH TỔNG GIỜ ===
        end_dt = datetime.strptime(end_date, "%Y-%m-%d")
        query_start_dt = end_dt.replace(day=1)  # NGÀY 1 CỦA THÁNG CUỐI
        query_start = query_start_dt.strftime("%Y-%m-%d")

        # === 3. LẤY DỮ LIỆU TỪ ĐẦU THÁNG ĐẾN end_date ===
        search = request.args.get("search", "").strip()
        query = build_attendance_query("custom", query_start, end_date, search, username=username)
        data = list(collection.find(query, {"_id": 0}))

        # === 4. TÍNH GIỜ LÀM THEO NGÀY (CỘT 14) ===
        daily_hours_map = {}
        emp_data = {}
        for rec in data:
            emp_id = rec.get("EmployeeId")
            if emp_id:
                emp_data.setdefault(emp_id, []).append(rec)

        for emp_id, records in emp_data.items():
            daily_groups = {}
            for rec in records:
                date_str = rec.get("CheckinDate")
                if date_str:
                    daily_groups.setdefault(date_str, []).append(rec)
            for date_str, day_records in daily_groups.items():
                checkins = [r["Timestamp"] for r in day_records if r.get("CheckType") == "checkin" and r.get("Timestamp")]
                checkouts = [r["Timestamp"] for r in day_records if r.get("CheckType") == "checkout" and r.get("Timestamp")]
                checkins = [t if isinstance(t, datetime) else datetime.strptime(t, "%Y-%m-%d %H:%M:%S") for t in checkins]
                checkouts = [t if isinstance(t, datetime) else datetime.strptime(t, "%Y-%m-%d %H:%M:%S") for t in checkouts]
                checkins = [t for t in checkins if isinstance(t, datetime)]
                checkouts = [t for t in checkouts if isinstance(t, datetime)]
                daily_seconds = 0
                if checkins and checkouts:
                    first_in = min(checkins)
                    last_out = max(checkouts)
                    if last_out > first_in:
                        daily_seconds = (last_out - first_in).total_seconds()
                daily_hours_map[(emp_id, date_str)] = daily_seconds

        # === 5. TÍNH TỔNG GIỜ TỪ ĐẦU THÁNG (CỘT 15) ===
        monthly_hours_map = {}
        for emp_id, records in emp_data.items():
            # Sắp xếp theo ngày
            sorted_records = sorted(records, key=lambda r: datetime.strptime(r.get("CheckinDate", "01/01/1900"), "%d/%m/%Y"))
            running_total = 0
            for rec in sorted_records:
                date_str = rec.get("CheckinDate")
                daily_sec = daily_hours_map.get((emp_id, date_str), 0)
                running_total += daily_sec
                h, rem = divmod(running_total, 3600)
                m, s = divmod(rem, 60)
                monthly_hours_map[(emp_id, date_str)] = f"{int(h)}h {int(m)}m {int(s)}s" if running_total > 0 else ""

        # === 6. GHI VÀO EXCEL ===
        grouped = {}
        for d in data:
            if d.get("CheckinDate", "").startswith(end_dt.strftime("%d/%m/%Y").split("/")[0]):  # Chỉ lấy ngày trong khoảng xuất
                key = (d.get("EmployeeId", ""), d.get("EmployeeName", ""), d.get("CheckinDate"))
                grouped.setdefault(key, []).append(d)

        template_path = "templates/Copy of Form chấm công.xlsx"
        wb = load_workbook(template_path)
        ws = wb.active
        border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        start_row = 2

        for i, ((emp_id, emp_name, date_str), records) in enumerate(grouped.items()):
            row = start_row + i
            ws.cell(row=row, column=1, value=emp_id)
            ws.cell(row=row, column=2, value=emp_name)
            ws.cell(row=row, column=3, value=date_str)

            # Cột 14: Giờ ngày
            daily_sec = daily_hours_map.get((emp_id, date_str), 0)
            h_d, rem_d = divmod(daily_sec, 3600)
            m_d, s_d = divmod(rem_d, 60)
            ws.cell(row=row, column=14, value=f"{int(h_d)}h {int(m_d)}m {int(s_d)}s" if daily_sec > 0 else "")

            # Cột 15: Tổng giờ tháng
            ws.cell(row=row, column=15, value=monthly_hours_map.get((emp_id, date_str), ""))

            # Ghi checkin/checkout
            checkin_counter = 0
            for rec in sorted(records, key=lambda x: x.get("Timestamp") or datetime.min):
                if rec.get("CheckType") == "checkin" and checkin_counter < 9:
                    time_str = rec["Timestamp"]
                    if isinstance(time_str, str):
                        time_str = datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S")
                    time_str = time_str.astimezone(VN_TZ).strftime("%H:%M:%S") if time_str else ""
                    tasks_str = ", ".join(rec.get("Tasks", [])) if isinstance(rec.get("Tasks"), list) else str(rec.get("Tasks", ""))
                    cell_value = "; ".join(filter(None, [time_str, rec.get("ProjectId", ""), tasks_str, rec.get("Address", ""), rec.get("CheckinNote", "")]))
                    ws.cell(row=row, column=4 + checkin_counter, value=cell_value)
                    checkin_counter += 1
                elif rec.get("CheckType") == "checkout":
                    time_str = rec["Timestamp"]
                    if isinstance(time_str, str):
                        time_str = datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S")
                    time_str = time_str.astimezone(VN_TZ).strftime("%H:%M:%S") if time_str else ""
                    tasks_str = ", ".join(rec.get("Tasks", [])) if isinstance(rec.get("Tasks"), list) else str(rec.get("Tasks", ""))
                    cell_value = "; ".join(filter(None, [time_str, rec.get("ProjectId", ""), tasks_str, rec.get("Address", ""), rec.get("CheckinNote", "")]))
                    ws.cell(row=row, column=13, value=cell_value)

            for col in range(1, 16):
                cell = ws.cell(row=row, column=col)
                cell.border = border
                cell.alignment = align_left

        export_date_str = datetime.now(VN_TZ).strftime('%d-%m-%Y')
        filename = get_export_filename("Chấm công", start_date, end_date, export_date_str)
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return send_file(output, as_attachment=True, download_name=filename,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        print(f"Lỗi export: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

# ---- API xuất Excel cho nghỉ phép ----
@app.route("/api/export-leaves-excel", methods=["GET"])
def export_leaves_to_excel():
    try:
        email = request.args.get("email")
        if not email:
            return jsonify({"error": "Thiếu email"}), 400
        admin = admins.find_one({"email": email})
        user_doc = users.find_one({"email": email})
        username = None if admin else user_doc.get("username") if user_doc else None
        if not admin and not user_doc:
            return jsonify({"error": "Email không tồn tại"}), 403

        start_date, end_date = get_export_date_range()
        if not start_date or not end_date:
            return jsonify({"error": "Thiếu thông tin ngày xuất"}), 400

        # --- Xác định export_year, export_month ---
        try:
            start_dt = datetime.strptime(start_date, "%Y-%m-%d")
            export_year = start_dt.year
            export_month = start_dt.month
        except:
            today = datetime.now(VN_TZ)
            export_year = today.year
            export_month = today.month

        regex_leave = re.compile("Nghỉ phép", re.IGNORECASE)
        conditions = [{"$or": [{"Tasks": regex_leave}, {"Reason": {"$exists": True}}]}]
        search = request.args.get("search", "").strip()
        if search:
            regex = re.compile(search, re.IGNORECASE)
            conditions.append({"$or": [{"EmployeeId": regex}, {"EmployeeName": regex}]})
        if username:
            conditions.append({"EmployeeName": username})
        query = {"$and": conditions}
        all_leaves_data = list(collection.find(query, {"_id": 0}))

        template_path = "templates/Copy of Form nghỉ phép.xlsx"
        wb = load_workbook(template_path)
        ws = wb.active
        ws['A1'] = "Mã NV"; ws['B1'] = "Tên NV"; ws['C1'] = "Ngày Nghỉ"; ws['D1'] = "Số ngày nghỉ"
        ws['E1'] = "Ngày tạo đơn"; ws['F1'] = "Lý do"; ws['G1'] = "Ngày Duyệt/Từ chối Lần đầu"
        ws['H1'] = "Trạng thái Lần đầu"; ws['I1'] = "Ngày Duyệt/Từ chối Lần cuối"
        ws['J1'] = "Trạng thái Lần cuối"; ws['K1'] = "Ghi chú"

        border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        current_row = 2

        for rec in all_leaves_data:
            leave_days, is_overlap = calculate_leave_days_for_month(rec, export_year, export_month)
            if not is_overlap:
                continue

            # --- XỬ LÝ HIỂN THỊ NGÀY NGHỈ ĐƠN ---
            start_dt = end_dt = None
            session = rec.get("Session", "").strip()

            if rec.get("StartDate"):
                try: start_dt = datetime.strptime(rec["StartDate"], "%Y-%m-%d")
                except: pass
            if rec.get("EndDate"):
                try: end_dt = datetime.strptime(rec["EndDate"], "%Y-%m-%d")
                except: pass
            if rec.get("LeaveDate") and not start_dt:
                try: start_dt = end_dt = datetime.strptime(rec["LeaveDate"], "%Y-%m-%d")
                except: pass

            if start_dt and end_dt and start_dt == end_dt:
                day_str = start_dt.strftime("%d/%m/%Y")
                if session:
                    day_str += f" ({session})"
                display_date = day_str
            elif start_dt and end_dt:
                display_date = f"Từ {start_dt.strftime('%d/%m/%Y')} đến {end_dt.strftime('%d/%m/%Y')}"
            else:
                display_date = rec.get("DisplayDate", "")
                if display_date:
                    display_date = re.sub(r"\d{4}-\d{2}-\d{2}", lambda m: datetime.strptime(m.group(0), "%Y-%m-%d").strftime("%d/%m/%Y"), display_date)

            # --- Ghi dữ liệu ---
            ws.cell(row=current_row, column=1, value=rec.get("EmployeeId", ""))
            ws.cell(row=current_row, column=2, value=rec.get("EmployeeName", ""))
            ws.cell(row=current_row, column=3, value=display_date)
            ws.cell(row=current_row, column=4, value=leave_days)

            timestamp_str = ""
            if rec.get("CreationTime"):
                try:
                    ct = rec['CreationTime']
                    dt = ct if isinstance(ct, datetime) else datetime.fromisoformat(ct.replace('Z', '+00:00'))
                    timestamp_str = dt.astimezone(VN_TZ).strftime('%d/%m/%Y %H:%M:%S')
                except:
                    timestamp_str = str(rec.get("CreationTime"))
            ws.cell(row=current_row, column=5, value=timestamp_str)

            tasks_str = (", ".join(rec.get("Tasks", [])) if isinstance(rec.get("Tasks"), list) else str(rec.get("Tasks", ""))).replace("Nghỉ phép: ", "")
            ws.cell(row=current_row, column=6, value=rec.get("Reason") or tasks_str)
            ws.cell(row=current_row, column=7, value=get_formatted_approval_date(rec.get("ApprovalDate1")))
            ws.cell(row=current_row, column=8, value=rec.get("Status1", ""))
            ws.cell(row=current_row, column=9, value=get_formatted_approval_date(rec.get("ApprovalDate2")))
            ws.cell(row=current_row, column=10, value=rec.get("Status2", ""))
            ws.cell(row=current_row, column=11, value=rec.get("LeaveNote", ""))

            for col_idx in range(1, 12):
                cell = ws.cell(row=current_row, column=col_idx)
                cell.border = border
                cell.alignment = align_left
            current_row += 1

        # --- Tên file ---
        export_date_str = datetime.now(VN_TZ).strftime('%d-%m-%Y')
        filename = get_export_filename("Nghỉ phép", start_date, end_date, export_date_str)

        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return send_file(output, as_attachment=True, download_name=filename,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        print(f"Lỗi export leaves: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

@app.route("/api/export-combined-excel", methods=["GET"])
def export_combined_to_excel():
    try:
        email = request.args.get("email")
        if not email:
            return jsonify({"error": "Thiếu email"}), 400
        admin = admins.find_one({"email": email})
        user = users.find_one({"email": email})
        if not admin and not user:
            return jsonify({"error": "Email không tồn tại"}), 403
        username = None if admin else user.get("username")

        # === 1. XÁC ĐỊNH K0ẢNG XUẤT ===
        start_date, end_date = get_export_date_range()
        if not start_date or not end_date:
            return jsonify({"error": "Thiếu thông tin ngày xuất"}), 400

        # === 2. XÁC ĐỊNH NGÀY 1 CỦA THÁNG CUỐI ĐỂ TÍNH TỔNG GIỜ ===
        end_dt = datetime.strptime(end_date, "%Y-%m-%d")
        query_start_dt = end_dt.replace(day=1)  # Luôn lấy từ ngày 1 của tháng
        query_start = query_start_dt.strftime("%Y-%m-%d")

        # === 3. LẤY DỮ LIỆU TỪ ĐẦU THÁNG ĐẾN end_date ===
        search = request.args.get("search", "").strip()
        attendance_query = build_attendance_query("custom", query_start, end_date, search, username=username)
        regex_leave = re.compile("Nghỉ phép", re.IGNORECASE)
        leave_conditions = [{"$or": [{"Tasks": regex_leave}, {"Reason": {"$exists": True}}]}]
        if search:
            regex = re.compile(search, re.IGNORECASE)
            leave_conditions.append({"$or": [{"EmployeeId": regex}, {"EmployeeName": regex}]})
        if username:
            leave_conditions.append({"EmployeeName": username})
        leave_query = {"$and": leave_conditions}

        attendance_data = list(collection.find(attendance_query, {"_id": 0}))
        leave_data = list(collection.find(leave_query, {"_id": 0}))

        # === 4. TẠO WORKBOOK ===
        template_path = "templates/Form kết hợp.xlsx"
        wb = load_workbook(template_path)
        border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

        # ================= SHEET: ĐIỂM DANH =================
        ws_att = wb["Điểm danh"]

        # --- Tính giờ làm theo ngày (Cột 14) ---
        daily_hours_map = {}
        emp_data = {}
        for rec in attendance_data:
            emp_id = rec.get("EmployeeId")
            if emp_id:
                emp_data.setdefault(emp_id, []).append(rec)

        for emp_id, records in emp_data.items():
            daily_groups = {}
            for rec in records:
                date_str = rec.get("CheckinDate")
                if date_str:
                    daily_groups.setdefault(date_str, []).append(rec)

            for date_str, day_records in daily_groups.items():
                checkins = []
                checkouts = []
                for r in day_records:
                    ts = r.get("Timestamp")
                    if not ts:
                        continue
                    try:
                        if isinstance(ts, str):
                            ts = datetime.strptime(ts, "%Y-%m-%d %H:%M:%S")
                        if r.get("CheckType") == "checkin":
                            checkins.append(ts)
                        elif r.get("CheckType") == "checkout":
                            checkouts.append(ts)
                    except:
                        continue
                checkins = [t for t in checkins if isinstance(t, datetime)]
                checkouts = [t for t in checkouts if isinstance(t, datetime)]
                daily_seconds = 0
                if checkins and checkouts:
                    first_in = min(checkins)
                    last_out = max(checkouts)
                    if last_out > first_in:
                        daily_seconds = (last_out - first_in).total_seconds()
                daily_hours_map[(emp_id, date_str)] = daily_seconds

        # --- Tính tổng giờ tích lũy từ đầu tháng (Cột 15) ---
        monthly_hours_map = {}
        for emp_id, records in emp_data.items():
            # Sắp xếp theo ngày (CheckinDate)
            sorted_records = sorted(
                records,
                key=lambda r: datetime.strptime(r.get("CheckinDate", "01/01/1900"), "%d/%m/%Y")
            )
            running_total = 0
            for rec in sorted_records:
                date_str = rec.get("CheckinDate")
                daily_sec = daily_hours_map.get((emp_id, date_str), 0)
                running_total += daily_sec
                h, rem = divmod(running_total, 3600)
                m, s = divmod(rem, 60)
                monthly_hours_map[(emp_id, date_str)] = f"{int(h)}h {int(m)}m {int(s)}s" if running_total > 0 else ""

        # --- Ghi dữ liệu điểm danh (chỉ ngày trong khoảng xuất) ---
        grouped = {}
        for d in attendance_data:
            date_str = d.get("CheckinDate", "")
            # Chỉ lấy các ngày nằm trong khoảng start_date → end_date
            try:
                rec_date = datetime.strptime(date_str, "%d/%m/%Y")
                start_dt = datetime.strptime(start_date, "%Y-%m-%d")
                end_dt = datetime.strptime(end_date, "%Y-%m-%d")
                if start_dt <= rec_date <= end_dt:
                    key = (d.get("EmployeeId", ""), d.get("EmployeeName", ""), date_str)
                    grouped.setdefault(key, []).append(d)
            except:
                continue

        start_row = 2
        for i, ((emp_id, emp_name, date_str), records) in enumerate(grouped.items()):
            row = start_row + i
            ws_att.cell(row=row, column=1, value=emp_id)
            ws_att.cell(row=row, column=2, value=emp_name)
            ws_att.cell(row=row, column=3, value=date_str)

            # Cột 14: Giờ làm trong ngày
            daily_sec = daily_hours_map.get((emp_id, date_str), 0)
            h_d, rem_d = divmod(daily_sec, 3600)
            m_d, s_d = divmod(rem_d, 60)
            ws_att.cell(row=row, column=14, value=f"{int(h_d)}h {int(m_d)}m {int(s_d)}s" if daily_sec > 0 else "")

            # Cột 15: Tổng giờ từ đầu tháng đến ngày đó
            ws_att.cell(row=row, column=15, value=monthly_hours_map.get((emp_id, date_str), ""))

            # Ghi checkin/checkout
            checkin_counter = 0
            for rec in sorted(records, key=lambda x: x.get("Timestamp") or datetime.min):
                if rec.get("CheckType") == "checkin" and checkin_counter < 9:
                    time_str = ""
                    if rec.get("Timestamp"):
                        try:
                            ts = rec["Timestamp"]
                            if isinstance(ts, str):
                                ts = datetime.strptime(ts, "%Y-%m-%d %H:%M:%S")
                            time_str = ts.astimezone(VN_TZ).strftime("%H:%M:%S")
                        except:
                            time_str = ""
                    tasks_str = ", ".join(rec.get("Tasks", [])) if isinstance(rec.get("Tasks"), list) else str(rec.get("Tasks", ""))
                    cell_value = "; ".join(filter(None, [time_str, rec.get("ProjectId", ""), tasks_str, rec.get("Address", ""), rec.get("CheckinNote", "")]))
                    ws_att.cell(row=row, column=4 + checkin_counter, value=cell_value)
                    checkin_counter += 1
                elif rec.get("CheckType") == "checkout":
                    time_str = ""
                    if rec.get("Timestamp"):
                        try:
                            ts = rec["Timestamp"]
                            if isinstance(ts, str):
                                ts = datetime.strptime(ts, "%Y-%m-%d %H:%M:%S")
                            time_str = ts.astimezone(VN_TZ).strftime("%H:%M:%S")
                        except:
                            time_str = ""
                    tasks_str = ", ".join(rec.get("Tasks", [])) if isinstance(rec.get("Tasks"), list) else str(rec.get("Tasks", ""))
                    cell_value = "; ".join(filter(None, [time_str, rec.get("ProjectId", ""), tasks_str, rec.get("Address", ""), rec.get("CheckinNote", "")]))
                    ws_att.cell(row=row, column=13, value=cell_value)

            for col in range(1, 16):
                cell = ws_att.cell(row=row, column=col)
                cell.border = border
                cell.alignment = align_left

        # ================= SHEET: NGHỈ PHÉP =================
        ws_leave = wb["Nghỉ phép"]
        ws_leave['A1'] = "Mã NV"; ws_leave['B1'] = "Tên NV"; ws_leave['C1'] = "Ngày Nghỉ"; ws_leave['D1'] = "Số ngày nghỉ"
        ws_leave['E1'] = "Ngày tạo đơn"; ws_leave['F1'] = "Lý do"; ws_leave['G1'] = "Ngày Duyệt/Từ chối Lần đầu"
        ws_leave['H1'] = "Trạng thái Lần đầu"; ws_leave['I1'] = "Ngày Duyệt/Từ chối Lần cuối"
        ws_leave['J1'] = "Trạng thái Lần cuối"; ws_leave['K1'] = "Ghi chú"

        current_row_leave = 2
        export_year = end_dt.year
        export_month = end_dt.month

        for rec in leave_data:
            leave_days, is_overlap = calculate_leave_days_for_month(rec, export_year, export_month)
            if not is_overlap:
                continue

            # Xử lý hiển thị ngày nghỉ
            start_dt = end_dt = None
            session = rec.get("Session", "").strip()
            if rec.get("StartDate"):
                try: start_dt = datetime.strptime(rec["StartDate"], "%Y-%m-%d")
                except: pass
            if rec.get("EndDate"):
                try: end_dt = datetime.strptime(rec["EndDate"], "%Y-%m-%d")
                except: pass
            if rec.get("LeaveDate") and not start_dt:
                try: start_dt = end_dt = datetime.strptime(rec["LeaveDate"], "%Y-%m-%d")
                except: pass

            if start_dt and end_dt and start_dt == end_dt:
                day_str = start_dt.strftime("%d/%m/%Y")
                if session:
                    day_str += f" ({session})"
                display_date = day_str
            elif start_dt and end_dt:
                display_date = f"Từ {start_dt.strftime('%d/%m/%Y')} đến {end_dt.strftime('%d/%m/%Y')}"
            else:
                display_date = rec.get("DisplayDate", "")
                if display_date:
                    display_date = re.sub(r"\d{4}-\d{2}-\d{2}", lambda m: datetime.strptime(m.group(0), "%Y-%m-%d").strftime("%d/%m/%Y"), display_date)

            # Ghi dữ liệu
            ws_leave.cell(row=current_row_leave, column=1, value=rec.get("EmployeeId", ""))
            ws_leave.cell(row=current_row_leave, column=2, value=rec.get("EmployeeName", ""))
            ws_leave.cell(row=current_row_leave, column=3, value=display_date)
            ws_leave.cell(row=current_row_leave, column=4, value=leave_days)

            timestamp_str = ""
            if rec.get("CreationTime"):
                try:
                    ct = rec['CreationTime']
                    dt = ct if isinstance(ct, datetime) else datetime.fromisoformat(ct.replace('Z', '+00:00'))
                    timestamp_str = dt.astimezone(VN_TZ).strftime('%d/%m/%Y %H:%M:%S')
                except:
                    timestamp_str = str(rec.get("CreationTime"))
            ws_leave.cell(row=current_row_leave, column=5, value=timestamp_str)

            tasks_str = (", ".join(rec.get("Tasks", [])) if isinstance(rec.get("Tasks"), list) else str(rec.get("Tasks", ""))).replace("Nghỉ phép: ", "")
            ws_leave.cell(row=current_row_leave, column=6, value=rec.get("Reason") or tasks_str)
            ws_leave.cell(row=current_row_leave, column=7, value=get_formatted_approval_date(rec.get("ApprovalDate1")))
            ws_leave.cell(row=current_row_leave, column=8, value=rec.get("Status1", ""))
            ws_leave.cell(row=current_row_leave, column=9, value=get_formatted_approval_date(rec.get("ApprovalDate2")))
            ws_leave.cell(row=current_row_leave, column=10, value=rec.get("Status2", ""))
            ws_leave.cell(row=current_row_leave, column=11, value=rec.get("LeaveNote", ""))

            for col in range(1, 12):
                cell = ws_leave.cell(row=current_row_leave, column=col)
                cell.border = border
                cell.alignment = align_left

            current_row_leave += 1

        # === TẠO TÊN FILE ===
        export_date_str = datetime.now(VN_TZ).strftime('%d-%m-%Y')
        filename = get_export_filename("Báo cáo tổng hợp", start_date, end_date, export_date_str)

        # === XUẤT FILE ===
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        print(f"Lỗi export combined: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

# if __name__ == "__main__":
#     app.run(host="0.0.0.0", port=5000, debug=False)







