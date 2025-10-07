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
import smtplib
import threading
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

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

# ---- SMTP Config for Email ----
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))
SMTP_USERNAME = os.getenv("SMTP_USERNAME", "sun.automation.sys@gmail.com")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "ihgzxunefndizeub")
APP_URL = os.getenv("APP_URL", "https://read-gps-data.vercel.app")

# ---- Kết nối MongoDB ----
client = MongoClient(MONGO_URI)
db = client[DB_NAME]

# Các collection sử dụng
admins = db["admins"]
collection = db["alt_checkins"]
reset_tokens = db["reset_tokens"]

# ---- Quản lý kết nối SMTP ----
smtp_server = None

def get_smtp_server(timeout=10):
    global smtp_server
    if smtp_server is None:
        try:
            smtp_server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=timeout)
            smtp_server.starttls()
            smtp_server.login(SMTP_USERNAME, SMTP_PASSWORD)
            print("DEBUG: SMTP connection established")
        except Exception as e:
            print(f"❌ Error establishing SMTP connection: {e}")
            smtp_server = None
            raise
    return smtp_server

def close_smtp_server():
    global smtp_server
    if smtp_server is not None:
        try:
            smtp_server.quit()
            print("DEBUG: SMTP connection closed")
        except Exception as e:
            print(f"❌ Error closing SMTP connection: {e}")
        smtp_server = None

# ---- Gửi email bất đồng bộ ----
def send_reset_email_async(email, token, retries=3, timeout=10):
    def send_email():
        msg = MIMEMultipart()
        msg["From"] = SMTP_USERNAME
        msg["To"] = email
        msg["Subject"] = "Đặt lại mật khẩu"
        reset_link = f"{APP_URL}/reset-password?token={token}"
        body = f"""
        Xin chào,
        Bạn đã yêu cầu đặt lại mật khẩu. Vui lòng nhấp vào liên kết dưới đây để đặt lại mật khẩu:
        {reset_link}
        Liên kết này sẽ hết hạn sau 1 giờ. Nếu bạn không yêu cầu đặt lại mật khẩu, vui lòng bỏ qua email này.
        Trân trọng,
        Đội ngũ hỗ trợ
        """
        msg.attach(MIMEText(body, "plain"))
        for attempt in range(retries):
            try:
                start_time = time.time()
                server = get_smtp_server(timeout=timeout)
                server.send_message(msg)
                end_time = time.time()
                print(f"DEBUG: Reset email sent to {email} in {end_time - start_time:.2f} seconds")
                return True
            except Exception as e:
                print(f"❌ Attempt {attempt + 1} failed: {e}")
                global smtp_server
                smtp_server = None  # Reset kết nối nếu lỗi
                if attempt < retries - 1:
                    time.sleep(2)  # Chờ trước khi thử lại
                continue
        print(f"❌ Failed to send reset email to {email} after {retries} attempts")
        return False

    thread = threading.Thread(target=send_email)
    thread.start()

# ---- Trang chủ (đăng nhập chính) ----
@app.route("/")
def index():
    success = request.args.get("success")
    return render_template("index.html", success=success)

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
    if not admin or not check_password_hash(admin.get("password", ""), password):
        return jsonify({"success": False, "message": "🚫 Email hoặc mật khẩu không đúng!"}), 401
    return jsonify({
        "success": True,
        "message": "✅ Đăng nhập thành công",
        "username": admin["username"],
        "email": admin["email"]
    })

# ---- Gửi email quên mật khẩu ----
@app.route("/forgot-password", methods=["GET", "POST"])
def forgot_password():
    if request.method == "GET":
        return """
        <!DOCTYPE html>
        <html lang="vi">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Quên mật khẩu</title>
            <style>
                body { font-family: Arial, sans-serif; background: #f4f6f9; margin: 0; padding: 20px; }
                .container { max-width: 400px; margin: 100px auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
                input { width: 100%; padding: 10px; margin: 10px 0; box-sizing: border-box; border: 1px solid #ddd; border-radius: 4px; }
                button { background: #28a745; color: white; padding: 12px; width: 100%; border: none; border-radius: 4px; cursor: pointer; font-size: 16px; }
                button:hover { background: #218838; }
                .success { color: #28a745; text-align: center; }
            </style>
        </head>
        <body>
            <div class="container">
                <h2>🔒 Quên mật khẩu</h2>
                <form method="POST">
                    <input type="email" name="email" placeholder="Email" required>
                    <button type="submit">Gửi liên kết đặt lại mật khẩu</button>
                    <a href="/">Quay về trang chủ</a>
                </form>
            </div>
        </body>
        </html>
        """
    if request.method == "POST":
        email = request.form.get("email")
        if not email:
            return jsonify({"success": False, "message": "❌ Vui lòng nhập email"}), 400
        
        admin = admins.find_one({"email": email})
        if not admin:
            return jsonify({"success": False, "message": "🚫 Email không tồn tại!"}), 404
        
        # Generate reset token
        token = secrets.token_urlsafe(32)
        expiry = datetime.now(VN_TZ) + timedelta(hours=1)
        
        # Store token in reset_tokens collection
        reset_tokens.delete_one({"email": email})
        reset_tokens.insert_one({
            "email": email,
            "token": token,
            "expiry": expiry
        })
        
        # Gửi email bất đồng bộ
        send_reset_email_async(email, token)
        
        return """
        <!DOCTYPE html>
        <html lang="vi">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Quên mật khẩu</title>
            <style>
                body { font-family: Arial, sans-serif; background: #f4f6f9; margin: 0; padding: 20px; }
                .container { max-width: 400px; margin: 100px auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); text-align: center; }
                .success { color: #28a745; font-size: 18px; margin-bottom: 20px; }
                button { background: #28a745; color: white; padding: 12px 20px; border: none; border-radius: 4px; cursor: pointer; font-size: 16px; }
                button:hover { background: #218838; }
            </style>
        </head>
        <body>
            <div class="container">
                <div class="success">✅ Liên kết đặt lại mật khẩu đang được gửi đến email của bạn!</div>
                <button onclick="window.location.href='/'">Quay về trang chủ</button>
            </div>
        </body>
        </html>
        """

# ---- Đặt lại mật khẩu ----
@app.route("/reset-password", methods=["GET", "POST"])
def reset_password():
    if request.method == "GET":
        token = request.args.get("token")
        if not token:
            return jsonify({"success": False, "message": "❌ Thiếu token"}), 400
        
        token_data = reset_tokens.find_one({"token": token})
        if not token_data:
            return jsonify({"success": False, "message": "🚫 Token không hợp lệ hoặc đã hết hạn"}), 400
        
        if token_data["expiry"] < datetime.now(VN_TZ):
            reset_tokens.delete_one({"token": token})
            return jsonify({"success": False, "message": "🚫 Token đã hết hạn"}), 400
        
        return """
        <!DOCTYPE html>
        <html lang="vi">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Đặt lại mật khẩu</title>
            <style>
                body { font-family: Arial, sans-serif; background: #f4f6f9; margin: 0; padding: 20px; }
                .container { max-width: 400px; margin: 100px auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
                input { width: 100%; padding: 10px; margin: 10px 0; box-sizing: border-box; border: 1px solid #ddd; border-radius: 4px; }
                button { background: #28a745; color: white; padding: 12px; width: 100%; border: none; border-radius: 4px; cursor: pointer; font-size: 16px; }
                button:hover { background: #218838; }
            </style>
        </head>
        <body>
            <div class="container">
                <h2>🔒 Đặt lại mật khẩu</h2>
                <form method="POST">
                    <input type="hidden" name="token" value="{}">
                    <input type="password" name="new_password" placeholder="Mật khẩu mới" required>
                    <input type="password" name="confirm_password" placeholder="Xác nhận mật khẩu" required>
                    <button type="submit">Cập nhật mật khẩu</button>
                    <a href="/">Quay về trang chủ</a>
                </form>
            </div>
        </body>
        </html>
        """.format(token)
    
    if request.method == "POST":
        token = request.form.get("token")
        new_password = request.form.get("new_password")
        confirm_password = request.form.get("confirm_password")
        
        if not token or not new_password or not confirm_password:
            return jsonify({"success": False, "message": "❌ Vui lòng điền đầy đủ thông tin"}), 400
        
        if new_password != confirm_password:
            return jsonify({"success": False, "message": "❌ Mật khẩu xác nhận không khớp"}), 400
        
        token_data = reset_tokens.find_one({"token": token})
        if not token_data:
            return jsonify({"success": False, "message": "🚫 Token không hợp lệ hoặc đã hết hạn"}), 400
        
        if token_data["expiry"] < datetime.now(VN_TZ):
            reset_tokens.delete_one({"token": token})
            return jsonify({"success": False, "message": "🚫 Token đã hết hạn"}), 400
        
        email = token_data["email"]
        admin = admins.find_one({"email": email})
        if not admin:
            return jsonify({"success": False, "message": "🚫 Email không tồn tại!"}), 404
        
        # Update password
        hashed_pw = generate_password_hash(new_password)
        admins.update_one({"email": email}, {"$set": {"password": hashed_pw}})
        
        # Remove used token
        reset_tokens.delete_one({"token": token})
        
        return redirect(url_for("index", success=1))

# ---- Build attendance query (không thay đổi) ----
def build_attendance_query(filter_type, start_date, end_date, search):
    today = datetime.now(VN_TZ)
    regex_leave = re.compile("Nghỉ phép", re.IGNORECASE)
    conditions = []
    date_filter = {}
    if filter_type == "custom" and start_date and end_date:
        date_filter = {"CheckinDate": {"$gte": start_date, "$lte": end_date}}
    elif filter_type == "hôm nay":
        date_filter = {"CheckinDate": today.strftime("%Y-%m-%d")}
    elif filter_type == "tuần":
        start = (today - timedelta(days=today.weekday())).strftime("%Y-%m-%d")
        end = (today + timedelta(days=6 - today.weekday())).strftime("%Y-%m-%d")
        date_filter = {"CheckinDate": {"$gte": start, "$lte": end}}
    elif filter_type == "tháng":
        start = today.replace(day=1).strftime("%Y-%m-%d")
        end = today.replace(day=calendar.monthrange(today.year, today.month)[1]).strftime("%Y-%m-%d")
        date_filter = {"CheckinDate": {"$gte": start, "$lte": end}}
    elif filter_type == "năm":
        date_filter = {"CheckinDate": {"$regex": f"^{today.year}"}}
    if date_filter:
        conditions.append(date_filter)
    not_leave_or = {
        "$or": [
            {"Tasks": {"$not": regex_leave}},
            {"Tasks": {"$exists": False}},
            {"Tasks": None}
        ]
    }
    conditions.append(not_leave_or)
    if search:
        regex = re.compile(search, re.IGNORECASE)
        search_or = {
            "$or": [
                {"EmployeeId": {"$regex": regex}},
                {"EmployeeName": {"$regex": regex}}
            ]
        }
        conditions.append(search_or)
    if len(conditions) == 1:
        return conditions[0]
    else:
        return {"$and": conditions}

# ---- Build leave query (không thay đổi) ----
def build_leave_query(filter_type, start_date, end_date, search):
    today = datetime.now(VN_TZ)
    regex_leave = re.compile("Nghỉ phép", re.IGNORECASE)
    conditions = []
    leave_or = {
        "$or": [
            {"Tasks": {"$regex": regex_leave}},
            {"Status": {"$regex": regex_leave}},
            {"OtherNote": {"$regex": regex_leave}}
        ]
    }
    conditions.append(leave_or)
    date_filter = {}
    if filter_type == "custom" and start_date and end_date:
        start_dt = datetime.strptime(start_date, "%Y-%m-%d").strftime("%d/%m/%Y")
        end_dt = datetime.strptime(end_date, "%Y-%m-%d").strftime("%d/%m/%Y")
        date_filter = {
            "CheckinTime": {
                "$gte": f"{start_dt} 00:00:00",
                "$lte": f"{end_dt} 23:59:59"
            }
        }
    elif filter_type == "hôm nay":
        today_str = today.strftime("%d/%m/%Y")
        date_filter = {
            "CheckinTime": {
                "$gte": f"{today_str} 00:00:00",
                "$lte": f"{today_str} 23:59:59"
            }
        }
    elif filter_type == "tuần":
        week_start = (today - timedelta(days=today.weekday())).strftime("%d/%m/%Y")
        week_end = (today + timedelta(days=6 - today.weekday())).strftime("%d/%m/%Y")
        date_filter = {
            "CheckinTime": {
                "$gte": f"{week_start} 00:00:00",
                "$lte": f"{week_end} 23:59:59"
            }
        }
    elif filter_type == "tháng":
        month = f"{today.month:02d}"
        year = str(today.year)
        start_day = "01"
        end_day = str(calendar.monthrange(today.year, today.month)[1])
        date_filter = {
            "CheckinTime": {
                "$gte": f"{start_day}/{month}/{year} 00:00:00",
                "$lte": f"{end_day}/{month}/{year} 23:59:59"
            }
        }
    elif filter_type == "năm":
        year = str(today.year)
        date_filter = {
            "CheckinTime": {
                "$gte": f"01/01/{year} 00:00:00",
                "$lte": f"31/12/{year} 23:59:59"
            }
        }
    if date_filter:
        conditions.append(date_filter)
    if search:
        regex = re.compile(search, re.IGNORECASE)
        search_or = {
            "$or": [
                {"EmployeeId": {"$regex": regex}},
                {"EmployeeName": {"$regex": regex}}
            ]
        }
        conditions.append(search_or)
    if len(conditions) == 1:
        return conditions[0]
    else:
        return {"$and": conditions}

# ---- API lấy dữ liệu chấm công (không thay đổi) ----
@app.route("/api/attendances", methods=["GET"])
def get_attendances():
    try:
        email = request.args.get("email")
        if not email:
            return jsonify({"error": "❌ Thiếu email"}), 400
        admin = admins.find_one({"email": email}, {"_id": 0, "username": 1})
        if not admin:
            return jsonify({"error": "🚫 Email không hợp lệ (không có quyền truy cập)"}), 403
        filter_type = request.args.get("filter", "hôm nay").lower()
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()
        query = build_attendance_query(filter_type, start_date, end_date, search)
        data = list(collection.find(query, {"_id": 0}))
        for item in data:
            ghi_chu_parts = []
            if item.get('ProjectId'):
                ghi_chu_parts.append(f"Project: {item['ProjectId']}")
            if item.get('Tasks'):
                tasks_str = ', '.join(item['Tasks']) if isinstance(item['Tasks'], list) else str(item['Tasks'])
                ghi_chu_parts.append(f"Tasks: {tasks_str}")
            if item.get('OtherNote'):
                ghi_chu_parts.append(f"Note: {item['OtherNote']}")
            item['GhiChu'] = '; '.join(ghi_chu_parts) if ghi_chu_parts else ''
        print(f"DEBUG: Fetched {len(data)} records for email {email} with filter {filter_type}")
        return jsonify(data)
    except Exception as e:
        print(f"❌ Error in get_attendances: {e}")
        return jsonify({"error": str(e)}), 500

# ---- API lấy dữ liệu nghỉ phép (không thay đổi) ----
@app.route("/api/leaves", methods=["GET"])
def get_leaves():
    try:
        email = request.args.get("email")
        if not email:
            return jsonify({"error": "❌ Thiếu email"}), 400
        admin = admins.find_one({"email": email}, {"_id": 0, "username": 1})
        if not admin:
            return jsonify({"error": "🚫 Email không hợp lệ (không có quyền truy cập)"}), 403
        filter_type = request.args.get("filter", "tất cả").lower()
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()
        query = build_leave_query(filter_type, start_date, end_date, search)
        data = list(collection.find(query, {
            "_id": 0,
            "EmployeeId": 1,
            "EmployeeName": 1,
            "CheckinDate": 1,
            "CheckinTime": 1,
            "Tasks": 1,
            "Status": 1,
            "ApprovalDate": 1,
            "ApprovedBy": 1,
            "ApproveNote": 1
        }))
        for item in data:
            approval_date = item.get("ApprovalDate")
            if approval_date:
                if isinstance(approval_date, datetime):
                    item["ApprovalDate"] = approval_date.astimezone(VN_TZ).strftime("%d/%m/%Y %H:%M:%S")
                elif isinstance(approval_date, str) and approval_date.strip():
                    try:
                        parsed = datetime.strptime(approval_date, "%d/%m/%Y %H:%M:%S")
                        item["ApprovalDate"] = parsed.astimezone(VN_TZ).strftime("%d/%m/%Y %H:%M:%S")
                    except Exception:
                        item["ApprovalDate"] = approval_date
            else:
                item["ApprovalDate"] = None
            item["ApprovedBy"] = item.get("ApprovedBy", "")
            item["ApproveNote"] = item.get("ApproveNote", "")
        print(f"DEBUG: Fetched {len(data)} leave records for email {email} with filter {filter_type}")
        return jsonify(data)
    except Exception as e:
        print(f"❌ Error in get_leaves: {e}")
        return jsonify({"error": str(e)}), 500

# ---- API xuất Excel cho nghỉ phép (không thay đổi) ----
@app.route("/api/export-leaves-excel", methods=["GET"])
def export_leaves_to_excel():
    try:
        email = request.args.get("email")
        if not email:
            return jsonify({"error": "❌ Thiếu email"}), 400
        admin = admins.find_one({"email": email}, {"_id": 0, "username": 1})
        if not admin:
            return jsonify({"error": "🚫 Email không hợp lệ (không có quyền truy cập)"}), 403
        filter_type = request.args.get("filter", "tất cả").lower()
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()
        query = build_leave_query(filter_type, start_date, end_date, search)
        data = list(collection.find(query, {
            "_id": 0,
            "EmployeeId": 1,
            "EmployeeName": 1,
            "CheckinDate": 1,
            "CheckinTime": 1,
            "ApprovalDate": 1,
            "Tasks": 1,
            "Status": 1,
            "ApprovedBy": 1,
            "ApproveNote": 1
        }))
        grouped = {}
        for d in data:
            emp_id = d.get("EmployeeId", "")
            emp_name = d.get("EmployeeName", "")
            date = d.get("CheckinDate", "")
            key = (emp_id, emp_name, date)
            grouped.setdefault(key, []).append(d)
        template_path = "templates/Copy of Form chấm công.xlsx"
        wb = load_workbook(template_path)
        ws = wb.active
        border = Border(
            left=Side(style="thin", color="000000"),
            right=Side(style="thin", color="000000"),
            top=Side(style="thin", color="000000"),
            bottom=Side(style="thin", color="000000"),
        )
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        start_row = 2
        for i, ((emp_id, emp_name, date), records) in enumerate(grouped.items(), start=0):
            row = start_row + i
            ws.cell(row=row, column=1, value=emp_id)
            ws.cell(row=row, column=2, value=emp_name)
            ws.cell(row=row, column=3, value=date)
            for j, rec in enumerate(records[:10], start=1):
                checkin_time = rec.get("CheckinTime")
                full_datetime_str = ""
                if isinstance(checkin_time, datetime):
                    full_datetime_str = checkin_time.astimezone(VN_TZ).strftime("%d/%m/%Y %H:%M:%S")
                elif isinstance(checkin_time, str) and checkin_time.strip():
                    try:
                        parsed = datetime.strptime(checkin_time, "%d/%m/%Y %H:%M:%S")
                        full_datetime_str = parsed.strftime("%d/%m/%Y %H:%M:%S")
                    except Exception:
                        full_datetime_str = checkin_time
                tasks = rec.get("Tasks")
                tasks_str = ", ".join(tasks) if isinstance(tasks, list) else str(tasks or "")
                leave_task = tasks_str.strip()
                leave_reason = ""
                if ":" in leave_task:
                    split_task = leave_task.split(":", 1)
                    leave_task = split_task[0].strip()
                    leave_reason = split_task[1].strip()
                else:
                    leave_reason = rec.get("ApproveNote", "") or ""
                approved_by = rec.get("ApprovedBy", "")
                approval_date = rec.get("ApprovalDate")
                approval_date_str = ""
                if isinstance(approval_date, datetime):
                    approval_date_str = approval_date.astimezone(VN_TZ).strftime("%d/%m/%Y %H:%M:%S")
                elif isinstance(approval_date, str) and approval_date.strip():
                    try:
                        parsed = datetime.strptime(approval_date, "%d/%m/%Y %H:%M:%S")
                        approval_date_str = parsed.strftime("%d/%m/%Y %H:%M:%S")
                    except Exception:
                        approval_date_str = approval_date
                approval_status = f"Đã duyệt bởi {approved_by} lúc {approval_date_str}" if approved_by and approval_date_str else "Chưa duyệt"
                entry = f"{full_datetime_str}; {leave_task}; {leave_reason}; {approval_status}"
                ws.cell(row=row, column=3 + j, value=entry)
            for col in range(1, 14):
                cell = ws.cell(row=row, column=col)
                cell.border = border
                cell.alignment = align_left
            max_lines = max(
                (str(ws.cell(row=row, column=col).value).count("\n") + 1 if ws.cell(row=row, column=col).value else 1)
                for col in range(1, 14)
            )
            ws.row_dimensions[row].height = max_lines * 20
        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    length = len(str(cell.value).split("\n")[0])
                    max_length = max(max_length, length)
            ws.column_dimensions[col_letter].width = min(max_length + 2, 70)
        today_str = datetime.now(VN_TZ).strftime("%d-%m-%Y")
        filename = f"Danh sách nghỉ phép_{today_str}.xlsx"
        if search:
            filename = f"Danh sách nghỉ phép theo tìm kiếm_{today_str}.xlsx"
        elif filter_type == "custom" and start_date and end_date:
            filename = f"Danh sách nghỉ phép từ {start_date} đến {end_date}_{today_str}.xlsx"
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
        print("❌ Lỗi export leaves:", e)
        return jsonify({"error": str(e)}), 500

# ---- API xuất Excel kết hợp chấm công và nghỉ phép (không thay đổi) ----
@app.route("/api/export-combined-excel", methods=["GET"])
def export_combined_to_excel():
    try:
        email = request.args.get("email")
        if not email:
            return jsonify({"error": "❌ Thiếu email"}), 400
        admin = admins.find_one({"email": email}, {"_id": 0, "username": 1})
        if not admin:
            return jsonify({"error": "🚫 Email không hợp lệ (không có quyền truy cập)"}), 403
        filter_type = request.args.get("filter", "hôm nay").lower()
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()
        attendance_query = build_attendance_query(filter_type, start_date, end_date, search)
        leave_query = build_leave_query(filter_type, start_date, end_date, search)
        attendance_data = list(collection.find(attendance_query, {
            "_id": 0,
            "EmployeeId": 1, "EmployeeName": 1, "ProjectId": 1, "Tasks": 1,
            "OtherNote": 1, "Address": 1, "CheckinTime": 1, "CheckinDate": 1,
            "Status": 1, "ApprovedBy": 1, "Latitude": 1, "Longitude": 1
        }))
        leave_data = list(collection.find(leave_query, {
            "_id": 0,
            "EmployeeId": 1, "EmployeeName": 1, "CheckinDate": 1, "CheckinTime": 1,
            "Tasks": 1, "Status": 1, "ApprovedBy": 1, "ApproveNote": 1, "ApprovalDate": 1
        }))
        attendance_grouped = {}
        for d in attendance_data:
            key = (d.get("EmployeeId", ""), d.get("EmployeeName", ""), d.get("CheckinDate", ""))
            attendance_grouped.setdefault(key, []).append(d)
        leave_grouped = {}
        for d in leave_data:
            key = (d.get("EmployeeId", ""), d.get("EmployeeName", ""), d.get("CheckinDate", ""))
            leave_grouped.setdefault(key, []).append(d)
        template_path = "templates/Form kết hợp.xlsx"
        wb = load_workbook(template_path)
        ws_attendance = wb["Điểm danh"] if "Điểm danh" in wb.sheetnames else wb.create_sheet("Điểm danh")
        ws_leaves = wb["Nghỉ phép"] if "Nghỉ phép" in wb.sheetnames else wb.create_sheet("Nghỉ phép")
        headers = ["Mã NV", "Tên NV", "Ngày", "Check 1", "Check 2", "Check 3", "Check 4", "Check 5",
                   "Check 6", "Check 7", "Check 8", "Check 9", "Check 10"]
        for col, header in enumerate(headers, start=1):
            ws_attendance.cell(row=1, column=col, value=header)
            ws_leaves.cell(row=1, column=col, value=header)
        border = Border(
            left=Side(style="thin", color="000000"), right=Side(style="thin", color="000000"),
            top=Side(style="thin", color="000000"), bottom=Side(style="thin", color="000000"),
        )
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        start_row = 2
        for i, ((emp_id, emp_name, date), records) in enumerate(attendance_grouped.items(), start=0):
            row = start_row + i
            ws_attendance.cell(row=row, column=1, value=emp_id)
            ws_attendance.cell(row=row, column=2, value=emp_name)
            ws_attendance.cell(row=row, column=3, value=date)
            for j, rec in enumerate(records[:10], start=1):
                checkin_time = rec.get("CheckinTime")
                time_str = ""
                if isinstance(checkin_time, datetime):
                    time_str = checkin_time.astimezone(VN_TZ).strftime("%H:%M:%S")
                elif isinstance(checkin_time, str) and checkin_time.strip():
                    try:
                        parsed = datetime.strptime(checkin_time, "%d/%m/%Y %H:%M:%S")
                        time_str = parsed.strftime("%H:%M:%S")
                    except Exception:
                        time_str = checkin_time
                parts = []
                tasks = rec.get("Tasks")
                tasks_str = ", ".join(tasks) if isinstance(tasks, list) else str(tasks or "")
                if "nghỉ phép" in tasks_str.lower():
                    entry = "NGHỈ PHÉP (xem chi tiết ở sheet Nghỉ phép)"
                else:
                    if time_str: parts.append(time_str)
                    if rec.get("ProjectId"): parts.append(str(rec["ProjectId"]))
                    if tasks_str: parts.append(tasks_str)
                    if rec.get("Status"): parts.append(rec["Status"])
                    if rec.get("OtherNote"): parts.append(rec["OtherNote"])
                    if rec.get("Address"): parts.append(rec["Address"])
                    entry = "; ".join(parts)
                ws_attendance.cell(row=row, column=3 + j, value=entry)
            for col in range(1, 14):
                cell = ws_attendance.cell(row=row, column=col)
                cell.border = border
                cell.alignment = align_left
            max_lines = max(
                (str(ws_attendance.cell(row=row, column=col).value).count("\n") + 1 if ws_attendance.cell(row=row, column=col).value else 1)
                for col in range(1, 14)
            )
            ws_attendance.row_dimensions[row].height = max_lines * 20
        for col in ws_attendance.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    length = len(str(cell.value).split("\n")[0])
                    max_length = max(max_length, length)
            ws_attendance.column_dimensions[col_letter].width = min(max_length + 2, 70)
        start_row = 2
        for i, ((emp_id, emp_name, date), records) in enumerate(leave_grouped.items(), start=0):
            row = start_row + i
            ws_leaves.cell(row=row, column=1, value=emp_id)
            ws_leaves.cell(row=row, column=2, value=emp_name)
            ws_leaves.cell(row=row, column=3, value=date)
            for j, rec in enumerate(records[:10], start=1):
                checkin_time = rec.get("CheckinTime")
                full_datetime_str = ""
                if isinstance(checkin_time, datetime):
                    full_datetime_str = checkin_time.astimezone(VN_TZ).strftime("%d/%m/%Y %H:%M:%S")
                elif isinstance(checkin_time, str) and checkin_time.strip():
                    try:
                        parsed = datetime.strptime(checkin_time, "%d/%m/%Y %H:%M:%S")
                        full_datetime_str = parsed.strftime("%d/%m/%Y %H:%M:%S")
                    except Exception:
                        full_datetime_str = checkin_time
                tasks = rec.get("Tasks")
                tasks_str = ", ".join(tasks) if isinstance(tasks, list) else str(tasks or "")
                leave_task = tasks_str.strip()
                leave_reason = ""
                if ":" in leave_task:
                    split_task = leave_task.split(":", 1)
                    leave_task = split_task[0].strip()
                    leave_reason = split_task[1].strip()
                else:
                    leave_reason = rec.get("ApproveNote", "") or ""
                approved_by = rec.get("ApprovedBy", "")
                approval_date = rec.get("ApprovalDate")
                approval_date_str = ""
                if isinstance(approval_date, datetime):
                    approval_date_str = approval_date.astimezone(VN_TZ).strftime("%d/%m/%Y %H:%M:%S")
                elif isinstance(approval_date, str) and approval_date.strip():
                    try:
                        parsed = datetime.strptime(approval_date, "%d/%m/%Y %H:%M:%S")
                        approval_date_str = parsed.strftime("%d/%m/%Y %H:%M:%S")
                    except Exception:
                        approval_date_str = approval_date
                approval_status = f"Đã duyệt bởi {approved_by} lúc {approval_date_str}" if approved_by and approval_date_str else "Chưa duyệt"
                entry = f"{full_datetime_str}; {leave_task}; {leave_reason}; {approval_status}"
                ws_leaves.cell(row=row, column=3 + j, value=entry)
            for col in range(1, 14):
                cell = ws_leaves.cell(row=row, column=col)
                cell.border = border
                cell.alignment = align_left
            max_lines = max(
                (str(ws_leaves.cell(row=row, column=col).value).count("\n") + 1 if ws_leaves.cell(row=row, column=col).value else 1)
                for col in range(1, 14)
            )
            ws_leaves.row_dimensions[row].height = max_lines * 20
        for col in ws_leaves.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    length = len(str(cell.value).split("\n")[0])
                    max_length = max(max_length, length)
            ws_leaves.column_dimensions[col_letter].width = min(max_length + 2, 70)
        today_str = datetime.now(VN_TZ).strftime("%d-%m-%Y")
        filename = f"Danh sách chấm công và nghỉ phép_{today_str}.xlsx"
        if search:
            filename = f"Danh sách chấm công và nghỉ phép theo tìm kiếm_{today_str}.xlsx"
        elif filter_type == "custom" and start_date and end_date:
            filename = f"Danh sách chấm công và nghỉ phép từ {start_date} đến {end_date}_{today_str}.xlsx"
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
        print("❌ Lỗi export combined:", e)
        return jsonify({"error": str(e)}), 500

# ---- Đóng kết nối SMTP khi ứng dụng tắt ----
@app.teardown_appcontext
def cleanup(exception=None):
    close_smtp_server()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
