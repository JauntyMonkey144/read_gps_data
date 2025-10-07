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
import smtplib
from email.mime.text import MIMEText
from itsdangerous import URLSafeTimedSerializer as Serializer, SignatureExpired

app = Flask(__name__, template_folder="templates")
CORS(app, methods=["GET", "POST"])

# ---- Cấu hình chung ----
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'mot_key_bi_mat_va_dai_duoc_giu_kin')
VN_TZ = timezone(timedelta(hours=7))

# ---- MongoDB Config ----
MONGO_URI = os.getenv(
    "MONGO_URI",
    "mongodb+srv://banhbaobeo2205:lm2hiCLXp6B0D7hq@cluster0.festnla.mongodb.net/?retryWrites=true&w=majority"
)
DB_NAME = os.getenv("DB_NAME", "Sun_Database_1")
client = MongoClient(MONGO_URI)
db = client[DB_NAME]
admins = db["admins"]
collection = db["alt_checkins"]

# ---- SMTP Config ----
SMTP_SERVER = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
SMTP_PORT = os.getenv('SMTP_PORT', 587)
SMTP_USERNAME = os.getenv('SMTP_USERNAME', 'banhbaobeo2205@gmail.com')
SMTP_PASSWORD = os.getenv('SMTP_PASSWORD', 'vynqvvvmbcigpdvy')  # Thay bằng mật khẩu ứng dụng Gmail
SMTP_FROM = os.getenv('SMTP_FROM', 'it.trankhanhvinh@gmail.com')

# ---- ItsDangerous Serializer ----
s = Serializer(app.config['SECRET_KEY'])

# ---- HÀM TIỆN ÍCH MAIL ----
def get_reset_token(email, expires_sec=1800):  # 30 phút
    return s.dumps({'user_email': email}).decode('utf-8')

def verify_reset_token(token):
    try:
        data = s.loads(token, max_age=1800)
        return data['user_email']
    except SignatureExpired:
        return None
    except Exception:
        return None

def send_reset_email(admin):
    token = get_reset_token(admin['email'])
    reset_url = url_for('reset_password', token=token, _external=True)
    msg = MIMEText(f"""
    <p>Xin chào {admin['username']},</p>
    <p>Bạn (hoặc ai đó) đã yêu cầu đặt lại mật khẩu cho tài khoản admin.</p>
    <p>Vui lòng nhấp vào đường link sau để <strong>ĐẶT LẠI MẬT KHẨU</strong>: <a href="{reset_url}">{reset_url}</a></p>
    <p style="color: red;"><strong>Link này sẽ hết hạn sau 30 phút.</strong></p>
    <p>Nếu bạn không yêu cầu điều này, hãy bỏ qua email này.</p>
    <p>Trân trọng,</p>
    <p>Hệ thống Admin</p>
    """, 'html')
    msg['Subject'] = 'Yêu cầu Đặt lại Mật khẩu'
    msg['From'] = SMTP_FROM
    msg['To'] = admin['email']
    try:
        print(f"DEBUG: Trying to connect to {SMTP_SERVER}:{SMTP_PORT} with {SMTP_USERNAME}")
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=10) as server:
            server.starttls()  # Bật TLS
            print(f"DEBUG: Logging in with {SMTP_USERNAME}")
            server.login(SMTP_USERNAME, SMTP_PASSWORD)
            print(f"DEBUG: Sending email to {admin['email']}")
            server.send_message(msg)
        print(f"✅ Email gửi thành công đến {admin['email']} tại {datetime.now(VN_TZ).strftime('%H:%M:%S %d/%m/%Y')}")
        return True
    except smtplib.SMTPAuthenticationError as e:
        print(f"❌ Lỗi xác thực SMTP: {str(e)} tại {datetime.now(VN_TZ).strftime('%H:%M:%S %d/%m/%Y')}")
        return False
    except smtplib.SMTPConnectError as e:
        print(f"❌ Lỗi kết nối SMTP: {str(e)} tại {datetime.now(VN_TZ).strftime('%H:%M:%S %d/%m/%Y')}")
        return False
    except smtplib.SMTPException as e:
        print(f"❌ Lỗi SMTP: {str(e)} tại {datetime.now(VN_TZ).strftime('%H:%M:%S %d/%m/%Y')}")
        return False
    except Exception as e:
        print(f"❌ Lỗi gửi email: {str(e)} tại {datetime.now(VN_TZ).strftime('%H:%M:%S %d/%m/%Y')}")
        return False

# ---- ROUTES ỨNG DỤNG ----
@app.route("/")
def index():
    success = request.args.get("success")
    message = None
    if success == '1':
        message = "✅ Đặt lại mật khẩu thành công! Vui lòng đăng nhập."
    return render_template("index.html", success=success, message=message)

# ---- Trang quên mật khẩu ----
@app.route("/forgot-password", methods=["GET", "POST"])
def forgot_password():
    if request.method == "POST":
        email = request.form.get("email")
        admin = db.admins.find_one({"email": email})
        if admin:
            reset_token = str(uuid.uuid4())
            db.admins.update_one({"_id": admin["_id"]}, {"$set": {"reset_token": reset_token, "reset_expiry": datetime.now(VN_TZ) + timedelta(hours=1)}})
            reset_url = f"{request.host_url.rstrip('/')}/reset-password/{reset_token}"
            body = f"""
            <!DOCTYPE html>
            <html><head><meta charset="UTF-8"></head><body>
                <h2>🔄 Đặt lại mật khẩu</h2>
                <p>Nhấp vào link để đặt lại: <a href="{reset_url}">Đặt lại mật khẩu</a></p>
                <p>Token hết hạn sau 1 giờ.</p>
            </body></html>
            """
            send_email(email, "Đặt lại mật khẩu Admin", body)
            flash("Email đặt lại mật khẩu đã gửi!", "success")
        else:
            flash("Email không tồn tại!", "error")
        return redirect(url_for("login"))
    return f"""
    <!DOCTYPE html>
    <html lang="vi">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Quên mật khẩu</title>
        <style>
            body {{ font-family: Arial, sans-serif; background: #f4f6f9; margin: 0; padding: 20px; }}
            .container {{ max-width: 400px; margin: 100px auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }}
            input {{ width: 100%; padding: 10px; margin: 10px 0; box-sizing: border-box; border: 1px solid #ddd; border-radius: 4px; }}
            button {{ background: #007bff; color: white; padding: 12px; width: 100%; border: none; border-radius: 4px; cursor: pointer; font-size: 16px; }}
            button:hover {{ background: #0056b3; }}
            a {{ color: #007bff; text-decoration: none; display: block; margin-top: 10px; text-align: center; }}
            .flash {{ margin: 10px 0; padding: 10px; border-radius: 4px; }}
        </style>
    </head>
    <body>
        <div class="container">
            <h2>🔒 Quên mật khẩu</h2>
            {render_flash_messages()}
            <form method="POST">
                <input type="email" name="email" placeholder="Email" required>
                <button type="submit">Gửi email đặt lại</button>
            </form>
            <a href="/login">Quay về đăng nhập</a>
        </div>
    </body>
    </html>
    """

# ---- Trang đặt lại mật khẩu ----
@app.route("/reset-password/<token>", methods=["GET", "POST"])
def reset_password(token):
    admin = db.admins.find_one({"reset_token": token, "reset_expiry": {"$gt": datetime.now(VN_TZ)}})
    if not admin:
        flash("Token không hợp lệ hoặc hết hạn!", "error")
        return redirect(url_for("login"))
    if request.method == "POST":
        new_password = request.form.get("password")
        hashed_pw = generate_password_hash(new_password)
        db.admins.update_one({"_id": admin["_id"]}, {"$set": {"password": hashed_pw, "reset_token": None, "reset_expiry": None}})
        flash("Mật khẩu đã được cập nhật!", "success")
        return redirect(url_for("login"))
    return f"""
    <!DOCTYPE html>
    <html lang="vi">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Đặt lại mật khẩu</title>
        <style>
            body {{ font-family: Arial, sans-serif; background: #f4f6f9; margin: 0; padding: 20px; }}
            .container {{ max-width: 400px; margin: 100px auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }}
            input {{ width: 100%; padding: 10px; margin: 10px 0; box-sizing: border-box; border: 1px solid #ddd; border-radius: 4px; }}
            button {{ background: #28a745; color: white; padding: 12px; width: 100%; border: none; border-radius: 4px; cursor: pointer; font-size: 16px; }}
            button:hover {{ background: #218838; }}
            a {{ color: #007bff; text-decoration: none; display: block; margin-top: 10px; text-align: center; }}
            .flash {{ margin: 10px 0; padding: 10px; border-radius: 4px; }}
        </style>
    </head>
    <body>
        <div class="container">
            <h2>🔄 Đặt lại mật khẩu</h2>
            {render_flash_messages()}
            <form method="POST">
                <input type="password" name="password" placeholder="Mật khẩu mới" required>
                <button type="submit">Cập nhật</button>
            </form>
            <a href="/login">Quay về đăng nhập</a>
        </div>
    </body>
    </html>
    """

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
    return {"$and": conditions} if len(conditions) > 1 else conditions[0]

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
            "ApprovalDate": 1,
            "Tasks": 1,
            "Status": 1
        }))
        print(f"DEBUG: Fetched {len(data)} leave records for email {email} with filter {filter_type}")
        return jsonify(data)
    except Exception as e:
        print(f"❌ Error in get_leaves: {e}")
        return jsonify({"error": str(e)}), 500

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
                status = rec.get("Status", "")
                if rec.get("ApprovedBy"):
                    status = f"Đã duyệt bởi {rec['ApprovedBy']}"
                entry = f"{full_datetime_str}; {leave_task}: {leave_reason}; {status}"
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

@app.route("/api/export-excel", methods=["GET"])
def export_to_excel():
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
        data = list(collection.find(query, {
            "_id": 0,
            "EmployeeId": 1,
            "EmployeeName": 1,
            "ProjectId": 1,
            "Tasks": 1,
            "OtherNote": 1,
            "Address": 1,
            "CheckinTime": 1,
            "CheckinDate": 1,
            "Status": 1,
            "ApprovedBy": 1,
            "Latitude": 1,
            "Longitude": 1
        }))
        grouped = {}
        for d in data:
            emp_id = d.get("EmployeeId", "")
            emp_name = d.get("EmployeeName", "")
            date = d.get("CheckinDate", "")
            key = (emp_id, emp_name, date)
            grouped.setdefault(key, []).append(d)
        template_path = "templates/Form kết hợp.xlsx"
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
                leave_reason = ""
                if "nghỉ phép" in tasks_str.lower():
                    if ":" in tasks_str:
                        split_task = tasks_str.split(":", 1)
                        tasks_str = split_task[0].strip()
                        leave_reason = split_task[1].strip()
                    else:
                        tasks_str = tasks_str.strip()
                    status = rec.get("Status", "")
                    approve_date = ""
                    if rec.get("ApprovedBy"):
                        approve_date = datetime.now(VN_TZ).strftime("%d/%m/%Y") if isinstance(checkin_time, str) else checkin_time.astimezone(VN_TZ).strftime("%d/%m/%Y")
                    entry = f"{date}; Nghỉ phép; {leave_reason}; {status}; {approve_date}"
                else:
                    if time_str:
                        parts.append(time_str)
                    if rec.get("ProjectId"):
                        parts.append(str(rec["ProjectId"]))
                    if tasks_str:
                        parts.append(tasks_str)
                    if rec.get("Status"):
                        parts.append(rec["Status"])
                    if rec.get("OtherNote"):
                        parts.append(rec["OtherNote"])
                    if rec.get("Address"):
                        parts.append(rec["Address"])
                    entry = "; ".join(parts)
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
                    length = len(str(cell.value))
                    max_length = max(max_length, length)
            ws.column_dimensions[col_letter].width = max_length + 2
        today_str = datetime.now(VN_TZ).strftime("%d-%m-%Y")
        filename = f"Danh sách chấm công_{today_str}.xlsx"
        if search:
            filename = f"Danh sách chấm công theo tìm kiếm_{today_str}.xlsx"
        elif filter_type == "custom" and start_date and end_date:
            filename = f"Danh sách chấm công từ {start_date} đến {end_date}_{today_str}.xlsx"
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
        print("❌ Lỗi export:", e)
        return jsonify({"error": str(e)}), 500

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
            "Tasks": 1, "Status": 1, "ApprovedBy": 1, "ApproveNote": 1
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
                status = rec.get("Status", "")
                if rec.get("ApprovedBy"):
                    status = f"Đã duyệt bởi {rec['ApprovedBy']}"
                entry = f"{full_datetime_str}; {leave_task}: {leave_reason}; {status}"
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

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
