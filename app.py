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

# Imports mới cho chức năng gửi mail và token
from flask_mail import Mail, Message
from itsdangerous import URLSafeTimedSerializer as Serializer, SignatureExpired # Thư viện tạo token bảo mật

app = Flask(__name__, template_folder="templates")
CORS(app, methods=["GET", "POST"])

# ---- Cấu hình chung ----
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'mot_key_bi_mat_va_dai_duoc_giu_kin') # Rất quan trọng cho token!

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

# Các collection sử dụng
admins = db["admins"]
collection = db["alt_checkins"]

# ---- Flask-Mail Config ----
# Cần thay đổi các giá trị này bằng thông tin SMTP của bạn
app.config['MAIL_SERVER'] = os.getenv('MAIL_SERVER', 'smtp.gmail.com') # Ví dụ: SMTP của Gmail
app.config['MAIL_PORT'] = os.getenv('MAIL_PORT', 587)
app.config['MAIL_USE_TLS'] = os.getenv('MAIL_USE_TLS', True)
# Thay thế bằng email và mật khẩu ứng dụng của bạn!
app.config['MAIL_USERNAME'] = os.getenv('MAIL_USERNAME', 'your_email@gmail.com')
app.config['MAIL_PASSWORD'] = os.getenv('MAIL_PASSWORD', 'your_app_password')
app.config['MAIL_DEFAULT_SENDER'] = ('Admin', app.config['MAIL_USERNAME'])

mail = Mail(app)

# ---- ItsDangerous Serializer ----
s = Serializer(app.config['SECRET_KEY'])

# --------------------------
# ---- HÀM TIỆN ÍCH MAIL ----
# --------------------------

def get_reset_token(email, expires_sec=1800): # 30 phút
    """Tạo token đặt lại mật khẩu với thời hạn (mặc định 30 phút)"""
    # Dùng email làm payload, mã hóa và trả về chuỗi token
    return s.dumps({'user_email': email}).decode('utf-8')

def verify_reset_token(token):
    """Xác minh token và trả về email nếu hợp lệ, ngược lại trả về None"""
    try:
        # Giải mã token và kiểm tra thời hạn
        data = s.loads(token, max_age=1800)
        return data['user_email']
    except SignatureExpired:
        return None # Token hết hạn
    except Exception:
        return None # Lỗi khác (token sai)

def send_reset_email(admin):
    """Gửi email chứa link đặt lại mật khẩu"""
    token = get_reset_token(admin['email'])
    
    # Tạo đường link đặt lại mật khẩu
    reset_url = url_for('reset_password', token=token, _external=True)
    
    msg = Message(
        'Yêu cầu Đặt lại Mật khẩu',
        recipients=[admin['email']],
        html=f"""
        <p>Xin chào {admin['username']},</p>
        <p>Bạn (hoặc ai đó) đã yêu cầu đặt lại mật khẩu cho tài khoản admin.</p>
        <p>Vui lòng nhấp vào đường link sau để **ĐẶT LẠI MẬT KHẨU**: <a href="{reset_url}">{reset_url}</a></p>
        <p style="color: red;">**Link này sẽ hết hạn sau 30 phút.**</p>
        <p>Nếu bạn không yêu cầu điều này, hãy bỏ qua email này.</p>
        <p>Trân trọng,</p>
        <p>Hệ thống Admin</p>
        """
    )
    try:
        mail.send(msg)
        return True
    except Exception as e:
        print(f"❌ Lỗi gửi email: {e}")
        return False

# --------------------------
# ---- ROUTES ỨNG DỤNG ----
# --------------------------

# ---- Trang chủ (đăng nhập chính) ----
@app.route("/")
def index():
    success = request.args.get("success")
    message = None
    if success == '1':
        message = "✅ Đặt lại mật khẩu thành công! Vui lòng đăng nhập."
    
    # Giữ nguyên `success` cho logic ban đầu của bạn nếu có, thêm `message` cho thông báo cụ thể
    return render_template("index.html", success=success, message=message)

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

# ---- Yêu cầu Quên mật khẩu (Gửi email) ----
@app.route("/forgot-password", methods=["GET", "POST"])
def forgot_password():
    if request.method == "GET":
        status = request.args.get('status', 'info')
        message = request.args.get('message')
        
        # HTML form để người dùng nhập email
        return f"""
        <!DOCTYPE html>
        <html lang="vi">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Quên mật khẩu</title>
            <style>
                body {{ font-family: Arial, sans-serif; background: #f4f6f9; margin: 0; padding: 20px; text-align: center; }}
                .container {{ max-width: 400px; margin: 100px auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }}
                input {{ width: 100%; padding: 10px; margin: 10px 0; box-sizing: border-box; border: 1px solid #ddd; border-radius: 4px; }}
                button {{ background: #ffc107; color: black; padding: 12px; width: 100%; border: none; border-radius: 4px; cursor: pointer; font-size: 16px; font-weight: bold; }}
                button:hover {{ background: #e0a800; }}
                p.success {{ color: green; font-weight: bold; }}
                p.error {{ color: red; font-weight: bold; }}
            </style>
        </head>
        <body>
            <div class="container">
                <h2>🔒 Yêu cầu Đặt lại Mật khẩu</h2>
                <p>Nhập email của bạn để nhận link đặt lại.</p>
                {f'<p class="{status}">{message}</p>' if message else ''}
                <form method="POST">
                    <input type="email" name="email" placeholder="Email đăng nhập" required>
                    <button type="submit">Gửi Link Đặt lại</button>
                    <p style="margin-top: 15px;"><a href="/">Quay về trang chủ</a></p>
                </form>
            </div>
        </body>
        </html>
        """

    if request.method == "POST":
        email = request.form.get("email")
        if not email:
            return redirect(url_for("forgot_password", message="❌ Vui lòng nhập email", status="error"))

        admin = admins.find_one({"email": email})
        
        # Luôn trả về thông báo thành công chung để tránh lộ thông tin user
        message_redirect = redirect(url_for(
            "forgot_password", 
            message="✅ Nếu email tồn tại, một link đặt lại đã được gửi đến hộp thư của bạn (hết hạn sau 30 phút).", 
            status="success"
        ))

        if admin:
            # Chỉ gửi mail nếu email tồn tại
            send_reset_email(admin)
        
        return message_redirect

# ---- Đặt lại mật khẩu (Xử lý token và mật khẩu mới) ----
@app.route("/reset-password/<token>", methods=["GET", "POST"])
def reset_password(token):
    email = verify_reset_token(token)

    if not email:
        return """
        <p style="color:red; text-align:center; margin-top: 50px; font-size: 20px;">
            🚫 Token không hợp lệ hoặc đã hết hạn (30 phút). 
            Vui lòng yêu cầu đặt lại mật khẩu <a href="/forgot-password">tại đây</a>.
        </p>
        """, 403
    
    admin = admins.find_one({"email": email})
    if not admin:
        return """
        <p style="color:red; text-align:center; margin-top: 50px; font-size: 20px;">
            🚫 Tài khoản không tồn tại.
        </p>
        """, 404

    if request.method == "GET":
        # HTML form để người dùng nhập mật khẩu mới
        return f"""
        <!DOCTYPE html>
        <html lang="vi">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Đặt lại mật khẩu</title>
            <style>
                body {{ font-family: Arial, sans-serif; background: #f4f6f9; margin: 0; padding: 20px; text-align: center; }}
                .container {{ max-width: 400px; margin: 100px auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }}
                input {{ width: 100%; padding: 10px; margin: 10px 0; box-sizing: border-box; border: 1px solid #ddd; border-radius: 4px; }}
                button {{ background: #28a745; color: white; padding: 12px; width: 100%; border: none; border-radius: 4px; cursor: pointer; font-size: 16px; }}
                button:hover {{ background: #218838; }}
            </style>
        </head>
        <body>
            <div class="container">
                <h2>✅ Đặt lại Mật khẩu cho {email}</h2>
                <form method="POST">
                    <input type="password" name="new_password" placeholder="Mật khẩu mới" required>
                    <button type="submit">Cập nhật mật khẩu</button>
                    <p style="margin-top: 15px;"><a href="/">Quay về trang chủ</a></p>
                </form>
            </div>
        </body>
        </html>
        """

    if request.method == "POST":
        new_password = request.form.get("new_password")
        
        if not new_password:
            return "❌ Vui lòng nhập mật khẩu mới", 400
        
        # Cập nhật mật khẩu
        hashed_pw = generate_password_hash(new_password)
        admins.update_one({"email": email}, {"$set": {"password": hashed_pw}})

        # Chuyển hướng về trang đăng nhập với thông báo thành công
        return redirect(url_for("index", success=1)) # success=1: thông báo đặt lại mật khẩu thành công


def build_attendance_query(filter_type, start_date, end_date, search):
    today = datetime.now(VN_TZ)
    regex_leave = re.compile("Nghỉ phép", re.IGNORECASE)
    conditions = []

    # --- Bộ lọc thời gian ---
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

    # --- Bộ lọc nghỉ phép ---
    not_leave_or = {
        "$or": [
            {"Tasks": {"$not": regex_leave}},
            {"Tasks": {"$exists": False}},
            {"Tasks": None}
        ]
    }
    conditions.append(not_leave_or)

    # --- Bộ lọc tìm kiếm ---
    if search:
        regex = re.compile(search, re.IGNORECASE)
        search_or = {
            "$or": [
                {"EmployeeId": {"$regex": regex}},
                {"EmployeeName": {"$regex": regex}}
            ]
        }
        conditions.append(search_or)

    # Kết hợp tất cả với $and
    if len(conditions) == 1:
        return conditions[0]
    else:
        return {"$and": conditions}

def build_leave_query(filter_type, start_date, end_date, search):
    today = datetime.now(VN_TZ)
    regex_leave = re.compile("Nghỉ phép", re.IGNORECASE)
    conditions = []

    # Luôn lọc cho nghỉ phép
    leave_or = {
        "$or": [
            {"Tasks": {"$regex": regex_leave}},
            {"Status": {"$regex": regex_leave}},
            {"OtherNote": {"$regex": regex_leave}}
        ]
    }
    conditions.append(leave_or)

    # --- Bộ lọc thời gian ---
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

    # --- Bộ lọc tìm kiếm ---
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

# ---- API lấy dữ liệu chấm công ----
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

# ---- API lấy dữ liệu nghỉ phép ----
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
            "Status": 1
        }))
        print(f"DEBUG: Fetched {len(data)} leave records for email {email} with filter {filter_type}")
        return jsonify(data)
    except Exception as e:
        print(f"❌ Error in get_leaves: {e}")
        return jsonify({"error": str(e)}), 500

# ---- API xuất Excel cho nghỉ phép ----
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
        
        # Chỉ lấy các trường cần thiết
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
            # Vẫn nhóm theo Ngày nghỉ, nhưng logic xuất sẽ khác
            key = (emp_id, emp_name, date)
            grouped.setdefault(key, []).append(d)
        
        template_path = "templates/Copy of Form chấm công.xlsx" # Giả sử template này được dùng
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
                
                # --- LOGIC MỚI BẮT ĐẦU ---
                
                # 1. Chuyển đổi CheckinTime (thời gian tạo đơn) sang format đầy đủ
                full_datetime_str = ""
                if isinstance(checkin_time, datetime):
                    # Nếu là đối tượng datetime (ví dụ: từ MongoDB), chuyển về múi giờ VN và format
                    full_datetime_str = checkin_time.astimezone(VN_TZ).strftime("%d/%m/%Y %H:%M:%S")
                elif isinstance(checkin_time, str) and checkin_time.strip():
                    try:
                        # Thử phân tích chuỗi CheckinTime cũ
                        parsed = datetime.strptime(checkin_time, "%d/%m/%Y %H:%M:%S")
                        full_datetime_str = parsed.strftime("%d/%m/%Y %H:%M:%S")
                    except Exception:
                        # Giữ nguyên nếu không parse được
                        full_datetime_str = checkin_time

                # 2. Phân tích Tasks và Lý do nghỉ (nếu có)
                tasks = rec.get("Tasks")
                tasks_str = ", ".join(tasks) if isinstance(tasks, list) else str(tasks or "")
                
                leave_task = tasks_str.strip()
                leave_reason = ""
                
                # Phân tách Tasks và Lý do: TáchTasks: Lý do
                if ":" in leave_task:
                    split_task = leave_task.split(":", 1)
                    leave_task = split_task[0].strip()
                    leave_reason = split_task[1].strip()
                
                # 3. Lấy Status (kèm ApprovedBy nếu có)
                status = rec.get("Status", "")
                if rec.get("ApprovedBy"):
                    status = f"Đã duyệt bởi {rec['ApprovedBy']}"
                
                # 4. Tạo chuỗi entry theo format yêu cầu: {CheckInTime}; {Task}: {Lý do}; {Status}
                entry = f"{full_datetime_str}; {leave_task}: {leave_reason}; {status}"

                # --- LOGIC MỚI KẾT THÚC ---

                ws.cell(row=row, column=3 + j, value=entry)
                
            # Áp dụng style và tính chiều cao dòng
            for col in range(1, 14):
                cell = ws.cell(row=row, column=col)
                cell.border = border
                cell.alignment = align_left
            max_lines = max(
                (str(ws.cell(row=row, column=col).value).count("\n") + 1 if ws.cell(row=row, column=col).value else 1)
                for col in range(1, 14)
            )
            ws.row_dimensions[row].height = max_lines * 20
        
        # Tự động điều chỉnh độ rộng cột
        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    # Giới hạn độ dài tối đa để tránh cột quá rộng
                    length = len(str(cell.value).split("\n")[0]) # Chỉ tính chiều dài của dòng đầu tiên
                    max_length = max(max_length, length)
            ws.column_dimensions[col_letter].width = min(max_length + 2, 70) # Giới hạn max width 70
        
        # Xuất file
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

# ---- API xuất Excel cho chấm công ----
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
                    if leave_reason:
                        parts.append(leave_reason)
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

# ---- API xuất Excel kết hợp chấm công và nghỉ phép ----
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

        # Xác định bộ lọc
        attendance_query = build_attendance_query(filter_type, start_date, end_date, search)
        leave_query = build_leave_query(filter_type, start_date, end_date, search)

        # Lấy dữ liệu
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

        # Nhóm dữ liệu
        attendance_grouped = {}
        for d in attendance_data:
            key = (d.get("EmployeeId", ""), d.get("EmployeeName", ""), d.get("CheckinDate", ""))
            attendance_grouped.setdefault(key, []).append(d)

        leave_grouped = {}
        for d in leave_data:
            key = (d.get("EmployeeId", ""), d.get("EmployeeName", ""), d.get("CheckinDate", ""))
            leave_grouped.setdefault(key, []).append(d)

        # Load template Excel
        template_path = "templates/Form kết hợp.xlsx"
        wb = load_workbook(template_path)
        ws_attendance = wb["Điểm danh"] if "Điểm danh" in wb.sheetnames else wb.create_sheet("Điểm danh")
        ws_leaves = wb["Nghỉ phép"] if "Nghỉ phép" in wb.sheetnames else wb.create_sheet("Nghỉ phép")

        # Ghi tiêu đề (giữ nguyên)
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

        # Điền dữ liệu chấm công (Giữ nguyên logic cũ)
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
                
                # Logic cũ (giữ nguyên cho sheet chấm công)
                if "nghỉ phép" in tasks_str.lower():
                    # Nếu là nghỉ phép, chỉ ghi thông tin tóm tắt cho sheet chấm công
                    entry = "NGHỈ PHÉP (xem chi tiết ở sheet Nghỉ phép)"
                else:
                    # Nếu là chấm công bình thường
                    if time_str: parts.append(time_str)
                    if rec.get("ProjectId"): parts.append(str(rec["ProjectId"]))
                    if tasks_str: parts.append(tasks_str)
                    if rec.get("Status"): parts.append(rec["Status"])
                    if rec.get("OtherNote"): parts.append(rec["OtherNote"])
                    if rec.get("Address"): parts.append(rec["Address"])
                    entry = "; ".join(parts)
                    
                ws_attendance.cell(row=row, column=3 + j, value=entry)
                
            # Áp dụng style và tính chiều cao dòng
            for col in range(1, 14):
                cell = ws_attendance.cell(row=row, column=col)
                cell.border = border
                cell.alignment = align_left
            max_lines = max(
                (str(ws_attendance.cell(row=row, column=col).value).count("\n") + 1 if ws_attendance.cell(row=row, column=col).value else 1)
                for col in range(1, 14)
            )
            ws_attendance.row_dimensions[row].height = max_lines * 20
        
        # Tự động điều chỉnh độ rộng cột
        for col in ws_attendance.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    length = len(str(cell.value).split("\n")[0])
                    max_length = max(max_length, length)
            ws_attendance.column_dimensions[col_letter].width = min(max_length + 2, 70)

        # Điền dữ liệu nghỉ phép (ÁP DỤNG LOGIC MỚI)
        start_row = 2
        for i, ((emp_id, emp_name, date), records) in enumerate(leave_grouped.items(), start=0):
            row = start_row + i
            ws_leaves.cell(row=row, column=1, value=emp_id)
            ws_leaves.cell(row=row, column=2, value=emp_name)
            ws_leaves.cell(row=row, column=3, value=date)
            
            for j, rec in enumerate(records[:10], start=1):
                checkin_time = rec.get("CheckinTime")
                
                # 1. Chuyển đổi CheckinTime (thời gian tạo đơn) sang format đầy đủ
                full_datetime_str = ""
                if isinstance(checkin_time, datetime):
                    full_datetime_str = checkin_time.astimezone(VN_TZ).strftime("%d/%m/%Y %H:%M:%S")
                elif isinstance(checkin_time, str) and checkin_time.strip():
                    try:
                        parsed = datetime.strptime(checkin_time, "%d/%m/%Y %H:%M:%S")
                        full_datetime_str = parsed.strftime("%d/%m/%Y %H:%M:%S")
                    except Exception:
                        full_datetime_str = checkin_time

                # 2. Phân tích Tasks và Lý do nghỉ (nếu có)
                tasks = rec.get("Tasks")
                tasks_str = ", ".join(tasks) if isinstance(tasks, list) else str(tasks or "")
                
                leave_task = tasks_str.strip()
                leave_reason = ""
                
                if ":" in leave_task:
                    split_task = leave_task.split(":", 1)
                    leave_task = split_task[0].strip()
                    leave_reason = split_task[1].strip()
                
                # 3. Lấy Status (kèm ApprovedBy nếu có)
                status = rec.get("Status", "")
                if rec.get("ApprovedBy"):
                    status = f"Đã duyệt bởi {rec['ApprovedBy']}"
                
                # 4. Tạo chuỗi entry theo format yêu cầu: {CheckInTime}; {Task}: {Lý do}; {Status}
                entry = f"{full_datetime_str}; {leave_task}: {leave_reason}; {status}"

                ws_leaves.cell(row=row, column=3 + j, value=entry)
                
            # Áp dụng style và tính chiều cao dòng
            for col in range(1, 14):
                cell = ws_leaves.cell(row=row, column=col)
                cell.border = border
                cell.alignment = align_left
            max_lines = max(
                (str(ws_leaves.cell(row=row, column=col).value).count("\n") + 1 if ws_leaves.cell(row=row, column=col).value else 1)
                for col in range(1, 14)
            )
            ws_leaves.row_dimensions[row].height = max_lines * 20
            
        # Tự động điều chỉnh độ rộng cột
        for col in ws_leaves.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    length = len(str(cell.value).split("\n")[0])
                    max_length = max(max_length, length)
            ws_leaves.column_dimensions[col_letter].width = min(max_length + 2, 70)
            
        # Xuất file
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
