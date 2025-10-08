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

# Các collection sử dụng
admins = db["admins"]
users = db["users"]
collection = db["alt_checkins"]


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
            "success": True,
            "message": "✅ Đăng nhập thành công",
            "username": admin["username"],
            "email": admin["email"],
            "role": "admin"
        })

    user = users.find_one({"email": email})
    if user and check_password_hash(user.get("password", ""), password):
        return jsonify({
            "success": True,
            "message": "✅ Đăng nhập thành công",
            "username": user["username"],
            "email": user["email"],
            "role": "user"
        })

    return jsonify({"success": False, "message": "🚫 Email hoặc mật khẩu không đúng!"}), 401


# ---- Reset mật khẩu ----
@app.route("/forgot-password", methods=["GET", "POST"])
def forgot_password():
    if request.method == "GET":
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
                .success { color: #28a745; text-align: center; }
            </style>
        </head>
        <body>
            <div class="container">
                <h2>🔒 Đặt lại mật khẩu</h2>
                <form method="POST">
                    <input type="email" name="email" placeholder="Email" required>
                    <input type="password" name="new_password" placeholder="Mật khẩu mới" required>
                    <input type="password" name="confirm_password" placeholder="Xác nhận mật khẩu" required>
                    <button type="submit">Cập nhật mật khẩu</button>
                    <a href="/">Quay về trang chủ</a>
                </form>
            </div>
        </body>
        </html>
        """
    if request.method == "POST":
        email = request.form.get("email")
        new_password = request.form.get("new_password")
        confirm_password = request.form.get("confirm_password")

        if not email or not new_password or not confirm_password:
            return jsonify({"success": False, "message": "❌ Vui lòng điền đầy đủ thông tin"}), 400

        if new_password != confirm_password:
            return jsonify({"success": False, "message": "❌ Mật khẩu xác nhận không khớp"}), 400

        admin = admins.find_one({"email": email})
        user = None
        if not admin:
            user = users.find_one({"email": email})
            if not user:
                return jsonify({"success": False, "message": "🚫 Email không tồn tại!"}), 404

        # Update password
        hashed_pw = generate_password_hash(new_password)
        if admin:
            admins.update_one({"email": email}, {"$set": {"password": hashed_pw}})
        else:
            users.update_one({"email": email}, {"$set": {"password": hashed_pw}})

        # Hiển thị thông báo thành công với nút quay về trang chủ
        return """
        <!DOCTYPE html>
        <html lang="vi">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Thay đổi mật khẩu thành công</title>
            <style>
                body { font-family: Arial, sans-serif; background: #f4f6f9; margin: 0; padding: 20px; }
                .container { max-width: 400px; margin: 100px auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
                .success { color: #28a745; text-align: center; font-size: 18px; margin-bottom: 20px; }
                button { background: #28a745; color: white; padding: 12px; width: 100%; border: none; border-radius: 4px; cursor: pointer; font-size: 16px; }
                button:hover { background: #218838; }
            </style>
        </head>
        <body>
            <div class="container">
                <div class="success">✅ Thay đổi mật khẩu thành công</div>
                <a href="/"><button>Quay về trang chủ</button></a>
            </div>
        </body>
        </html>
        """


# ---- Build attendance query ----
def build_attendance_query(filter_type, start_date, end_date, search, username=None):
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

    if username:
        conditions.append({"EmployeeName": username})

    if not conditions:
        return {}
    elif len(conditions) == 1:
        return conditions[0]
    else:
        return {"$and": conditions}


# ---- Build leave query ----
def build_leave_query(filter_type, start_date, end_date, search, username=None):
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

    if username:
        conditions.append({"EmployeeName": username})

    if len(conditions) == 1:
        return conditions[0]
    else:
        return {"$and": conditions}

# ---- Helper function để tính số ngày nghỉ ----
def calculate_leave_days(task_string):
    if not isinstance(task_string, str):
        return 1.0  # Mặc định là 1 nếu dữ liệu không phải chuỗi

    task_string_lower = task_string.lower()
    num_days = 1.0

    # 1. Xác định hệ số nhân theo buổi (sáng/chiều/cả ngày)
    multiplier = 1.0
    if 'sáng' in task_string_lower or 'chieu' in task_string_lower or 'chiều' in task_string_lower:
        multiplier = 0.5
    
    # 2. Trích xuất ngày để tính khoảng thời gian nghỉ
    date_pattern = r'\d{2}/\d{2}/\d{4}'
    dates_found = re.findall(date_pattern, task_string)
    
    # Kiểm tra nếu là nghỉ nhiều ngày (có từ "đến")
    if 'đến' in task_string_lower and len(dates_found) >= 2:
        try:
            start_date_str = dates_found[0]
            end_date_str = dates_found[1]
            
            start_date = datetime.strptime(start_date_str, "%d/%m/%Y")
            end_date = datetime.strptime(end_date_str, "%d/%m/%Y")
            
            # Tính số ngày (bao gồm cả ngày bắt đầu và kết thúc)
            delta = (end_date - start_date).days + 1
            if delta > 0:
                num_days = float(delta)
        except (ValueError, IndexError) as e:
            # Nếu không parse được ngày, quay về mặc định
            print(f"Không thể phân tích ngày từ '{task_string}': {e}")
            num_days = 1.0
            
    # Tính toán cuối cùng
    total_leave_days = num_days * multiplier
    return total_leave_days

# ---- API lấy dữ liệu chấm công ----
@app.route("/api/attendances", methods=["GET"])
def get_attendances():
    try:
        email = request.args.get("email")
        if not email:
            return jsonify({"error": "❌ Thiếu email"}), 400

        admin = admins.find_one({"email": email}, {"_id": 0, "username": 1})
        if admin:
            username = None
            log_msg = "admin"
        else:
            user = users.find_one({"email": email}, {"_id": 0, "username": 1})
            if not user:
                return jsonify({"error": "🚫 Email không tồn tại"}), 403
            username = user["username"]
            log_msg = f"user: {username}"

        filter_type = request.args.get("filter", "hôm nay").lower()
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()

        query = build_attendance_query(filter_type, start_date, end_date, search, username=username)
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

        print(f"DEBUG: Fetched {len(data)} records for email {email} ({log_msg}) with filter {filter_type}")
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
        if admin:
            username = None
            log_msg = "admin"
        else:
            user = users.find_one({"email": email}, {"_id": 0, "username": 1})
            if not user:
                return jsonify({"error": "🚫 Email không tồn tại"}), 403
            username = user["username"]
            log_msg = f"user: {username}"

        filter_type = request.args.get("filter", "tất cả").lower()
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()

        query = build_leave_query(filter_type, start_date, end_date, search, username=username)
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

        print(f"DEBUG: Fetched {len(data)} leave records for email {email} ({log_msg}) with filter {filter_type}")
        return jsonify(data)
    except Exception as e:
        print(f"❌ Error in get_leaves: {e}")
        return jsonify({"error": str(e)}), 500


# ---- API xuất Excel ----
@app.route("/api/export-excel", methods=["GET"])
def export_to_excel():
    try:
        email = request.args.get("email")
        if not email:
            return jsonify({"error": "❌ Thiếu email"}), 400

        admin = admins.find_one({"email": email}, {"_id": 0, "username": 1})
        if admin:
            username = None
        else:
            user = users.find_one({"email": email}, {"_id": 0, "username": 1})
            if not user:
                return jsonify({"error": "🚫 Email không tồn tại"}), 403
            username = user["username"]

        filter_type = request.args.get("filter", "hôm nay").lower()
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()

        query = build_attendance_query(filter_type, start_date, end_date, search, username=username)
        data = list(collection.find(query, {
            "_id": 0, "EmployeeId": 1, "EmployeeName": 1, "ProjectId": 1,
            "Tasks": 1, "OtherNote": 1, "Address": 1, "CheckinTime": 1,
            "CheckinDate": 1, "Status": 1, "ApprovedBy": 1, "Latitude": 1, "Longitude": 1
        }))

        grouped = {}
        for d in data:
            emp_id = d.get("EmployeeId", "")
            emp_name = d.get("EmployeeName", "")
            date = d.get("CheckinDate")
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
                if time_str:
                    parts.append(time_str)
                if rec.get("ProjectId"):
                    parts.append(str(rec["ProjectId"]))
                tasks = rec.get("Tasks")
                if tasks:
                    tasks_str = ", ".join(tasks) if isinstance(tasks, list) else str(tasks)
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

        today_str = datetime.now(VN_TZ).strftime("%d-%m-%Y")
        if search:
            filename = f"Danh sách chấm công theo tìm kiếm_{today_str}.xlsx"
        elif filter_type == "hôm nay":
            filename = f"Danh sách chấm công_{today_str}.xlsx"
        elif filter_type == "custom" and start_date and end_date:
            filename = f"Danh sách chấm công từ {start_date} đến {end_date}_{today_str}.xlsx"
        else:
            filename = f"Danh sách chấm công_{today_str}.xlsx"

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


# ---- API xuất Excel cho nghỉ phép ----
@app.route("/api/export-leaves-excel", methods=["GET"])
def export_leaves_to_excel():
    try:
        email = request.args.get("email")
        if not email:
            return jsonify({"error": "❌ Thiếu email"}), 400

        admin = admins.find_one({"email": email})
        if admin:
            username = None
        else:
            user = users.find_one({"email": email})
            if not user:
                return jsonify({"error": "🚫 Email không tồn tại"}), 403
            username = user["username"]

        filter_type = request.args.get("filter", "tất cả").lower()
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()

        query = build_leave_query(filter_type, start_date, end_date, search, username=username)
        data = list(collection.find(query, {
            "_id": 0, "EmployeeId": 1, "EmployeeName": 1, "CheckinDate": 1,
            "CheckinTime": 1, "ApprovalDate": 1, "Tasks": 1, "Status": 1,
            "ApprovedBy": 1, "ApproveNote": 1
        }))

        template_path = "templates/Copy of Form nghỉ phép.xlsx"
        wb = load_workbook(template_path)
        ws = wb.active
        
        border = Border(
            left=Side(style="thin", color="000000"), right=Side(style="thin", color="000000"),
            top=Side(style="thin", color="000000"), bottom=Side(style="thin", color="000000"),
        )
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        start_row = 2

        for i, rec in enumerate(data, start=0):
            row = start_row + i
            
            # Cột 1: Mã NV
            ws.cell(row=row, column=1, value=rec.get("EmployeeId", ""))
            
            # Cột 2: Tên NV
            ws.cell(row=row, column=2, value=rec.get("EmployeeName", ""))
            
            # Cột 3: Ngày Nghỉ
            ws.cell(row=row, column=3, value=rec.get("CheckinDate", ""))
            
            # Lấy chuỗi task để xử lý
            tasks = rec.get("Tasks")
            tasks_str = ", ".join(tasks) if isinstance(tasks, list) else str(tasks or "")

            # Cột 4: Số ngày nghỉ (Tính toán tự động)
            leave_days = calculate_leave_days(tasks_str)
            ws.cell(row=row, column=4, value=leave_days)
            
            # Cột 5: Ngày tạo đơn
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
            ws.cell(row=row, column=5, value=full_datetime_str)
            
            # Cột 6: Lý do
            leave_reason = ""
            if ":" in tasks_str:
                leave_reason = tasks_str.split(":", 1)[1].strip()
            else:
                leave_reason = rec.get("ApproveNote", "") or ""
            ws.cell(row=row, column=6, value=leave_reason)

            # Cột 7: Trạng thái
            approved_by = rec.get("ApprovedBy", "")
            approval_date = rec.get("ApprovalDate")
            approval_date_str = ""
            if isinstance(approval_date, datetime):
                approval_date_str = approval_date.astimezone(VN_TZ).strftime("%d/%m/%Y %H:%M:%S")
            elif isinstance(approval_date, str) and approval_date.strip():
                 approval_date_str = approval_date
            
            approval_status = f"Đã duyệt bởi {approved_by} lúc {approval_date_str}" if approved_by and approval_date_str else "Chưa duyệt"
            ws.cell(row=row, column=7, value=approval_status)

            # Áp dụng style cho các ô
            for col_idx in range(1, 8):
                cell = ws.cell(row=row, column=col_idx)
                cell.border = border
                cell.alignment = align_left
        
        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            adjusted_width = max_length + 2
            ws.column_dimensions[col_letter].width = adjusted_width

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


# ---- API xuất Excel kết hợp chấm công và nghỉ phép ----
@app.route("/api/export-combined-excel", methods=["GET"])
def export_combined_to_excel():
    try:
        email = request.args.get("email")
        if not email:
            return jsonify({"error": "❌ Thiếu email"}), 400
        
        admin = admins.find_one({"email": email})
        if admin:
            username = None
        else:
            user = users.find_one({"email": email})
            if not user:
                return jsonify({"error": "🚫 Email không tồn tại"}), 403
            username = user["username"]

        filter_type = request.args.get("filter", "hôm nay").lower()
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()

        attendance_query = build_attendance_query(filter_type, start_date, end_date, search, username=username)
        leave_query = build_leave_query(filter_type, start_date, end_date, search, username=username)

        attendance_data = list(collection.find(attendance_query, {
            "_id": 0, "EmployeeId": 1, "EmployeeName": 1, "ProjectId": 1, "Tasks": 1,
            "OtherNote": 1, "Address": 1, "CheckinTime": 1, "CheckinDate": 1,
            "Status": 1
        }))
        leave_data = list(collection.find(leave_query, {
            "_id": 0, "EmployeeId": 1, "EmployeeName": 1, "CheckinDate": 1, "CheckinTime": 1,
            "Tasks": 1, "Status": 1, "ApprovedBy": 1, "ApproveNote": 1, "ApprovalDate": 1
        }))

        template_path = "templates/Form kết hợp.xlsx"
        wb = load_workbook(template_path)

        # ---- Xử lý sheet Điểm danh ----
        ws_attendance = wb["Điểm danh"] if "Điểm danh" in wb.sheetnames else wb.create_sheet("Điểm danh")
        attendance_grouped = {}
        for d in attendance_data:
            key = (d.get("EmployeeId", ""), d.get("EmployeeName", ""), d.get("CheckinDate", ""))
            attendance_grouped.setdefault(key, []).append(d)

        border = Border(
            left=Side(style="thin", color="000000"), right=Side(style="thin", color="000000"),
            top=Side(style="thin", color="000000"), bottom=Side(style="thin", color="000000"),
        )
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        
        start_row_att = 2
        for i, ((emp_id, emp_name, date), records) in enumerate(attendance_grouped.items()):
            row = start_row_att + i
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
                if time_str: parts.append(time_str)
                if rec.get("ProjectId"): parts.append(str(rec["ProjectId"]))
                if rec.get("Tasks"): parts.append(str(rec["Tasks"]))
                if rec.get("Status"): parts.append(rec["Status"])
                if rec.get("OtherNote"): parts.append(rec["OtherNote"])
                if rec.get("Address"): parts.append(rec["Address"])
                entry = "; ".join(parts)
                ws_attendance.cell(row=row, column=3 + j, value=entry)

            for col in range(1, 14):
                cell = ws_attendance.cell(row=row, column=col)
                cell.border = border
                cell.alignment = align_left

        # ---- Xử lý sheet Nghỉ phép ----
        ws_leaves = wb["Nghỉ phép"] if "Nghỉ phép" in wb.sheetnames else wb.create_sheet("Nghỉ phép")
        
        leave_headers = ["Mã NV", "Tên NV", "Ngày Nghỉ", "Số ngày nghỉ", "Ngày tạo đơn", "Lý do", "Trạng thái"]
        for col, header in enumerate(leave_headers, start=1):
            ws_leaves.cell(row=1, column=col, value=header)

        start_row_leaves = 2
        for i, rec in enumerate(leave_data, start=0):
            row = start_row_leaves + i
            
            ws_leaves.cell(row=row, column=1, value=rec.get("EmployeeId"))
            ws_leaves.cell(row=row, column=2, value=rec.get("EmployeeName"))
            ws_leaves.cell(row=row, column=3, value=rec.get("CheckinDate"))
            
            tasks = rec.get("Tasks")
            tasks_str = ", ".join(tasks) if isinstance(tasks, list) else str(tasks or "")
            
            leave_days = calculate_leave_days(tasks_str)
            ws_leaves.cell(row=row, column=4, value=leave_days)
            
            checkin_time = rec.get("CheckinTime")
            full_datetime_str = ""
            if isinstance(checkin_time, datetime):
                full_datetime_str = checkin_time.astimezone(VN_TZ).strftime("%d/%m/%Y %H:%M:%S")
            elif isinstance(checkin_time, str) and checkin_time.strip():
                 full_datetime_str = checkin_time
            ws_leaves.cell(row=row, column=5, value=full_datetime_str)
            
            leave_reason = tasks_str.split(":", 1)[1].strip() if ":" in tasks_str else rec.get("ApproveNote", "")
            ws_leaves.cell(row=row, column=6, value=leave_reason)

            approved_by = rec.get("ApprovedBy", "")
            approval_date = rec.get("ApprovalDate")
            approval_date_str = ""
            if isinstance(approval_date, datetime):
                approval_date_str = approval_date.astimezone(VN_TZ).strftime("%d/%m/%Y %H:%M:%S")
            elif isinstance(approval_date, str) and approval_date.strip():
                approval_date_str = approval_date
            approval_status = f"Đã duyệt bởi {approved_by} lúc {approval_date_str}" if approved_by and approval_date_str else "Chưa duyệt"
            ws_leaves.cell(row=row, column=7, value=approval_status)

            for col in range(1, 8):
                cell = ws_leaves.cell(row=row, column=col)
                cell.border = border
                cell.alignment = align_left

        for ws in [ws_attendance, ws_leaves]:
            for col in ws.columns:
                max_length = 0
                col_letter = col[0].column_letter
                for cell in col:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                adjusted_width = min(max_length + 2, 70)
                ws.column_dimensions[col_letter].width = adjusted_width

        today_str = datetime.now(VN_TZ).strftime("%d-%m-%Y")
        filename = f"Báo cáo chấm công và nghỉ phép_{today_str}.xlsx"
        if search:
            filename = f"Báo cáo theo tìm kiếm_{today_str}.xlsx"
        elif filter_type == "custom" and start_date and end_date:
            filename = f"Báo cáo từ {start_date} đến {end_date}_{today_str}.xlsx"

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
