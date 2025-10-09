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

        hashed_pw = generate_password_hash(new_password)
        if admin:
            admins.update_one({"email": email}, {"$set": {"password": hashed_pw}})
        else:
            users.update_one({"email": email}, {"$set": {"password": hashed_pw}})

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


# ---- Build leave query (lọc theo CheckinTime)----
def build_leave_query(filter_type, start_date, end_date, search, username=None):
    today = datetime.now(VN_TZ)
    regex_leave = re.compile("Nghỉ phép", re.IGNORECASE)
    conditions = [{"$or": [{"Tasks": {"$regex": regex_leave}}, {"Status": {"$regex": regex_leave}}]}]

    date_filter = {}
    
    def get_time_range_str(start_dt_obj, end_dt_obj):
        start_str = start_dt_obj.strftime("%d/%m/%Y %H:%M:%S")
        end_str = end_dt_obj.strftime("%d/%m/%Y %H:%M:%S")
        return {"CheckinTime": {"$gte": start_str, "$lte": end_str}}

    if filter_type == "custom" and start_date and end_date:
        start_dt = datetime.strptime(start_date, "%Y-%m-%d")
        end_dt = datetime.strptime(end_date, "%Y-%m-%d").replace(hour=23, minute=59, second=59)
        date_filter = get_time_range_str(start_dt, end_dt)
    
    elif filter_type == "hôm nay":
        start_dt = today.replace(hour=0, minute=0, second=0, microsecond=0)
        end_dt = today.replace(hour=23, minute=59, second=59, microsecond=999)
        date_filter = get_time_range_str(start_dt, end_dt)

    elif filter_type == "tuần":
        start_dt = (today - timedelta(days=today.weekday())).replace(hour=0, minute=0, second=0)
        end_dt = (start_dt + timedelta(days=6)).replace(hour=23, minute=59, second=59)
        date_filter = get_time_range_str(start_dt, end_dt)

    elif filter_type == "tháng":
        start_dt = today.replace(day=1, hour=0, minute=0, second=0)
        _, last_day = calendar.monthrange(today.year, today.month)
        end_dt = today.replace(day=last_day, hour=23, minute=59, second=59)
        date_filter = get_time_range_str(start_dt, end_dt)
        
    elif filter_type == "năm":
        start_dt = today.replace(month=1, day=1, hour=0, minute=0, second=0)
        end_dt = today.replace(month=12, day=31, hour=23, minute=59, second=59)
        date_filter = get_time_range_str(start_dt, end_dt)

    if date_filter:
        conditions.append(date_filter)

    if search:
        regex = re.compile(search, re.IGNORECASE)
        conditions.append({"$or": [{"EmployeeId": {"$regex": regex}}, {"EmployeeName": {"$regex": regex}}]})

    if username:
        conditions.append({"EmployeeName": username})
    
    return {"$and": conditions} if len(conditions) > 1 else conditions[0]


# ---- Helper function để tính số ngày nghỉ từ record ----
def calculate_leave_days_from_record(record):
    if 'StartDate' in record and 'EndDate' in record and record.get('StartDate') and record.get('EndDate'):
        try:
            start_date = datetime.strptime(record['StartDate'], "%Y-%m-%d")
            end_date = datetime.strptime(record['EndDate'], "%Y-%m-%d")
            delta = (end_date - start_date).days + 1
            return float(delta)
        except (ValueError, TypeError):
            return 1.0

    if 'LeaveDate' in record and record.get('LeaveDate'):
        session = record.get('Session', '').lower()
        if 'sáng' in session or 'chiều' in session:
            return 0.5
        return 1.0

    return 1.0

# ---- Helper function để format ngày duyệt ----
def get_formatted_approval_date(approval_date):
    if not approval_date:
        return ""
    
    formatted_date = ""
    try:
        if isinstance(approval_date, str):
            if 'Z' in approval_date:
                parsed_date = datetime.fromisoformat(approval_date.replace('Z', '+00:00'))
            else:
                parsed_date = datetime.strptime(approval_date, "%d/%m/%Y %H:%M:%S")
        elif isinstance(approval_date, datetime):
            parsed_date = approval_date
        else:
            return ""
        
        formatted_date = parsed_date.astimezone(VN_TZ).strftime("%d/%m/%Y %H:%M:%S")
    except (ValueError, TypeError):
        return str(approval_date)
        
    return formatted_date

# ---- API lấy dữ liệu chấm công ----
@app.route("/api/attendances", methods=["GET"])
def get_attendances():
    try:
        email = request.args.get("email")
        if not email:
            return jsonify({"error": "❌ Thiếu email"}), 400

        admin = admins.find_one({"email": email})
        username = None if admin else users.find_one({"email": email})["username"]
        if not admin and not username:
            return jsonify({"error": "🚫 Email không tồn tại"}), 403

        filter_type = request.args.get("filter", "hôm nay").lower()
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()

        query = build_attendance_query(filter_type, start_date, end_date, search, username=username)
        data = list(collection.find(query, {"_id": 0}))

        for item in data:
            ghi_chu_parts = []
            if item.get('ProjectId'): ghi_chu_parts.append(f"Project: {item['ProjectId']}")
            if item.get('Tasks'):
                tasks_str = ', '.join(item['Tasks']) if isinstance(item['Tasks'], list) else str(item['Tasks'])
                ghi_chu_parts.append(f"Tasks: {tasks_str}")
            if item.get('OtherNote'): ghi_chu_parts.append(f"Note: {item['OtherNote']}")
            item['GhiChu'] = '; '.join(ghi_chu_parts) if ghi_chu_parts else ''
        
        return jsonify(data)
    except Exception as e:
        print(f"❌ Lỗi tại get_attendances: {e}")
        return jsonify({"error": str(e)}), 500


# ---- API lấy dữ liệu nghỉ phép ----
@app.route("/api/leaves", methods=["GET"])
def get_leaves():
    try:
        email = request.args.get("email")
        if not email:
            return jsonify({"error": "❌ Thiếu email"}), 400

        admin = admins.find_one({"email": email})
        username = None if admin else users.find_one({"email": email})["username"]
        if not admin and not username:
            return jsonify({"error": "🚫 Email không tồn tại"}), 403

        filter_type = request.args.get("filter", "tất cả").lower()
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()

        query = build_leave_query(filter_type, start_date, end_date, search, username=username)
        
        data = list(collection.find(query, {
            "_id": 0, "EmployeeId": 1, "EmployeeName": 1, "CheckinDate": 1,
            "CheckinTime": 1, "Tasks": 1, "Status": 1, "ApprovalDate": 1,
            "ApprovedBy": 1, "ApproveNote": 1, "StartDate": 1, "EndDate": 1,
            "LeaveDate": 1, "Session": 1
        }))

        for item in data:
            item["ApprovalDate"] = get_formatted_approval_date(item.get("ApprovalDate"))

        return jsonify(data)
    except Exception as e:
        print(f"❌ Lỗi tại get_leaves: {e}")
        return jsonify({"error": str(e)}), 500


# ---- API xuất Excel Chấm công ----
@app.route("/api/export-excel", methods=["GET"])
def export_to_excel():
    try:
        email = request.args.get("email")
        if not email:
            return jsonify({"error": "❌ Thiếu email"}), 400

        admin = admins.find_one({"email": email})
        username = None if admin else users.find_one({"email": email})["username"]
        if not admin and not username:
            return jsonify({"error": "🚫 Email không tồn tại"}), 403

        filter_type = request.args.get("filter", "hôm nay").lower()
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()

        query = build_attendance_query(filter_type, start_date, end_date, search, username=username)
        data = list(collection.find(query, {"_id": 0}))

        grouped = {}
        for d in data:
            key = (d.get("EmployeeId", ""), d.get("EmployeeName", ""), d.get("CheckinDate"))
            grouped.setdefault(key, []).append(d)

        template_path = "templates/Copy of Form chấm công.xlsx"
        wb = load_workbook(template_path)
        ws = wb.active

        border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
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
                elif isinstance(checkin_time, str):
                    try:
                        time_str = datetime.strptime(checkin_time, "%d/%m/%Y %H:%M:%S").strftime("%H:%M:%S")
                    except ValueError:
                        time_str = checkin_time

                parts = [part for part in [
                    time_str,
                    rec.get("ProjectId"),
                    ", ".join(rec.get("Tasks", [])) if isinstance(rec.get("Tasks"), list) else rec.get("Tasks"),
                    rec.get("Status"),
                    rec.get("OtherNote"),
                    rec.get("Address")
                ] if part]
                ws.cell(row=row, column=3 + j, value="; ".join(map(str, parts)))

            for col in range(1, 14):
                cell = ws.cell(row=row, column=col)
                cell.border = border
                cell.alignment = align_left

        today_str = datetime.now(VN_TZ).strftime("%d-%m-%Y")
        filename = f"Danh sách chấm công_{filter_type}_{today_str}.xlsx"

        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return send_file(output, as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        print("❌ Lỗi export:", e)
        return jsonify({"error": str(e)}), 500


# ---- API xuất Excel cho nghỉ phép (FIXED) ----
@app.route("/api/export-leaves-excel", methods=["GET"])
def export_leaves_to_excel():
    try:
        email = request.args.get("email")
        if not email: return jsonify({"error": "❌ Thiếu email"}), 400

        admin = admins.find_one({"email": email})
        username = None if admin else users.find_one({"email": email})["username"]
        if not admin and not username: return jsonify({"error": "🚫 Email không tồn tại"}), 403

        filter_type = request.args.get("filter", "tất cả").lower()
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()

        query = build_leave_query(filter_type, start_date, end_date, search, username=username)
        data = list(collection.find(query, {"_id": 0}))

        template_path = "templates/Copy of Form nghỉ phép.xlsx"
        wb = load_workbook(template_path)
        ws = wb.active
        
        # Cập nhật tiêu đề để khớp với file mẫu mới
        ws['E1'] = "Ngày tạo đơn"
        ws['F1'] = "Lý do"
        ws['G1'] = "Trạng thái"
        ws['H1'] = "Ngày duyệt đơn"

        border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        start_row = 2

        for i, rec in enumerate(data, start=0):
            row = start_row + i
            
            # Cột 1-4: Thông tin cơ bản
            ws.cell(row=row, column=1, value=rec.get("EmployeeId", ""))
            ws.cell(row=row, column=2, value=rec.get("EmployeeName", ""))
            ws.cell(row=row, column=3, value=rec.get("CheckinDate", ""))
            ws.cell(row=row, column=4, value=calculate_leave_days_from_record(rec))
            
            # Cột 5: Ngày tạo đơn
            ws.cell(row=row, column=5, value=rec.get("CheckinTime", ""))
            
            # Cột 6: Lý do
            tasks = rec.get("Tasks", [])
            tasks_str = (", ".join(tasks) if isinstance(tasks, list) else str(tasks or "")).replace("Nghỉ phép: ", "")
            ws.cell(row=row, column=6, value=tasks_str)
            
            # Cột 7: Trạng thái
            ws.cell(row=row, column=7, value=rec.get("Status", "Chưa duyệt"))

            # Cột 8: Ngày duyệt đơn
            ws.cell(row=row, column=8, value=get_formatted_approval_date(rec.get("ApprovalDate")))

            # Áp dụng style
            for col_idx in range(1, 9):
                ws.cell(row=row, column=col_idx).border = border
                ws.cell(row=row, column=col_idx).alignment = align_left
        
        today_str = datetime.now(VN_TZ).strftime("%d-%m-%Y")
        filename = f"Danh sách nghỉ phép_{filter_type}_{today_str}.xlsx"

        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return send_file(output, as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        print("❌ Lỗi export leaves:", e)
        return jsonify({"error": str(e)}), 500


# ---- API xuất Excel kết hợp (FIXED) ----
@app.route("/api/export-combined-excel", methods=["GET"])
def export_combined_to_excel():
    try:
        email = request.args.get("email")
        if not email: return jsonify({"error": "❌ Thiếu email"}), 400

        admin = admins.find_one({"email": email})
        username = None if admin else users.find_one({"email": email})["username"]
        if not admin and not username: return jsonify({"error": "🚫 Email không tồn tại"}), 403

        filter_type = request.args.get("filter", "hôm nay").lower()
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()

        attendance_query = build_attendance_query(filter_type, start_date, end_date, search, username=username)
        leave_query = build_leave_query(filter_type, start_date, end_date, search, username=username)

        attendance_data = list(collection.find(attendance_query, {"_id": 0}))
        leave_data = list(collection.find(leave_query, {"_id": 0}))
        
        template_path = "templates/Form kết hợp.xlsx"
        wb = load_workbook(template_path)
        border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

        # ---- Xử lý sheet Điểm danh ----
        ws_attendance = wb["Điểm danh"]
        attendance_grouped = {}
        for d in attendance_data:
            key = (d.get("EmployeeId", ""), d.get("EmployeeName", ""), d.get("CheckinDate"))
            attendance_grouped.setdefault(key, []).append(d)
        
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
                elif isinstance(checkin_time, str):
                    try:
                        time_str = datetime.strptime(checkin_time, "%d/%m/%Y %H:%M:%S").strftime("%H:%M:%S")
                    except ValueError:
                        time_str = checkin_time
                parts = [part for part in [time_str, rec.get("ProjectId"), ", ".join(rec.get("Tasks",[])) if isinstance(rec.get("Tasks"), list) else rec.get("Tasks"), rec.get("Status"), rec.get("OtherNote"), rec.get("Address")] if part]
                ws_attendance.cell(row=row, column=3 + j, value="; ".join(map(str, parts)))

            for col in range(1, 14):
                 ws_attendance.cell(row=row, column=col).border = border
                 ws_attendance.cell(row=row, column=col).alignment = align_left

        # ---- Xử lý sheet Nghỉ phép ----
        ws_leaves = wb["Nghỉ phép"]
        ws_leaves['E1'] = "Ngày tạo đơn"
        ws_leaves['F1'] = "Lý do"
        ws_leaves['G1'] = "Trạng thái"
        ws_leaves['H1'] = "Ngày duyệt đơn"
        
        start_row_leaves = 2
        for i, rec in enumerate(leave_data, start=0):
            row = start_row_leaves + i
            
            ws_leaves.cell(row=row, column=1, value=rec.get("EmployeeId"))
            ws_leaves.cell(row=row, column=2, value=rec.get("EmployeeName"))
            ws_leaves.cell(row=row, column=3, value=rec.get("CheckinDate"))
            ws_leaves.cell(row=row, column=4, value=calculate_leave_days_from_record(rec))
            ws_leaves.cell(row=row, column=5, value=rec.get("CheckinTime", ""))
            
            tasks = rec.get("Tasks", [])
            tasks_str = (", ".join(tasks) if isinstance(tasks, list) else str(tasks or "")).replace("Nghỉ phép: ", "")
            ws_leaves.cell(row=row, column=6, value=tasks_str)
            
            ws_leaves.cell(row=row, column=7, value=rec.get("Status", "Chưa duyệt"))
            ws_leaves.cell(row=row, column=8, value=get_formatted_approval_date(rec.get("ApprovalDate")))

            for col in range(1, 9):
                ws_leaves.cell(row=row, column=col).border = border
                ws_leaves.cell(row=row, column=col).alignment = align_left

        today_str = datetime.now(VN_TZ).strftime("%d-%m-%Y")
        filename = f"Báo cáo chấm công và nghỉ phép_{filter_type}_{today_str}.xlsx"

        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return send_file(output, as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        print("❌ Lỗi export combined:", e)
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
