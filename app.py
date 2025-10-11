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

# ---- Reset mật khẩu ----
@app.route("/forgot-password", methods=["GET", "POST"])
def forgot_password():
    if request.method == "GET":
        return """
        <!DOCTYPE html><html lang="vi"><head><meta charset="UTF-8"><title>Đặt lại mật khẩu</title>
        <style>body{font-family:Arial,sans-serif;background:#f4f6f9;padding:20px}.container{max-width:400px;margin:100px auto;background:white;padding:30px;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,.1)}input{width:100%;padding:10px;margin:10px 0;box-sizing:border-box;border:1px solid #ddd;border-radius:4px}button{background:#28a745;color:white;padding:12px;width:100%;border:none;border-radius:4px;cursor:pointer;font-size:16px}</style>
        </head><body><div class="container"><h2>🔒 Đặt lại mật khẩu</h2><form method="POST">
        <input type="email" name="email" placeholder="Email" required>
        <input type="password" name="new_password" placeholder="Mật khẩu mới" required>
        <input type="password" name="confirm_password" placeholder="Xác nhận mật khẩu" required>
        <button type="submit">Cập nhật mật khẩu</button><a href="/">Quay về trang chủ</a></form></div></body></html>"""
    if request.method == "POST":
        email = request.form.get("email")
        new_password = request.form.get("new_password")
        confirm_password = request.form.get("confirm_password")
        if not all([email, new_password, confirm_password]):
            return jsonify({"success": False, "message": "❌ Vui lòng điền đầy đủ thông tin"}), 400
        if new_password != confirm_password:
            return jsonify({"success": False, "message": "❌ Mật khẩu xác nhận không khớp"}), 400
        account = admins.find_one({"email": email}) or users.find_one({"email": email})
        if not account:
            return jsonify({"success": False, "message": "🚫 Email không tồn tại!"}), 404
        hashed_pw = generate_password_hash(new_password)
        collection_to_update = admins if "username" in account else users
        collection_to_update.update_one({"email": email}, {"$set": {"password": hashed_pw}})
        return """
        <!DOCTYPE html><html lang="vi"><head><title>Thay đổi mật khẩu thành công</title>
        <style>body{font-family:Arial,sans-serif;background:#f4f6f9;padding:20px}.container{max-width:400px;margin:100px auto;background:white;padding:30px;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,.1)}.success{color:#28a745;text-align:center;font-size:18px;margin-bottom:20px}button{background:#28a745;color:white;padding:12px;width:100%;border:none;border-radius:4px;cursor:pointer;font-size:16px}</style>
        </head><body><div class="container"><div class="success">✅ Thay đổi mật khẩu thành công</div>
        <a href="/"><button>Quay về trang chủ</button></a></div></body></html>"""

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

# ---- Build leave query (lọc theo Timestamp)----
def build_leave_query(filter_type, start_date_str, end_date_str, search, username=None):
    today = datetime.now(VN_TZ)
    regex_leave = re.compile("Nghỉ phép", re.IGNORECASE)
    conditions = [{"$or": [{"Tasks": regex_leave}, {"Reason": {"$exists": True}}]}]
    date_filter = {}
    if filter_type == "custom" and start_date_str and end_date_str:
        start_dt = datetime.strptime(start_date_str, "%Y-%m-%d").replace(tzinfo=VN_TZ)
        end_dt = datetime.strptime(end_date_str, "%Y-%m-%d").replace(hour=23, minute=59, second=59, tzinfo=VN_TZ)
        date_filter = {"Timestamp": {"$gte": start_dt, "$lte": end_dt}}
    else:
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
        if filter_type != "tất cả":
            date_filter = {"Timestamp": {"$gte": start_dt, "$lte": end_dt}}
   
    if date_filter: conditions.append(date_filter)
    if search:
        regex = re.compile(search, re.IGNORECASE)
        conditions.append({"$or": [{"EmployeeId": regex}, {"EmployeeName": regex}]})
    if username:
        conditions.append({"EmployeeName": username})
    return {"$and": conditions}

# ---- Helper functions ----
def calculate_leave_days_from_record(record):
    if 'StartDate' in record and 'EndDate' in record:
        try: return float((datetime.strptime(record['EndDate'], "%Y-%m-%d") - datetime.strptime(record['StartDate'], "%Y-%m-%d")).days + 1)
        except: return 1.0
    if 'LeaveDate' in record: return 0.5 if record.get('Session','').lower() in ['sáng', 'chiều'] else 1.0
    return 1.0

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
       
        for item in all_relevant_data:
            emp_id, date_str = item.get("EmployeeId"), item.get("CheckinDate")
            daily_sec = daily_hours_map.get((emp_id, date_str), 0)
            h, rem = divmod(daily_sec, 3600); m, _ = divmod(rem, 60)
            item['DailyHours'], item['_dailySeconds'] = (f"{int(h)}h {int(m)}m" if daily_sec > 0 else ""), daily_sec
            monthly_sec = monthly_hours_map.get((emp_id, date_str), 0)
            h, rem = divmod(monthly_sec, 3600); m, _ = divmod(rem, 60)
            item['MonthlyHours'], item['_monthlySeconds'] = (f"{int(h)}h {int(m)}m" if monthly_sec > 0 else ""), monthly_sec
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
        query = build_leave_query(
            request.args.get("filter", "tất cả").lower(),
            request.args.get("startDate"), request.args.get("endDate"),
            request.args.get("search", "").strip(), username=username
        )
        data = list(collection.find(query, {"_id": 0}))
        for item in data:
            item["ApprovalDate1"] = get_formatted_approval_date(item.get("ApprovalDate1"))
            item["ApprovalDate2"] = get_formatted_approval_date(item.get("ApprovalDate2"))
            item["Status1"] = item.get("Status1", "")
            item["Status2"] = item.get("Status2", "")
            item["Note"] = item.get("LeaveNote", "")  # Giả sử Note là LeaveNote
            if item.get('Timestamp'):
                item['CheckinTime'] = item['Timestamp'].astimezone(VN_TZ).strftime('%d/%m/%Y %H:%M:%S')
           
            if item.get('StartDate') and item.get('EndDate'):
                start = datetime.strptime(item['StartDate'], '%Y-%m-%d').strftime('%d/%m/%Y')
                end = datetime.strptime(item['EndDate'], '%Y-%m-%d').strftime('%d/%m/%Y')
                item['CheckinDate'] = f"Từ {start} đến {end}"
            elif item.get('LeaveDate'):
                leave_date = datetime.strptime(item['LeaveDate'], '%Y-%m-%d').strftime('%d/%m/%Y')
                item['CheckinDate'] = f"{leave_date} ({item.get('Session', '')})"
        return jsonify(data)
    except Exception as e:
        print(f"❌ Lỗi tại get_leaves: {e}")
        return jsonify({"error": str(e)}), 500

# ---- API xuất Excel Chấm công ----
@app.route("/api/export-excel", methods=["GET"])
def export_to_excel():
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
        data = list(collection.find(query, {"_id": 0}))
        grouped = {}
        daily_hours_map, monthly_hours_map = {}, {}
        for d in data:
            key = (d.get("EmployeeId", ""), d.get("EmployeeName", ""), d.get("CheckinDate"))
            grouped.setdefault(key, []).append(d)
            emp_id = d.get("EmployeeId")
            date_str = d.get("CheckinDate")
            if emp_id and date_str:
                daily_groups = {}
                for rec in grouped[key]:
                    if date_str: daily_groups.setdefault(date_str, []).append(rec)
               
                # Handle Timestamp (string or datetime)
                checkins = []
                for r in daily_groups.get(date_str, []):
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
                for r in daily_groups.get(date_str, []):
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
               
                try:
                    month_key = datetime.strptime(date_str, "%d/%m/%Y").strftime("%Y-%m")
                except:
                    continue
                monthly_groups = {}
                monthly_groups.setdefault(month_key, []).append((date_str, daily_seconds))
               
                for month, days in monthly_groups.items():
                    sorted_days = sorted(days, key=lambda x: datetime.strptime(x[0], "%d/%m/%Y"))
                    running_total = 0
                    for date_str, daily_seconds in sorted_days:
                        running_total += daily_seconds
                        monthly_hours_map[(emp_id, date_str)] = running_total
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
            ws.cell(row=row, column=3, value=date_str) # Giữ nguyên format DD/MM/YYYY
           
            daily_sec = daily_hours_map.get((emp_id, date_str), 0)
            h, rem = divmod(daily_sec, 3600); m, _ = divmod(rem, 60)
            daily_hours = f"{int(h)}h {int(m)}m" if daily_sec > 0 else "0h 0m"
            ws.cell(row=row, column=14, value=daily_hours) # Assuming column 14 for DailyHours
           
            monthly_sec = monthly_hours_map.get((emp_id, date_str), 0)
            h, rem = divmod(monthly_sec, 3600); m, _ = divmod(rem, 60)
            monthly_hours = f"{int(h)}h {int(m)}m" if monthly_sec > 0 else "0h 0m"
            ws.cell(row=row, column=15, value=monthly_hours) # Assuming column 15 for MonthlyHours
           
            checkin_counter, checkin_start_col, checkout_col = 0, 4, 13
            sorted_records = sorted(records, key=lambda x: (
                datetime.strptime(x['Timestamp'], "%Y-%m-%d %H:%M:%S")
                if isinstance(x.get('Timestamp'), str) and x.get('Timestamp')
                else x['Timestamp']
                if isinstance(x.get('Timestamp'), datetime)
                else datetime.min
            ))
            for rec in sorted_records:
                time_str = ""
                if rec.get('Timestamp'):
                    try:
                        if isinstance(rec['Timestamp'], str):
                            time_str = datetime.strptime(rec['Timestamp'], "%Y-%m-%d %H:%M:%S").astimezone(VN_TZ).strftime("%H:%M:%S")
                        elif isinstance(rec['Timestamp'], datetime):
                            time_str = rec['Timestamp'].astimezone(VN_TZ).strftime("%H:%M:%S")
                    except (ValueError, TypeError):
                        time_str = ""
                tasks_str = ", ".join(rec.get("Tasks", [])) if isinstance(rec.get("Tasks"), list) else str(rec.get("Tasks", ""))
                cell_value = f"{time_str}; {rec.get('ProjectId','')}; {tasks_str}; {rec.get('Address','')}; {rec.get('CheckinNote','')}"
               
                if rec.get('CheckType') == 'checkin' and checkin_counter < 9:
                    ws.cell(row=row, column=checkin_start_col + checkin_counter, value=cell_value)
                    checkin_counter += 1
                elif rec.get('CheckType') == 'checkout':
                    ws.cell(row=row, column=checkout_col, value=cell_value)
            for col in range(1, 16): # Adjusted to include DailyHours and MonthlyHours columns
                ws.cell(row=row, column=col).border = border
                ws.cell(row=row, column=col).alignment = align_left
        filename = f"Danh sách chấm công_{request.args.get('filter')}_{datetime.now(VN_TZ).strftime('%d-%m-%Y')}.xlsx"
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return send_file(output, as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        print(f"❌ Lỗi export: {e}")
        return jsonify({"error": str(e)}), 500

# ---- API xuất Excel cho nghỉ phép ----
@app.route("/api/export-leaves-excel", methods=["GET"])
def export_leaves_to_excel():
    try:
        email = request.args.get("email")
        if not email: return jsonify({"error": "❌ Thiếu email"}), 400
        admin = admins.find_one({"email": email})
        username = None if admin else users.find_one({"email": email})["username"]
        if not admin and not username: return jsonify({"error": "🚫 Email không tồn tại"}), 403
        query = build_leave_query(
            request.args.get("filter", "tất cả").lower(),
            request.args.get("startDate"), request.args.get("endDate"),
            request.args.get("search", "").strip(), username=username
        )
        data = list(collection.find(query, {"_id": 0}))
        template_path = "templates/Copy of Form nghỉ phép.xlsx"
        wb = load_workbook(template_path)
        ws = wb.active
        # Cập nhật tiêu đề cột theo thứ tự yêu cầu
        ws['A1'], ws['B1'], ws['C1'], ws['D1'], ws['E1'], ws['F1'], ws['G1'], ws['H1'], ws['I1'], ws['J1'], ws['K1'] = (
            "Mã NV", "Tên NV", "Ngày Nghỉ", "Số ngày nghỉ", "Ngày tạo đơn", "Lý do",
            "Ngày Duyệt/Từ chối 1", "Trạng thái 1", "Ngày Duyệt/Từ chối 2", "Trạng thái 2", "Note"
        )
        border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        for i, rec in enumerate(data, start=2):
            display_date = rec.get("CheckinDate", "")
            if rec.get('StartDate') and rec.get('EndDate'):
                start = datetime.strptime(rec['StartDate'], '%Y-%m-%d').strftime('%d/%m/%Y')
                end = datetime.strptime(rec['EndDate'], '%Y-%m-%d').strftime('%d/%m/%Y')
                display_date = f"Từ {start} đến {end}"
            elif rec.get('LeaveDate'):
                leave_date = datetime.strptime(rec['LeaveDate'], '%Y-%m-%d').strftime('%d/%m/%Y')
                display_date = f"{leave_date} ({rec.get('Session', '')})"
            ws.cell(row=i, column=1, value=rec.get("EmployeeId", ""))
            ws.cell(row=i, column=2, value=rec.get("EmployeeName", ""))
            ws.cell(row=i, column=3, value=display_date)
            ws.cell(row=i, column=4, value=calculate_leave_days_from_record(rec))
            ws.cell(row=i, column=5, value=rec.get("Timestamp").astimezone(VN_TZ).strftime('%d/%m/%Y %H:%M:%S') if rec.get("Timestamp") else "")
            tasks = rec.get("Tasks", [])
            tasks_str = (", ".join(tasks) if isinstance(tasks, list) else str(tasks or "")).replace("Nghỉ phép: ", "")
            ws.cell(row=i, column=6, value=rec.get("Reason") or tasks_str)
            ws.cell(row=i, column=7, value=get_formatted_approval_date(rec.get("ApprovalDate1")))
            ws.cell(row=i, column=8, value=rec.get("Status1", ""))
            ws.cell(row=i, column=9, value=get_formatted_approval_date(rec.get("ApprovalDate2")))
            ws.cell(row=i, column=10, value=rec.get("Status2", ""))
            ws.cell(row=i, column=11, value=rec.get("LeaveNote", ""))
            for col_idx in range(1, 12):  # Cập nhật số cột đến 11
                ws.cell(row=i, column=col_idx).border = border
                ws.cell(row=i, column=col_idx).alignment = align_left
       
        filename = f"Danh sách nghỉ phép_{request.args.get('filter')}_{datetime.now(VN_TZ).strftime('%d-%m-%Y')}.xlsx"
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return send_file(output, as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        print(f"❌ Lỗi export leaves: {e}")
        return jsonify({"error": str(e)}), 500

# ---- API xuất Excel kết hợp ----
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
       
        daily_hours_map, monthly_hours_map = {}, {}
        for d in attendance_data:
            emp_id = d.get("EmployeeId")
            date_str = d.get("CheckinDate")
            if emp_id and date_str:
                daily_groups = {}
                daily_groups.setdefault(date_str, []).append(d)
               
                # Handle Timestamp (string or datetime)
                checkins = []
                for r in daily_groups.get(date_str, []):
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
                for r in daily_groups.get(date_str, []):
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
               
                try:
                    month_key = datetime.strptime(date_str, "%d/%m/%Y").strftime("%Y-%m")
                except:
                    continue
                monthly_groups = {}
                monthly_groups.setdefault(month_key, []).append((date_str, daily_seconds))
               
                for month, days in monthly_groups.items():
                    sorted_days = sorted(days, key=lambda x: datetime.strptime(x[0], "%d/%m/%Y"))
                    running_total = 0
                    for date_str, daily_seconds in sorted_days:
                        running_total += daily_seconds
                        monthly_hours_map[(emp_id, date_str)] = running_total
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
        for i, ((emp_id, emp_name, date_str), records) in enumerate(attendance_grouped.items()):
            row = start_row_att + i
            ws_attendance.cell(row=row, column=1, value=emp_id)
            ws_attendance.cell(row=row, column=2, value=emp_name)
            ws_attendance.cell(row=row, column=3, value=date_str)
           
            daily_sec = daily_hours_map.get((emp_id, date_str), 0)
            h, rem = divmod(daily_sec, 3600); m, _ = divmod(rem, 60)
            daily_hours = f"{int(h)}h {int(m)}m" if daily_sec > 0 else "0h 0m"
            ws_attendance.cell(row=row, column=14, value=daily_hours) # Assuming column 14 for DailyHours
           
            monthly_sec = monthly_hours_map.get((emp_id, date_str), 0)
            h, rem = divmod(monthly_sec, 3600); m, _ = divmod(rem, 60)
            monthly_hours = f"{int(h)}h {int(m)}m" if monthly_sec > 0 else "0h 0m"
            ws_attendance.cell(row=row, column=15, value=monthly_hours) # Assuming column 15 for MonthlyHours
           
            checkin_counter, checkin_start_col, checkout_col = 0, 4, 13
            sorted_records = sorted(records, key=lambda x: (
                datetime.strptime(x['Timestamp'], "%Y-%m-%d %H:%M:%S")
                if isinstance(x.get('Timestamp'), str) and x.get('Timestamp')
                else x['Timestamp']
                if isinstance(x.get('Timestamp'), datetime)
                else datetime.min
            ))
           
            for rec in sorted_records:
                time_str = ""
                if rec.get('Timestamp'):
                    try:
                        if isinstance(rec['Timestamp'], str):
                            time_str = datetime.strptime(rec['Timestamp'], "%Y-%m-%d %H:%M:%S").astimezone(VN_TZ).strftime("%H:%M:%S")
                        elif isinstance(rec['Timestamp'], datetime):
                            time_str = rec['Timestamp'].astimezone(VN_TZ).strftime("%H:%M:%S")
                    except (ValueError, TypeError):
                        time_str = ""
                tasks_str = ", ".join(rec.get("Tasks", [])) if isinstance(rec.get("Tasks"), list) else str(rec.get("Tasks", ""))
                cell_value = f"{time_str}; {rec.get('ProjectId','')}; {tasks_str}; {rec.get('Address','')}; {rec.get('CheckinNote','')}"
                if rec.get('CheckType') == 'checkin' and checkin_counter < 9:
                    ws_attendance.cell(row=row, column=checkin_start_col + checkin_counter, value=cell_value)
                    checkin_counter += 1
                elif rec.get('CheckType') == 'checkout':
                    ws_attendance.cell(row=row, column=checkout_col, value=cell_value)
            for col in range(1, 16): # Adjusted to include DailyHours and MonthlyHours columns
                ws_attendance.cell(row=row, column=col).border = border
                ws_attendance.cell(row=row, column=col).alignment = align_left
        # ---- Xử lý sheet Nghỉ phép ----
        ws_leaves = wb["Nghỉ phép"]
        ws_leaves['A1'], ws_leaves['B1'], ws_leaves['C1'], ws_leaves['D1'], ws_leaves['E1'], ws_leaves['F1'], ws_leaves['G1'], ws_leaves['H1'], ws_leaves['I1'], ws_leaves['J1'], ws_leaves['K1'] = (
            "Mã NV", "Tên NV", "Ngày Nghỉ", "Số ngày nghỉ", "Ngày tạo đơn", "Lý do",
            "Ngày Duyệt/Từ chối 1", "Trạng thái 1", "Ngày Duyệt/Từ chối 2", "Trạng thái 2", "Note"
        )
       
        for i, rec in enumerate(leave_data, start=2):
            display_date = rec.get("CheckinDate", "")
            if rec.get('StartDate') and rec.get('EndDate'):
                start = datetime.strptime(rec['StartDate'], '%Y-%m-%d').strftime('%d/%m/%Y')
                end = datetime.strptime(rec['EndDate'], '%Y-%m-%d').strftime('%d/%m/%Y')
                display_date = f"Từ {start} đến {end}"
            elif rec.get('LeaveDate'):
                leave_date = datetime.strptime(rec['LeaveDate'], '%Y-%m-%d').strftime('%d/%m/%Y')
                display_date = f"{leave_date} ({rec.get('Session', '')})"
            ws_leaves.cell(row=i, column=1, value=rec.get("EmployeeId"))
            ws_leaves.cell(row=i, column=2, value=rec.get("EmployeeName"))
            ws_leaves.cell(row=i, column=3, value=display_date)
            ws_leaves.cell(row=i, column=4, value=calculate_leave_days_from_record(rec))
            # Handle Timestamp for leave data
            timestamp_str = ""
            if rec.get("Timestamp"):
                try:
                    if isinstance(rec['Timestamp'], str):
                        timestamp_str = datetime.strptime(rec['Timestamp'], "%Y-%m-%d %H:%M:%S").astimezone(VN_TZ).strftime('%d/%m/%Y %H:%M:%S')
                    elif isinstance(rec['Timestamp'], datetime):
                        timestamp_str = rec['Timestamp'].astimezone(VN_TZ).strftime('%d/%m/%Y %H:%M:%S')
                except (ValueError, TypeError):
                    timestamp_str = ""
            ws_leaves.cell(row=i, column=5, value=timestamp_str)
            tasks = rec.get("Tasks", [])
            tasks_str = (", ".join(tasks) if isinstance(tasks, list) else str(tasks or "")).replace("Nghỉ phép: ", "")
            ws_leaves.cell(row=i, column=6, value=rec.get("Reason") or tasks_str)
            ws_leaves.cell(row=i, column=7, value=get_formatted_approval_date(rec.get("ApprovalDate1")))
            ws_leaves.cell(row=i, column=8, value=rec.get("Status1", ""))
            ws_leaves.cell(row=i, column=9, value=get_formatted_approval_date(rec.get("ApprovalDate2")))
            ws_leaves.cell(row=i, column=10, value=rec.get("Status2", ""))
            ws_leaves.cell(row=i, column=11, value=rec.get("LeaveNote", ""))
            for col in range(1, 12):  # Cập nhật số cột đến 11
                ws_leaves.cell(row=i, column=col).border = border
                ws_leaves.cell(row=i, column=col).alignment = align_left
        filename = f"Báo cáo tổng hợp_{filter_type}_{datetime.now(VN_TZ).strftime('%d-%m-%Y')}.xlsx"
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return send_file(output, as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        print(f"❌ Lỗi export combined: {e}")
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
