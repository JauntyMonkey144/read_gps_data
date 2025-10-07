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
collection = db["alt_checkins"]  

# ---- Trang chủ (đăng nhập chính) ----
@app.route("/")
def index():
    success = request.args.get("success")  # nếu =1 -> hiển thị thông báo
    return render_template("index.html", success=success)
# ---- Đăng nhập API ----
@app.route("/login", methods=["POST", "GET"])
def login():
    if request.method == "GET":
        # Không hiển thị form nữa, chuyển về trang chủ
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
# ---- Quên mật khẩu ----
@app.route("/forgot-password", methods=["GET", "POST"])
def forgot_password():
    if request.method == "GET":
        # Form HTML đơn giản cho reset mật khẩu
        return render_template("index.html", success=success)

    if request.method == "POST":
        email = request.form.get("email")
        new_password = request.form.get("new_password")
        if not email or not new_password:
            return jsonify({"success": False, "message": "❌ Vui lòng nhập email và mật khẩu mới"}), 400
        admin = admins.find_one({"email": email})
        if not admin:
            return jsonify({"success": False, "message": "🚫 Email không tồn tại!"}), 404
        hashed_pw = generate_password_hash(new_password)
        admins.update_one({"email": email}, {"$set": {"password": hashed_pw}})
        # ✅ Chuyển về trang chủ có thông báo thành công
        return redirect(url_for("index", success=1))
def build_attendance_query(filter_type, start_date, end_date, search):
    query = {}
    today = datetime.now(VN_TZ)
    regex_leave = re.compile("Nghỉ phép", re.IGNORECASE)
    # --- Bộ lọc thời gian ---
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
    # --- Bộ lọc nghỉ phép ---
    if filter_type == "nghỉ phép":
        query["$or"] = [
            {"Tasks": {"$regex": regex_leave}},
            {"Status": {"$regex": regex_leave}},
            {"OtherNote": {"$regex": regex_leave}}
        ]
    else:
        # Các filter bình thường: loại bỏ bản ghi có “Nghỉ phép”
        query["$and"] = [
            {"$or": [
                {"Tasks": {"$not": regex_leave}},
                {"Tasks": {"$exists": False}},
                {"Tasks": None}
            ]}
        ]
    # --- Bộ lọc tìm kiếm ---
    if search:
        regex = re.compile(search, re.IGNORECASE)
        query["$and"] = query.get("$and", []) + [
            {"$or": [
                {"EmployeeId": {"$regex": regex}},
                {"EmployeeName": {"$regex": regex}}
            ]}
        ]
    return query

def build_leave_query(filter_type, start_date, end_date, search):
    query = {}
    today = datetime.now(VN_TZ)
    regex_leave = re.compile("Nghỉ phép", re.IGNORECASE)
    # Luôn lọc cho nghỉ phép
    leave_or = {
        "$or": [
            {"Tasks": {"$regex": regex_leave}},
            {"Status": {"$regex": regex_leave}},
            {"OtherNote": {"$regex": regex_leave}}
        ]
    }
    # --- Bộ lọc thời gian (dùng CheckinDate làm ngày nghỉ) ---
    date_filter = {}
    if filter_type == "custom" and start_date and end_date:
        date_filter["CheckinDate"] = {"$gte": start_date, "$lte": end_date}
    elif filter_type == "hôm nay":
        date_filter["CheckinDate"] = today.strftime("%Y-%m-%d")
    elif filter_type == "tuần":
        start = (today - timedelta(days=today.weekday())).strftime("%Y-%m-%d")
        end = (today + timedelta(days=6 - today.weekday())).strftime("%Y-%m-%d")
        date_filter["CheckinDate"] = {"$gte": start, "$lte": end}
    elif filter_type == "tháng":
        start = today.replace(day=1).strftime("%Y-%m-%d")
        end = today.replace(day=calendar.monthrange(today.year, today.month)[1]).strftime("%Y-%m-%d")
        date_filter["CheckinDate"] = {"$gte": start, "$lte": end}
    elif filter_type == "năm":
        date_filter["CheckinDate"] = {"$regex": f"^{today.year}"}
    
    # Kết hợp leave_or và date_filter
    if date_filter:
        query = {"$and": [leave_or, date_filter]}
    else:
        query = leave_or
    
    # --- Bộ lọc tìm kiếm ---
    if search:
        regex = re.compile(search, re.IGNORECASE)
        search_or = {
            "$or": [
                {"EmployeeId": {"$regex": regex}},
                {"EmployeeName": {"$regex": regex}}
            ]
        }
        if "$and" in query:
            query["$and"].append(search_or)
        else:
            query = {"$and": [query, search_or]}
    return query
# ---- API lấy dữ liệu chấm công (validate email từ admins) ----
@app.route("/api/attendances", methods=["GET"])
def get_attendances():
    try:
        email = request.args.get("email")  # ✅ Trùng key với front-end
        if not email:
            return jsonify({"error": "❌ Thiếu email"}), 400
        # ✅ Validate email tồn tại trong admins (không cần password lại)
        admin = admins.find_one({"email": email}, {"_id": 0, "username": 1})
        if not admin:
            return jsonify({"error": "🚫 Email không hợp lệ (không có quyền truy cập)"}), 403
        filter_type = request.args.get("filter", "hôm nay").lower()  # Default "hôm nay"
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()
        query = build_attendance_query(filter_type, start_date, end_date, search)
        # Fetch TẤT CẢ dữ liệu matching filter (không filter theo user)
        data = list(collection.find(query, {"_id": 0}))
        print(f"DEBUG: Fetched {len(data)} records for email {email} with filter {filter_type}")  # Log debug
        return jsonify(data)
    except Exception as e:
        print(f"❌ Error in get_attendances: {e}")
        return jsonify({"error": str(e)}), 500

# ---- API lấy dữ liệu nghỉ phép (validate email từ admins) ----
@app.route("/api/leaves", methods=["GET"])
def get_leaves():
    try:
        email = request.args.get("email")
        if not email:
            return jsonify({"error": "❌ Thiếu email"}), 400
        # ✅ Validate email tồn tại trong admins
        admin = admins.find_one({"email": email}, {"_id": 0, "username": 1})
        if not admin:
            return jsonify({"error": "🚫 Email không hợp lệ (không có quyền truy cập)"}), 403
        filter_type = request.args.get("filter", "tất cả").lower()
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()
        query = build_leave_query(filter_type, start_date, end_date, search)
        # Fetch dữ liệu nghỉ phép với fields phù hợp
        data = list(collection.find(query, {
            "_id": 0,
            "EmployeeId": 1,
            "EmployeeName": 1,
            "CheckinDate": 1,  # Ngày nghỉ
            "CheckinTime": 1,  # Ngày tạo đơn
            "Tasks": 1,        # Ghi chú
            "Status": 1
        }))
        print(f"DEBUG: Fetched {len(data)} leave records for email {email} with filter {filter_type}")
        return jsonify(data)
    except Exception as e:
        print(f"❌ Error in get_leaves: {e}")
        return jsonify({"error": str(e)}), 500
# ---- API xuất Excel (validate email từ admins) ----
@app.route("/api/export-excel", methods=["GET"])
def export_to_excel():
    try:
        email = request.args.get("email")
        if not email:
            return jsonify({"error": "❌ Thiếu email"}), 400
        # ✅ Kiểm tra quyền admin
        admin = admins.find_one({"email": email}, {"_id": 0, "username": 1})
        if not admin:
            return jsonify({"error": "🚫 Email không hợp lệ (không có quyền truy cập)"}), 403
        # ---- Tham số lọc ----
        filter_type = request.args.get("filter", "hôm nay").lower()
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()
        # ---- Tạo query ----
        query = build_attendance_query(filter_type, start_date, end_date, search)
        # ---- Lấy dữ liệu ----
        data = list(db.alt_checkins.find(query, {
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
        # ---- Nhóm theo nhân viên + ngày ----
        grouped = {}
        for d in data:
            emp_id = d.get("EmployeeId", "")
            emp_name = d.get("EmployeeName", "")
            date = d.get("CheckinDate") or (
                d["CheckinTime"].astimezone(VN_TZ).strftime("%Y-%m-%d")
                if isinstance(d.get("CheckinTime"), datetime) else ""
            )
            key = (emp_id, emp_name, date)
            grouped.setdefault(key, []).append(d)
        # ---- Load template Excel ----
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
        # ---- Điền dữ liệu ----
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
                # ---- Xử lý Tasks ----
                tasks = rec.get("Tasks")
                if isinstance(tasks, list):
                    tasks_str = ", ".join(tasks)
                else:
                    tasks_str = str(tasks or "")
                leave_reason = ""
                if "nghỉ phép" in tasks_str.lower():
                    if ":" in tasks_str:
                        split_task = tasks_str.split(":", 1)
                        tasks_str = split_task[0].strip()       # → "Nghỉ phép"
                        leave_reason = split_task[1].strip()    # → Lý do
                    else:
                        tasks_str = tasks_str.strip()
                status = rec.get("Status", "")
                # ---- Nếu là nghỉ phép thì format đặc biệt ----
                if "nghỉ phép" in tasks_str.lower():
                    approve_date = ""
                    if rec.get("ApprovedBy"):
                        if isinstance(checkin_time, datetime):
                            approve_date = checkin_time.astimezone(VN_TZ).strftime("%d/%m/%Y")
                        else:
                            approve_date = datetime.now(VN_TZ).strftime("%d/%m/%Y")
                    entry = f"{date}; Nghỉ phép; {leave_reason}; {status}; {approve_date}"
                else:
                    # ---- Build nội dung export mặc định ----
                    if time_str:
                        parts.append(time_str)
                    if rec.get("ProjectId"):
                        parts.append(str(rec["ProjectId"]))
                    if tasks_str:
                        parts.append(tasks_str)
                    if leave_reason:
                        parts.append(leave_reason)
                    if status:
                        parts.append(status)
                    if rec.get("OtherNote"):
                        parts.append(rec["OtherNote"])
                    if rec.get("Address"):
                        parts.append(rec["Address"])
                    entry = "; ".join(parts)
                ws.cell(row=row, column=3 + j, value=entry)
            # ---- Border + căn chỉnh ----
            for col in range(1, 14):
                cell = ws.cell(row=row, column=col)
                cell.border = border
                cell.alignment = align_left
            # ---- Auto-fit row height ----
            max_lines = max(
                (str(ws.cell(row=row, column=col).value).count("\n") + 1 if ws.cell(row=row, column=col).value else 1)
                for col in range(1, 14)
            )
            ws.row_dimensions[row].height = max_lines * 20
        # ---- Auto-fit column width ----
        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    length = len(str(cell.value))
                    max_length = max(max_length, length)
            ws.column_dimensions[col_letter].width = max_length + 2
        # ---- Xuất file ----
        today_str = datetime.now(VN_TZ).strftime("%d-%m-%Y")
        if search:
            filename = f"Danh sách chấm công theo tìm kiếm_{today_str}.xlsx"
        elif filter_type == "hôm nay":
            filename = f"Danh sách chấm công_{today_str}.xlsx"
        elif filter_type == "custom" and start_date and end_date:
            filename = f"Danh sách chấm công từ {start_date} đến {end_date}_{today_str}.xlsx"
        elif filter_type == "custom" and start_date and end_date:
            filename = f"Danh sách đơn nghỉ phép_{today_str}.xlsx"
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
        # ✅ Kiểm tra quyền admin
        admin = admins.find_one({"email": email}, {"_id": 0, "username": 1})
        if not admin:
            return jsonify({"error": "🚫 Email không hợp lệ (không có quyền truy cập)"}), 403
        # ---- Tham số lọc ----
        filter_type = request.args.get("filter", "tất cả").lower()
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()
        # ---- Tạo query ----
        query = build_leave_query(filter_type, start_date, end_date, search)
        # ---- Lấy dữ liệu nghỉ phép ----
        data = list(db.alt_checkins.find(query, {
            "_id": 0,
            "EmployeeId": 1,
            "EmployeeName": 1,
            "CheckinDate": 1,  # Ngày nghỉ
            "CheckinTime": 1,  # Ngày tạo đơn
            "Tasks": 1,        # Ghi chú
            "Status": 1
        }))
        # ---- Load template Excel (sử dụng cùng template, nhưng điền theo cột nghỉ phép) ----
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
        # ---- Điền dữ liệu nghỉ phép ----
        start_row = 2
        for i, rec in enumerate(data, start=0):
            row = start_row + i
            # Cột 1: Mã NV
            ws.cell(row=row, column=1, value=rec.get("EmployeeId", ""))
            # Cột 2: Tên nhân viên
            ws.cell(row=row, column=2, value=rec.get("EmployeeName", ""))
            # Cột 3: Ngày nghỉ
            ws.cell(row=row, column=3, value=rec.get("CheckinDate", ""))
            # Cột 4: Ngày tạo đơn (format CheckinTime)
            checkin_time = rec.get("CheckinTime")
            time_str = ""
            if isinstance(checkin_time, datetime):
                time_str = checkin_time.astimezone(VN_TZ).strftime("%d/%m/%Y %H:%M:%S")
            elif isinstance(checkin_time, str):
                time_str = checkin_time
            ws.cell(row=row, column=4, value=time_str)
            # Cột 5: Ghi chú (Tasks)
            tasks = rec.get("Tasks")
            tasks_str = ""
            if isinstance(tasks, list):
                tasks_str = ", ".join(tasks)
            else:
                tasks_str = str(tasks or "")
            ws.cell(row=row, column=5, value=tasks_str)
            # Cột 6: Trạng thái
            ws.cell(row=row, column=6, value=rec.get("Status", ""))
            # ---- Border + căn chỉnh cho 6 cột chính ----
            for col in range(1, 7):
                cell = ws.cell(row=row, column=col)
                cell.border = border
                cell.alignment = align_left
            # ---- Auto-fit row height ----
            max_lines = max(
                (str(ws.cell(row=row, column=col).value).count("\n") + 1 if ws.cell(row=row, column=col).value else 1)
                for col in range(1, 7)
            )
            ws.row_dimensions[row].height = max_lines * 20
        # ---- Auto-fit column width ----
        for col in ws.columns[:6]:  # Chỉ fit 6 cột đầu
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    length = len(str(cell.value))
                    max_length = max(max_length, length)
            ws.column_dimensions[col_letter].width = max_length + 2
        # ---- Xuất file ----
        today_str = datetime.now(VN_TZ).strftime("%d-%m-%Y")
        if search:
            filename = f"Danh sách nghỉ phép theo tìm kiếm_{today_str}.xlsx"
        elif filter_type == "hôm nay":
            filename = f"Danh sách nghỉ phép hôm nay_{today_str}.xlsx"
        elif filter_type == "custom" and start_date and end_date:
            filename = f"Danh sách nghỉ phép từ {start_date} đến {end_date}_{today_str}.xlsx"
        else:
            filename = f"Danh sách nghỉ phép_{today_str}.xlsx"
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
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
