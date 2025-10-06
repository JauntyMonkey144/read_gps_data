from flask import Flask, render_template, jsonify, request, redirect, url_for
from pymongo import MongoClient
from flask_cors import CORS
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime, timedelta, timezone
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import uuid

# ---- Cấu hình SMTP ----
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))
SMTP_USER = os.getenv("SMTP_USER", "sun.automation.sys@gmail.com")  # email hệ thống
SMTP_PASS = os.getenv("SMTP_PASS", "ihgzxunefndizeub")          # App Password Gmail
app = Flask(__name__, template_folder="templates")
CORS(app, methods=["GET", "POST"])
reset_tokens = db["reset_tokens"]

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
        return """
        <!DOCTYPE html>
        <html lang="vi">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Quên mật khẩu</title>
            <style>
                body { font-family: Arial; background: #f4f6f9; margin: 0; padding: 20px; }
                .container { max-width: 400px; margin: 100px auto; background: white; padding: 30px;
                             border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
                input, button { width: 100%; padding: 10px; margin: 10px 0; border-radius: 4px; border: 1px solid #ddd; }
                button { background: #28a745; color: white; border: none; font-size: 16px; cursor: pointer; }
                button:hover { background: #218838; }
            </style>
        </head>
        <body>
            <div class="container">
                <h2>🔒 Quên mật khẩu</h2>
                <form method="POST">
                    <input type="email" name="email" placeholder="Nhập email của bạn" required>
                    <button type="submit">Gửi link đặt lại mật khẩu</button>
                    <a href="/">⬅ Quay lại trang đăng nhập</a>
                </form>
            </div>
        </body>
        </html>
        """

    # POST -> gửi email
    email = request.form.get("email")
    if not email:
        return jsonify({"success": False, "message": "❌ Vui lòng nhập email"}), 400

    admin = admins.find_one({"email": email})
    if not admin:
        return jsonify({"success": False, "message": "🚫 Email không tồn tại"}), 404

    # Tạo token reset
    token = str(uuid.uuid4())
    expire_time = datetime.now(VN_TZ) + timedelta(hours=1)  # hết hạn sau 1h

    reset_tokens.insert_one({
        "email": email,
        "token": token,
        "expire_at": expire_time
    })

    reset_link = f"http://localhost:5000/reset-password?token={token}"

    # Gửi mail
    try:
        msg = MIMEMultipart()
        msg["From"] = SMTP_USER
        msg["To"] = email
        msg["Subject"] = "🔐 Đặt lại mật khẩu tài khoản của bạn"

        body = f"""
        <p>Xin chào <b>{admin.get('username', email)}</b>,</p>
        <p>Bạn đã yêu cầu đặt lại mật khẩu. Nhấn vào link bên dưới để tiếp tục:</p>
        <p><a href="{reset_link}" target="_blank">👉 Đặt lại mật khẩu</a></p>
        <p>Link này chỉ có hiệu lực trong 1 giờ.</p>
        <p>Nếu bạn không yêu cầu, vui lòng bỏ qua email này.</p>
        """

        msg.attach(MIMEText(body, "html"))

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASS)
            server.send_message(msg)

        return f"""
        <html><body style="font-family:Arial;text-align:center;padding:50px;">
        <h2>📩 Email khôi phục đã được gửi!</h2>
        <p>Vui lòng kiểm tra hộp thư đến của bạn: <b>{email}</b></p>
        <a href="/">⬅ Quay lại trang đăng nhập</a>
        </body></html>
        """

    except Exception as e:
        print("❌ Lỗi gửi mail:", e)
        return jsonify({"success": False, "message": f"Lỗi gửi mail: {e}"}), 500
        
@app.route("/reset-password", methods=["GET", "POST"])
def reset_password():
    token = request.args.get("token")

    if not token:
        return "<h3>❌ Token không hợp lệ</h3>"

    record = reset_tokens.find_one({"token": token})
    if not record:
        return "<h3>🚫 Token không tồn tại hoặc đã hết hạn</h3>"

    if record["expire_at"] < datetime.now(VN_TZ):
        return "<h3>⏰ Token đã hết hạn</h3>"

    if request.method == "GET":
        return f"""
        <html><body style="font-family:Arial; text-align:center; padding:50px;">
        <h2>🔑 Đặt lại mật khẩu mới</h2>
        <form method="POST">
            <input type="hidden" name="token" value="{token}">
            <input type="password" name="new_password" placeholder="Nhập mật khẩu mới" required><br><br>
            <button type="submit">Cập nhật mật khẩu</button>
        </form>
        </body></html>
        """

    # POST -> đổi mật khẩu
    new_pw = request.form.get("new_password")
    if not new_pw:
        return "<h3>❌ Mật khẩu không được để trống</h3>"

    email = record["email"]
    hashed_pw = generate_password_hash(new_pw)
    admins.update_one({"email": email}, {"$set": {"password": hashed_pw}})

    # Xóa token sau khi dùng
    reset_tokens.delete_one({"token": token})

    return """
    <html><body style="font-family:Arial;text-align:center;padding:50px;">
    <h2>✅ Mật khẩu đã được cập nhật thành công!</h2>
    <a href="/">⬅ Quay lại trang đăng nhập</a>
    </body></html>
    """
    
# ---- Hàm dựng query lọc ----
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
    elif filter_type == "tất cả":
        pass

    if search:
        regex = re.compile(search, re.IGNORECASE)
        query["$or"] = [
            {"EmployeeId": {"$regex": regex}},
            {"EmployeeName": {"$regex": regex}}
        ]
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

        query = build_query(filter_type, start_date, end_date, search)
        # Fetch TẤT CẢ dữ liệu matching filter (không filter theo user)

        data = list(collection.find(query, {"_id": 0}))
        print(f"DEBUG: Fetched {len(data)} records for email {email} with filter {filter_type}")  # Log debug
        return jsonify(data)
    except Exception as e:
        print(f"❌ Error in get_attendances: {e}")
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
        query = build_query(filter_type, start_date, end_date, search)

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



if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
