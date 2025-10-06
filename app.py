from flask import Flask, render_template, jsonify, send_file, request, session, redirect, url_for, flash
from pymongo import MongoClient
from flask_cors import CORS
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime, timedelta, timezone
from flask import get_flashed_messages  # Để render flash messages
import os
import smtplib
import uuid
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from io import BytesIO
import calendar
import re
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment

app = Flask(__name__, template_folder="templates")
app.secret_key = os.getenv("SECRET_KEY", "dev-secret-key-change-in-prod")  # Thêm secret key cho session
CORS(app)

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
collection = db["alt_checkins"]
idx_collection = db["idx_collection"]

# ---- SMTP Config (thêm để gửi email reset) ----
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))
SMTP_USER = os.getenv("SMTP_USER", "banhbaobeo2205@gmail.com")
SMTP_PASS = os.getenv("SMTP_PASS", "vynqvvvmbcigpdvy")  # App Password Gmail

# ---- Thời gian hết hạn token reset (1 giờ) ----
RESET_EXPIRY_HOURS = 1

# ---- Hàm gửi email ----
def send_email(to_email, subject, body):
    try:
        msg = MIMEMultipart()
        msg["From"] = SMTP_USER
        msg["To"] = to_email
        msg["Subject"] = subject

        # Gửi body như HTML
        msg.attach(MIMEText(body, "html", "utf-8"))

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASS)
            server.sendmail(SMTP_USER, [to_email], msg.as_string())

        print(f"✅ Email đã gửi thành công đến {to_email}")
    except Exception as e:
        print("❌ Lỗi gửi email:", e)

# ---- Decorator login required ----
def login_required(f):
    def wrap(*args, **kwargs):
        if 'email' not in session:
            return jsonify({"success": False, "message": "🔒 Vui lòng đăng nhập lại"}), 401
        return f(*args, **kwargs)
    wrap.__name__ = f.__name__
    return wrap

# ---- Helper render flash messages ----
def render_flash_messages():
    messages_html = ""
    for category, message in get_flashed_messages(with_categories=True):
        color = "green" if category == "success" else "red"
        messages_html += f'<p style="color: {color}; margin: 10px 0; padding: 5px; background: #f8f9fa; border-radius: 4px;">{message}</p>'
    return messages_html

# ---- Trang chủ ----
@app.route("/")
def index():
    return render_template("index.html")


# ---- Trang đăng nhập (POST với email + password) ----
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "GET":
        # Render form login nếu GET (cho forgot/reset fallback)
        return f"""
        <!DOCTYPE html>
        <html lang="vi">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Đăng nhập</title>
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
                <h2>🔐 Đăng nhập</h2>
                {render_flash_messages()}
                <form method="POST">
                    <input type="email" name="email" placeholder="Email" required>
                    <input type="password" name="password" placeholder="Mật khẩu" required>
                    <button type="submit">Đăng nhập</button>
                </form>
                <a href="/forgot-password">Quên mật khẩu?</a>
            </div>
        </body>
        </html>
        """

    if request.method == "POST":
        email = request.form.get("email")
        password = request.form.get("password")
        if not email or not password:
            flash("❌ Vui lòng nhập email và mật khẩu", "error")
            return redirect(url_for("login"))

        emp = idx_collection.find_one({"Email": email})
        if not emp or not check_password_hash(emp.get("Password", ""), password):
            flash("🚫 Email hoặc mật khẩu không đúng!", "error")
            return redirect(url_for("login"))

        # Set session
        session["email"] = email
        session["employee_id"] = emp["EmployeeId"]
        session["employee_name"] = emp["EmployeeName"]
        flash("✅ Đăng nhập thành công!", "success")
        return redirect(url_for("index"))  # Redirect về trang chủ sau login


# ---- Trang quên mật khẩu ----
@app.route("/forgot-password", methods=["GET", "POST"])
def forgot_password():
    if request.method == "POST":
        email = request.form.get("email")
        if not email:
            flash("❌ Vui lòng nhập email", "error")
            return redirect(url_for("forgot_password"))

        emp = idx_collection.find_one({"Email": email})
        if not emp:
            flash("🚫 Email không tồn tại!", "error")
            return redirect(url_for("forgot_password"))

        # Tạo reset token và expiry
        reset_token = str(uuid.uuid4())
        expiry = datetime.now(VN_TZ) + timedelta(hours=RESET_EXPIRY_HOURS)
        idx_collection.update_one(
            {"_id": emp["_id"]},
            {"$set": {"reset_token": reset_token, "reset_expiry": expiry}}
        )

        # Gửi email reset
        reset_url = f"{request.host_url.rstrip('/')}/reset-password/{reset_token}"
        body = f"""
        <!DOCTYPE html>
        <html><head><meta charset="UTF-8"></head><body>
            <h2>🔄 Đặt lại mật khẩu</h2>
            <p>Nhấp vào link để đặt lại: <a href="{reset_url}">Đặt lại mật khẩu</a></p>
            <p>Token hết hạn sau {RESET_EXPIRY_HOURS} giờ.</p>
        </body></html>
        """
        send_email(email, "Đặt lại mật khẩu", body)
        flash("📧 Email đặt lại mật khẩu đã gửi!", "success")
        return redirect(url_for("login"))

    # GET: Render form
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
    emp = idx_collection.find_one({
        "reset_token": token,
        "reset_expiry": {"$gt": datetime.now(VN_TZ)}
    })
    if not emp:
        flash("Token không hợp lệ hoặc hết hạn!", "error")
        return redirect(url_for("login"))

    if request.method == "POST":
        new_password = request.form.get("password")
        if not new_password:
            flash("❌ Vui lòng nhập mật khẩu mới", "error")
            return redirect(url_for("reset_password", token=token))

        hashed_pw = generate_password_hash(new_password)
        idx_collection.update_one(
            {"_id": emp["_id"]},
            {"$set": {"Password": hashed_pw, "reset_token": None, "reset_expiry": None}}
        )
        flash("✅ Mật khẩu đã được cập nhật!", "success")
        return redirect(url_for("login"))

    # GET: Render form
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


# ---- Logout ----
@app.route("/logout")
def logout():
    session.clear()
    flash("👋 Đăng xuất thành công!", "success")
    return redirect(url_for("login"))


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

    if search:
        regex = re.compile(search, re.IGNORECASE)
        query["$or"] = [
            {"EmployeeId": {"$regex": regex}},
            {"EmployeeName": {"$regex": regex}}
        ]
    return query


# ---- API lấy dữ liệu chấm công (protected) ----
@app.route("/api/attendances", methods=["GET"])
@login_required
def get_attendances():
    try:
        # ✅ Sử dụng session email thay vì param (an toàn hơn)
        email = session.get("email")
        if not email:
            return jsonify({"error": "🔒 Vui lòng đăng nhập lại"}), 401

        filter_type = request.args.get("filter", "hôm nay").lower()  # Default "hôm nay"
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()

        query = build_query(filter_type, start_date, end_date, search)

        data = list(collection.find(query, {"_id": 0}))
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ---- API xuất Excel (protected) ----
@app.route("/api/export-excel", methods=["GET"])
@login_required
def export_to_excel():
    try:
        # Sử dụng session email
        email = session.get("email")
        if not email:
            return jsonify({"error": "🔒 Vui lòng đăng nhập lại"}), 401

        # Lấy info user từ session hoặc DB (không cần cho export, nhưng giữ để validate)
        emp = idx_collection.find_one({"Email": email}, {"EmployeeId": 1, "EmployeeName": 1, "_id": 0})
        if not emp:
            return jsonify({"error": "🚫 Email không tồn tại"}), 403

        emp_id = emp["EmployeeId"]
        emp_name = emp["EmployeeName"]

        # ---- Tham số lọc ----
        filter_type = request.args.get("filter", "hôm nay").lower()  # Default "hôm nay" cho export
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()

        # ---- Tạo query ----
        query = {}
        if filter_type == "hôm nay":
            query["CheckinDate"] = datetime.now(VN_TZ).strftime("%Y-%m-%d")
        elif filter_type == "custom" and start_date and end_date:
            query["CheckinDate"] = {"$gte": start_date, "$lte": end_date}
        elif filter_type == "tuần":
            today = datetime.now(VN_TZ)
            start = (today - timedelta(days=today.weekday())).strftime("%Y-%m-%d")
            end = (today + timedelta(days=6 - today.weekday())).strftime("%Y-%m-%d")
            query["CheckinDate"] = {"$gte": start, "$lte": end}
        elif filter_type == "tháng":
            today = datetime.now(VN_TZ)
            start = today.replace(day=1).strftime("%Y-%m-%d")
            end = today.replace(day=calendar.monthrange(today.year, today.month)[1]).strftime("%Y-%m-%d")
            query["CheckinDate"] = {"$gte": start, "$lte": end}
        elif filter_type == "năm":
            today = datetime.now(VN_TZ)
            query["CheckinDate"] = {"$regex": f"^{today.year}"}
        if search:
            query["$or"] = [
                {"EmployeeName": {"$regex": search, "$options": "i"}},
                {"EmployeeId": {"$regex": search, "$options": "i"}},
                {"Tasks": {"$regex": search, "$options": "i"}},
                {"ProjectId": {"$regex": search, "$options": "i"}},
            ]

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
            "Status": 1
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

                # ---- Nếu là nghỉ phép có lý do dạng "Nghỉ phép: xxx" ----
                leave_reason = ""
                if "nghỉ phép" in tasks_str.lower():
                    if ":" in tasks_str:
                        split_task = tasks_str.split(":", 1)
                        tasks_str = split_task[0].strip()       # → "Nghỉ phép"
                        leave_reason = split_task[1].strip()    # → "Sức khoẻ"
                    else:
                        tasks_str = tasks_str.strip()
                else:
                    tasks_str = tasks_str.strip()

                status = rec.get("Status", "")

                # ---- Build nội dung xuất Excel ----
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
