from flask import Flask, render_template, jsonify, request, redirect, url_for, send_file
from pymongo import MongoClient
from flask_cors import CORS
from werkzeug.security import generate_password_hash, check_password_hash
from itsdangerous import URLSafeTimedSerializer, SignatureExpired, BadSignature
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment
from datetime import datetime, timedelta, timezone
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from io import BytesIO
from dotenv import load_dotenv
import smtplib, os, re, calendar

# =====================================
# ⚙️ Cấu hình Flask + MongoDB
# =====================================
load_dotenv()

app = Flask(__name__, template_folder="templates")
CORS(app, methods=["GET", "POST"])

# ---- Timezone VN ----
VN_TZ = timezone(timedelta(hours=7))
app.secret_key = os.getenv("SECRET_KEY", "supersecret")

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
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SMTP_USER = os.getenv("SMTP_USER", "banhbaobeo2205@gmail.com")
SMTP_PASS = os.getenv("SMTP_PASS", "vynqvvvmbcigpdvy")

# ---- Token Serializer ----
serializer = URLSafeTimedSerializer(app.secret_key)


# =====================================
# 🏠 Trang chính hiển thị bảng chấm công
# =====================================
@app.route("/")
def index():
    success = request.args.get("success")
    return render_template("index.html", success=success)


# =====================================
# 🔐 API đăng nhập
# =====================================
@app.route("/login", methods=["POST"])
def login():
    data = request.get_json(force=True)
    email = data.get("email")
    password = data.get("password")

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


# =====================================
# ✉️ Quên mật khẩu
# =====================================
@app.route("/forgot-password", methods=["GET", "POST"])
def forgot_password():
    if request.method == "GET":
        return render_template("forgot_password.html")

    email = request.form.get("email")
    admin = admins.find_one({"email": email})

    if not admin:
        return render_template("message.html", msg="🚫 Email không tồn tại!", color="red")

    token = serializer.dumps(email, salt="password-reset")
    reset_link = f"{request.host_url}reset-password/{token}"

    msg = MIMEMultipart("alternative")
    msg["Subject"] = "🔒 Đặt lại mật khẩu - Sun Automation"
    msg["From"] = SMTP_USER
    msg["To"] = email

    body = f"""
    <html><body>
    <h3>Xin chào {admin.get("username","Admin")},</h3>
    <p>Bạn vừa yêu cầu đặt lại mật khẩu. Link hợp lệ trong 15 phút:</p>
    <a href="{reset_link}">{reset_link}</a>
    <br><br><b>Sun Automation System</b>
    </body></html>
    """
    msg.attach(MIMEText(body, "html"))

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASS)
            server.sendmail(SMTP_USER, email, msg.as_string())
        return render_template("message.html", msg=f"✅ Link đã gửi đến {email}", color="green")
    except Exception as e:
        print("❌ Lỗi gửi email:", e)
        return render_template("message.html", msg=f"Lỗi khi gửi email: {e}", color="red")


# =====================================
# 🔁 Reset mật khẩu
# =====================================
@app.route("/reset-password/<token>", methods=["GET", "POST"])
def reset_password(token):
    try:
        email = serializer.loads(token, salt="password-reset", max_age=900)
    except SignatureExpired:
        return render_template("message.html", msg="⏰ Link đã hết hạn!", color="red")
    except BadSignature:
        return render_template("message.html", msg="🚫 Link không hợp lệ!", color="red")

    if request.method == "GET":
        return render_template("reset_password.html", email=email)

    new_password = request.form.get("new_password")
    if not new_password:
        return render_template("message.html", msg="❌ Vui lòng nhập mật khẩu mới!", color="red")

    hashed_pw = generate_password_hash(new_password)
    admins.update_one({"email": email}, {"$set": {"password": hashed_pw}})
    return redirect(url_for("index", success=1))


# =====================================
# 🧩 API lấy dữ liệu chấm công JSON
# =====================================
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
            {"OtherNote": {"$regex": regex}},
        ]

    if search:
        regex = re.compile(search, re.IGNORECASE)
        query["$or"] = [{"EmployeeId": {"$regex": regex}}, {"EmployeeName": {"$regex": regex}}]
    return query


@app.route("/api/attendances", methods=["GET"])
def get_attendances():
    try:
        filter_type = request.args.get("filter", "hôm nay").lower()
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()

        query = build_query(filter_type, start_date, end_date, search)
        data = list(collection.find(query, {"_id": 0}))
        return jsonify({"success": True, "count": len(data), "data": data})
    except Exception as e:
        print("❌ Lỗi get_attendances:", e)
        return jsonify({"success": False, "error": str(e)}), 500


# =====================================
# 📊 API xuất Excel
# =====================================
@app.route("/api/export-excel", methods=["GET"])
def export_to_excel():
    try:
        email = request.args.get("email")
        admin = admins.find_one({"email": email})
        if not admin:
            return jsonify({"error": "🚫 Email không có quyền truy cập"}), 403

        filter_type = request.args.get("filter", "hôm nay").lower()
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()
        query = build_query(filter_type, start_date, end_date, search)

        data = list(collection.find(query, {"_id": 0}))
        template_path = "templates/Copy of Form chấm công.xlsx"
        wb = load_workbook(template_path)
        ws = wb.active

        border = Border(left=Side(style="thin"), right=Side(style="thin"),
                        top=Side(style="thin"), bottom=Side(style="thin"))
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

        row = 2
        for d in data:
            ws.cell(row=row, column=1, value=d.get("EmployeeId"))
            ws.cell(row=row, column=2, value=d.get("EmployeeName"))
            ws.cell(row=row, column=3, value=d.get("ProjectId"))
            ws.cell(row=row, column=4, value=d.get("Tasks"))
            ws.cell(row=row, column=5, value=d.get("Address"))
            ws.cell(row=row, column=6, value=d.get("Status"))
            ws.cell(row=row, column=7, value=d.get("CheckinDate"))
            for col in range(1, 8):
                c = ws.cell(row=row, column=col)
                c.border = border
                c.alignment = align_left
            row += 1

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        filename = f"ChamCong_{datetime.now(VN_TZ).strftime('%d-%m-%Y')}.xlsx"
        return send_file(output, as_attachment=True,
                         download_name=filename,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        print("❌ Lỗi export:", e)
        return jsonify({"error": str(e)}), 500


# =====================================
# 🚀 Run server
# =====================================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
