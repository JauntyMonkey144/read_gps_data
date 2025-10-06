# app.py
from flask import (
    Flask, render_template_string, jsonify, request, redirect,
    url_for, send_file, session, flash
)
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

load_dotenv()

# ---------- Cấu hình cơ bản ----------
app = Flask(__name__)
CORS(app, methods=["GET", "POST"])
app.secret_key = os.getenv("SECRET_KEY", "supersecret_dev_key")

# Timezone VN
VN_TZ = timezone(timedelta(hours=7))

# MongoDB
MONGO_URI = os.getenv("MONGO_URI", "mongodb://localhost:27017")
DB_NAME = os.getenv("DB_NAME", "Sun_Database_1")
client = MongoClient(MONGO_URI)
db = client[DB_NAME]
admins = db["admins"]
collection = db["alt_checkins"]

# SMTP
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))
SMTP_USER = os.getenv("SMTP_USER", "")
SMTP_PASS = os.getenv("SMTP_PASS", "")

# Serializer
serializer = URLSafeTimedSerializer(app.secret_key)

# ---------- Helper: login required ----------
from functools import wraps

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("user_email"):
            return redirect(url_for("login_page", next=request.path))
        return f(*args, **kwargs)
    return decorated

# ---------- Simple templates (for quick local test) ----------
INDEX_HTML = """
<!doctype html>
<html>
  <head><meta charset="utf-8"><title>Sun Automation - Dashboard</title></head>
  <body>
    <h2>Dashboard - Sun Automation</h2>
    {% if session.username %}
      <p>Chào, {{ session.username }} (<a href="{{ url_for('logout') }}">Logout</a>)</p>
    {% endif %}
    <p><a href="{{ url_for('export_to_excel') }}">Export Excel (demo)</a></p>
    <div id="content"></div>
  </body>
</html>
"""

LOGIN_HTML = """
<!doctype html>
<html>
  <head><meta charset="utf-8"><title>Đăng nhập</title></head>
  <body>
    <h2>Đăng nhập</h2>
    {% with messages = get_flashed_messages() %}
      {% if messages %}
        <ul>{% for m in messages %}<li style="color:red">{{m}}</li>{% endfor %}</ul>
      {% endif %}
    {% endwith %}
    <form method="post" action="{{ url_for('login') }}">
      <label>Email: <input name="email" type="email" required></label><br>
      <label>Password: <input name="password" type="password" required></label><br>
      <button type="submit">Đăng nhập</button>
    </form>
    <p><a href="{{ url_for('forgot_password') }}">Quên mật khẩu?</a></p>
  </body>
</html>
"""

FORGOT_PW_HTML = """
<!doctype html>
<html>
  <head><meta charset="utf-8"><title>Quên mật khẩu</title></head>
  <body>
    <h2>Quên mật khẩu</h2>
    <form method="post">
      <label>Email: <input name="email" type="email" required></label><br>
      <button type="submit">Gửi link đặt lại</button>
    </form>
  </body>
</html>
"""

RESET_PW_HTML = """
<!doctype html>
<html>
  <head><meta charset="utf-8"><title>Reset mật khẩu</title></head>
  <body>
    <h2>Đặt lại mật khẩu cho {{ email }}</h2>
    <form method="post">
      <label>Mật khẩu mới: <input name="new_password" type="password" required></label><br>
      <button type="submit">Xác nhận</button>
    </form>
  </body>
</html>
"""

MESSAGE_HTML = """
<!doctype html>
<html>
  <head><meta charset="utf-8"><title>Message</title></head>
  <body>
    <h3 style="color:{{ color }}">{{ msg }}</h3>
    <p><a href="{{ url_for('index') }}">Về trang chính</a></p>
  </body>
</html>
"""

# ---------- Routes ----------
@app.route("/")
@login_required
def index():
    # Hiển thị dashboard (tạm simple)
    return render_template("index.html")

# Render login page (GET) and handle login POST
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "GET":
        return render_template_string(LOGIN_HTML)

    # POST (form submit)
    email = request.form.get("email")
    password = request.form.get("password")

    if not email or not password:
        flash("Vui lòng cung cấp email và mật khẩu")
        return redirect(url_for("login"))

    admin = admins.find_one({"email": email})
    if not admin:
        flash("Email không tồn tại")
        return redirect(url_for("login"))

    hashed_pw = admin.get("password", "")
    # ensure using werkzeug.check_password_hash correctly
    if not hashed_pw or not check_password_hash(hashed_pw, password):
        flash("Email hoặc mật khẩu không đúng")
        return redirect(url_for("login"))

    # set session
    session["user_email"] = admin["email"]
    session["username"] = admin.get("username", "Admin")
    return redirect(url_for("index"))

# Separate API-style login (JSON) like bạn có trước
@app.route("/api/login", methods=["POST"])
def api_login():
    try:
        data = request.get_json(force=True)
        email = data.get("email")
        password = data.get("password")
        if not email or not password:
            return jsonify({"success": False, "message": "Thiếu email hoặc mật khẩu"}), 400
        admin = admins.find_one({"email": email})
        if not admin or not check_password_hash(admin.get("password",""), password):
            return jsonify({"success": False, "message": "Email hoặc mật khẩu không đúng"}), 401
        # For API, return success but do NOT set session (unless you want sessions for API)
        return jsonify({"success": True, "username": admin.get("username"), "email": admin.get("email")})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("index"))

# Forgot password (GET shows form, POST sends email)
@app.route("/forgot-password", methods=["GET", "POST"])
def forgot_password():
    if request.method == "GET":
        return render_template_string(FORGOT_PW_HTML)

    email = request.form.get("email")
    if not email:
        return render_template_string(MESSAGE_HTML, msg="Vui lòng nhập email", color="red")

    admin = admins.find_one({"email": email})
    if not admin:
        return render_template_string(MESSAGE_HTML, msg="Email không tồn tại", color="red")

    token = serializer.dumps(email, salt="password-reset")
    reset_link = f"{request.host_url}reset-password/{token}"

    msg = MIMEMultipart("alternative")
    msg["Subject"] = "🔒 Đặt lại mật khẩu - Sun Automation"
    msg["From"] = SMTP_USER or "no-reply@example.com"
    msg["To"] = email
    body = f"""
    <html><body>
    <h3>Xin chào {admin.get('username','Admin')},</h3>
    <p>Bạn vừa yêu cầu đặt lại mật khẩu. Link hợp lệ trong 15 phút:</p>
    <a href="{reset_link}">{reset_link}</a>
    <p>Nếu bạn không yêu cầu, có thể bỏ qua email này.</p>
    <br><b>Sun Automation System</b>
    </body></html>
    """
    msg.attach(MIMEText(body, "html"))

    # Try send mail (if SMTP config provided)
    if not SMTP_USER or not SMTP_PASS:
        # For local testing, just show the link instead of sending email
        return render_template_string(MESSAGE_HTML, msg=f"Link reset (dev): {reset_link}", color="orange")

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASS)
            server.sendmail(msg["From"], [email], msg.as_string())
        return render_template_string(MESSAGE_HTML, msg=f"Link đã gửi đến {email}", color="green")
    except Exception as e:
        return render_template_string(MESSAGE_HTML, msg=f"Lỗi gửi email: {e}", color="red")

# Reset password
@app.route("/reset-password/<token>", methods=["GET", "POST"])
def reset_password(token):
    try:
        email = serializer.loads(token, salt="password-reset", max_age=900)  # 15 phút
    except SignatureExpired:
        return render_template_string(MESSAGE_HTML, msg="Link đã hết hạn", color="red")
    except BadSignature:
        return render_template_string(MESSAGE_HTML, msg="Link không hợp lệ", color="red")

    if request.method == "GET":
        return render_template_string(RESET_PW_HTML, email=email)

    new_password = request.form.get("new_password")
    if not new_password:
        return render_template_string(MESSAGE_HTML, msg="Vui lòng nhập mật khẩu mới", color="red")

    hashed_pw = generate_password_hash(new_password)
    admins.update_one({"email": email}, {"$set": {"password": hashed_pw}})
    return render_template_string(MESSAGE_HTML, msg="Đổi mật khẩu thành công. Vui lòng đăng nhập lại", color="green")

# ---------- Build query function (giữ nguyên logic) ----------
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

# API lấy attendances (bảo vệ bằng login_required)
@app.route("/api/attendances", methods=["GET"])
@login_required
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
        return jsonify({"success": False, "error": str(e)}), 500

# Export excel (bảo vệ)
@app.route("/api/export-excel", methods=["GET"])
@login_required
def export_to_excel():
    try:
        # Tùy theo UI bạn có thể dùng session["user_email"]
        email = session.get("user_email")
        admin = admins.find_one({"email": email})
        if not admin:
            return jsonify({"error": "Không có quyền"}), 403

        filter_type = request.args.get("filter", "hôm nay").lower()
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()
        query = build_query(filter_type, start_date, end_date, search)

        data = list(collection.find(query, {"_id": 0}))
        # template path: nếu bạn dùng file trên disk, đảm bảo tồn tại
        template_path = os.path.join("templates", "Copy of Form chấm công.xlsx")
        if not os.path.exists(template_path):
            # fallback: tạo workbook cơ bản nếu không có template
            from openpyxl import Workbook
            wb = Workbook()
            ws = wb.active
            headers = ["EmployeeId", "EmployeeName", "ProjectId", "Tasks", "Address", "Status", "CheckinDate"]
            for idx, h in enumerate(headers, start=1):
                ws.cell(row=1, column=idx, value=h)
        else:
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
        return jsonify({"error": str(e)}), 500

# ---------- Admin helper route: tạo admin test (LOCAL ONLY) ----------
@app.route("/create-admin", methods=["POST"])
def create_admin():
    """
    Route helper để tạo admin test (LOCAL dev).
    Sử dụng:
      curl -X POST -d "email=you@local&username=You&password=123456" http://localhost:5000/create-admin
    NOTE: Xóa route này khi deploy production.
    """
    try:
        # Chỉ cho phép local hoặc khi biến env DEV=true
        dev_mode = os.getenv("DEV", "true").lower() in ("1", "true", "yes")
        if not dev_mode:
            return jsonify({"error": "Not allowed"}), 403

        email = request.form.get("email")
        username = request.form.get("username", "Admin")
        password = request.form.get("password", "123456")
        if not email or not password:
            return jsonify({"error": "email và password required"}), 400

        if admins.find_one({"email": email}):
            return jsonify({"error": "Email đã tồn tại"}), 400

        hashed = generate_password_hash(password)
        admins.insert_one({"email": email, "username": username, "password": hashed})
        return jsonify({"success": True, "email": email, "username": username})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ---------- Run ----------
if __name__ == "__main__":
    # debug mode controlled by env var
    debug = os.getenv("FLASK_DEBUG", "true").lower() in ("1", "true", "yes")
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", 5000)), debug=debug)
