from flask import Flask, jsonify, send_file, request, render_template_string, redirect, url_for
from pymongo import MongoClient
from flask_cors import CORS
from io import BytesIO
from datetime import datetime, timedelta, timezone
from werkzeug.security import generate_password_hash, check_password_hash
import calendar
import re
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment
import os
import uuid

# ====== Cấu hình Flask ======
app = Flask(__name__)
CORS(app)

# ====== Kết nối MongoDB ======
MONGO_URI = os.getenv("MONGO_URI", "mongodb://localhost:27017/")
client = MongoClient(MONGO_URI)
db = client["gps_attendance"]
idx_collection = db["idx_collection"]
alt_checkins = db["alt_checkins"]
admins_collection = db["admins"]

# ====== Múi giờ Việt Nam ======
VN_TZ = timezone(timedelta(hours=7))

# ==========================================================
# 🧩 QUÊN MẬT KHẨU — HIỂN THỊ FORM ĐẶT LẠI NGAY
# ==========================================================

# Trang nhập email quên mật khẩu
@app.route("/forgot-password", methods=["GET", "POST"])
def forgot_password():
    if request.method == "GET":
        return render_template_string("""
            <html>
            <head><title>Quên mật khẩu</title></head>
            <body style="font-family: sans-serif; text-align: center; margin-top: 80px;">
                <h2>🔑 Quên mật khẩu</h2>
                <form method="POST">
                    <input type="email" name="email" placeholder="Nhập email của bạn" required style="padding: 8px; width: 250px;"><br><br>
                    <button type="submit" style="padding: 8px 20px;">Tiếp tục</button>
                </form>
            </body>
            </html>
        """)
    else:
        email = request.form.get("email")
        admin = admins_collection.find_one({"email": email})
        if not admin:
            return "<h3 style='color:red;text-align:center;'>❌ Không tìm thấy tài khoản với email này</h3>"

        # Tạo token tạm lưu vào DB
        token = str(uuid.uuid4())
        expiry = datetime.now(VN_TZ) + timedelta(minutes=10)
        admins_collection.update_one({"email": email}, {"$set": {"reset_token": token, "reset_expiry": expiry}})

        # Chuyển hướng đến trang reset mật khẩu
        return redirect(url_for("reset_password", token=token))


# Trang đặt lại mật khẩu (tự động hiện sau khi xác thực email)
@app.route("/reset-password", methods=["GET", "POST"])
def reset_password():
    token = request.args.get("token")
    user = admins_collection.find_one({"reset_token": token})

    if not user:
        return "<h3 style='color:red;text-align:center;'>❌ Token không hợp lệ hoặc đã hết hạn</h3>"

    # Kiểm tra hạn token
    if datetime.now(VN_TZ) > user.get("reset_expiry", datetime.min):
        return "<h3 style='color:red;text-align:center;'>⚠️ Token đã hết hạn, vui lòng làm lại bước quên mật khẩu</h3>"

    if request.method == "GET":
        return render_template_string(f"""
            <html>
            <head><title>Đặt lại mật khẩu</title></head>
            <body style="font-family: sans-serif; text-align: center; margin-top: 80px;">
                <h2>🔒 Đặt lại mật khẩu cho {user['email']}</h2>
                <form method="POST">
                    <input type="hidden" name="token" value="{token}">
                    <input type="password" name="new_password" placeholder="Nhập mật khẩu mới" required style="padding:8px; width:250px;"><br><br>
                    <button type="submit" style="padding:8px 20px;">Cập nhật mật khẩu</button>
                </form>
            </body>
            </html>
        """)
    else:
        new_password = request.form.get("new_password")
        hashed_pw = generate_password_hash(new_password)

        admins_collection.update_one({"_id": user["_id"]}, {
            "$set": {"password": hashed_pw},
            "$unset": {"reset_token": "", "reset_expiry": ""}
        })
        return "<h3 style='color:green;text-align:center;'>✅ Mật khẩu đã được đặt lại thành công!</h3>"


# ==========================================================
# 🧾 API XUẤT EXCEL (KHÔNG ĐỔI)
# ==========================================================
@app.route("/api/export-excel", methods=["GET"])
def export_to_excel():
    try:
        email = request.args.get("email")
        if not email:
            return jsonify({"error": "❌ Thiếu email người dùng"}), 400

        # --- Tìm admin theo email ---
        admin = admins_collection.find_one({"email": email})
        if not admin:
            return jsonify({"error": "❌ Không tìm thấy tài khoản admin với email này"}), 404

        # --- Lấy thông tin nhân viên từ collection chính ---
        employees = list(idx_collection.find({}, {"EmployeeId": 1, "EmployeeName": 1, "_id": 0}))

        # --- Mở template Excel ---
        template_path = os.path.join("templates", "template.xlsx")
        wb = load_workbook(template_path)
        ws = wb.active

        # --- Thông tin xuất file ---
        now = datetime.now(VN_TZ)
        current_month = now.month
        current_year = now.year
        month_name = calendar.month_name[current_month]

        thin = Side(border_style="thin", color="000000")
        border = Border(top=thin, left=thin, right=thin, bottom=thin)
        center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

        row = 4
        for emp in employees:
            emp_id = emp.get("EmployeeId")
            emp_name = emp.get("EmployeeName")

            records = list(
                alt_checkins.find(
                    {
                        "EmployeeId": emp_id,
                        "CheckinTime": {
                            "$gte": datetime(current_year, current_month, 1, tzinfo=VN_TZ),
                            "$lt": datetime(
                                current_year,
                                current_month + 1 if current_month < 12 else 1,
                                1 if current_month < 12 else 1,
                                tzinfo=VN_TZ,
                            ),
                        },
                    }
                ).sort("CheckinTime", 1)
            )

            if not records:
                continue

            ws.cell(row=row, column=1, value=emp_name)
            ws.cell(row=row, column=2, value=emp_id)

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
                if isinstance(tasks, list):
                    tasks_str = ", ".join(tasks)
                else:
                    tasks_str = str(tasks or "")

                leave_reason = ""
                if "nghỉ phép" in tasks_str.lower():
                    if ":" in tasks_str:
                        split_task = tasks_str.split(":", 1)
                        tasks_str = split_task[0].strip()
                        leave_reason = split_task[1].strip()
                    else:
                        tasks_str = tasks_str.strip()

                status = rec.get("Status", "")
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
                ws.cell(row=row, column=2 + j, value=entry)

            row += 1

        for r in ws.iter_rows(min_row=4, max_row=row - 1, min_col=1, max_col=15):
            for cell in r:
                cell.border = border
                cell.alignment = center_align

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        filename = f"ChamCong_{month_name}_{current_year}.xlsx"
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ====== Chạy thử Flask ======
if __name__ == "__main__":
    app.run(debug=True, port=5000)
