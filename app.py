from flask import Flask, render_template, jsonify, send_file, request
from pymongo import MongoClient
from flask_cors import CORS
import os
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta, timezone
import calendar
import re
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment
from openpyxl.utils import get_column_letter
from collections import defaultdict

app = Flask(__name__, template_folder="templates")
CORS(app)

# ---- Timezone VN ----
VN_TZ = timezone(timedelta(hours=7))

# ---- Load MONGO_URI ----
MONGO_URI = os.getenv(
    "MONGO_URI",
    "mongodb+srv://banhbaobeo2205:lm2hiCLXp6B0D7hq@cluster0.festnla.mongodb.net/?retryWrites=true&w=majority"
)
DB_NAME = os.getenv("DB_NAME", "Sun_Database_1")

if not MONGO_URI or MONGO_URI.strip() == "":
    raise ValueError("❌ Lỗi: MONGO_URI chưa được cấu hình!")

# ---- Kết nối MongoDB ----
try:
    client = MongoClient(MONGO_URI)
    db = client[DB_NAME]
    collection = db["alt_checkins"]
    idx_collection = db["idx_collection"]
except Exception as e:
    raise RuntimeError(f"❌ Không thể kết nối MongoDB: {e}")

# ---- Danh sách NV được phép vào trang xem dữ liệu ----
ALLOWED_IDS = {"Admin", "Admin01", "Admin02", "Admin03"}


# ---- API: Trang index ----
@app.route("/")
def index():
    return render_template("index.html")


# ---- API: Login ----
@app.route("/login", methods=["GET"])
def login():
    emp_id = request.args.get("empId")
    if not emp_id:
        return jsonify({"success": False, "message": "❌ Bạn cần nhập mã nhân viên"}), 400

    if emp_id in ALLOWED_IDS:
        emp = idx_collection.find_one({"EmployeeId": emp_id}, {"_id": 0, "EmployeeName": 1})
        emp_name = emp["EmployeeName"] if emp else emp_id
        return jsonify({
            "success": True,
            "message": "✅ Đăng nhập thành công",
            "EmployeeId": emp_id,
            "EmployeeName": emp_name
        })
    else:
        return jsonify({"success": False, "message": "🚫 Mã nhân viên của bạn không có quyền truy cập"}), 403


# ---- Xây dựng query cho filter ----
def build_query(filter_type, start_date, end_date, search):
    query = {}
    today = datetime.now(VN_TZ)

    # ---- Lọc theo thời gian ----
    if filter_type == "custom" and start_date and end_date:
        query["CheckinDate"] = {"$gte": start_date, "$lte": end_date}
    elif filter_type == "today":
        today_str = today.strftime("%Y-%m-%d")
        query["CheckinDate"] = today_str
    elif filter_type == "week":
        start = (today - timedelta(days=today.weekday())).strftime("%Y-%m-%d")
        end = (today + timedelta(days=6 - today.weekday())).strftime("%Y-%m-%d")
        query["CheckinDate"] = {"$gte": start, "$lte": end}
    elif filter_type == "month":
        start = today.replace(day=1).strftime("%Y-%m-%d")
        last_day = calendar.monthrange(today.year, today.month)[1]
        end = today.replace(day=last_day).strftime("%Y-%m-%d")
        query["CheckinDate"] = {"$gte": start, "$lte": end}
    elif filter_type == "year":
        start = today.replace(month=1, day=1).strftime("%Y-%m-%d")
        end = today.replace(month=12, day=31).strftime("%Y-%m-%d")
        query["CheckinDate"] = {"$gte": start, "$lte": end}

    # ---- Lọc theo NV (cả mã & tên) ----
    if search:
        regex = re.compile(search, re.IGNORECASE)
        query["$or"] = [
            {"EmployeeId": {"$regex": regex}},
            {"EmployeeName": {"$regex": regex}}
        ]

    return query


# ---- API: Lấy danh sách chấm công ----
@app.route("/api/attendances", methods=["GET"])
def get_attendances():
    try:
        emp_id = request.args.get("empId")
        if emp_id not in ALLOWED_IDS:
            return jsonify({"error": "🚫 Không có quyền truy cập!"}), 403

        filter_type = request.args.get("filter", "all")
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()

        query = build_query(filter_type, start_date, end_date, search)

        data = list(collection.find(query, {
            "_id": 0,
            "EmployeeId": 1,
            "EmployeeName": 1,
            "ProjectId": 1,
            "Tasks": 1,
            "OtherNote": 1,
            "Address": 1,
            "CheckinTime": 1,
            "Status": 1,
            "FaceImage": 1
        }))

        # Convert datetime -> string
        for d in data:
            if isinstance(d.get("CheckinTime"), datetime):
                d["CheckinTime"] = d["CheckinTime"].astimezone(VN_TZ).strftime("%d/%m/%Y %H:%M:%S")

        return jsonify(data), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500


ch", "").strip()
        query = build_query(filter_type, start_date, end_date, search)

        # Query dữ liệu
        data = list(collection.find(query, {
            "_id": 0,
            "EmployeeId": 1,
            "EmployeeName": 1,
            "ProjectId": 1,
            "Tasks": 1,
            "OtherNote": 1,
            "Address": 1,
            "CheckinTime": 1
        }))

        # ---- Group theo EmployeeId + CheckinDate ----
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

        # Load template
        template_path = "templates/Copy of Form chấm công.xlsx"
        wb = load_workbook(template_path)
        ws = wb.active

        border = Border(
            left=Side(border_style="thin", color="000000"),
            right=Side(border_style="thin", color="000000"),
            top=Side(border_style="thin", color="000000"),
            bottom=Side(border_style="thin", color="000000")
        )
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

        start_row = 2
        for i, ((emp_id, emp_name, date), records) in enumerate(grouped.items(), start=0):
            row = start_row + i
            ws.cell(row=row, column=1, value=emp_id)
            ws.cell(row=row, column=2, value=emp_name)
            ws.cell(row=row, column=3, value=date)

            # Thêm dữ liệu check 1–10
            for j, rec in enumerate(records[:10], start=1):
                time_str = ""
                if isinstance(rec.get("CheckinTime"), datetime):
                    time_str = rec["CheckinTime"].astimezone(VN_TZ).strftime("%H:%M:%S")

                entry = f"{time_str}"
                if rec.get("ProjectId"):
                    entry += f" ; ID: {rec['ProjectId']}"
                if rec.get("Tasks"):
                    tasks = ", ".join(rec["Tasks"]) if isinstance(rec["Tasks"], list) else rec["Tasks"]
                    entry += f" ; Công việc: {tasks}"
                if rec.get("OtherNote"):
                    entry += f" ; Ghi chú khác: {rec['OtherNote']}"
                if rec.get("Address"):
                    entry += f" ; Địa chỉ: {rec['Address']}"

                ws.cell(row=row, column=3 + j, value=entry)

            # ---- Áp dụng border + align cho cả row đến Check10 ----
            for col in range(1, 13):  # 1..12 = Mã NV → Check 10
                cell = ws.cell(row=row, column=col)
                cell.alignment = align_left
                cell.border = border

        # Xuất file
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        filename = f"Cham_cong_{filter_type}_{datetime.now().strftime('%d-%m-%Y')}.xlsx"
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500




if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
