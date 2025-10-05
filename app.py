from flask import Flask, render_template, jsonify, send_file, request
from pymongo import MongoClient
from flask_cors import CORS
import os
from io import BytesIO
from datetime import datetime, timedelta, timezone
import calendar
import re
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment

app = Flask(__name__, template_folder="templates")
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

# ---- Trang chủ ----
@app.route("/")
def index():
    return render_template("index.html")


# ---- Đăng nhập bằng Email ----
@app.route("/login", methods=["GET"])
def login():
    email = request.args.get("Email")
    if not email:
        return jsonify({"success": False, "message": "❌ Vui lòng nhập email"}), 400

    emp = idx_collection.find_one({"Email": email}, {"_id": 0, "EmployeeName": 1, "EmployeeId": 1, "Email": 1})
    if not emp:
        return jsonify({"success": False, "message": "🚫 Email không tồn tại trong hệ thống"}), 404

    return jsonify({
        "success": True,
        "message": "✅ Đăng nhập thành công",
        "EmployeeId": emp["EmployeeId"],
        "EmployeeName": emp["EmployeeName"],
        "Email": emp["Email"]
    })


# ---- Xây dựng query lọc ----
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


# ---- API lấy dữ liệu chấm công ----
@app.route("/api/attendances", methods=["GET"])
def get_attendances():
    try:
        email = request.args.get("Email")
        if not email:
            return jsonify({"error": "❌ Thiếu email"}), 400

        emp = idx_collection.find_one({"Email": email}, {"EmployeeId": 1, "_id": 0})
        if not emp:
            return jsonify({"error": "🚫 Email không tồn tại"}), 403

        emp_id = emp["EmployeeId"]

        filter_type = request.args.get("filter", "all").lower()
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()

        query = build_query(filter_type, start_date, end_date, search)
        query["EmployeeId"] = emp_id  # Chỉ lấy dữ liệu của chính nhân viên đó

        data = list(collection.find(query, {"_id": 0}))
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ---- API xuất Excel ----
@app.route("/api/export-excel", methods=["GET"])
def export_to_excel():
    try:
        email = request.args.get("Email")
        if not email:
            return jsonify({"error": "❌ Thiếu email"}), 400

        emp = idx_collection.find_one({"Email": email}, {"EmployeeId": 1, "EmployeeName": 1, "_id": 0})
        if not emp:
            return jsonify({"error": "🚫 Email không tồn tại"}), 403

        emp_id = emp["EmployeeId"]
        emp_name = emp["EmployeeName"]

        filter_type = request.args.get("filter", "all").lower()
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()

        query = build_query(filter_type, start_date, end_date, search)
        query["EmployeeId"] = emp_id

        data = list(collection.find(query, {"_id": 0}))

        # ---- Load Excel Template ----
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
        for i, rec in enumerate(data, start=0):
            row = start_row + i
            ws.cell(row=row, column=1, value=emp_id)
            ws.cell(row=row, column=2, value=emp_name)
            ws.cell(row=row, column=3, value=rec.get("CheckinDate"))
            ws.cell(row=row, column=4, value=rec.get("Address"))
            ws.cell(row=row, column=5, value=rec.get("Status"))

            for col in range(1, 6):
                cell = ws.cell(row=row, column=col)
                cell.border = border
                cell.alignment = align_left
            ws.row_dimensions[row].height = 25

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        filename = f"ChamCong_{emp_name}_{datetime.now(VN_TZ).strftime('%d-%m-%Y')}.xlsx"

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
