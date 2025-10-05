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
    email = request.args.get("email")  # ✅ trùng key với front-end
    if not email:
        return jsonify({"success": False, "message": "❌ Vui lòng nhập email"}), 400

    emp = idx_collection.find_one(
        {"Email": email},
        {"_id": 0, "EmployeeName": 1, "EmployeeId": 1, "Email": 1}
    )

    if not emp:
        return jsonify({"success": False, "message": "🚫 Email không tồn tại trong hệ thống"}), 404

    return jsonify({
        "success": True,
        "message": "✅ Đăng nhập thành công",
        "EmployeeId": emp["EmployeeId"],
        "EmployeeName": emp["EmployeeName"],
        "Email": emp["Email"]
    })


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


# ---- API lấy dữ liệu chấm công ----
@app.route("/api/attendances", methods=["GET"])
def get_attendances():
    try:
        email = request.args.get("email")  # ✅ trùng key với front-end
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
        query["EmployeeId"] = emp_id  # chỉ lấy dữ liệu của chính nhân viên

        data = list(collection.find(query, {"_id": 0}))
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


from flask import send_file, jsonify, request
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment
from io import BytesIO
from datetime import datetime

@app.route("/api/export-excel", methods=["GET"])
def export_to_excel():
    try:
        email = request.args.get("email")  # email người xuất file
        if not email:
            return jsonify({"error": "❌ Thiếu email"}), 400

        emp = db.idx_collection.find_one({"Email": email}, {"EmployeeId": 1, "EmployeeName": 1, "_id": 0})
        if not emp:
            return jsonify({"error": "🚫 Email không tồn tại"}), 403

        emp_id = emp["EmployeeId"]
        emp_name = emp["EmployeeName"]

        # ---- Tham số lọc ----
        filter_type = request.args.get("filter", "all").lower()
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()

        # ---- Tạo query ----
        query = {}
        if filter_type == "hôm nay":
            query["CheckinDate"] = datetime.now(VN_TZ).strftime("%Y-%m-%d")
        elif filter_type == "custom" and start_date and end_date:
            query["CheckinDate"] = {"$gte": start_date, "$lte": end_date}
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
                # ---- Parse giờ ----
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

                # ---- Build nội dung ô ----
                parts = []

                tasks = rec.get("Tasks")
                if isinstance(tasks, list):
                    tasks_str = ", ".join(tasks)
                else:
                    tasks_str = str(tasks or "")

                status = rec.get("Status", "")
                is_leave = "nghỉ phép" in tasks_str.lower()

                if is_leave:
                    # ---- Thêm số hiệu trạng thái ----
                    if "đã duyệt" in status.lower():
                        parts.append("1")
                    elif "từ chối" in status.lower():
                        parts.append("3")
                    else:
                        parts.append("2")  # Chờ duyệt

                # ---- Giờ ----
                if time_str:
                    parts.append(time_str)

                # ---- Project ID ----
                if rec.get("ProjectId"):
                    parts.append(str(rec["ProjectId"]))

                # ---- Tasks ----
                if tasks_str:
                    parts.append(tasks_str)

                # ---- Status nếu là nghỉ phép ----
                if is_leave and status:
                    parts.append(status)

                # ---- Ghi chú khác ----
                if rec.get("OtherNote"):
                    parts.append(rec["OtherNote"])

                # ---- Địa chỉ ----
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
