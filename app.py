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

if not MONGO_URI.strip():
    raise ValueError("❌ Lỗi: MONGO_URI chưa được cấu hình!")

# ---- Kết nối MongoDB ----
try:
    client = MongoClient(MONGO_URI)
    db = client[DB_NAME]
    collection = db["alt_checkins"]
    idx_collection = db["idx_collection"]
except Exception as e:
    raise RuntimeError(f"❌ Không thể kết nối MongoDB: {e}")

# ---- Danh sách được phép truy cập ----
ALLOWED_IDS = {"it.trankhanhvinh@gmail.com", "thinhnv@sunautomation.com.vn", "kimcuong@sunautomation.com.vn"}

# ---- Trang chủ ----
@app.route("/")
def index():
    return render_template("index.html")

# ---- Login ----
@app.route("/login", methods=["GET"])
def login():
    emp_id = request.args.get("Email")
    if not emp_id:
        return jsonify({"success": False, "message": "❌ Bạn cần nhập email"}), 400

    if emp_id in ALLOWED_IDS:
        emp = idx_collection.find_one({"Email": Email}, {"_id": 0, "EmployeeName": 1})
        emp_name = emp["EmployeeName"] if emp else emp_id
        return jsonify({
            "success": True,
            "message": "✅ Đăng nhập thành công",
            "EmployeeId": emp_id,
            "EmployeeName": emp_name
        })
    else:
        return jsonify({"success": False, "message": "🚫 Không có quyền"}), 403


# ---- Xây dựng query ----
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
        last_day = calendar.monthrange(today.year, today.month)[1]
        end = today.replace(day=last_day).strftime("%Y-%m-%d")
        query["CheckinDate"] = {"$gte": start, "$lte": end}
    elif filter_type == "năm":
        start = today.replace(month=1, day=1).strftime("%Y-%m-%d")
        end = today.replace(month=12, day=31).strftime("%Y-%m-%d")
        query["CheckinDate"] = {"$gte": start, "$lte": end}

    if search:
        regex = re.compile(search, re.IGNORECASE)
        query["$or"] = [
            {"EmployeeId": {"$regex": regex}},
            {"EmployeeName": {"$regex": regex}}
        ]
    return query


# ---- API lấy dữ liệu ----
@app.route("/api/attendances", methods=["GET"])
def get_attendances():
    try:
        emp_id = request.args.get("empId")
        if emp_id not in ALLOWED_IDS:
            return jsonify({"error": "🚫 Không có quyền!"}), 403

        filter_type = request.args.get("filter", "all").lower()
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()

        query = build_query(filter_type, start_date, end_date, search)
        data = list(collection.find(query, {"_id": 0}))
        return jsonify(data), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ---- API xuất Excel ----
@app.route("/api/export-excel", methods=["GET"])
def export_to_excel():
    try:
        emp_id = request.args.get("Email")
        if emp_id not in ALLOWED_IDS:
            return jsonify({"error": "🚫 Không có quyền xuất Excel!"}), 403

        filter_type = request.args.get("filter", "all").lower()
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()

        query = build_query(filter_type, start_date, end_date, search)
        data = list(collection.find(query, {"_id": 0}))

        # ---- Group theo nhân viên + ngày ----
        grouped = {}
        for d in data:
            emp_id = d.get("EmployeeId", "")
            emp_name = d.get("EmployeeName", "")
            date = d.get("CheckinDate")
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

        start_row = 2
        for i, ((emp_id, emp_name, date), records) in enumerate(grouped.items(), start=0):
            row = start_row + i
            ws.cell(row=row, column=1, value=emp_id)
            ws.cell(row=row, column=2, value=emp_name)
            ws.cell(row=row, column=3, value=date)

            # ---- Kiểm tra nghỉ phép ----
            leave_entries = [r for r in records if "Nghỉ phép" in (r.get("Tasks") or [])]
            if leave_entries:
                r = leave_entries[-1]  # bản ghi nghỉ phép mới nhất
                parts = ["Nghỉ phép"]
                if r.get("OtherNote"):
                    parts.append(f"Lý do: {r['OtherNote']}")
                parts.append(r.get("Status", "Chờ duyệt"))
                if r.get("ApprovedBy"):
                    parts.append(f"Duyệt bởi {r['ApprovedBy']}")
                if r.get("Address"):
                    parts.append(r["Address"])
                ws.cell(row=row, column=4, value="; ".join(parts))
            else:
                # ---- Các bản ghi thường ----
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
                    if time_str:
                        parts.append(f"{time_str}")
                    if rec.get("ProjectId"):
                        parts.append(rec["ProjectId"])
                    if rec.get("Tasks"):
                        tasks = ", ".join(rec["Tasks"]) if isinstance(rec["Tasks"], list) else rec["Tasks"]
                        parts.append(tasks)
                    if rec.get("OtherNote"):
                        parts.append(rec["OtherNote"])
                    if rec.get("Address"):
                        parts.append(rec["Address"])

                    ws.cell(row=row, column=3 + j, value="; ".join(parts))

            # ---- Border + căn lề ----
            for col in range(1, 14):
                cell = ws.cell(row=row, column=col)
                cell.border = border
                cell.alignment = align_left

            # ---- Tự căn chiều cao dòng ----
            ws.row_dimensions[row].height = 25

        # ---- Auto fit độ rộng cột ----
        for col in ws.columns:
            max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 80)

        # ---- Xuất file ----
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        today_str = datetime.now(VN_TZ).strftime("%d-%m-%Y")
        filename = f"ChamCong_{today_str}.xlsx"

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
