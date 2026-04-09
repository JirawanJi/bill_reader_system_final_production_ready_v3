import os
import uuid
from datetime import datetime
from functools import wraps

from flask import (
    Flask, render_template, request, redirect, url_for,
    flash, session, send_from_directory, jsonify
)
from werkzeug.utils import secure_filename

from database import db, User, UploadHistory, init_db
from parser_pea import parse_pea_pdf
from parser_mea import parse_mea_pdf
from fv60_export import (
    load_cost_center_mapping,
    export_dynamic_fv60_excel,
)

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
EXPORT_FOLDER = os.path.join(BASE_DIR, "exports")
MAPPING_FILE = os.path.join(BASE_DIR, "Cost_Center_2026_PROFIT_CENTER.xlsx")

ALLOWED_EXTENSIONS = {"pdf"}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(EXPORT_FOLDER, exist_ok=True)

app = Flask(__name__)
app.config["SECRET_KEY"] = os.getenv("SECRET_KEY", "dev-secret-key")

database_url = os.getenv("DATABASE_URL", f"sqlite:///{os.path.join(BASE_DIR, 'bill_reader.db')}")
app.config["SQLALCHEMY_DATABASE_URI"] = database_url
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db.init_app(app)
init_db(app)


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def login_required(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        if "user_id" not in session:
            flash("กรุณาเข้าสู่ระบบก่อน", "warning")
            return redirect(url_for("login"))
        return func(*args, **kwargs)
    return wrapper


def get_mapping():
    if os.path.exists(MAPPING_FILE):
        return load_cost_center_mapping(MAPPING_FILE)
    return {}


@app.route("/")
def home():
    if "user_id" in session:
        return redirect(url_for("dashboard"))
    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()

        if not username or not password:
            flash("กรุณากรอก Username และ Password", "danger")
            return render_template("login.html")

        user = User.query.filter_by(username=username).first()

        if not user:
            user = User(username=username)
            user.set_password(password)
            db.session.add(user)
            db.session.commit()
        else:
            if not user.check_password(password):
                flash("รหัสผ่านไม่ถูกต้อง", "danger")
                return render_template("login.html")

        session["user_id"] = user.id
        session["username"] = user.username

        flash("เข้าสู่ระบบสำเร็จ", "success")
        return redirect(url_for("dashboard"))

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    flash("ออกจากระบบแล้ว", "info")
    return redirect(url_for("login"))


@app.route("/dashboard")
@login_required
def dashboard():
    histories = (
        UploadHistory.query
        .filter_by(user_id=session["user_id"])
        .order_by(UploadHistory.created_at.desc())
        .limit(50)
        .all()
    )
    return render_template(
        "dashboard.html",
        histories=histories,
        username=session.get("username", "")
    )


@app.route("/pea")
@login_required
def pea_page():
    return render_template("pea.html")


@app.route("/mea")
@login_required
def mea_page():
    return render_template("mea.html")


@app.route("/upload/<bill_type>", methods=["POST"])
@login_required
def upload_bill(bill_type):
    if bill_type not in ["pea", "mea"]:
        return jsonify({
            "success": False,
            "message": "ประเภทบิลไม่ถูกต้อง"
        }), 400

    files = request.files.getlist("files")
    if not files:
        return jsonify({
            "success": False,
            "message": "กรุณาเลือกไฟล์ PDF"
        }), 400

    mapping = get_mapping()
    rows_for_export = []
    rows_for_display = []
    errors = []

    for file in files:
        if not file or file.filename == "":
            continue

        original_name = secure_filename(file.filename)

        if not allowed_file(original_name):
            errors.append(f"{original_name}: รองรับเฉพาะไฟล์ PDF")
            continue

        unique_name = f"{uuid.uuid4().hex}_{original_name}"
        save_path = os.path.join(UPLOAD_FOLDER, unique_name)
        file.save(save_path)

        try:
            if bill_type == "pea":
                row = parse_pea_pdf(
                    save_path,
                    mapping,
                    original_filename=original_name
                )
            else:
                row = parse_mea_pdf(
                    save_path,
                    mapping,
                    original_filename=original_name
                )

            # debug แบบสั้น ไม่พิมพ์ raw_text_preview ยาว ๆ
            print(
                "APP ROW DEBUG =",
                {
                    "reference": row.get("reference", ""),
                    "invoice_date": row.get("invoice_date", ""),
                    "posting_date": row.get("posting_date", ""),
                    "amount": row.get("amount", ""),
                    "store_id": row.get("store_id", ""),
                    "cost_center": row.get("cost_center", ""),
                    "profit_center": row.get("profit_center", ""),
                    "filename": row.get("filename", original_name),
                }
            )

            rows_for_export.append(row)

            rows_for_display.append({
                "source_file": original_name,
                "filename": row.get("filename", original_name),
                "reference": row.get("reference", ""),
                "invoice_date": row.get("invoice_date", ""),
                "amount": row.get("amount", ""),
                "store_id": row.get("store_id", ""),
                "cost_center": row.get("cost_center", ""),
                "profit_center": row.get("profit_center", ""),
            })

            print("APP DISPLAY DEBUG =", rows_for_display[-1])

        except Exception as e:
            errors.append(f"{original_name}: {str(e)}")

    if not rows_for_export:
        return jsonify({
            "success": False,
            "message": "ไม่สามารถอ่านข้อมูลจากไฟล์ได้",
            "errors": errors
        }), 400

    export_filename = f"{bill_type}_dynamic_fv60_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    export_path = os.path.join(EXPORT_FOLDER, export_filename)

    try:
        export_dynamic_fv60_excel(rows_for_export, export_path)
    except Exception as e:
        return jsonify({
            "success": False,
            "message": f"สร้างไฟล์ Excel ไม่สำเร็จ: {str(e)}"
        }), 500

    history = UploadHistory(
        user_id=session["user_id"],
        bill_type=bill_type.upper(),
        total_files=len(files),
        success_files=len(rows_for_export),
        failed_files=len(errors),
        export_file=export_filename,
        error_message="\n".join(errors) if errors else None
    )
    db.session.add(history)
    db.session.commit()

    return jsonify({
        "success": True,
        "message": f"ประมวลผลสำเร็จ {len(rows_for_export)} ไฟล์",
        "download_url": url_for("download_export", filename=export_filename),
        "rows": rows_for_display,
        "errors": errors
    })


@app.route("/download/<path:filename>")
@login_required
def download_export(filename):
    return send_from_directory(EXPORT_FOLDER, filename, as_attachment=True)


@app.route("/api/history")
@login_required
def api_history():
    histories = (
        UploadHistory.query
        .filter_by(user_id=session["user_id"])
        .order_by(UploadHistory.created_at.desc())
        .all()
    )

    data = []
    for item in histories:
        data.append({
            "id": item.id,
            "bill_type": item.bill_type,
            "total_files": item.total_files,
            "success_files": item.success_files,
            "failed_files": item.failed_files,
            "export_file": item.export_file,
            "error_message": item.error_message,
            "created_at": item.created_at.strftime("%Y-%m-%d %H:%M:%S")
        })

    return jsonify(data)


if __name__ == "__main__":
    app.run(debug=True)