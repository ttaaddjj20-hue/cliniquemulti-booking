"""
تطبيق حجز المرضى — عيادة الحرش الطبية
=========================================
هذا الجزء فقط مخصص للعمل عبر المتصفح (Chrome) — يحتوي فقط على صفحة
الحجز العمومية (/) واستقبال الطلب (/book). كل ما يخص الاستقبال
والأطباء والفواتير أصبح في تطبيق سطح المكتب المنفصل (desktop_app).

يشارك هذا التطبيق نفس قاعدة البيانات (SQLite) مع تطبيق سطح المكتب عبر
مجلد "data" المشترك — انظر config.py.
"""

import os
from datetime import datetime
from pathlib import Path

import sqlite3
from flask import (
    Flask, render_template, request, redirect, url_for, flash, g
)
from openpyxl import Workbook, load_workbook

from config import DB_PATH, UPLOADS_DIR, EXPORTS_DIR

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "clinique-el-harrach-demo-secret-change-me")

SPECIALTIES = {
    "radiologie": "الأشعة",
    "cardiologie": "القلب",
    "medecine_generale": "الطب العام",
    "ophtalmologie": "العيون",
}

BOOKING_OPEN_HOUR = 6
BOOKING_CLOSE_HOUR = 24

DEMO_PASSWORD = "demo1234"


def specialty_label(key):
    return SPECIALTIES.get(key, key)


app.jinja_env.globals["specialty_label"] = specialty_label


# ---------------------------------------------------------------------------
# Database helpers (نفس مخطط تطبيق سطح المكتب — قاعدة بيانات مشتركة)
# ---------------------------------------------------------------------------

def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
        g.db.execute("PRAGMA journal_mode = WAL")  # يسمح بوصول متزامن من التطبيقين
    return g.db


@app.teardown_appcontext
def close_db(exception=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db():
    """ينشئ الجداول إذا لم تكن موجودة (نفس المخطط المستعمل في تطبيق سطح المكتب).
    لا يُعيد البذر التجريبي إن كانت البيانات موجودة أصلاً (أي تطبيق يشتغل أولاً يُنشئها)."""
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA foreign_keys = ON")
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            role TEXT NOT NULL,
            fullname TEXT NOT NULL,
            specialty TEXT,
            email TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            phone TEXT
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS patients (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fullname TEXT NOT NULL,
            phone TEXT,
            age INTEGER,
            sexe TEXT,
            created_at TEXT NOT NULL
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS exam_categories (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            specialty TEXT NOT NULL,
            name TEXT NOT NULL,
            daily_limit INTEGER
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS exam_subtypes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            category_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            price REAL NOT NULL,
            FOREIGN KEY(category_id) REFERENCES exam_categories(id)
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS appointments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            patient_id INTEGER NOT NULL,
            doctor_id INTEGER,
            specialty TEXT NOT NULL,
            exam TEXT NOT NULL,
            subtype TEXT,
            price REAL DEFAULT 0,
            paid REAL DEFAULT 0,
            remaining REAL DEFAULT 0,
            status TEXT DEFAULT 'en_attente',
            prescription TEXT,
            invoice_number TEXT,
            invoice_date TEXT,
            created_at TEXT NOT NULL,
            FOREIGN KEY(patient_id) REFERENCES patients(id),
            FOREIGN KEY(doctor_id) REFERENCES users(id)
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS prescriptions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            patient_id INTEGER NOT NULL,
            doctor_id INTEGER NOT NULL,
            appointment_id INTEGER,
            date TEXT NOT NULL,
            medicaments TEXT,
            notes TEXT
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS invoices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            invoice_number TEXT UNIQUE,
            appointment_id INTEGER,
            total_ht REAL,
            tva REAL,
            total_ttc REAL,
            created_at TEXT
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS clinic_settings (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            name TEXT, address TEXT, phone TEXT, city TEXT,
            matricule_fiscal TEXT, matricule_stat TEXT,
            article_imposition TEXT, compte_bancaire TEXT
        )
    """)
    conn.commit()

    cur.execute("SELECT COUNT(*) FROM users")
    if cur.fetchone()[0] == 0:
        seed_demo_data(cur)
        conn.commit()

    conn.close()


def seed_demo_data(cur):
    cur.execute(
        "INSERT INTO clinic_settings (id, name, address, phone, city, matricule_fiscal, "
        "matricule_stat, article_imposition, compte_bancaire) VALUES (1,?,?,?,?,?,?,?,?)",
        (
            "Clinique El Harrach Médicale",
            "QUARTIER TIMAKRET CO4 ET CO5 - METLILI CHAAMBA",
            "020.71.86.82 / 020.71.87.29",
            "METLILI CHAAMBA",
            "184470500605171",
            "198447050060526",
            "47050014916",
            "00300299000043330082",
        ),
    )
    staff = [
        ("reception", "الاستقبال", None, "demo.reception@clinique.com", DEMO_PASSWORD, ""),
        ("medecin", "Dr LAHRECHE HOUARI", "radiologie", "demo.radio@clinique.com", DEMO_PASSWORD, ""),
        ("medecin", "Dr AMINE BELABBES", "cardiologie", "demo.cardio@clinique.com", DEMO_PASSWORD, ""),
        ("medecin", "Dr SARAH MERABET", "medecine_generale", "demo.generale@clinique.com", DEMO_PASSWORD, ""),
        ("medecin", "Dr YOUCEF ZEROUKI", "ophtalmologie", "demo.ophta@clinique.com", DEMO_PASSWORD, ""),
    ]
    cur.executemany(
        "INSERT INTO users (role, fullname, specialty, email, password, phone) VALUES (?,?,?,?,?,?)",
        staff,
    )
    catalog = [
        ("radiologie", "Scanner", 2, [("Brain", 7000), ("Abdomen", 6500), ("Thorax", 6000)]),
        ("radiologie", "Échographie", 2, [("Échographie Abdominal", 2000), ("Échographie Pelvic", 3500)]),
        ("radiologie", "Mammographie", 1, [("Mammographie", 5000)]),
        ("radiologie", "Radiographie", None, [("Standard", 2000), ("Spine", 3000)]),
        ("radiologie", "Panoramique", None, [("Panoramique Dentaire", 2500)]),
        ("cardiologie", "Électrocardiogramme", None, [("ECG standard", 1500)]),
        ("cardiologie", "Échocardiographie", 5, [("Échocardiographie transthoracique", 4000), ("Échocardiographie de stress", 6000)]),
        ("cardiologie", "Test d'effort", 3, [("Épreuve d'effort", 5000)]),
        ("cardiologie", "Holter ECG", 2, [("Holter 24h", 4500)]),
        ("cardiologie", "Consultation cardiologique", None, [("Consultation", 2000)]),
        ("medecine_generale", "Consultation générale", None, [("Consultation", 1500)]),
        ("medecine_generale", "Consultation de contrôle", None, [("Contrôle", 1000)]),
        ("medecine_generale", "Certificat médical", None, [("Certificat", 800)]),
        ("medecine_generale", "Vaccination", None, [("Injection / Vaccin", 500)]),
        ("ophtalmologie", "Consultation ophtalmologique", None, [("Consultation", 2000)]),
        ("ophtalmologie", "Examen de la vue", None, [("Réfraction", 1500)]),
        ("ophtalmologie", "Fond d'œil", 10, [("Fond d'œil", 2500)]),
        ("ophtalmologie", "Champ visuel", 5, [("Champ visuel", 3000)]),
    ]
    for specialty, name, limit, subtypes in catalog:
        cur.execute(
            "INSERT INTO exam_categories (specialty, name, daily_limit) VALUES (?,?,?)",
            (specialty, name, limit),
        )
        category_id = cur.lastrowid
        for sub_name, price in subtypes:
            cur.execute(
                "INSERT INTO exam_subtypes (category_id, name, price) VALUES (?,?,?)",
                (category_id, sub_name, price),
            )


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------

def is_open():
    now = datetime.now()
    return BOOKING_OPEN_HOUR <= now.hour < BOOKING_CLOSE_HOUR


def exams_by_specialty_dict():
    db = get_db()
    rows = db.execute("SELECT specialty, name FROM exam_categories ORDER BY id").fetchall()
    result = {key: [] for key in SPECIALTIES}
    for r in rows:
        result.setdefault(r["specialty"], []).append(r["name"])
    return result


def get_category(specialty, exam_name):
    db = get_db()
    return db.execute(
        "SELECT * FROM exam_categories WHERE specialty=? AND name=?", (specialty, exam_name)
    ).fetchone()


def get_default_price(category_id):
    db = get_db()
    row = db.execute(
        "SELECT price FROM exam_subtypes WHERE category_id=? ORDER BY id LIMIT 1", (category_id,)
    ).fetchone()
    return row["price"] if row else 0


def get_exam_usage_today(specialty):
    db = get_db()
    today = datetime.now().date().isoformat()
    rows = db.execute(
        "SELECT exam, COUNT(*) c FROM appointments WHERE specialty=? AND date(created_at)=? "
        "AND status != 'annule' GROUP BY exam",
        (specialty, today),
    ).fetchall()
    return {r["exam"]: r["c"] for r in rows}


def get_daily_folder(base):
    today = datetime.now()
    folder = Path(base) / str(today.year) / f"{today.month:02d}" / f"{today.day:02d}"
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def clean_filename(name):
    return "".join(c for c in name if c.isalnum() or c in (" ", "_")).strip().replace(" ", "_")


def find_or_create_patient(fullname, phone, age):
    db = get_db()
    existing = db.execute("SELECT * FROM patients WHERE phone=?", (phone,)).fetchone()
    if existing:
        return existing["id"]
    cur = db.execute(
        "INSERT INTO patients (fullname, phone, age, created_at) VALUES (?,?,?,?)",
        (fullname, phone, age, datetime.now().isoformat(sep=" ")),
    )
    db.commit()
    return cur.lastrowid


def save_to_excel(fullname, phone, specialty, exam, price, paid, remaining, created_at):
    folder = get_daily_folder(EXPORTS_DIR)
    file_path = folder / "patients.xlsx"
    if file_path.exists():
        wb = load_workbook(file_path)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.append(["Name", "Phone", "Specialty", "Exam", "Price", "Paid", "Remaining", "Date"])
    ws.append([fullname, phone, specialty_label(specialty), exam, price, paid, remaining, created_at])
    wb.save(file_path)


# ---------------------------------------------------------------------------
# Public booking (الصفحتان الوحيدتان في هذا التطبيق)
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template(
        "index.html", open=is_open(),
        specialties=SPECIALTIES, exams_by_specialty=exams_by_specialty_dict(),
    )


@app.route("/book", methods=["POST"])
def book():
    if not is_open():
        flash("❌ الحجز مغلق حالياً", "error")
        return redirect(url_for("index"))

    db = get_db()
    fullname = request.form["fullname"]
    phone = request.form["phone"]
    specialty = request.form["specialty"]
    exam = request.form["exam"]
    age = request.form.get("age", 0) or 0

    if specialty not in SPECIALTIES:
        flash("❌ تخصص غير صالح", "error")
        return redirect(url_for("index"))

    category = get_category(specialty, exam)
    if not category:
        flash("❌ فحص غير صالح", "error")
        return redirect(url_for("index"))

    if category["daily_limit"] is not None:
        usage = get_exam_usage_today(specialty).get(exam, 0)
        if usage >= category["daily_limit"]:
            flash(f"❌ تم الوصول للحد الأقصى لهذا الفحص اليوم ({exam})", "error")
            return redirect(url_for("index"))

    price = get_default_price(category["id"])
    now = datetime.now().isoformat(sep=" ")
    patient_id = find_or_create_patient(fullname, phone, age)

    filename = None
    file = request.files.get("prescription")
    if file and file.filename != "":
        folder = get_daily_folder(UPLOADS_DIR)
        safe_name = clean_filename(fullname)
        ext = file.filename.rsplit(".", 1)[-1].lower()
        filepath = folder / f"{safe_name}_{int(datetime.now().timestamp())}.{ext}"
        file.save(str(filepath))
        filename = str(filepath.relative_to(UPLOADS_DIR))

    db.execute("""
        INSERT INTO appointments (patient_id, specialty, exam, price, paid, remaining,
                                    status, prescription, created_at)
        VALUES (?, ?, ?, ?, 0, ?, 'en_attente', ?, ?)
    """, (patient_id, specialty, exam, price, price, filename, now))
    db.commit()

    save_to_excel(fullname, phone, specialty, exam, price, 0, price, now)
    flash("✅ تم إرسال طلبك بنجاح، سيتم استدعاؤكم من طرف الاستقبال")
    return redirect(url_for("index"))


# تهيئة قاعدة البيانات عند تشغيل التطبيق
# يعمل سواءً شُغّل التطبيق مباشرة أو بواسطة Gunicorn على Render
init_db()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)