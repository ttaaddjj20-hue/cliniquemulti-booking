"""
إعدادات مشتركة — يحدد مكان قاعدة البيانات المشتركة بين تطبيق الحجز (الويب)
وتطبيق سطح المكتب (الاستقبال / الأطباء).

الافتراضي: مجلد "data" بجانب مجلد المشروع (../data بالنسبة لهذا الملف).
يمكن تغييره عبر متغير البيئة CLINIQUE_DATA_DIR (مثلاً لتوجيه التطبيقين
نحو مجلد مشترك على الشبكة إذا كانا يعملان على جهازين مختلفين).
"""

import os

_THIS_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_DIR = os.environ.get("CLINIQUE_DATA_DIR") or os.path.join(_THIS_DIR, "..", "data")
DATA_DIR = os.path.abspath(DATA_DIR)

DB_PATH = os.path.join(DATA_DIR, "database.db")
UPLOADS_DIR = os.path.join(DATA_DIR, "uploads")
EXPORTS_DIR = os.path.join(DATA_DIR, "exports")

os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(UPLOADS_DIR, exist_ok=True)
os.makedirs(EXPORTS_DIR, exist_ok=True)
