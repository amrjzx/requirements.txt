# -*- coding: utf-8 -*-
"""
Aumet | لوحة تحكم ومتابعة تنفيذ مشاريع الرعاية الصحية — الإصدار v5 (SQLite Live Engine)
=========================================================================================
يتطلب التشغيل: streamlit, pandas, folium, streamlit-folium, openpyxl
أمر التشغيل:    streamlit run aumet_dashboard_sqlite.py

أهم تغييرات هذا الإصدار:
  • التحول الكامل من Excel كمصدر بيانات أساسي إلى قاعدة بيانات SQLite حية (aumet_system.db)
    تدعم عدة مستخدمين في نفس الوقت بأمان (WAL mode + busy timeout).
  • جدول تدقيق مستقل (activity_logs) يسجل كل عملية: التاريخ والوقت، المستخدم، نوع الحركة، المنشأة.
  • زر "مزامنة وعكس ملف التفعيل لايف" يقرأ ملف Health Care Center Activation_2.xlsx ديناميكياً
    ويحدّث حالة المراكز في القاعدة مباشرة (SQL UPDATE) باستخدام خوارزمية super_fuzzy_match.
  • حماية تعديل يدوي (is_manually_locked) تمنع المزامنة التلقائية من المساس بأي مركز عدّله المسؤول يدوياً.
  • واجهة مرتبطة بالكامل بقراءات مباشرة وسريعة من القاعدة (KPIs / Health Score / الخريطة / آخر العمليات).
  • حماية كاملة من الأعطال: busy_timeout، معالجة NaN، وزر إعادة ضبط شامل للنظام.
"""

import os
import re
import io
import sqlite3
import shutil
from datetime import datetime
from contextlib import contextmanager

import pandas as pd
import streamlit as st
import folium
from folium.plugins import MarkerCluster
from streamlit_folium import st_folium

# ══════════════════════════════════════════════════════════════════════
# ── ثوابت الإعداد العام
# ══════════════════════════════════════════════════════════════════════

ADMIN_PASSWORD = "Aumet@2025"   # ⚠️ كلمة مرور المسؤول المعتمدة

DB_PATH = "aumet_system.db"

FILES = {
    "hospitals":       "وزارة_الصحة_المستشفيات.xlsx",
    "subcenters":      "وزارة_الصحة_المراكز_الفرعية.xlsx",
    "old_centers":     "وزارة_الصحة_(1).xlsx",
    "activation_live": "Health Care Center Activation_2.xlsx",   # ملف التفعيل الحي المعتمد للمزامنة
    "activation_fallback": "Health Care Center Activation.xlsx", # نسخة احتياطية إن لم يوجد الملف أعلاه
    "logo":            "aumet_logo.jpg",
    "backup_dir":       "backups",
}

# ── الهوية البصرية المعتمدة لـ Aumet ──────────────────────────────────────────
AUMET_TEAL   = "#00C9A7"
AUMET_NAVY   = "#3D4B63"
AUMET_DARK   = "#2C3A4F"
AUMET_LIGHT  = "#F0FBF8"
AUMET_WHITE  = "#FFFFFF"
AUMET_GRAY   = "#8A9BB0"
AUMET_BORDER = "#D1E8E2"

PHASES = [
    {"key": "medicines",   "ar": "الأدوية",          "full": "مرحلة الأدوية",          "en": "Medicines",           "icon": "💊"},
    {"key": "consumables", "ar": "المستهلكات الطبية", "full": "مرحلة المستهلكات الطبية", "en": "Medical Consumables", "icon": "🧰"},
    {"key": "vaccines",    "ar": "المطاعيم",          "full": "مرحلة المطاعيم",          "en": "Vaccines",            "icon": "💉"},
]
PHASE_KEYS = [p["key"] for p in PHASES]

def status_col(key):  return f"status_{key}"
def percent_col(key): return f"percent_{key}"
def notes_col(key):   return f"notes_{key}"

STATUS_OPTIONS = [
    "لم يبدأ بعد / Not Started",
    "قيد التنفيذ / In Progress",
    "مفعّل ويعمل / Live",
    "متوقف / On Hold",
]
STATUS_COLORS = {
    STATUS_OPTIONS[0]: "#4A5568",
    STATUS_OPTIONS[1]: "#D97706",
    STATUS_OPTIONS[2]: "#15803D",
    STATUS_OPTIONS[3]: "#B91C1C",
}
STATUS_ICONS = {
    STATUS_OPTIONS[0]: "⚪",
    STATUS_OPTIONS[1]: "🔄",
    STATUS_OPTIONS[2]: "✅",
    STATUS_OPTIONS[3]: "⏸️",
}

TYPE_LABELS = {
    "مستشفيات وزارة الصحة": "🏥 مستشفى",
    "فرعي":                 "🩺 مركز فرعي",
    "شامل":                 "🏥 مركز شامل",
    "أولي":                 "🩺 مركز أولي",
    "عيادة":                "🏢 عيادة صحية",
}

# ══════════════════════════════════════════════════════════════════════
# ── إعداد الصفحة والـ CSS
# ══════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Aumet | لوحة المتابعة التنفيذية الموحدة (Live DB)",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght=400;500;700;800&display=swap');

  html, body, [class*="css"], .stMarkdown p, p, span, label, .stText {{
    font-family: 'Tajawal', sans-serif !important;
    color: #1A202C !important;
  }}

  [data-testid="stSidebar"] {{ background-color: {AUMET_DARK} !important; border-right: 3px solid {AUMET_TEAL}; }}
  [data-testid="stSidebar"] .stMarkdown p, [data-testid="stSidebar"] label, [data-testid="stSidebar"] span {{
    color: #FFFFFF !important;
    font-weight: 600;
  }}

  .aumet-header {{
    background: linear-gradient(135deg, {AUMET_DARK} 0%, {AUMET_NAVY} 60%, {AUMET_TEAL} 150%);
    border-radius: 14px; padding: 22px 28px; margin-bottom: 20px; text-align: center; direction: rtl;
  }}
  .aumet-header h1 {{ color: #FFFFFF !important; font-size: 26px; font-weight: 800; margin: 0; }}
  .aumet-header .sub {{ color: #E6FFFA !important; font-size: 13px; margin-top: 6px; }}

  .kpi-card {{ background: #FFFFFF !important; border-radius: 12px; padding: 20px; border-top: 5px solid {AUMET_TEAL}; box-shadow: 0 4px 15px rgba(0,0,0,.06); text-align: center; direction: rtl; margin-bottom: 15px; }}
  .kpi-title {{ font-size: 15px; font-weight: 700; color: #4A5568 !important; margin-bottom: 8px; }}
  .kpi-num {{ font-size: 34px; font-weight: 800; color: {AUMET_DARK} !important; }}

  .phase-card {{ background: #FFFFFF !important; border-radius: 12px; padding: 16px; border: 1px solid #E2E8F0; direction: rtl; margin-bottom: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.02); }}
  .phase-title {{ font-weight: 700; font-size: 14px; color: #2D3748 !important; margin-bottom: 8px; }}
  .progress-track {{ background: #EDF2F7 !important; border-radius: 8px; height: 12px; overflow: hidden; margin-bottom: 6px; }}
  .progress-fill {{ height: 100%; border-radius: 8px; }}
  .progress-text {{ font-size: 12.5px; font-weight: 700; color: #2D3748 !important; text-align: left; }}

  .section-title {{ font-size: 17px; font-weight: 700; color: {AUMET_DARK} !important; border-right: 4px solid {AUMET_TEAL}; padding-right: 10px; margin: 20px 0 12px; direction: rtl; }}
  .status-badge {{ display: inline-block; border-radius: 8px; padding: 4px 12px; font-size: 12px; font-weight: 700; text-align: center; }}
  .lock-badge {{ display:inline-block; border-radius:8px; padding:2px 10px; font-size:11px; font-weight:700; background:#FEF3C7; color:#92400E; border:1px solid #F59E0B50; }}
  .db-pill {{ display:inline-block; background:rgba(255,255,255,.12); color:#fff !important; font-size:11px; border-radius:20px; padding:4px 12px; margin-top:8px; }}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════
# ── طبقة قاعدة البيانات (SQLite Live Engine) — حماية كاملة من الأعطال والتعارض
# ══════════════════════════════════════════════════════════════════════

@contextmanager
def get_connection():
    """اتصال آمن بقاعدة البيانات مع حماية من الـ Database Lock أثناء العمل المشترك."""
    conn = sqlite3.connect(DB_PATH, timeout=30, check_same_thread=False)
    try:
        conn.execute("PRAGMA journal_mode=WAL;")
        conn.execute("PRAGMA busy_timeout=30000;")
        conn.row_factory = sqlite3.Row
        yield conn
    finally:
        conn.close()


def init_db():
    """إنشاء جدولي medical_centers و activity_logs إذا لم يكونا موجودين."""
    with get_connection() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS medical_centers (
                uid                 TEXT PRIMARY KEY,
                source              TEXT,
                name                TEXT,
                type_raw            TEXT,
                type_display        TEXT,
                governorate         TEXT,
                liwa                TEXT,
                qadaa               TEXT,
                municipality        TEXT,
                lat                 REAL,
                lon                 REAL,
                status_medicines    TEXT,
                percent_medicines   INTEGER DEFAULT 0,
                notes_medicines     TEXT,
                status_consumables  TEXT,
                percent_consumables INTEGER DEFAULT 0,
                notes_consumables   TEXT,
                status_vaccines     TEXT,
                percent_vaccines    INTEGER DEFAULT 0,
                notes_vaccines      TEXT,
                is_manually_locked  INTEGER DEFAULT 0,
                last_updated        TEXT
            );
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS activity_logs (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp     TEXT,
                actor         TEXT,
                action_type   TEXT,
                facility_name TEXT,
                details       TEXT
            );
        """)
        conn.commit()


def log_activity(actor: str, action_type: str, facility_name: str, details: str):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        with get_connection() as conn:
            conn.execute(
                "INSERT INTO activity_logs (timestamp, actor, action_type, facility_name, details) VALUES (?,?,?,?,?)",
                (ts, actor or "غير معروف", action_type, facility_name, details),
            )
            conn.commit()
    except sqlite3.OperationalError:
        st.toast("⚠️ تعذر تسجيل العملية في سجل التدقيق (قاعدة البيانات مشغولة)", icon="⚠️")


def db_row_count() -> int:
    try:
        with get_connection() as conn:
            return conn.execute("SELECT COUNT(*) FROM medical_centers").fetchone()[0]
    except sqlite3.OperationalError:
        return 0


# ── دوال التطهير والمطابقة الذكية فائقة الأداء (Super Fuzzy Match) ────────────

def clean_and_tokenize(s):
    if pd.isna(s):
        return set()
    s = str(s).lower().strip()
    s = re.split(r'[/\\_-]', s)[0].strip()
    s = s.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا").replace("ة", "ه").replace("ى", "ي")
    ignore_words = {
        "مركز", "صحي", "مستشفى", "عيادة", "الصحى", "الصحي", "م", "الاولى", "الاول",
        "الشامل", "الفرعي", "حكومي", "رئيسي", "الرئيسي", "وزارة", "الصحة", "مستودع", "مستودعات",
    }
    tokens = set(re.findall(r'\w+', s))
    return tokens - ignore_words


def super_fuzzy_match(name_source, target_set) -> bool:
    """خوارزمية مطابقة الكلمات المتقدمة لضمان ربط الأسماء مهما اختلف ترتيبها أو حشوها."""
    tokens_source = clean_and_tokenize(name_source)
    if not tokens_source:
        return False
    for target_name in target_set:
        tokens_target = clean_and_tokenize(target_name)
        if not tokens_target:
            continue
        intersection = tokens_source.intersection(tokens_target)
        if not intersection:
            continue
        min_len = min(len(tokens_source), len(tokens_target))
        if len(intersection) >= min_len and min_len >= 1:
            return True
        if len(intersection) >= 2:
            return True
    return False


def parse_coords(coord_str):
    if pd.isna(coord_str) or str(coord_str).strip() == "":
        return None, None
    try:
        clean = re.sub(r"[^\d\.,\-]", "", str(coord_str))
        parts = [p.strip() for p in clean.split(",") if p.strip()]
        if len(parts) >= 2:
            lat, lon = float(parts[0]), float(parts[1])
            if 29.0 <= lat <= 34.0 and 34.0 <= lon <= 40.0:
                return lat, lon
    except (ValueError, IndexError):
        pass
    return None, None


# ══════════════════════════════════════════════════════════════════════
# ── قراءة ملفات الوزارة الأصلية وتغذية القاعدة (Seeding)
# ══════════════════════════════════════════════════════════════════════

def read_ministry_excel(path: str):
    """يقرأ ملف وزارة خام، يطهّر الإحداثيات والقيم الفارغة، ويعيد DataFrame موحّد."""
    df = pd.read_excel(path)
    n_empty = int(df.isna().sum().sum())

    coord_col = None
    for c in ["الاحداثيات", "الإحداثيات"]:
        if c in df.columns:
            coord_col = c
            break
    if coord_col:
        df[["lat", "lon"]] = df[coord_col].apply(lambda x: pd.Series(parse_coords(x)))
    else:
        df["lat"], df["lon"] = None, None
    df = df.dropna(subset=["lat", "lon"]).reset_index(drop=True)

    type_col = "نوع النقطة" if "نوع النقطة" in df.columns else ("التصنيف" if "التصنيف" in df.columns else None)
    df["type_raw"] = df[type_col].astype(str) if type_col else "غير محدد"
    df["type_display"] = df["type_raw"].map(TYPE_LABELS).fillna(df["type_raw"]) if type_col else "🩺 مركز صحي عام"

    name_col = "الاسم" if "الاسم" in df.columns else df.columns[0]
    df["name"] = df[name_col].fillna("منشأة غير مسماة").astype(str)

    for c in ["المحافظة", "اللواء", "القضاء", "البلدية"]:
        df[c] = df[c].fillna("غير محدد").astype(str) if c in df.columns else "غير محدد"

    return df, n_empty


def seed_database_from_excel(actor: str = "النظام") -> int:
    """يمسح جدول المراكز ويعيد تغذيته من ملفات الوزارة الثلاثة المعتمدة."""
    rows = []
    total_empty = 0
    missing_files = []

    for source_key in ["hospitals", "subcenters", "old_centers"]:
        path = FILES[source_key]
        if not os.path.exists(path):
            missing_files.append(path)
            continue
        try:
            df, n_empty = read_ministry_excel(path)
            total_empty += n_empty
        except Exception as e:
            st.toast(f"⚠️ تعذرت قراءة الملف {path}: {e}", icon="⚠️")
            continue

        now_ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for i, r in df.iterrows():
            uid = f"{source_key}_{i}"
            rows.append((
                uid, source_key, r["name"], r["type_raw"], r["type_display"],
                r["المحافظة"], r["اللواء"], r["القضاء"], r["البلدية"],
                float(r["lat"]), float(r["lon"]),
                STATUS_OPTIONS[0], 0, "",
                STATUS_OPTIONS[0], 0, "",
                STATUS_OPTIONS[0], 0, "",
                0, now_ts,
            ))

    if missing_files:
        st.toast(f"⚠️ ملفات غير موجودة وتم تجاوزها: {', '.join(missing_files)}", icon="⚠️")
    if total_empty:
        st.toast(f"⚠️ تم تجاوز {total_empty} خلية فارغة تلقائياً أثناء التغذية", icon="⚠️")

    if not rows:
        return 0

    try:
        with get_connection() as conn:
            conn.execute("DELETE FROM medical_centers")
            conn.executemany("""
                INSERT INTO medical_centers (
                    uid, source, name, type_raw, type_display, governorate, liwa, qadaa, municipality,
                    lat, lon, status_medicines, percent_medicines, notes_medicines,
                    status_consumables, percent_consumables, notes_consumables,
                    status_vaccines, percent_vaccines, notes_vaccines, is_manually_locked, last_updated
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, rows)
            conn.commit()
    except sqlite3.OperationalError as e:
        st.error(f"⚠️ تعذرت الكتابة إلى قاعدة البيانات (مشغولة حالياً): {e}")
        return 0

    log_activity(actor, "تهيئة قاعدة البيانات", "كافة المنشآت",
                 f"تمت تغذية القاعدة بـ {len(rows)} منشأة من ملفات الوزارة الثلاثة.")
    return len(rows)


# ══════════════════════════════════════════════════════════════════════
# ── العكس والمزامنة الحية لملف التفعيل (Activation Live Sync)
# ══════════════════════════════════════════════════════════════════════

def find_activation_file():
    for key in ["activation_live", "activation_fallback"]:
        path = FILES.get(key, "")
        if path and os.path.exists(path):
            return path, (key == "activation_fallback")
    return None, False


def parse_activation_columns(df_raw: pd.DataFrame):
    """بحث ديناميكي عن أعمدة Scope / Entity Name / system، حتى لو لم تكن في صف الهيدر الأول."""
    scope_col, name_col, sys_col = None, None, None

    def scan(df):
        s, n, y = None, None, None
        for col in df.columns:
            col_str = str(col).strip().lower()
            if 'scope' in col_str: s = col
            if 'entity name' in col_str: n = col
            if 'system' in col_str: y = col
        return s, n, y

    scope_col, name_col, sys_col = scan(df_raw)

    if not scope_col or not name_col:
        for idx, row in df_raw.head(10).iterrows():
            row_vals = [str(v).strip().lower() for v in row.values]
            if any('scope' in v for v in row_vals):
                df_raw.columns = df_raw.iloc[idx]
                df_raw = df_raw.iloc[idx + 1:].reset_index(drop=True)
                break
        scope_col, name_col, sys_col = scan(df_raw)

    return df_raw, scope_col, name_col, sys_col


def sync_activation_file_live(actor: str = "Admin"):
    """يقرأ ملف التفعيل الحي، يطبّق super_fuzzy_match، ويحدّث القاعدة مباشرة (SQL UPDATE)
    متجاوزاً أي مركز محمي بقفل التعديل اليدوي."""
    path, used_fallback = find_activation_file()
    if not path:
        return False, "⚠️ لم يتم العثور على ملف التفعيل (Health Care Center Activation_2.xlsx)", {}

    try:
        df_raw = pd.read_excel(path)
    except Exception as e:
        return False, f"⚠️ تعذرت قراءة ملف التفعيل: {e}", {}

    n_empty = int(df_raw.isna().sum().sum())
    df_raw.dropna(how='all', inplace=True)
    df_raw, scope_col, name_col, sys_col = parse_activation_columns(df_raw)

    if not scope_col or not name_col:
        return False, "⚠️ لم يتم العثور على أعمدة Scope / Entity Name داخل الملف", {}

    if n_empty:
        st.toast(f"⚠️ تم تجاوز {n_empty} خلية فارغة تلقائياً في ملف التفعيل", icon="⚠️")

    activated_drugs, activated_consumables = set(), set()
    for _, row in df_raw.iterrows():
        scope = str(row.get(scope_col, '')).strip().lower()
        n1 = str(row.get(name_col, '')).strip()
        n2 = str(row.get(sys_col, '')).strip() if sys_col else ""
        if 'drug' in scope:
            if n1 and n1.lower() != 'nan': activated_drugs.add(n1)
            if n2 and n2.lower() != 'nan': activated_drugs.add(n2)
        if 'consumable' in scope or 'mats' in scope:
            if n1 and n1.lower() != 'nan': activated_consumables.add(n1)
            if n2 and n2.lower() != 'nan': activated_consumables.add(n2)

    if not activated_drugs and not activated_consumables:
        return False, "⚠️ لم يتم العثور على أي قيم صالحة للمطابقة (Drug/Consumable) في ملف التفعيل", {}

    try:
        with get_connection() as conn:
            centers = pd.read_sql_query(
                "SELECT uid, name, is_manually_locked, status_medicines, status_consumables FROM medical_centers",
                conn,
            )
    except sqlite3.OperationalError as e:
        return False, f"⚠️ قاعدة البيانات مشغولة حالياً، حاول مرة أخرى ({e})", {}

    meds_updates, cons_updates = [], []
    skipped_locked = 0
    now_ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for _, r in centers.iterrows():
        hit_drug = super_fuzzy_match(r["name"], activated_drugs)
        hit_cons = super_fuzzy_match(r["name"], activated_consumables)
        if not hit_drug and not hit_cons:
            continue
        if int(r["is_manually_locked"]) == 1:
            skipped_locked += 1
            continue
        if hit_drug and r["status_medicines"] != STATUS_OPTIONS[2]:
            meds_updates.append((STATUS_OPTIONS[2], 100, now_ts, r["uid"]))
        if hit_cons and r["status_consumables"] != STATUS_OPTIONS[2]:
            cons_updates.append((STATUS_OPTIONS[2], 100, now_ts, r["uid"]))

    try:
        with get_connection() as conn:
            cur = conn.cursor()
            if meds_updates:
                cur.executemany(
                    "UPDATE medical_centers SET status_medicines=?, percent_medicines=?, last_updated=? WHERE uid=?",
                    meds_updates,
                )
            if cons_updates:
                cur.executemany(
                    "UPDATE medical_centers SET status_consumables=?, percent_consumables=?, last_updated=? WHERE uid=?",
                    cons_updates,
                )
            conn.commit()
    except sqlite3.OperationalError as e:
        return False, f"⚠️ تعذر تحديث قاعدة البيانات (مشغولة حالياً): {e}", {}

    touched_uids = set([u[3] for u in meds_updates] + [u[3] for u in cons_updates])
    fallback_note = " (تم استخدام الملف الاحتياطي لعدم توفر الملف الحي)" if used_fallback else ""
    details = (f"تحديث {len(meds_updates)} مركز لمرحلة الأدوية، {len(cons_updates)} مركز لمرحلة المستهلكات، "
               f"تجاوز {skipped_locked} مركز محمي بقفل يدوي.{fallback_note}")
    log_activity(actor, "♻️ مزامنة جماعية لايف (Activation Sync)", "كافة المنشآت", details)

    return True, "✅ تمت المزامنة والعكس بنجاح", {
        "meds": len(meds_updates), "cons": len(cons_updates),
        "skipped": skipped_locked, "total": len(touched_uids), "fallback": used_fallback,
    }


# ══════════════════════════════════════════════════════════════════════
# ── إعادة ضبط النظام بالكامل
# ══════════════════════════════════════════════════════════════════════

def reset_system_completely(actor: str = "Admin"):
    try:
        with get_connection() as conn:
            conn.execute("DELETE FROM medical_centers")
            conn.execute("DELETE FROM activity_logs")
            conn.commit()
    except sqlite3.OperationalError as e:
        return False, f"⚠️ تعذرت إعادة الضبط (قاعدة البيانات مشغولة): {e}"

    n = seed_database_from_excel(actor=actor)
    log_activity(actor, "🗑️ إعادة ضبط النظام بالكامل", "كافة المنشآت",
                 f"تم مسح القاعدة وإعادة تغذيتها بـ {n} منشأة من ملفات الإكسل الأصلية.")
    return True, f"✅ تمت إعادة ضبط النظام بنجاح ({n} منشأة)"


# ══════════════════════════════════════════════════════════════════════
# ── دوال الاستعلام المباشر من القاعدة (Fast Live Queries)
# ══════════════════════════════════════════════════════════════════════

def get_centers_df(search_q: str = "", types=None, govs=None) -> pd.DataFrame:
    query = "SELECT * FROM medical_centers WHERE 1=1"
    params = []
    if search_q:
        query += " AND name LIKE ?"
        params.append(f"%{search_q.strip()}%")
    if types:
        query += f" AND type_display IN ({','.join(['?'] * len(types))})"
        params.extend(types)
    if govs:
        query += f" AND governorate IN ({','.join(['?'] * len(govs))})"
        params.extend(govs)
    try:
        with get_connection() as conn:
            return pd.read_sql_query(query, conn, params=params)
    except sqlite3.OperationalError as e:
        st.error(f"⚠️ تعذر الاستعلام من القاعدة (مشغولة حالياً): {e}")
        return pd.DataFrame()


def get_distinct(column: str):
    try:
        with get_connection() as conn:
            rows = conn.execute(f"SELECT DISTINCT {column} FROM medical_centers ORDER BY {column}").fetchall()
        return [r[0] for r in rows if r[0]]
    except sqlite3.OperationalError:
        return []


def get_recent_logs(limit: int = 20) -> pd.DataFrame:
    try:
        with get_connection() as conn:
            return pd.read_sql_query(
                "SELECT timestamp, actor, action_type, facility_name, details FROM activity_logs ORDER BY id DESC LIMIT ?",
                conn, params=(limit,),
            )
    except sqlite3.OperationalError:
        return pd.DataFrame()


def overall_status(row) -> str:
    statuses = [row[status_col(k)] for k in PHASE_KEYS]
    if all(s == STATUS_OPTIONS[2] for s in statuses):
        return STATUS_OPTIONS[2]
    if any(s == STATUS_OPTIONS[3] for s in statuses):
        return STATUS_OPTIONS[3]
    if any(s == STATUS_OPTIONS[1] for s in statuses):
        return STATUS_OPTIONS[1]
    return STATUS_OPTIONS[0]


# ══════════════════════════════════════════════════════════════════════
# ── دوال الواجهة المساعدة (HTML/Map/Report)
# ══════════════════════════════════════════════════════════════════════

def status_badge_html(status: str) -> str:
    color = STATUS_COLORS.get(status, STATUS_COLORS[STATUS_OPTIONS[0]])
    icon = STATUS_ICONS.get(status, "⚪")
    return f'<span class="status-badge" style="background:{color}18;color:{color};border:1px solid {color}50;">{icon} {status.split(" / ")[0]}</span>'


def progress_bar_html(label: str, pct: float, color: str) -> str:
    return f"""
    <div class="phase-card">
        <div class="phase-title">{label}</div>
        <div class="progress-track">
            <div class="progress-fill" style="width:{pct}%;background:{color};"></div>
        </div>
        <div class="progress-text"><span>{pct:.1f}% مكتمل (Live من القاعدة)</span></div>
    </div>
    """


def health_score_html(score: float) -> str:
    if score >= 80: color = STATUS_COLORS[STATUS_OPTIONS[2]]
    elif score >= 50: color = STATUS_COLORS[STATUS_OPTIONS[1]]
    else: color = STATUS_COLORS[STATUS_OPTIONS[3]]
    return f"""
    <div class="kpi-card" style="border-top-color:{color};">
        <div class="kpi-title">🎯 مؤشر الجاهزية الكلي للمشروع (Health Score)</div>
        <div class="kpi-num" style="color:{color} !important;">{score:.1f}%</div>
    </div>
    """


def build_map(df: pd.DataFrame, phase_view: str) -> folium.Map:
    m = folium.Map(location=[31.95, 35.91], zoom_start=8, tiles="CartoDB positron", prefer_canvas=True)
    type_groups = {}
    for disp in df["type_display"].dropna().unique():
        type_groups[disp] = (
            folium.FeatureGroup(name=str(disp), show=True),
            MarkerCluster(options={"disableClusteringAtZoom": 13, "maxClusterRadius": 50}),
        )
    for _, row in df.iterrows():
        status = overall_status(row) if phase_view == "overall" else row.get(status_col(phase_view), STATUS_OPTIONS[0])
        clr = STATUS_COLORS.get(status, STATUS_COLORS[STATUS_OPTIONS[0]])
        disp = str(row.get("type_display", "")).strip()
        type_raw = str(row.get("type_raw", ""))
        is_hospital = any(k in type_raw for k in ["مستشفى", "مستشفيات", "شامل"]) or any(k in disp for k in ["🏥"])
        emoji = "🏥" if is_hospital else "🩺"
        lock_icon = " 🔒" if int(row.get("is_manually_locked", 0)) == 1 else ""
        icon_html = f"""<div style="background:{clr};width:18px;height:18px;border-radius:50%;border:2px solid white;box-shadow:0 2px 6px rgba(0,0,0,.4);"></div>"""
        badges = "".join(
            f'<div style="margin:6px 0;display:flex;justify-content:space-between;align-items:center;">'
            f'<span style="color:#2D3748;font-weight:600;">{p["icon"]} {p["ar"]}</span>'
            f'{status_badge_html(row.get(status_col(p["key"]), STATUS_OPTIONS[0]))}</div>'
            for p in PHASES
        )
        popup_html = (
            f'<div dir="rtl" style="font-family:\'Tajawal\',Arial;width:280px;font-size:13px;color:#1A202C;">'
            f'<div style="background:linear-gradient(135deg,{AUMET_DARK},{AUMET_NAVY});color:white;padding:12px;'
            f'font-weight:700;border-radius:6px 6px 0 0;">{emoji} {row.get("name","")}{lock_icon}</div>'
            f'<div style="padding:12px;background:#FFFFFF;border:1px solid #E2E8F0;border-radius:0 0 6px 6px;">{badges}</div></div>'
        )
        if disp in type_groups:
            folium.Marker(
                location=[row["lat"], row["lon"]],
                popup=folium.Popup(popup_html, max_width=300),
                tooltip=f"{emoji} {row.get('name','')}",
                icon=folium.DivIcon(html=icon_html, icon_size=(20, 20)),
            ).add_to(type_groups[disp][1])
    for fg, cluster in type_groups.values():
        cluster.add_to(fg)
        fg.add_to(m)
    folium.LayerControl(position="topright", collapsed=False).add_to(m)
    return m


REPORT_LABELS = {
    "name": "اسم المنشأة", "type_display": "النوع", "governorate": "المحافظة", "liwa": "اللواء",
    "status_medicines": "حالة الأدوية", "status_consumables": "حالة المستهلكات", "status_vaccines": "حالة المطاعيم",
    "percent_medicines": "% الأدوية", "percent_consumables": "% المستهلكات", "percent_vaccines": "% المطاعيم",
    "is_manually_locked": "مقفل يدوياً", "last_updated": "آخر تحديث",
}


def generate_excel_report(df: pd.DataFrame) -> bytes:
    cols = [c for c in REPORT_LABELS if c in df.columns]
    export_df = df[cols].rename(columns=REPORT_LABELS)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        export_df.to_excel(writer, sheet_name="تقرير المملكة الموحد", index=False)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════
# ── تهيئة النظام عند أول تشغيل (Auto Init + Seed)
# ══════════════════════════════════════════════════════════════════════

init_db()
if db_row_count() == 0:
    with st.spinner("🔄 جاري تهيئة قاعدة البيانات لأول مرة وتغذيتها من ملفات الوزارة..."):
        n_seeded = seed_database_from_excel()
    if n_seeded:
        st.toast(f"✅ تم إنشاء aumet_system.db وتغذيتها بـ {n_seeded} منشأة", icon="✅")
    else:
        st.warning("⚠️ لم يتم العثور على ملفات الوزارة الأساسية بنفس المجلد — تأكد من وجودها لتغذية القاعدة.")

if "role" not in st.session_state: st.session_state.role = "viewer"
if "admin_name" not in st.session_state: st.session_state.admin_name = "Admin"


# ══════════════════════════════════════════════════════════════════════
# ── القائمة الجانبية (SIDEBAR)
# ══════════════════════════════════════════════════════════════════════
with st.sidebar:
    if os.path.exists(FILES["logo"]):
        st.image(FILES["logo"], use_container_width=True)
    st.markdown(f'<div style="text-align:center;color:{AUMET_TEAL};font-size:20px;font-weight:800;margin-bottom:4px;">AUMET SMART MONITOR</div>', unsafe_allow_html=True)
    st.markdown(f'<div style="text-align:center;"><span class="db-pill">🗄️ SQLite Live — {db_row_count()} منشأة</span></div>', unsafe_allow_html=True)

    role_choice = st.radio("نوع الحساب", ["👁️ مشاهد (Viewer)", "🔑 مسؤول (Admin)"])
    if "Admin" in role_choice:
        pwd = st.text_input("كلمة مرور المسؤول", type="password")
        if pwd == ADMIN_PASSWORD:
            st.session_state.role = "admin"
            st.session_state.admin_name = st.text_input("اسمك في السجل", value=st.session_state.admin_name)
        else:
            st.session_state.role = "viewer"
    else:
        st.session_state.role = "viewer"

    st.markdown("<hr style='border-color:rgba(255,255,255,0.1);'>", unsafe_allow_html=True)
    search_q = st.text_input("🔍 بحث باسم المنشأة")
    sel_types = st.multiselect("📂 نوع النقطة", get_distinct("type_display"))
    sel_govs = st.multiselect("📍 المحافظة", get_distinct("governorate"))

    st.markdown("<hr style='border-color:rgba(255,255,255,0.1);'>", unsafe_allow_html=True)

    if st.session_state.role == "admin":
        st.markdown("**⚙️ أدوات المسؤول**")
        if st.button("♻️ مزامنة وعكس ملف التفعيل لايف", use_container_width=True):
            with st.spinner("🔄 جاري قراءة ملف التفعيل والمزامنة المباشرة مع القاعدة..."):
                ok, msg, info = sync_activation_file_live(actor=st.session_state.admin_name)
            if ok:
                st.success(msg)
                st.toast(
                    f"تم تحديث {info['meds']} (أدوية) + {info['cons']} (مستهلكات) = {info['total']} مركز · "
                    f"تجاوز {info['skipped']} مركز مقفل يدوياً",
                    icon="✅",
                )
                st.rerun()
            else:
                st.error(msg)

        with st.expander("🗑️ إعادة ضبط النظام بالكامل (خطر)"):
            st.caption("سيتم مسح كل بيانات القاعدة وإعادة تغذيتها من ملفات الإكسل الأصلية. لا يمكن التراجع عن هذا الإجراء.")
            confirm_reset = st.checkbox("أؤكد أنني أريد مسح كل البيانات الحالية بالكامل")
            if st.button("🗑️ تنفيذ إعادة الضبط الآن", disabled=not confirm_reset, use_container_width=True):
                with st.spinner("🔄 جاري إعادة ضبط النظام..."):
                    ok, msg = reset_system_completely(actor=st.session_state.admin_name)
                (st.success if ok else st.error)(msg)
                st.rerun()
    else:
        st.info("🔐 سجّل دخولك كمسؤول لتفعيل أدوات المزامنة والتعديل.")


# ══════════════════════════════════════════════════════════════════════
# ── جلب البيانات المفلترة مباشرة من القاعدة (Live Query)
# ══════════════════════════════════════════════════════════════════════
filtered = get_centers_df(search_q, sel_types, sel_govs)

# ── هيدر لوحة التحكم
st.markdown(
    f'<div class="aumet-header"><h1>🏥 لوحة تحكم ومتابعة تنفيذ مشاريع Aumet الموحدة للرعاية الصحية</h1>'
    f'<div class="sub">مدعومة بقاعدة بيانات حية SQLite — متعددة المستخدمين وآمنة من التعارض</div></div>',
    unsafe_allow_html=True,
)

if filtered.empty:
    st.warning("⚠️ لا توجد بيانات مطابقة للفلاتر الحالية، أو قاعدة البيانات لا تزال فارغة.")

# ── كروت الـ KPI + مؤشر الجاهزية الكلي
k1, k2, k3, k4 = st.columns(4)
with k1:
    st.markdown(f'<div class="kpi-card"><div class="kpi-title">📍 إجمالي المنشآت الطبية</div><div class="kpi-num">{len(filtered)}</div></div>', unsafe_allow_html=True)
with k2:
    live_count = int((filtered.apply(overall_status, axis=1) == STATUS_OPTIONS[2]).sum()) if len(filtered) else 0
    st.markdown(f'<div class="kpi-card"><div class="kpi-title">✅ جاهز ومفعّل بالكامل (Live)</div><div class="kpi-num" style="color:{STATUS_COLORS[STATUS_OPTIONS[2]]} !important;">{live_count}</div></div>', unsafe_allow_html=True)
with k3:
    prog_count = int((filtered.apply(overall_status, axis=1) == STATUS_OPTIONS[1]).sum()) if len(filtered) else 0
    st.markdown(f'<div class="kpi-card"><div class="kpi-title">🔄 قيد التنفيذ والربط حالياً</div><div class="kpi-num" style="color:{STATUS_COLORS[STATUS_OPTIONS[1]]} !important;">{prog_count}</div></div>', unsafe_allow_html=True)
with k4:
    if len(filtered):
        health_score = round(filtered[[percent_col(p["key"]) for p in PHASES]].mean().mean(), 1)
    else:
        health_score = 0.0
    st.markdown(health_score_html(health_score), unsafe_allow_html=True)

tabs = st.tabs(["🗺️ الخريطة التفاعلية والتحليلات الإدارية", "✏️ لوحة المسؤول والتعديل المباشر", "📋 التقارير وسجل العمليات"])

# ── التبويب 1: الخريطة والتحليلات ────────────────────────────────────────
with tabs[0]:
    col_map, col_kpi = st.columns([3, 2])
    with col_map:
        st.markdown('<div class="section-title">🗺️ التوزيع الجغرافي ونقاط الحالة المباشرة للمشروع</div>', unsafe_allow_html=True)
        phase_view = st.selectbox("عرض حالة الخريطة حسب المرحلة:", ["overall"] + PHASE_KEYS)
        if len(filtered):
            st_folium(build_map(filtered, phase_view), width=None, height=480, returned_objects=[], key="main_map")
    with col_kpi:
        st.markdown('<div class="section-title">📊 نسب الإنجاز العام لكل مرحلة تشغيلية</div>', unsafe_allow_html=True)
        for p in PHASES:
            live_pct = round((filtered[status_col(p["key"])] == STATUS_OPTIONS[2]).sum() / len(filtered) * 100, 1) if len(filtered) else 0
            st.markdown(progress_bar_html(f"{p['icon']} {p['full']}", live_pct, AUMET_TEAL), unsafe_allow_html=True)

        st.markdown('<div class="section-title">🏛️ المنشآت المفعلة (Live) حسب المحافظة</div>', unsafe_allow_html=True)
        if len(filtered) > 0:
            tmp = filtered.copy()
            tmp["is_live"] = tmp.apply(overall_status, axis=1) == STATUS_OPTIONS[2]
            gov_stats = tmp.groupby("governorate")["is_live"].sum().reset_index().sort_values(by="is_live", ascending=True)
            st.bar_chart(data=gov_stats.set_index("governorate"), y="is_live", color=AUMET_TEAL, use_container_width=True)

# ── التبويب 2: لوحة المسؤول والتعديل المباشر (مع حماية القفل اليدوي) ────────
with tabs[1]:
    if st.session_state.role == "admin":
        st.markdown('<div class="section-title">✏️ جدول البيانات التفاعلي للتعديل الجماعي الفوري والتحديث المباشر في القاعدة</div>', unsafe_allow_html=True)
        st.caption("🔒 أي تعديل يدوي على الحالة يقفل المركز تلقائياً ويحميه من المزامنة التلقائية اللاحقة. فك القفل عبر تبديل خانة 'مقفل يدوياً' وحفظ التغييرات.")

        if len(filtered):
            editor_cols = ["name", "type_display", "governorate"] + [status_col(p["key"]) for p in PHASES] + ["is_manually_locked"]
            base_df = filtered.set_index("uid")[editor_cols].copy()
            base_df["is_manually_locked"] = base_df["is_manually_locked"].astype(bool)

            col_config = {
                "name": st.column_config.TextColumn("اسم المنشأة", disabled=True),
                "type_display": st.column_config.TextColumn("النوع", disabled=True),
                "governorate": st.column_config.TextColumn("المحافظة", disabled=True),
                "is_manually_locked": st.column_config.CheckboxColumn("🔒 مقفل يدوياً"),
            }
            for p in PHASES:
                col_config[status_col(p["key"])] = st.column_config.SelectboxColumn(
                    f"{p['icon']} {p['ar']}", options=STATUS_OPTIONS, required=True
                )

            edited = st.data_editor(
                base_df, column_config=col_config, use_container_width=True, height=420, key="main_editor"
            )

            if st.button("💾 حفظ وتثبيت كافة التغييرات اليدوية", use_container_width=True):
                try:
                    with get_connection() as conn:
                        current = pd.read_sql_query("SELECT * FROM medical_centers", conn).set_index("uid")
                except sqlite3.OperationalError as e:
                    st.error(f"⚠️ تعذر قراءة القاعدة قبل الحفظ: {e}")
                    current = pd.DataFrame()

                changes = 0
                now_ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                try:
                    with get_connection() as conn:
                        cur = conn.cursor()
                        for uid in edited.index:
                            if uid not in current.index:
                                continue
                            orig = current.loc[uid]
                            new = edited.loc[uid]

                            set_parts, params, log_lines = [], [], []
                            status_changed = False

                            for p in PHASES:
                                sc = status_col(p["key"])
                                pc = percent_col(p["key"])
                                if str(new[sc]) != str(orig[sc]):
                                    status_changed = True
                                    set_parts.append(f"{sc}=?"); params.append(new[sc])
                                    if new[sc] == STATUS_OPTIONS[2]:
                                        set_parts.append(f"{pc}=?"); params.append(100)
                                    elif new[sc] == STATUS_OPTIONS[0]:
                                        set_parts.append(f"{pc}=?"); params.append(0)
                                    log_lines.append(f"تعديل يدوي لحالة {p['ar']}: {new[sc]}")

                            new_lock = bool(new["is_manually_locked"])
                            orig_lock = bool(int(orig["is_manually_locked"]))
                            if status_changed:
                                new_lock = True  # حماية تلقائية عند أي تعديل يدوي على الحالة
                            lock_changed = new_lock != orig_lock

                            if lock_changed or status_changed:
                                set_parts.append("is_manually_locked=?"); params.append(int(new_lock))
                                if lock_changed and not status_changed:
                                    log_lines.append("🔒 تم قفل المركز يدوياً" if new_lock else "🔓 تم فك القفل اليدوي")

                            if set_parts:
                                set_parts.append("last_updated=?"); params.append(now_ts)
                                params.append(uid)
                                cur.execute(f"UPDATE medical_centers SET {', '.join(set_parts)} WHERE uid=?", params)
                                log_activity(st.session_state.admin_name, "تعديل يدوي", orig["name"], "؛ ".join(log_lines))
                                changes += 1
                        conn.commit()
                except sqlite3.OperationalError as e:
                    st.error(f"⚠️ قاعدة البيانات مشغولة حالياً، حاول الحفظ مرة أخرى بعد لحظات ({e})")
                    changes = 0

                if changes:
                    st.success(f"✅ تم حفظ {changes} تعديل بنجاح ومباشرة في القاعدة!")
                    st.rerun()
                else:
                    st.info("ℹ️ لم يتم رصد أي تغييرات لحفظها.")
        else:
            st.info("لا توجد بيانات مطابقة للفلاتر الحالية لعرضها في جدول التعديل.")
    else:
        st.warning("🔐 يرجى الانتقال إلى حساب مسؤول (Admin) من القائمة الجانبية لتتمكن من تعديل البيانات وحفظها.")

# ── التبويب 3: التقارير وسجل العمليات (Live من جدول الـ logs) ────────────────
with tabs[2]:
    st.markdown('<div class="section-title">📋 الجدول الشامل للبيانات والتقارير التنفيذية للمملكة</div>', unsafe_allow_html=True)
    if len(filtered):
        show_cols = ["name", "type_display", "governorate", "liwa"] + [status_col(p["key"]) for p in PHASES] + ["is_manually_locked"]
        display_df = filtered[show_cols].rename(columns=REPORT_LABELS)
        st.dataframe(display_df, use_container_width=True)
        st.download_button(
            "⬇️ تحميل واستخراج التقرير الموحد الحالي بصيغة Excel المعتمدة",
            generate_excel_report(filtered),
            "Aumet_Executive_Report.xlsx",
            use_container_width=True,
        )
    else:
        st.info("لا توجد بيانات لعرضها أو تصديرها حالياً.")

    st.markdown('<div class="section-title">🕓 آخر 20 عملية تحديث تمت في النظام (Live من سجل التدقيق)</div>', unsafe_allow_html=True)
    logs_df = get_recent_logs(20)
    if not logs_df.empty:
        logs_df = logs_df.rename(columns={
            "timestamp": "التاريخ والوقت", "actor": "المستخدم", "action_type": "نوع الحركة",
            "facility_name": "المنشأة", "details": "تفاصيل العملية",
        })
        st.dataframe(logs_df, use_container_width=True, height=420)
    else:
        st.info("لا توجد عمليات مسجلة بعد في سجل التدقيق.")
