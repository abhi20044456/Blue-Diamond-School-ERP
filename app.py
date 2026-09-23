"""
BLUE DIAMOND PUBLIC SCHOOL - SCHOOL MANAGEMENT ERP
Hybrid PostgreSQL + Excel Backend
"""
import os
import json
import re
import calendar
import shutil
from datetime import datetime, date, timedelta
from functools import wraps
from pathlib import Path
from io import BytesIO

from flask import (
    Flask, request, jsonify, send_from_directory,
    session, send_file
)
from werkzeug.security import generate_password_hash, check_password_hash

from config import (
    DATABASE_URL, EXCEL_FILE, SECRET_KEY, SESSION_HOURS,
    MAX_LOGIN_ATTEMPTS, SCHOOL_NAME, SCHOOL_ADDRESS, DATA_DIR, BACKUP_DIR
)
from models import (
    db, User, Student, FeeTransaction, FeeStructure,
    Staff, Attendance, SalaryRecord, AuditLog, Setting, SyncQueue
)
from database import HybridDataManager, init_excel

# ── App Setup ─────────────────────────────────────────────
app = Flask(__name__, static_folder=None)
app.secret_key = SECRET_KEY
app.permanent_session_lifetime = timedelta(hours=SESSION_HOURS)
app.config['SQLALCHEMY_DATABASE_URI'] = DATABASE_URL
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SQLALCHEMY_ENGINE_OPTIONS'] = {
    'pool_size': 5,
    'max_overflow': 10,
    'pool_recycle': 280,       # Neon scale-to-zero
    'pool_pre_ping': True,
    'pool_timeout': 30,
    'connect_args': {
        'connect_timeout': 10,
        'application_name': 'blue_diamond_erp'
    }
}

db.init_app(app)
data_mgr = HybridDataManager(app, db)

# ── Permission System ─────────────────────────────────────
ALL_SECTIONS = [
    "dashboard", "students", "admission", "fee", "receipts", "pending",
    "structure", "reports", "users", "audit", "settings", "backup",
    "staff", "hiring", "attendance", "salary", "notifications"
]

ROLE_DEFAULTS = {
    "Super Admin": {s: "edit" for s in ALL_SECTIONS},
    "Admin": {s: "edit" for s in ALL_SECTIONS},
    "Accountant": {
        **{s: "none" for s in ALL_SECTIONS},
        "dashboard": "view", "students": "view", "admission": "edit",
        "fee": "edit", "receipts": "edit", "pending": "view",
        "reports": "view", "structure": "view", "salary": "view",
        "staff": "view", "notifications": "view"
    },
    "Data Entry": {
        **{s: "none" for s in ALL_SECTIONS},
        "dashboard": "view", "students": "edit", "admission": "edit",
        "fee": "edit", "receipts": "view", "pending": "view",
        "notifications": "view"
    },
    "Teacher": {
        **{s: "none" for s in ALL_SECTIONS},
        "dashboard": "view", "students": "view", "attendance": "edit",
        "notifications": "view"
    },
}
ROLES = list(ROLE_DEFAULTS.keys())

# ── Helpers ───────────────────────────────────────────────
def safe_float(v, d=0.0):
    try:
        if v is None or v == "": return d
        return float(v)
    except: return d

def safe_int(v, d=0):
    try:
        if v is None or v == "": return d
        return int(float(v))
    except: return d

def get_settings():
    """Return settings dict from PG or Excel."""
    rows = data_mgr.get_all(Setting)
    d = {}
    for r in rows:
        if isinstance(r, dict):
            d[str(r.get('Key', ''))] = r.get('Value', '')
        else:
            d[str(r.key)] = r.value
    return d

def set_setting(key, value):
    if data_mgr.pg_available:
        s = Setting.query.filter_by(key=key).first()
        if s:
            s.value = str(value)
        else:
            db.session.add(Setting(key=key, value=str(value)))
        db.session.commit()
    else:
        from database import append_to_excel
        append_to_excel("Settings", {"Key": key, "Value": value})

def add_audit(username, action, details):
    try:
        now = datetime.now()
        log = AuditLog(
            date=now.strftime('%Y-%m-%d'),
            time=now.strftime('%H:%M:%S'),
            username=username or 'system',
            action=action,
            details=details
        )
        data_mgr.save(log)
    except:
        pass

def get_perms(username):
    u = data_mgr.get_one(User, username=username)
    if not u: return {}
    try:
        p = json.loads(u.permissions or '{}')
    except:
        p = {}
    if not p:
        p = ROLE_DEFAULTS.get(u.role, {})
    return p

def perm_ok(perms, section, level="view"):
    order = {"none": 0, "view": 1, "edit": 2}
    return order.get(str(perms.get(section, "none")).lower(), 0) >= order.get(level, 1)

def login_required(fn):
    @wraps(fn)
    def wrapper(*a, **kw):
        if "username" not in session:
            return jsonify({"success": False, "message": "Login required."}), 401
        return fn(*a, **kw)
    return wrapper

def require_perm(section, level="view"):
    def deco(fn):
        @wraps(fn)
        def wrapper(*a, **kw):
            if "username" not in session:
                return jsonify({"success": False, "message": "Login required."}), 401
            perms = get_perms(session["username"])
            if not perm_ok(perms, section, level):
                return jsonify({"success": False,
                                "message": f"Permission denied: {section} ({level})"}), 403
            return fn(*a, **kw)
        return wrapper
    return deco

def session_months():
    """Return 12 month keys for academic session."""
    s = get_settings()
    sess = str(s.get('Academic Session', '2026-27'))
    m = re.findall(r'\d{4}', sess)
    try: sy = int(m[0])
    except: sy = datetime.now().year
    sm = max(1, min(12, safe_int(s.get('Session Start Month', 4), 4)))
    keys = []
    y, mo = sy, sm
    for _ in range(12):
        keys.append(f"{y:04d}-{mo:02d}")
        mo += 1
        if mo > 12:
            mo = 1; y += 1
    return keys

def elapsed_months():
    keys = session_months()
    cur = datetime.now().strftime('%Y-%m')
    return [k for k in keys if k <= cur]

def month_label(key):
    try:
        y, m = key.split('-')
        return datetime(int(y), int(m), 1).strftime('%B %Y')
    except:
        return key

def fee_structure_map():
    rows = data_mgr.get_all(FeeStructure)
    d = {}
    for r in rows:
        if isinstance(r, dict):
            d[str(r.get('Class', '')).strip()] = r
        else:
            d[str(r.class_name).strip()] = r.to_dict()
    return d

def build_ledger(student):
    """Build month-wise fee ledger for a student."""
    settings = get_settings()
    structure = fee_structure_map()
    txs = data_mgr.get_all(FeeTransaction)

    adm = str(student.get('Admission No', '')).strip()
    cls = str(student.get('Class', '')).strip()
    st = structure.get(cls, {})

    monthly = safe_float(st.get('Tuition Fee'))
    if str(student.get('Transport', 'No')).lower() in ('yes', 'y', 'true'):
        monthly += safe_float(st.get('Transport Fee'))

    onetime = sum(safe_float(st.get(k)) for k in
                  ['Admission Fee', 'Exam Fee', 'Annual Fee', 'Other Fee'])

    my_txs = [t for t in txs if str(t.get('Admission No', '')).strip().lower() == adm.lower()]
    paid_by_month = {}
    total_paid = 0.0
    total_late_paid = 0.0
    onetime_paid = 0.0

    for t in my_txs:
        amt = safe_float(t.get('Amount'))
        lf = safe_float(t.get('Late Fee'))
        ft = str(t.get('Fee Type', '')).strip()
        total_paid += amt
        total_late_paid += lf
        if ft in ('Admission Fee', 'Exam Fee', 'Annual Fee', 'Other Fee'):
            onetime_paid += amt
        mk = str(t.get('Month') or '').strip()
        if mk:
            paid_by_month[mk] = paid_by_month.get(mk, 0.0) + amt

    due_day = safe_int(settings.get('Due Day', 10), 10)
    grace = safe_int(settings.get('Grace Days', 0), 0)
    late_on = str(settings.get('Late Fee Enabled', 'Yes')).lower() in ('yes', 'true', '1')
    late_type = str(settings.get('Late Fee Type', 'Per Day'))
    late_amt = safe_float(settings.get('Late Fee Amount', 50))
    late_pct = safe_float(settings.get('Late Fee Percent', 2))
    max_late = safe_float(settings.get('Max Late Fee', 1000))

    today = date.today()
    months = elapsed_months()
    rows = []
    total_due = total_pending = total_late = overdue_amount = 0.0
    max_overdue = 0
    pending_months = 0

    for key in months:
        y, m = int(key[:4]), int(key[5:7])
        last_day = calendar.monthrange(y, m)[1]
        due_date = date(y, m, min(due_day, last_day))
        due_eff = due_date + timedelta(days=grace)
        paid = paid_by_month.get(key, 0.0)
        pending = max(monthly - paid, 0.0)

        days_over = 0
        if pending > 0 and today > due_eff:
            days_over = (today - due_eff).days

        late = 0.0
        if late_on and pending > 0 and days_over > 0:
            if late_type == 'Per Day':
                late = days_over * late_amt
            elif late_type == 'Fixed':
                late = late_amt
            elif late_type == 'Percent':
                late = pending * late_pct / 100
            if max_late > 0:
                late = min(late, max_late)

        if pending <= 0:
            status = 'Paid'
        elif days_over > 0:
            status = 'Overdue'
        else:
            status = 'Pending'

        rows.append({
            'month': key, 'label': month_label(key),
            'due': round(monthly, 2), 'paid': round(paid, 2),
            'pending': round(pending, 2),
            'due_date': due_date.strftime('%Y-%m-%d'),
            'days_overdue': days_over,
            'late_fee': round(late, 2),
            'status': status
        })
        total_due += monthly
        total_pending += pending
        total_late += late
        if days_over > 0:
            overdue_amount += pending
            max_overdue = max(max_overdue, days_over)
        if pending > 0:
            pending_months += 1

    onetime_pending = max(onetime - onetime_paid, 0.0)
    net_late = max(total_late - total_late_paid, 0.0)
    total_demand = total_due + onetime
    pending_total = total_pending + onetime_pending

    return {
        'admission_no': adm,
        'student_name': student.get('Student Name', ''),
        'class_name': cls,
        'section': student.get('Section', ''),
        'father_name': student.get('Father Name', ''),
        'mobile': student.get('Mobile', ''),
        'monthly_fee': round(monthly, 2),
        'months': rows,
        'total_due': round(total_demand, 2),
        'total_paid': round(total_paid, 2),
        'total_late_paid': round(total_late_paid, 2),
        'pending': round(pending_total, 2),
        'late_fee': round(net_late, 2),
        'total_payable': round(pending_total + net_late, 2),
        'overdue_amount': round(overdue_amount, 2),
        'overdue_days': max_overdue,
        'pending_months': pending_months,
    }

def build_all_ledgers():
    students = data_mgr.get_all(Student)
    out = {}
    for s in students:
        if str(s.get('Status', '')).strip().lower() != 'active':
            continue
        out[str(s.get('Admission No', '')).strip()] = build_ledger(s)
    return out, students

# ══════════════════════════════════════════════════════════
#  ROUTES
# ══════════════════════════════════════════════════════════

@app.route("/")
def home():
    return send_from_directory(Path(__file__).parent, "index.html")

@app.route("/favicon.ico")
def favicon():
    return ("", 204)

# ── Session ───────────────────────────────────────────────
@app.get("/api/session")
def api_session():
    if "username" not in session:
        return jsonify({"logged_in": False})
    return jsonify({
        "logged_in": True,
        "username": session["username"],
        "role": session.get("role", ""),
        "permissions": get_perms(session["username"])
    })

@app.get("/api/meta")
@login_required
def api_meta():
    settings = get_settings()
    classes = [str(r.get('Class', '')).strip() for r in data_mgr.get_all(FeeStructure) if r.get('Class')]
    students = data_mgr.get_all(Student)
    sections = sorted({str(s.get('Section', '')).strip() for s in students if s.get('Section')})
    return jsonify({
        "success": True,
        "settings": settings,
        "session_months": [{"key": k, "label": month_label(k)} for k in session_months()],
        "classes": classes or ["Nursery", "LKG", "UKG", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"],
        "sections": sections or ["A", "B", "C"],
        "roles": ROLES,
        "permission_sections": ALL_SECTIONS,
        "fee_types": ["Tuition Fee", "Admission Fee", "Transport Fee", "Exam Fee", "Annual Fee", "Other Fee"],
        "payment_modes": ["Cash", "UPI", "Bank Transfer", "Cheque", "Other"],
        "attendance_status": ["Present", "Absent", "Leave", "Half Day", "Paid Leave"],
        "security_questions": [
            "What is your mother's maiden name?",
            "What was the name of your first school?",
            "What is your favorite teacher's name?",
            "What is your pet's name?",
            "What city were you born in?",
            "What is your favorite food?",
            "What was your childhood nickname?",
            "What is your father's middle name?",
        ]
    })

# ══════════════════════════════════════════════════════════
#  LOGIN / AUTH  (Admin unlimited, others 3 attempts)
# ══════════════════════════════════════════════════════════
@app.post("/api/login")
def api_login():
    data = request.get_json() or {}
    username = str(data.get("username", "")).strip()
    password = str(data.get("password", ""))

    if not username or not password:
        return jsonify({"success": False, "message": "Username and password required."}), 400

    user = data_mgr.get_one(User, username=username)
    if not user:
        # Try case-insensitive
        if data_mgr.pg_available:
            user = User.query.filter(db.func.lower(User.username) == username.lower()).first()
        if not user:
            return jsonify({"success": False, "message": "Invalid username or password."}), 401

    if str(user.status).lower() != 'active':
        return jsonify({"success": False, "message": "This account is inactive."}), 403

    # Admin bypasses lockout
    is_admin = user.role in ('Super Admin', 'Admin')

    if user.locked and not is_admin:
        return jsonify({
            "success": False, "locked": True,
            "message": "Account LOCKED. Contact admin to unlock."
        }), 403

    if not user.check_password(password):
        if not is_admin:
            user.failed_attempts = (user.failed_attempts or 0) + 1
            if user.failed_attempts >= MAX_LOGIN_ATTEMPTS:
                user.locked = True
                data_mgr.save(user)
                add_audit(username, "ACCOUNT LOCKED",
                          f"Locked after {user.failed_attempts} failed attempts.")
                return jsonify({
                    "success": False, "locked": True,
                    "message": f"Account LOCKED after {MAX_LOGIN_ATTEMPTS} wrong attempts. Ask admin to unlock."
                }), 403
            data_mgr.save(user)
            remaining = MAX_LOGIN_ATTEMPTS - user.failed_attempts
            add_audit(username, "LOGIN FAILED",
                      f"Wrong password. Attempt {user.failed_attempts}/{MAX_LOGIN_ATTEMPTS}.")
            return jsonify({
                "success": False, "remaining": remaining,
                "message": f"Wrong password. {remaining} attempt(s) left before lock."
            }), 401
        else:
            add_audit(username, "LOGIN FAILED", "Admin wrong password.")
            return jsonify({"success": False, "message": "Invalid username or password."}), 401

    # Success
    user.failed_attempts = 0
    user.locked = False
    data_mgr.save(user)

    session.permanent = True
    session["username"] = user.username
    session["role"] = user.role

    add_audit(user.username, "LOGIN", "User logged in.")
    return jsonify({
        "success": True,
        "username": user.username,
        "role": user.role,
        "permissions": get_perms(user.username)
    })

@app.post("/api/logout")
@login_required
def api_logout():
    add_audit(session.get("username"), "LOGOUT", "User logged out.")
    session.clear()
    return jsonify({"success": True})

# ── Password Reset via Security Questions ────────────────
@app.post("/api/forgot-password/init")
def forgot_password_init():
    """Step 1: Get 3 random security questions for a username."""
    data = request.get_json() or {}
    username = str(data.get("username", "")).strip()
    user = data_mgr.get_one(User, username=username)
    if not user:
        return jsonify({"success": False, "message": "User not found."}), 404

    # Gather available questions
    questions = []
    for i in (1, 2, 3):
        q = getattr(user, f'security_q{i}', None)
        if q:
            questions.append({"index": i, "question": q})

    if not questions:
        return jsonify({"success": False,
                        "message": "No security questions set. Contact admin."}), 400

    # Pick 3 random (or all if less than 3)
    import random
    random.shuffle(questions)
    selected = questions[:3]
    return jsonify({
        "success": True,
        "username": username,
        "questions": selected
    })

@app.post("/api/forgot-password/reset")
def forgot_password_reset():
    """Step 2: Verify answers and reset password."""
    data = request.get_json() or {}
    username = str(data.get("username", "")).strip()
    answers = data.get("answers", {})  # {"1": "answer", "2": "answer", "3": "answer"}
    new_password = str(data.get("new_password", ""))

    if len(new_password) < 6:
        return jsonify({"success": False,
                        "message": "Password must be at least 6 characters."}), 400

    user = data_mgr.get_one(User, username=username)
    if not user:
        return jsonify({"success": False, "message": "User not found."}), 404

    # Verify each provided answer
    for idx_str, ans in answers.items():
        idx = int(idx_str)
        hashed = getattr(user, f'security_a{idx}', None)
        if not hashed:
            return jsonify({"success": False,
                            "message": f"Question {idx} not configured."}), 400
        if not user.verify_security_answer(str(ans), hashed):
            return jsonify({"success": False,
                            "message": "One or more security answers are incorrect."}), 401

    user.set_password(new_password)
    user.failed_attempts = 0
    user.locked = False
    data_mgr.save(user)

    add_audit(username, "PASSWORD RESET", "Reset via security questions.")
    return jsonify({"success": True, "message": "Password reset successfully. Please login."})

# ── Setup Security Questions (logged-in user) ─────────────
@app.post("/api/setup-security-questions")
@login_required
def setup_security_questions():
    data = request.get_json() or {}
    qa = data.get("qa", [])  # list of {"q": "...", "a": "..."}
    if len(qa) < 3:
        return jsonify({"success": False,
                        "message": "At least 3 security questions required."}), 400

    user = data_mgr.get_one(User, username=session["username"])
    if not user:
        return jsonify({"success": False, "message": "User not found."}), 404

    for i, item in enumerate(qa[:3], start=1):
        setattr(user, f'security_q{i}', str(item.get("q", ""))[:255])
        setattr(user, f'security_a{i}', user.set_security_answer(str(item.get("a", ""))))

    data_mgr.save(user)
    add_audit(session["username"], "SECURITY QUESTIONS", "Updated.")
    return jsonify({"success": True, "message": "Security questions saved."})

@app.post("/api/change-password")
@login_required
def change_password():
    data = request.get_json() or {}
    old = str(data.get("old_password", ""))
    new = str(data.get("new_password", ""))
    if len(new) < 6:
        return jsonify({"success": False,
                        "message": "Password must be at least 6 characters."}), 400
    user = data_mgr.get_one(User, username=session["username"])
    if not user:
        return jsonify({"success": False, "message": "User not found."}), 404
    if not user.check_password(old):
        return jsonify({"success": False, "message": "Old password is incorrect."}), 400
    user.set_password(new)
    data_mgr.save(user)
    add_audit(session["username"], "CHANGE PASSWORD", "Password changed.")
    return jsonify({"success": True, "message": "Password changed successfully."})

# ══════════════════════════════════════════════════════════
#  DASHBOARD
# ══════════════════════════════════════════════════════════
@app.get("/api/dashboard")
@require_perm("dashboard", "view")
def api_dashboard():
    students = data_mgr.get_all(Student)
    fees = data_mgr.get_all(FeeTransaction)
    staff = [s for s in data_mgr.get_all(Staff) if str(s.get('Status', '')).lower() == 'active']
    attendance = data_mgr.get_all(Attendance)

    active = [s for s in students if str(s.get('Status', '')).lower() == 'active']
    today = date.today()
    today_str = today.strftime('%Y-%m-%d')
    month_str = today.strftime('%Y-%m')

    total_collected = today_col = month_col = cash_total = online_total = 0.0
    monthly = {}

    for f in fees:
        amt = safe_float(f.get('Amount'))
        total_collected += amt
        fd = str(f.get('Date') or '')
        if fd == today_str: today_col += amt
        if fd.startswith(month_str): month_col += amt
        if str(f.get('Payment Mode', '')).lower() == 'cash':
            cash_total += amt
        else:
            online_total += amt
        if len(fd) >= 7:
            monthly[fd[:7]] = monthly.get(fd[:7], 0.0) + amt

    ledgers, _ = build_all_ledgers()
    total_demand = total_pending = total_late = overdue_amount = 0.0
    overdue_students = due_today = 0
    class_collection = {}

    for adm, l in ledgers.items():
        total_demand += l['total_due']
        total_pending += l['pending']
        total_late += l['late_fee']
        if l['overdue_amount'] > 0:
            overdue_students += 1
            overdue_amount += l['overdue_amount']
        cls = l['class_name'] or '—'
        class_collection.setdefault(cls, {'paid': 0.0, 'pending': 0.0})
        class_collection[cls]['paid'] += l['total_paid']
        class_collection[cls]['pending'] += l['pending']
        for m in l['months']:
            if m['due_date'] == today_str and m['pending'] > 0:
                due_today += 1

    today_att = [a for a in attendance if str(a.get('Date')) == today_str]
    present_today = sum(1 for a in today_att if str(a.get('Status', '')).lower() in ('present', 'half day'))
    absent_today = sum(1 for a in today_att if str(a.get('Status', '')).lower() == 'absent')
    leave_today = sum(1 for a in today_att if 'leave' in str(a.get('Status', '')).lower())
    monthly_payroll = sum(safe_float(s.get('Basic Salary')) + safe_float(s.get('Allowances')) for s in staff)

    monthly_list = [{'month': k, 'amount': round(v, 2)} for k, v in sorted(monthly.items())]
    class_list = [{'class_name': k, 'paid': round(v['paid'], 2), 'pending': round(v['pending'], 2)}
                  for k, v in sorted(class_collection.items())]

    recent = sorted(fees, key=lambda x: (str(x.get('Date') or ''), str(x.get('Time') or '')), reverse=True)[:10]
    top_def = sorted(
        [{'admission_no': a, 'student_name': l['student_name'], 'class_name': l['class_name'],
          'section': l['section'], 'pending': l['pending'], 'overdue_days': l['overdue_days'],
          'late_fee': l['late_fee']} for a, l in ledgers.items() if l['pending'] > 0],
        key=lambda x: x['pending'], reverse=True
    )[:8]

    return jsonify({
        "success": True,
        "data": {
            "total_students": len(students),
            "active_students": len(active),
            "inactive_students": len(students) - len(active),
            "total_demand": round(total_demand, 2),
            "total_collected": round(total_collected, 2),
            "pending": round(total_pending, 2),
            "late_fee_total": round(total_late, 2),
            "today_collection": round(today_col, 2),
            "month_collection": round(month_col, 2),
            "total_receipts": len(fees),
            "cash_collection": round(cash_total, 2),
            "online_collection": round(online_total, 2),
            "overdue_students": overdue_students,
            "overdue_amount": round(overdue_amount, 2),
            "due_today": due_today,
            "total_staff": len(staff),
            "present_today": present_today,
            "absent_today": absent_today,
            "leave_today": leave_today,
            "monthly_payroll": round(monthly_payroll, 2),
            "monthly_collection": monthly_list,
            "class_collection": class_list,
            "recent_transactions": recent,
            "top_defaulters": top_def,
            "db_mode": "PostgreSQL" if data_mgr.pg_available else "Excel"
        }
    })

# ══════════════════════════════════════════════════════════
#  STUDENTS + ADMISSION
# ══════════════════════════════════════════════════════════
@app.get("/api/students")
@require_perm("students", "view")
def api_students():
    return jsonify({"success": True, "students": data_mgr.get_all(Student)})

@app.post("/api/students")
@require_perm("admission", "edit")
def api_create_student():
    data = request.get_json() or {}
    adm = str(data.get('admission_no', '')).strip()
    name = str(data.get('student_name', '')).strip()
    if not adm or not name:
        return jsonify({"success": False,
                        "message": "Admission No and Student Name are required."}), 400

    existing = data_mgr.get_one(Student, admission_no=adm)
    if existing:
        return jsonify({"success": False, "message": "Admission number already exists."}), 409

    s = Student(
        admission_no=adm,
        student_name=name,
        father_name=data.get('father_name', ''),
        mother_name=data.get('mother_name', ''),
        class_name=data.get('class_name', ''),
        section=data.get('section', ''),
        roll_no=data.get('roll_no', ''),
        dob=data.get('dob', ''),
        gender=data.get('gender', ''),
        mobile=data.get('mobile', ''),
        address=data.get('address', ''),
        admission_date=data.get('admission_date', datetime.now().strftime('%Y-%m-%d')),
        session=data.get('session', '2026-27'),
        transport=data.get('transport', 'No'),
        status=data.get('status', 'Active')
    )
    data_mgr.save(s)
    add_audit(session["username"], "CREATE STUDENT", f"{adm} - {name}")
    return jsonify({"success": True, "message": "Student admitted successfully."})

@app.put("/api/students/<admission_no>")
@require_perm("students", "edit")
def api_update_student(admission_no):
    data = request.get_json() or {}
    s = data_mgr.get_one(Student, admission_no=admission_no)
    if not s:
        return jsonify({"success": False, "message": "Student not found."}), 404

    s.student_name = data.get('student_name', s.student_name)
    s.father_name = data.get('father_name', s.father_name)
    s.mother_name = data.get('mother_name', s.mother_name)
    s.class_name = data.get('class_name', s.class_name)
    s.section = data.get('section', s.section)
    s.roll_no = data.get('roll_no', s.roll_no)
    s.dob = data.get('dob', s.dob)
    s.gender = data.get('gender', s.gender)
    s.mobile = data.get('mobile', s.mobile)
    s.address = data.get('address', s.address)
    s.admission_date = data.get('admission_date', s.admission_date)
    s.session = data.get('session', s.session)
    s.transport = data.get('transport', s.transport)
    s.status = data.get('status', s.status)
    data_mgr.save(s)
    add_audit(session["username"], "UPDATE STUDENT", admission_no)
    return jsonify({"success": True, "message": "Student updated."})

@app.delete("/api/students/<admission_no>")
@require_perm("students", "edit")
def api_delete_student(admission_no):
    s = data_mgr.get_one(Student, admission_no=admission_no)
    if not s:
        return jsonify({"success": False, "message": "Student not found."}), 404
    s.status = 'Inactive'
    data_mgr.save(s)
    add_audit(session["username"], "DEACTIVATE STUDENT", admission_no)
    return jsonify({"success": True, "message": "Student deactivated."})

@app.get("/api/students/<admission_no>/history")
@require_perm("students", "view")
def api_student_history(admission_no):
    s = data_mgr.get_one(Student, admission_no=admission_no)
    if not s:
        return jsonify({"success": False, "message": "Student not found."}), 404
    ledger = build_ledger(s.to_dict())
    txns = [t for t in data_mgr.get_all(FeeTransaction)
            if str(t.get('Admission No', '')).lower() == admission_no.lower()]
    txns = sorted(txns, key=lambda x: (str(x.get('Date') or ''), str(x.get('Time') or '')), reverse=True)
    return jsonify({
        "success": True,
        "student": s.to_dict(),
        "transactions": txns,
        "total_paid": ledger['total_paid'],
        "ledger": ledger
    })

# ══════════════════════════════════════════════════════════
#  FEE SUBMISSION
# ══════════════════════════════════════════════════════════
def generate_receipt_no():
    txs = data_mgr.get_all(FeeTransaction)
    settings = get_settings()
    prefix = str(settings.get('Receipt Prefix', 'BDPS') or 'BDPS')
    today = datetime.now().strftime('%Y%m%d')
    nums = []
    for t in txs:
        m = re.search(r'(\d+)$', str(t.get('Receipt No') or ''))
        if m: nums.append(int(m.group(1)))
    nxt = (max(nums) + 1) if nums else 1
    return f"{prefix}-{today}-{nxt:05d}"

@app.post("/api/fees")
@require_perm("fee", "edit")
def api_submit_fee():
    data = request.get_json() or {}
    adm = str(data.get('admission_no', '')).strip()
    amount = safe_float(data.get('amount'))
    late_fee = safe_float(data.get('late_fee'))
    if not adm:
        return jsonify({"success": False, "message": "Admission number required."}), 400
    if amount <= 0:
        return jsonify({"success": False, "message": "Enter a valid fee amount."}), 400

    s = data_mgr.get_one(Student, admission_no=adm)
    if not s:
        return jsonify({"success": False, "message": "Student not found."}), 404

    now = datetime.now()
    txn = FeeTransaction(
        receipt_no=generate_receipt_no(),
        date=data.get('date') or now.strftime('%Y-%m-%d'),
        time=now.strftime('%H:%M:%S'),
        admission_no=adm,
        student_name=s.student_name,
        class_name=s.class_name,
        section=s.section,
        fee_type=data.get('fee_type', 'Tuition Fee'),
        month=data.get('month', ''),
        amount=amount,
        late_fee=late_fee,
        payment_mode=data.get('payment_mode', 'Cash'),
        transaction_no=data.get('transaction_no', ''),
        remarks=data.get('remarks', ''),
        received_by=session["username"]
    )
    data_mgr.save(txn)
    add_audit(session["username"], "FEE SUBMITTED",
              f"{txn.receipt_no} / {adm} / Rs.{amount} / Late Rs.{late_fee}")
    return jsonify({
        "success": True,
        "message": "Fee submitted successfully.",
        "receipt_no": txn.receipt_no
    })

@app.get("/api/fees")
@require_perm("receipts", "view")
def api_all_fees():
    return jsonify({"success": True, "fees": data_mgr.get_all(FeeTransaction)})

# ══════════════════════════════════════════════════════════
#  PENDING
# ══════════════════════════════════════════════════════════
@app.get("/api/pending")
@require_perm("pending", "view")
def api_pending():
    ledgers, _ = build_all_ledgers()
    result = []
    for adm, l in ledgers.items():
        if l['pending'] > 0:
            result.append({
                'admission_no': adm,
                'student_name': l['student_name'],
                'class_name': l['class_name'],
                'section': l['section'],
                'father_name': l['father_name'],
                'mobile': l['mobile'],
                'total_fee': l['total_due'],
                'paid': l['total_paid'],
                'pending': l['pending'],
                'late_fee': l['late_fee'],
                'total_payable': l['total_payable'],
                'overdue_days': l['overdue_days'],
                'pending_months': l['pending_months'],
                'months': [m for m in l['months'] if m['pending'] > 0]
            })
    result.sort(key=lambda x: x['pending'], reverse=True)
    return jsonify({"success": True, "pending": result})

# ══════════════════════════════════════════════════════════
#  FEE STRUCTURE
# ══════════════════════════════════════════════════════════
@app.get("/api/fee-structure")
@require_perm("structure", "view")
def api_get_structure():
    return jsonify({"success": True, "data": data_mgr.get_all(FeeStructure)})

@app.post("/api/fee-structure")
@require_perm("structure", "edit")
def api_save_structure():
    data = request.get_json() or {}
    cls = str(data.get('class_name', '')).strip()
    if not cls:
        return jsonify({"success": False, "message": "Class is required."}), 400

    obj = data_mgr.get_one(FeeStructure, class_name=cls)
    if not obj:
        obj = FeeStructure(class_name=cls)
        if data_mgr.pg_available:
            db.session.add(obj)
    obj.admission_fee = safe_float(data.get('admission_fee'))
    obj.tuition_fee = safe_float(data.get('tuition_fee'))
    obj.transport_fee = safe_float(data.get('transport_fee'))
    obj.exam_fee = safe_float(data.get('exam_fee'))
    obj.annual_fee = safe_float(data.get('annual_fee'))
    obj.other_fee = safe_float(data.get('other_fee'))
    data_mgr.save(obj)
    add_audit(session["username"], "UPDATE FEE STRUCTURE", cls)
    return jsonify({"success": True, "message": "Fee structure saved."})

# ══════════════════════════════════════════════════════════
#  STAFF / HIRING
# ══════════════════════════════════════════════════════════
@app.get("/api/staff")
@require_perm("staff", "view")
def api_staff():
    return jsonify({"success": True, "staff": data_mgr.get_all(Staff)})

@app.post("/api/staff")
@require_perm("hiring", "edit")
def api_create_staff():
    data = request.get_json() or {}
    sid = str(data.get('staff_id', '')).strip()
    name = str(data.get('name', '')).strip()
    if not sid or not name:
        return jsonify({"success": False,
                        "message": "Staff ID and Name are required."}), 400
    if data_mgr.get_one(Staff, staff_id=sid):
        return jsonify({"success": False, "message": "Staff ID already exists."}), 409

    s = Staff(
        staff_id=sid,
        name=name,
        designation=data.get('designation', ''),
        department=data.get('department', ''),
        mobile=data.get('mobile', ''),
        joining_date=data.get('joining_date', datetime.now().strftime('%Y-%m-%d')),
        basic_salary=safe_float(data.get('basic_salary')),
        allowances=safe_float(data.get('allowances')),
        status=data.get('status', 'Active')
    )
    data_mgr.save(s)
    add_audit(session["username"], "HIRE STAFF", f"{sid} - {name}")
    return jsonify({"success": True, "message": "Staff hired successfully."})

@app.put("/api/staff/<staff_id>")
@require_perm("hiring", "edit")
def api_update_staff(staff_id):
    data = request.get_json() or {}
    s = data_mgr.get_one(Staff, staff_id=staff_id)
    if not s:
        return jsonify({"success": False, "message": "Staff not found."}), 404
    s.name = data.get('name', s.name)
    s.designation = data.get('designation', s.designation)
    s.department = data.get('department', s.department)
    s.mobile = data.get('mobile', s.mobile)
    s.joining_date = data.get('joining_date', s.joining_date)
    s.basic_salary = safe_float(data.get('basic_salary'), s.basic_salary)
    s.allowances = safe_float(data.get('allowances'), s.allowances)
    s.status = data.get('status', s.status)
    data_mgr.save(s)
    add_audit(session["username"], "UPDATE STAFF", staff_id)
    return jsonify({"success": True, "message": "Staff updated."})

# ══════════════════════════════════════════════════════════
#  ATTENDANCE  (checkbox-based, backdate approval)
# ══════════════════════════════════════════════════════════
@app.get("/api/attendance")
@require_perm("attendance", "view")
def api_get_attendance():
    day = request.args.get('date') or date.today().strftime('%Y-%m-%d')
    staff = [s for s in data_mgr.get_all(Staff) if str(s.get('Status', '')).lower() == 'active']
    existing = {str(a.get('Staff ID', '')).strip(): a
                for a in data_mgr.get_all(Attendance) if str(a.get('Date')) == day}
    result = []
    for s in staff:
        sid = str(s.get('Staff ID', '')).strip()
        rec = existing.get(sid)
        result.append({
            'staff_id': sid,
            'name': s.get('Name', ''),
            'designation': s.get('Designation', ''),
            'department': s.get('Department', ''),
            'status': rec.get('Status', 'Present') if rec else 'Present',
            'remarks': rec.get('Remarks', '') if rec else '',
            'marked': bool(rec),
            'approval_status': rec.get('Approval Status', 'Approved') if rec else 'Approved'
        })
    summary = {}
    for r in result:
        summary[r['status']] = summary.get(r['status'], 0) + 1
    return jsonify({"success": True, "date": day, "attendance": result, "summary": summary})

@app.post("/api/attendance")
@require_perm("attendance", "edit")
def api_save_attendance():
    data = request.get_json() or {}
    day = str(data.get('date') or date.today().strftime('%Y-%m-%d'))
    records = data.get('records') or []
    is_backdate = day < date.today().strftime('%Y-%m-%d')

    # Check permission for backdate
    user = data_mgr.get_one(User, username=session["username"])
    can_backdate = False
    if user and user.role in ('Super Admin', 'Admin'):
        can_backdate = True
    else:
        settings = get_settings()
        if str(settings.get('Allow Backdate Attendance', 'No')).lower() in ('yes', 'true', '1'):
            can_backdate = True

    if is_backdate and not can_backdate:
        return jsonify({"success": False,
                        "message": "Backdate attendance requires admin approval. Contact admin."}), 403

    for rec in records:
        sid = str(rec.get('staff_id', '')).strip()
        s = data_mgr.get_one(Staff, staff_id=sid)
        # Remove existing
        existing = Attendance.query.filter_by(date=day, staff_id=sid).first() if data_mgr.pg_available else None
        if existing:
            existing.status = rec.get('status', 'Present')
            existing.remarks = rec.get('remarks', '')
            existing.marked_by = session["username"]
            existing.approval_status = 'Approved' if not is_backdate else 'Pending'
        else:
            att = Attendance(
                date=day, staff_id=sid,
                staff_name=s.name if s else '',
                status=rec.get('status', 'Present'),
                remarks=rec.get('remarks', ''),
                marked_by=session["username"],
                approval_status='Approved' if not is_backdate else 'Pending'
            )
            data_mgr.save(att)

    add_audit(session["username"], "MARK ATTENDANCE", f"{day} ({len(records)} staff)")
    return jsonify({"success": True, "message": "Attendance saved.",
                    "backdate": is_backdate})

@app.post("/api/attendance/approve")
@require_perm("attendance", "edit")
def api_approve_backdate():
    """Approve backdate attendance (admin only)."""
    user = data_mgr.get_one(User, username=session["username"])
    if not user or user.role not in ('Super Admin', 'Admin'):
        return jsonify({"success": False, "message": "Admin permission required."}), 403
    data = request.get_json() or {}
    day = data.get('date')
    if not day:
        return jsonify({"success": False, "message": "Date required."}), 400
    if data_mgr.pg_available:
        records = Attendance.query.filter_by(date=day, approval_status='Pending').all()
        for r in records:
            r.approval_status = 'Approved'
            r.approved_by = session["username"]
        db.session.commit()
    add_audit(session["username"], "APPROVE BACKDATE", day)
    return jsonify({"success": True, "message": "Backdate attendance approved."})

# ══════════════════════════════════════════════════════════
#  SALARY
# ══════════════════════════════════════════════════════════
def compute_salary(month):
    settings = get_settings()
    wd = safe_int(settings.get('Working Days Per Month', 26), 26)
    hf = safe_float(settings.get('Half Day Count', 0.5), 0.5)
    lop_on = str(settings.get('Salary LOP Deduction', 'Yes')).lower() in ('yes', 'true', '1')
    staff = [s for s in data_mgr.get_all(Staff) if str(s.get('Status', '')).lower() == 'active']
    att = [a for a in data_mgr.get_all(Attendance) if str(a.get('Date', '')).startswith(month)]
    by_staff = {}
    for a in att:
        sid = str(a.get('Staff ID', '')).strip()
        by_staff.setdefault(sid, []).append(str(a.get('Status', '')).strip())
    results = []
    for s in staff:
        sid = str(s.get('Staff ID', '')).strip()
        statuses = by_staff.get(sid, [])
        present = sum(1 for x in statuses if x == 'Present')
        absent = sum(1 for x in statuses if x == 'Absent')
        leave = sum(1 for x in statuses if x == 'Leave')
        paid_leave = sum(1 for x in statuses if x == 'Paid Leave')
        half_day = sum(1 for x in statuses if x == 'Half Day')
        basic = safe_float(s.get('Basic Salary'))
        allow = safe_float(s.get('Allowances'))
        per_day = (basic / wd) if wd else 0
        lop_days = (absent + half_day * hf) if lop_on else 0
        deduction = round(per_day * lop_days, 2)
        net = round(basic + allow - deduction, 2)
        results.append({
            'month': month, 'staff_id': sid, 'staff_name': s.get('Name', ''),
            'designation': s.get('Designation', ''), 'department': s.get('Department', ''),
            'working_days': wd, 'present': present, 'absent': absent,
            'leave': leave, 'paid_leave': paid_leave, 'half_day': half_day,
            'lop': round(lop_days, 2), 'basic_salary': round(basic, 2),
            'allowances': round(allow, 2), 'deductions': deduction,
            'advance': 0.0, 'net_salary': net
        })
    return results

@app.post("/api/salary/generate")
@require_perm("salary", "edit")
def api_generate_salary():
    data = request.get_json() or {}
    month = str(data.get('month') or datetime.now().strftime('%Y-%m'))
    results = compute_salary(month)
    # Clear old
    if data_mgr.pg_available:
        SalaryRecord.query.filter_by(month=month).delete()
        db.session.commit()
    for r in results:
        rec = SalaryRecord(
            month=r['month'], staff_id=r['staff_id'], staff_name=r['staff_name'],
            working_days=r['working_days'], present=r['present'], absent=r['absent'],
            leave=r['leave'], half_day=r['half_day'], lop=r['lop'],
            basic_salary=r['basic_salary'], allowances=r['allowances'],
            deductions=r['deductions'], advance=r['advance'], net_salary=r['net_salary'],
            generated_by=session["username"]
        )
        data_mgr.save(rec)
    add_audit(session["username"], "GENERATE SALARY", f"{month} ({len(results)} staff)")
    total = sum(r['net_salary'] for r in results)
    return jsonify({"success": True, "message": "Salary generated.",
                    "salary": results, "total": round(total, 2)})

@app.get("/api/salary")
@require_perm("salary", "view")
def api_get_salary():
    month = request.args.get('month') or datetime.now().strftime('%Y-%m')
    records = [r for r in data_mgr.get_all(SalaryRecord) if str(r.get('Month')) == month]
    if not records:
        records = compute_salary(month)
    total = sum(safe_float(r.get('Net Salary', r.get('net_salary', 0))) for r in records)
    return jsonify({"success": True, "month": month, "salary": records,
                    "total": round(total, 2),
                    "generated": bool([r for r in data_mgr.get_all(SalaryRecord) if str(r.get('Month')) == month])})

# ══════════════════════════════════════════════════════════
#  USERS + PERMISSIONS
# ══════════════════════════════════════════════════════════
@app.get("/api/users")
@require_perm("users", "view")
def api_users():
    users = data_mgr.get_all(User)
    return jsonify({"success": True, "users": users})

@app.post("/api/users")
@require_perm("users", "edit")
def api_create_user():
    data = request.get_json() or {}
    username = str(data.get('username', '')).strip()
    password = str(data.get('password', ''))
    role = str(data.get('role', 'Data Entry'))
    if not username or not password:
        return jsonify({"success": False,
                        "message": "Username and password required."}), 400
    if len(password) < 4:
        return jsonify({"success": False,
                        "message": "Password must be at least 4 characters."}), 400
    if data_mgr.get_one(User, username=username):
        return jsonify({"success": False, "message": "Username already exists."}), 409
    perms = data.get('permissions') or ROLE_DEFAULTS.get(role, {})
    u = User(username=username, role=role, status='Active',
             permissions=json.dumps(perms))
    u.set_password(password)
    data_mgr.save(u)
    add_audit(session["username"], "CREATE USER", f"{username} ({role})")
    return jsonify({"success": True, "message": "User created."})

@app.post("/api/users/<username>/reset-password")
@require_perm("users", "edit")
def api_admin_reset_password(username):
    data = request.get_json() or {}
    new = str(data.get('new_password', ''))
    if len(new) < 4:
        return jsonify({"success": False,
                        "message": "Password must be at least 4 characters."}), 400
    u = data_mgr.get_one(User, username=username)
    if not u:
        return jsonify({"success": False, "message": "User not found."}), 404
    u.set_password(new)
    u.failed_attempts = 0
    u.locked = False
    data_mgr.save(u)
    add_audit(session["username"], "RESET PASSWORD", username)
    return jsonify({"success": True, "message": f"Password reset for {username}."})

@app.post("/api/users/<username>/unlock")
@require_perm("users", "edit")
def api_unlock_user(username):
    u = data_mgr.get_one(User, username=username)
    if not u:
        return jsonify({"success": False, "message": "User not found."}), 404
    u.failed_attempts = 0
    u.locked = False
    data_mgr.save(u)
    add_audit(session["username"], "UNLOCK USER", username)
    return jsonify({"success": True, "message": f"{username} unlocked."})

@app.post("/api/users/<username>/status")
@require_perm("users", "edit")
def api_toggle_status(username):
    data = request.get_json() or {}
    status = data.get('status', 'Active')
    if status not in ('Active', 'Inactive'):
        return jsonify({"success": False, "message": "Invalid status."}), 400
    u = data_mgr.get_one(User, username=username)
    if not u:
        return jsonify({"success": False, "message": "User not found."}), 404
    if str(u.username).lower() == 'admin' and status == 'Inactive':
        return jsonify({"success": False,
                        "message": "Default admin cannot be deactivated."}), 400
    u.status = status
    data_mgr.save(u)
    add_audit(session["username"], "USER STATUS", f"{username} -> {status}")
    return jsonify({"success": True, "message": f"{username} is now {status}."})

@app.post("/api/users/<username>/permissions")
@require_perm("users", "edit")
def api_update_permissions(username):
    data = request.get_json() or {}
    perms = data.get('permissions') or {}
    clean = {}
    for sec in ALL_SECTIONS:
        val = str(perms.get(sec, 'none')).lower()
        clean[sec] = val if val in ('none', 'view', 'edit') else 'none'
    u = data_mgr.get_one(User, username=username)
    if not u:
        return jsonify({"success": False, "message": "User not found."}), 404
    u.permissions = json.dumps(clean)
    data_mgr.save(u)
    add_audit(session["username"], "UPDATE PERMISSIONS", username)
    return jsonify({"success": True, "message": "Permissions updated."})

# ══════════════════════════════════════════════════════════
#  SETTINGS
# ══════════════════════════════════════════════════════════
@app.get("/api/settings")
@require_perm("settings", "view")
def api_get_settings():
    return jsonify({"success": True, "settings": get_settings()})

@app.post("/api/settings")
@require_perm("settings", "edit")
def api_save_settings():
    data = request.get_json() or {}
    for k, v in data.items():
        set_setting(k, v)
    add_audit(session["username"], "UPDATE SETTINGS", json.dumps(data))
    return jsonify({"success": True, "message": "Settings saved."})

# ══════════════════════════════════════════════════════════
#  AUDIT
# ══════════════════════════════════════════════════════════
@app.get("/api/audit")
@require_perm("audit", "view")
def api_audit():
    logs = data_mgr.get_all(AuditLog)
    logs.reverse()
    return jsonify({"success": True, "logs": logs[:800]})

# ══════════════════════════════════════════════════════════
#  NOTIFICATIONS
# ══════════════════════════════════════════════════════════
@app.get("/api/notifications")
@require_perm("notifications", "view")
def api_notifications():
    today = date.today()
    alerts = []
    ledgers, _ = build_all_ledgers()
    for adm, l in ledgers.items():
        if l['overdue_days'] > 0 and l['pending'] > 0:
            alerts.append({
                'type': 'danger' if l['overdue_days'] > 15 else 'warning',
                'title': f"{l['student_name']} ({adm})",
                'message': f"Rs.{l['pending']:.0f} pending • {l['overdue_days']} days overdue • Late Rs.{l['late_fee']:.0f}",
                'section': 'pending'
            })
        for m in l['months']:
            if m['due_date'] == today.strftime('%Y-%m-%d') and m['pending'] > 0:
                alerts.append({
                    'type': 'info',
                    'title': f"Fee due today - {l['student_name']}",
                    'message': f"{m['label']} fee Rs.{m['pending']:.0f} due today.",
                    'section': 'fee'
                })
    for u in data_mgr.get_all(User):
        if u.get('locked') in (True, 'Yes', 'yes', 'true', '1'):
            alerts.append({
                'type': 'danger',
                'title': f"Account locked: {u.get('username')}",
                'message': "Admin must unlock this account.",
                'section': 'users'
            })
    today_str = today.strftime('%Y-%m-%d')
    absent = [a for a in data_mgr.get_all(Attendance)
              if str(a.get('Date')) == today_str and str(a.get('Status', '')).lower() == 'absent']
    if absent:
        alerts.append({
            'type': 'warning',
            'title': f"{len(absent)} staff absent today",
            'message': ', '.join(str(a.get('Staff Name', '')) for a in absent[:5]),
            'section': 'attendance'
        })
    return jsonify({"success": True, "alerts": alerts[:100], "count": len(alerts)})

# ══════════════════════════════════════════════════════════
#  REPORTS + EXPORT
# ══════════════════════════════════════════════════════════
@app.get("/api/reports/data")
@require_perm("reports", "view")
def api_reports_data():
    rtype = request.args.get('type', 'fees')
    if rtype == 'fees':
        fees = data_mgr.get_all(FeeTransaction)
        frm = request.args.get('from') or ''
        to = request.args.get('to') or ''
        cls = (request.args.get('class') or '').strip().lower()
        mode = (request.args.get('mode') or '').strip().lower()
        out = []
        for f in fees:
            fd = str(f.get('Date') or '')
            if frm and fd < frm: continue
            if to and fd > to: continue
            if cls and str(f.get('Class', '')).strip().lower() != cls: continue
            if mode and str(f.get('Payment Mode', '')).strip().lower() != mode: continue
            out.append(f)
        out.sort(key=lambda x: (str(x.get('Date') or ''), str(x.get('Time') or '')), reverse=True)
        total = sum(safe_float(f.get('Amount')) for f in out)
        return jsonify({"success": True, "type": rtype, "rows": out,
                        "total": round(total, 2), "count": len(out)})
    if rtype == 'pending':
        ledgers, _ = build_all_ledgers()
        rows = [{
            'Admission No': adm, 'Student Name': l['student_name'],
            'Class': l['class_name'], 'Section': l['section'],
            'Father Name': l['father_name'], 'Mobile': l['mobile'],
            'Total Fee': l['total_due'], 'Paid': l['total_paid'],
            'Pending': l['pending'], 'Late Fee': l['late_fee'],
            'Total Payable': l['total_payable'], 'Overdue Days': l['overdue_days'],
            'Pending Months': l['pending_months']
        } for adm, l in ledgers.items() if l['pending'] > 0]
        rows.sort(key=lambda x: x['Pending'], reverse=True)
        total = sum(r['Pending'] for r in rows)
        return jsonify({"success": True, "type": rtype, "rows": rows,
                        "total": round(total, 2), "count": len(rows)})
    if rtype == 'students':
        return jsonify({"success": True, "type": rtype,
                        "rows": data_mgr.get_all(Student),
                        "count": len(data_mgr.get_all(Student))})
    if rtype == 'staff':
        return jsonify({"success": True, "type": rtype,
                        "rows": data_mgr.get_all(Staff),
                        "count": len(data_mgr.get_all(Staff))})
    if rtype == 'attendance':
        month = request.args.get('month') or datetime.now().strftime('%Y-%m')
        rows = [a for a in data_mgr.get_all(Attendance) if str(a.get('Date', '')).startswith(month)]
        return jsonify({"success": True, "type": rtype, "rows": rows, "count": len(rows)})
    if rtype == 'salary':
        month = request.args.get('month') or datetime.now().strftime('%Y-%m')
        rows = [r for r in data_mgr.get_all(SalaryRecord) if str(r.get('Month')) == month]
        if not rows:
            rows = compute_salary(month)
        total = sum(safe_float(r.get('Net Salary', r.get('net_salary', 0))) for r in rows)
        return jsonify({"success": True, "type": rtype, "rows": rows,
                        "total": round(total, 2), "count": len(rows)})
    return jsonify({"success": False, "message": "Unknown report type."}), 400

@app.get("/api/reports/export")
@require_perm("reports", "view")
def api_export_report():
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    rtype = request.args.get('type', 'fees')
    rows = []
    headers = []

    if rtype == 'fees':
        rows = data_mgr.get_all(FeeTransaction)
        headers = ["Receipt No", "Date", "Time", "Admission No", "Student Name",
                   "Class", "Section", "Fee Type", "Month", "Amount", "Late Fee",
                   "Payment Mode", "Transaction No", "Remarks", "Received By"]
    elif rtype == 'pending':
        ledgers, _ = build_all_ledgers()
        rows = [{
            'Admission No': adm, 'Student Name': l['student_name'],
            'Class': l['class_name'], 'Section': l['section'],
            'Father Name': l['father_name'], 'Mobile': l['mobile'],
            'Total Fee': l['total_due'], 'Paid': l['total_paid'],
            'Pending': l['pending'], 'Late Fee': l['late_fee'],
            'Total Payable': l['total_payable'], 'Overdue Days': l['overdue_days'],
            'Pending Months': l['pending_months']
        } for adm, l in ledgers.items() if l['pending'] > 0]
        headers = list(rows[0].keys()) if rows else []
    elif rtype == 'students':
        rows = data_mgr.get_all(Student)
        headers = ["Admission No", "Student Name", "Father Name", "Mother Name",
                   "Class", "Section", "Roll No", "DOB", "Gender", "Mobile",
                   "Address", "Admission Date", "Session", "Transport", "Status"]
    elif rtype == 'staff':
        rows = data_mgr.get_all(Staff)
        headers = ["Staff ID", "Name", "Designation", "Department", "Mobile",
                   "Joining Date", "Basic Salary", "Allowances", "Status"]
    elif rtype == 'attendance':
        month = request.args.get('month') or datetime.now().strftime('%Y-%m')
        rows = [a for a in data_mgr.get_all(Attendance) if str(a.get('Date', '')).startswith(month)]
        headers = ["Date", "Staff ID", "Staff Name", "Status", "Remarks"]
    elif rtype == 'salary':
        month = request.args.get('month') or datetime.now().strftime('%Y-%m')
        rows = [r for r in data_mgr.get_all(SalaryRecord) if str(r.get('Month')) == month]
        if not rows:
            rows = compute_salary(month)
        headers = ["Month", "Staff ID", "Staff Name", "Working Days", "Present",
                   "Absent", "Leave", "Half Day", "LOP", "Basic Salary",
                   "Allowances", "Deductions", "Net Salary"]
        rows = [{h: r.get(h, r.get(h.lower().replace(' ', '_'), '')) for h in headers} for r in rows]
    else:
        return jsonify({"success": False, "message": "Unknown report type."}), 400

    wb = Workbook()
    ws = wb.active
    ws.title = rtype[:31]
    for c, h in enumerate(headers, 1):
        ws.cell(1, c, h)
    for r in rows:
        ws.append([r.get(h, '') for h in headers])
    # Style
    fill = PatternFill("solid", fgColor="173F73")
    font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center")
    for col in ws.columns:
        max_len = max((len(str(c.value or '')) for c in col), default=0)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(max_len + 3, 12), 35)

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    fname = f"{rtype}_report_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    add_audit(session["username"], "EXPORT REPORT", f"{rtype} -> {fname}")
    return send_file(bio, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ══════════════════════════════════════════════════════════
#  BACKUP / SYNC
# ══════════════════════════════════════════════════════════
@app.post("/api/backup")
@require_perm("backup", "edit")
def api_backup():
    if not EXCEL_FILE.exists():
        return jsonify({"success": False, "message": "Excel file not found."}), 404
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    dest = BACKUP_DIR / f"Blue_Diamond_Backup_{ts}.xlsx"
    shutil.copy2(EXCEL_FILE, dest)
    add_audit(session["username"], "BACKUP", dest.name)
    return jsonify({"success": True, "message": "Backup created.", "file": dest.name})

@app.get("/api/backups")
@require_perm("backup", "view")
def api_backups():
    files = sorted(BACKUP_DIR.glob("*.xlsx"), reverse=True)
    return jsonify({"success": True, "backups": [
        {"name": f.name, "size": round(f.stat().st_size / 1024, 1),
         "created": datetime.fromtimestamp(f.stat().st_mtime).strftime('%Y-%m-%d %H:%M:%S')}
        for f in files[:50]
    ]})

@app.get("/api/download-excel")
@require_perm("backup", "edit")
def api_download_excel():
    if not EXCEL_FILE.exists():
        return jsonify({"success": False, "message": "File not found."}), 404
    return send_from_directory(DATA_DIR, EXCEL_FILE.name, as_attachment=True)

@app.post("/api/sync")
@require_perm("backup", "edit")
def api_sync():
    """Manually trigger Excel → PostgreSQL sync."""
    if not data_mgr.pg_available:
        data_mgr._check_pg()
    if not data_mgr.pg_available:
        return jsonify({"success": False, "message": "PostgreSQL not reachable."}), 503
    result = data_mgr.sync_excel_to_pg()
    if result['success']:
        add_audit(session["username"], "SYNC", f"Excel → PG: {result['counts']}")
    return jsonify(result)

@app.get("/api/db-status")
@login_required
def api_db_status():
    data_mgr._check_pg()
    return jsonify({
        "success": True,
        "postgresql": data_mgr.pg_available,
        "mode": "PostgreSQL" if data_mgr.pg_available else "Excel",
        "message": "Connected to Neon PostgreSQL" if data_mgr.pg_available else "Using Excel fallback (offline mode)"
    })

# ══════════════════════════════════════════════════════════
#  STARTUP
# ══════════════════════════════════════════════════════════
def startup():
    """Initialize database and Excel."""
    init_excel()
    with app.app_context():
        try:
            db.create_all()
            # Seed default admin if no users
            if not User.query.first():
                admin = User(username='admin', role='Super Admin',
                             status='Active',
                             permissions=json.dumps(ROLE_DEFAULTS['Super Admin']))
                admin.set_password('admin123')
                db.session.add(admin)
            # Seed settings
            defaults = [
                ("School Name", SCHOOL_NAME),
                ("School Address", SCHOOL_ADDRESS),
                ("Academic Session", "2026-27"),
                ("Session Start Month", "4"),
                ("Currency", "INR"),
                ("Receipt Prefix", "BDPS"),
                ("Due Day", "10"),
                ("Grace Days", "0"),
                ("Late Fee Enabled", "Yes"),
                ("Late Fee Type", "Per Day"),
                ("Late Fee Amount", "50"),
                ("Late Fee Percent", "2"),
                ("Max Late Fee", "1000"),
                ("Max Login Attempts", "3"),
                ("Working Days Per Month", "26"),
                ("Salary LOP Deduction", "Yes"),
                ("Half Day Count", "0.5"),
                ("Allow Backdate Attendance", "No"),
            ]
            for k, v in defaults:
                if not Setting.query.filter_by(key=k).first():
                    db.session.add(Setting(key=k, value=str(v)))
            # Seed fee structure
            for cls in ["Nursery", "LKG", "UKG", "1", "2", "3", "4", "5",
                        "6", "7", "8", "9", "10", "11", "12"]:
                if not FeeStructure.query.filter_by(class_name=cls).first():
                    db.session.add(FeeStructure(class_name=cls))
            db.session.commit()
            data_mgr.pg_available = True
            print("✅ PostgreSQL connected. Database ready.")
            # Sync any existing Excel data
            result = data_mgr.sync_excel_to_pg()
            if result.get('success'):
                print(f"✅ Synced Excel → PostgreSQL: {result.get('counts')}")
        except Exception as e:
            print(f"⚠️  PostgreSQL not available: {e}")
            print("   Running in Excel-only mode.")
            data_mgr.pg_available = False

if __name__ == '__main__':
    startup()

    print("=" * 64)
    print(" BLUE DIAMOND PUBLIC SCHOOL - SCHOOL MANAGEMENT ERP")
    print("=" * 64)
    print(f" Database: {'PostgreSQL (Neon)' if data_mgr.pg_available else 'Excel (Offline)'}")
    print(" Starting Flask server...")
    print("=" * 64)

    import os

    port = int(os.environ.get("PORT", 5000))

    app.run(
        host="0.0.0.0",
        port=port,
        debug=False
    )