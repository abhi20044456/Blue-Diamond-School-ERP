from datetime import datetime, date
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash

db = SQLAlchemy()

class User(db.Model):
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False, index=True)
    password_hash = db.Column(db.String(256), nullable=False)
    role = db.Column(db.String(40), default='Data Entry')
    status = db.Column(db.String(20), default='Active')
    failed_attempts = db.Column(db.Integer, default=0)
    locked = db.Column(db.Boolean, default=False)
    permissions = db.Column(db.Text, default='{}')  # JSON string
    # Security questions (hashed answers)
    security_q1 = db.Column(db.String(255), nullable=True)
    security_a1 = db.Column(db.String(256), nullable=True)
    security_q2 = db.Column(db.String(255), nullable=True)
    security_a2 = db.Column(db.String(256), nullable=True)
    security_q3 = db.Column(db.String(255), nullable=True)
    security_a3 = db.Column(db.String(256), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    def set_security_answer(self, answer):
        return generate_password_hash(answer.strip().lower())

    def verify_security_answer(self, answer, hashed):
        return check_password_hash(hashed, answer.strip().lower())

    def to_dict(self):
        import json
        try:
            perms = json.loads(self.permissions or '{}')
        except:
            perms = {}
        return {
            'id': self.id,
            'username': self.username,
            'role': self.role,
            'status': self.status,
            'failed_attempts': self.failed_attempts,
            'locked': self.locked,
            'permissions': perms,
            'created_at': self.created_at.strftime('%Y-%m-%d %H:%M:%S') if self.created_at else '',
            'updated_at': self.updated_at.strftime('%Y-%m-%d %H:%M:%S') if self.updated_at else '',
            'has_security': bool(self.security_q1)
        }


class Student(db.Model):
    __tablename__ = 'students'
    id = db.Column(db.Integer, primary_key=True)
    admission_no = db.Column(db.String(50), unique=True, nullable=False, index=True)
    student_name = db.Column(db.String(150), nullable=False)
    father_name = db.Column(db.String(150))
    mother_name = db.Column(db.String(150))
    class_name = db.Column(db.String(50))
    section = db.Column(db.String(20))
    roll_no = db.Column(db.String(20))
    dob = db.Column(db.String(20))
    gender = db.Column(db.String(10))
    mobile = db.Column(db.String(20))
    address = db.Column(db.Text)
    admission_date = db.Column(db.String(20))  # YYYY-MM-DD
    session = db.Column(db.String(20), default='2026-27')
    transport = db.Column(db.String(10), default='No')
    status = db.Column(db.String(20), default='Active')
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    def to_dict(self):
        return {
            'Admission No': self.admission_no,
            'Student Name': self.student_name,
            'Father Name': self.father_name or '',
            'Mother Name': self.mother_name or '',
            'Class': self.class_name or '',
            'Section': self.section or '',
            'Roll No': self.roll_no or '',
            'DOB': self.dob or '',
            'Gender': self.gender or '',
            'Mobile': self.mobile or '',
            'Address': self.address or '',
            'Admission Date': self.admission_date or '',
            'Session': self.session or '2026-27',
            'Transport': self.transport or 'No',
            'Status': self.status or 'Active',
            'Created At': self.created_at.strftime('%Y-%m-%d %H:%M:%S') if self.created_at else '',
            'Updated At': self.updated_at.strftime('%Y-%m-%d %H:%M:%S') if self.updated_at else ''
        }


class FeeTransaction(db.Model):
    __tablename__ = 'fee_transactions'
    id = db.Column(db.Integer, primary_key=True)
    receipt_no = db.Column(db.String(50), unique=True, nullable=False, index=True)
    date = db.Column(db.String(20))
    time = db.Column(db.String(20))
    admission_no = db.Column(db.String(50), index=True)
    student_name = db.Column(db.String(150))
    class_name = db.Column(db.String(50))
    section = db.Column(db.String(20))
    fee_type = db.Column(db.String(50))
    month = db.Column(db.String(20))
    amount = db.Column(db.Float, default=0)
    late_fee = db.Column(db.Float, default=0)
    payment_mode = db.Column(db.String(30))
    transaction_no = db.Column(db.String(100))
    remarks = db.Column(db.Text)
    received_by = db.Column(db.String(80))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {
            'Receipt No': self.receipt_no,
            'Date': self.date or '',
            'Time': self.time or '',
            'Admission No': self.admission_no or '',
            'Student Name': self.student_name or '',
            'Class': self.class_name or '',
            'Section': self.section or '',
            'Fee Type': self.fee_type or '',
            'Month': self.month or '',
            'Amount': self.amount or 0,
            'Late Fee': self.late_fee or 0,
            'Payment Mode': self.payment_mode or '',
            'Transaction No': self.transaction_no or '',
            'Remarks': self.remarks or '',
            'Received By': self.received_by or ''
        }


class FeeStructure(db.Model):
    __tablename__ = 'fee_structure'
    id = db.Column(db.Integer, primary_key=True)
    class_name = db.Column(db.String(50), unique=True, nullable=False)
    admission_fee = db.Column(db.Float, default=0)
    tuition_fee = db.Column(db.Float, default=0)
    transport_fee = db.Column(db.Float, default=0)
    exam_fee = db.Column(db.Float, default=0)
    annual_fee = db.Column(db.Float, default=0)
    other_fee = db.Column(db.Float, default=0)

    def to_dict(self):
        return {
            'Class': self.class_name,
            'Admission Fee': self.admission_fee or 0,
            'Tuition Fee': self.tuition_fee or 0,
            'Transport Fee': self.transport_fee or 0,
            'Exam Fee': self.exam_fee or 0,
            'Annual Fee': self.annual_fee or 0,
            'Other Fee': self.other_fee or 0
        }


class Staff(db.Model):
    __tablename__ = 'staff'
    id = db.Column(db.Integer, primary_key=True)
    staff_id = db.Column(db.String(50), unique=True, nullable=False, index=True)
    name = db.Column(db.String(150), nullable=False)
    designation = db.Column(db.String(100))
    department = db.Column(db.String(100))
    mobile = db.Column(db.String(20))
    joining_date = db.Column(db.String(20))
    basic_salary = db.Column(db.Float, default=0)
    allowances = db.Column(db.Float, default=0)
    status = db.Column(db.String(20), default='Active')
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    def to_dict(self):
        return {
            'Staff ID': self.staff_id,
            'Name': self.name,
            'Designation': self.designation or '',
            'Department': self.department or '',
            'Mobile': self.mobile or '',
            'Joining Date': self.joining_date or '',
            'Basic Salary': self.basic_salary or 0,
            'Allowances': self.allowances or 0,
            'Status': self.status or 'Active',
            'Created At': self.created_at.strftime('%Y-%m-%d %H:%M:%S') if self.created_at else '',
            'Updated At': self.updated_at.strftime('%Y-%m-%d %H:%M:%S') if self.updated_at else ''
        }


class Attendance(db.Model):
    __tablename__ = 'staff_attendance'
    id = db.Column(db.Integer, primary_key=True)
    date = db.Column(db.String(20), nullable=False, index=True)
    staff_id = db.Column(db.String(50), nullable=False, index=True)
    staff_name = db.Column(db.String(150))
    status = db.Column(db.String(20), default='Present')
    remarks = db.Column(db.Text)
    marked_by = db.Column(db.String(80))
    approved_by = db.Column(db.String(80), nullable=True)
    approval_status = db.Column(db.String(20), default='Approved')  # Pending/Approved/Rejected
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    def to_dict(self):
        return {
            'Date': self.date,
            'Staff ID': self.staff_id,
            'Staff Name': self.staff_name or '',
            'Status': self.status or '',
            'Remarks': self.remarks or '',
            'Marked By': self.marked_by or '',
            'Approved By': self.approved_by or '',
            'Approval Status': self.approval_status or 'Approved'
        }


class SalaryRecord(db.Model):
    __tablename__ = 'salary_records'
    id = db.Column(db.Integer, primary_key=True)
    month = db.Column(db.String(20), nullable=False, index=True)
    staff_id = db.Column(db.String(50), nullable=False)
    staff_name = db.Column(db.String(150))
    working_days = db.Column(db.Integer, default=26)
    present = db.Column(db.Integer, default=0)
    absent = db.Column(db.Integer, default=0)
    leave = db.Column(db.Integer, default=0)
    half_day = db.Column(db.Integer, default=0)
    lop = db.Column(db.Float, default=0)
    basic_salary = db.Column(db.Float, default=0)
    allowances = db.Column(db.Float, default=0)
    deductions = db.Column(db.Float, default=0)
    advance = db.Column(db.Float, default=0)
    net_salary = db.Column(db.Float, default=0)
    generated_at = db.Column(db.DateTime, default=datetime.utcnow)
    generated_by = db.Column(db.String(80))

    def to_dict(self):
        return {
            'Month': self.month,
            'Staff ID': self.staff_id,
            'Staff Name': self.staff_name or '',
            'Working Days': self.working_days,
            'Present': self.present,
            'Absent': self.absent,
            'Leave': self.leave,
            'Half Day': self.half_day,
            'LOP': self.lop,
            'Basic Salary': self.basic_salary,
            'Allowances': self.allowances,
            'Deductions': self.deductions,
            'Advance': self.advance,
            'Net Salary': self.net_salary,
            'Generated At': self.generated_at.strftime('%Y-%m-%d %H:%M:%S') if self.generated_at else '',
            'Generated By': self.generated_by or ''
        }


class AuditLog(db.Model):
    __tablename__ = 'audit_log'
    id = db.Column(db.Integer, primary_key=True)
    date = db.Column(db.String(20))
    time = db.Column(db.String(20))
    username = db.Column(db.String(80))
    action = db.Column(db.String(80))
    details = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {
            'Date': self.date or '',
            'Time': self.time or '',
            'Username': self.username or '',
            'Action': self.action or '',
            'Details': self.details or ''
        }


class Setting(db.Model):
    __tablename__ = 'settings'
    id = db.Column(db.Integer, primary_key=True)
    key = db.Column(db.String(100), unique=True, nullable=False)
    value = db.Column(db.Text)

    def to_dict(self):
        return {'Key': self.key, 'Value': self.value}


class SyncQueue(db.Model):
    """Queue for Excel data waiting to be synced to PostgreSQL"""
    __tablename__ = 'sync_queue'
    id = db.Column(db.Integer, primary_key=True)
    table_name = db.Column(db.String(50), nullable=False)
    record_data = db.Column(db.Text, nullable=False)  # JSON
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    synced = db.Column(db.Boolean, default=False)