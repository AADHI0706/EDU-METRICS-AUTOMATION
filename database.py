"""
SQL database models (SQLAlchemy ORM, backed by SQLite by default).
Swap SQLALCHEMY_DATABASE_URI in app.py to point at Postgres/MySQL instead —
nothing else in this file needs to change.
"""
from datetime import datetime
from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()


class Report(db.Model):
    """One uploaded result-file (+ optional faculty-file) processing run."""
    __tablename__ = "reports"

    id = db.Column(db.Integer, primary_key=True)
    result_filename = db.Column(db.String(255), nullable=False)
    faculty_filename = db.Column(db.String(255))
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)
    student_count = db.Column(db.Integer, default=0)
    semester_count = db.Column(db.Integer, default=0)
    excel_path = db.Column(db.String(500))
    html_path = db.Column(db.String(500))


class Student(db.Model):
    __tablename__ = "students"

    reg_no = db.Column(db.String(64), primary_key=True)
    name = db.Column(db.String(255))


class Semester(db.Model):
    __tablename__ = "semesters"

    id = db.Column(db.Integer, primary_key=True)
    report_id = db.Column(db.Integer, db.ForeignKey("reports.id"), nullable=False)
    name = db.Column(db.String(128), nullable=False)


class Subject(db.Model):
    __tablename__ = "subjects"

    id = db.Column(db.Integer, primary_key=True)
    semester_id = db.Column(db.Integer, db.ForeignKey("semesters.id"), nullable=False)
    code = db.Column(db.String(128), nullable=False)
    credit = db.Column(db.Integer, default=3)


class Grade(db.Model):
    __tablename__ = "grades"

    id = db.Column(db.Integer, primary_key=True)
    report_id = db.Column(db.Integer, db.ForeignKey("reports.id"), nullable=False)
    reg_no = db.Column(db.String(64), db.ForeignKey("students.reg_no"), nullable=False)
    semester_id = db.Column(db.Integer, db.ForeignKey("semesters.id"), nullable=False)
    subject_code = db.Column(db.String(128), nullable=False)
    grade = db.Column(db.String(8))
    credit = db.Column(db.Integer)
    grade_point = db.Column(db.Integer)
    is_arrear = db.Column(db.Boolean, default=False)


class Faculty(db.Model):
    __tablename__ = "faculty"

    id = db.Column(db.Integer, primary_key=True)
    semester_id = db.Column(db.Integer, db.ForeignKey("semesters.id"), nullable=False)
    subject_code = db.Column(db.String(128), nullable=False)
    faculty_id = db.Column(db.String(64))
    faculty_name = db.Column(db.String(255))


class SgpaResult(db.Model):
    __tablename__ = "sgpa_results"

    id = db.Column(db.Integer, primary_key=True)
    report_id = db.Column(db.Integer, db.ForeignKey("reports.id"), nullable=False)
    reg_no = db.Column(db.String(64), db.ForeignKey("students.reg_no"), nullable=False)
    semester_id = db.Column(db.Integer, db.ForeignKey("semesters.id"), nullable=False)
    sgpa = db.Column(db.Float)
    arrear_count = db.Column(db.Integer, default=0)


class CgpaResult(db.Model):
    __tablename__ = "cgpa_results"

    id = db.Column(db.Integer, primary_key=True)
    report_id = db.Column(db.Integer, db.ForeignKey("reports.id"), nullable=False)
    reg_no = db.Column(db.String(64), db.ForeignKey("students.reg_no"), nullable=False)
    cgpa = db.Column(db.Float)
    rank = db.Column(db.Integer)


class StaffPerformance(db.Model):
    __tablename__ = "staff_performance"

    id = db.Column(db.Integer, primary_key=True)
    report_id = db.Column(db.Integer, db.ForeignKey("reports.id"), nullable=False)
    semester_id = db.Column(db.Integer, db.ForeignKey("semesters.id"), nullable=False)
    subject_code = db.Column(db.String(128))
    faculty_id = db.Column(db.String(64))
    faculty_name = db.Column(db.String(255))
    total_students = db.Column(db.Integer)
    pass_count = db.Column(db.Integer)
    fail_count = db.Column(db.Integer)
    pass_percent = db.Column(db.Float)
