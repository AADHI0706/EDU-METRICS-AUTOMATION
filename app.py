"""
Academic Result Analysis Portal — Flask backend.

Wires together:
  - analysis.py   (pandas logic, unchanged from the original script)
  - database.py   (SQLAlchemy models / SQL persistence)
  - templates/, static/ (frontend: HTML + vanilla JS + Bootstrap + Chart.js)

Run:
    python app.py
Then open http://localhost:5000
"""
import math
import sys
import threading
import webbrowser
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.utils import secure_filename

import analysis
from database import (CgpaResult, Faculty, Grade, Report, Semester,
                       SgpaResult, StaffPerformance, Student, Subject, db)

# When packaged with PyInstaller, bundled files (templates) live in sys._MEIPASS
# while user data (uploads, reports, database) must sit next to the .exe.
if getattr(sys, "frozen", False):
    RES_DIR = Path(sys._MEIPASS)
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    RES_DIR = BASE_DIR = Path(__file__).resolve().parent

UPLOAD_DIR = BASE_DIR / "uploads"
GENERATED_DIR = BASE_DIR / "generated"
UPLOAD_DIR.mkdir(exist_ok=True)
GENERATED_DIR.mkdir(exist_ok=True)

app = Flask(__name__, template_folder=str(RES_DIR / "templates"))
app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{BASE_DIR / 'academic_portal.db'}"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB uploads
db.init_app(app)

with app.app_context():
    db.create_all()


def clean(value):
    """Make a value JSON-safe (NaN/None -> None)."""
    if isinstance(value, float) and math.isnan(value):
        return None
    return value


def df_to_records(df):
    return [{k: clean(v) for k, v in row.items()} for row in df.to_dict(orient="records")]


# --------------------------------------------------------------------------
# Persistence: push the pandas results of one analysis run into SQL tables
# --------------------------------------------------------------------------
def persist_to_db(report, semester_data, semester_credit, faculty_data, results):
    semester_ids = {}
    for sem_name, df in semester_data.items():
        sem = Semester(report_id=report.id, name=sem_name)
        db.session.add(sem)
        db.session.flush()
        semester_ids[sem_name] = sem.id

        credits = semester_credit.get(sem_name, {})
        for subject_code, credit in credits.items():
            db.session.add(Subject(semester_id=sem.id, code=subject_code, credit=credit))

        subjects = list(df.columns[2:])
        for _, row in df.iterrows():
            reg_no = analysis.normalize_text(row.get(df.columns[0]))
            name = analysis.normalize_text(row.get(df.columns[1]))
            if not reg_no:
                continue
            if not db.session.get(Student, reg_no):
                db.session.add(Student(reg_no=reg_no, name=name))
                db.session.flush()

            for subject in subjects:
                grade = analysis.normalize_text(row.get(subject)).upper()
                if not grade:
                    continue
                credit = credits.get(subject, analysis.DEFAULT_CREDIT)
                db.session.add(Grade(
                    report_id=report.id, reg_no=reg_no, semester_id=sem.id,
                    subject_code=subject, grade=grade, credit=credit,
                    grade_point=analysis.GRADE_POINTS.get(grade, 0),
                    is_arrear=grade in analysis.FAIL_GRADES,
                ))

        if faculty_data and sem_name in faculty_data:
            fdf = faculty_data[sem_name]
            for _, frow in fdf.iterrows():
                db.session.add(Faculty(
                    semester_id=sem.id,
                    subject_code=analysis.normalize_text(frow.get("Subject Code")).upper(),
                    faculty_id=analysis.normalize_text(frow.get("Faculty ID")),
                    faculty_name=analysis.normalize_text(frow.get("Faculty Name")),
                ))

    # SGPA per student per semester (from Arrear_Count-less per-semester calc)
    for sem_name, df in semester_data.items():
        credits = semester_credit.get(sem_name, {})
        sgpa_df, _ = analysis.calculate_sgpa(df, credits)
        sem_id = semester_ids[sem_name]
        for _, row in sgpa_df.iterrows():
            reg_no = row["Reg No"]
            if not reg_no:
                continue
            db.session.add(SgpaResult(
                report_id=report.id, reg_no=reg_no, semester_id=sem_id,
                sgpa=row["SGPA"], arrear_count=int(row["Arrear"]),
            ))

    # CGPA + rank
    sgpa_table = results.get("SGPA_CGPA")
    if sgpa_table is not None:
        for _, row in sgpa_table.iterrows():
            reg_no = row.get("Reg No")
            if not reg_no:
                continue
            db.session.add(CgpaResult(
                report_id=report.id, reg_no=reg_no,
                cgpa=clean(row.get("CGPA")), rank=clean(row.get("Rank")),
            ))

    # Staff / subject-pass-percentage rows
    for sem_name, df in semester_data.items():
        sem_id = semester_ids[sem_name]
        staff_key = f"Faculty_{sem_name}"
        if staff_key in results:
            for _, row in results[staff_key].iterrows():
                db.session.add(StaffPerformance(
                    report_id=report.id, semester_id=sem_id,
                    subject_code=row.get("Subject"), faculty_id=row.get("Faculty ID"),
                    faculty_name=row.get("Faculty Name"),
                    total_students=int(row.get("Total Students", 0)),
                    pass_count=int(row.get("Pass", 0)), fail_count=int(row.get("Fail", 0)),
                    pass_percent=float(row.get("Pass %", 0)),
                ))

    db.session.commit()
    return semester_ids


# --------------------------------------------------------------------------
# Frontend
# --------------------------------------------------------------------------
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/dashboard")
def dashboard():
    # Alternative sidebar-style layout
    return render_template("dashboard.html")


# --------------------------------------------------------------------------
# API: upload & process
# --------------------------------------------------------------------------
@app.route("/api/upload", methods=["POST"])
def upload():
    result_file = request.files.get("result_file")
    faculty_file = request.files.get("faculty_file")

    if not result_file or result_file.filename == "":
        return jsonify({"error": "A result Excel file is required."}), 400

    result_name = secure_filename(result_file.filename)
    result_path = UPLOAD_DIR / result_name
    result_file.save(result_path)

    faculty_path = None
    faculty_name = None
    if faculty_file and faculty_file.filename:
        faculty_name = secure_filename(faculty_file.filename)
        faculty_path = UPLOAD_DIR / faculty_name
        faculty_file.save(faculty_path)

    try:
        # Try the department "master sheet" layout first, then the plain format.
        semester_data, semester_credit, sem_labels, dept_title = \
            analysis.load_consolidated_workbook(result_path)
        if not semester_data:
            semester_data, semester_credit = analysis.load_result_file(result_path)
            sem_labels, dept_title = {}, ""
        dept_title = (request.form.get("department_title") or dept_title
                      or "DEPARTMENT OF CSE")
        batch_range = request.form.get("batch_range", "")
        if not semester_data:
            return jsonify({"error": "No usable sheets found in the result file."}), 400

        faculty_data = analysis.load_faculty_file(faculty_path) if faculty_path else None
        results = analysis.analyze_results(semester_data, semester_credit, faculty_data)
    except Exception as exc:  # surface parsing errors to the frontend
        return jsonify({"error": f"Failed to process file(s): {exc}"}), 400

    report = Report(
        result_filename=result_name,
        faculty_filename=faculty_name,
        student_count=len(results.get("SGPA_CGPA", [])),
        semester_count=len(semester_data),
    )
    db.session.add(report)
    db.session.flush()  # get report.id before writing generated files

    excel_path = GENERATED_DIR / f"report_{report.id}.xlsx"
    html_path = GENERATED_DIR / f"report_{report.id}.html"
    analysis.write_excel_report(
        results, excel_path,
        semester_data=semester_data, semester_credit=semester_credit,
        department_title=dept_title, batch_range=batch_range,
        sem_labels=sem_labels,
    )
    html_path.write_text(analysis.build_html_report(results), encoding="utf-8")
    report.excel_path = str(excel_path)
    report.html_path = str(html_path)
    db.session.add(report)

    persist_to_db(report, semester_data, semester_credit, faculty_data, results)

    return jsonify({
        "report_id": report.id,
        "semesters": list(semester_data.keys()),
        "student_count": report.student_count,
        "semester_count": report.semester_count,
    })


# --------------------------------------------------------------------------
# API: read-back from SQL
# --------------------------------------------------------------------------
@app.route("/api/reports", methods=["GET"])
def list_reports():
    reports = Report.query.order_by(Report.uploaded_at.desc()).all()
    return jsonify([{
        "id": r.id, "result_filename": r.result_filename,
        "faculty_filename": r.faculty_filename,
        "uploaded_at": r.uploaded_at.isoformat() if r.uploaded_at else None,
        "student_count": r.student_count, "semester_count": r.semester_count,
    } for r in reports])


@app.route("/api/reports/<int:report_id>/sgpa", methods=["GET"])
def sgpa_cgpa(report_id):
    semesters = Semester.query.filter_by(report_id=report_id).order_by(Semester.id).all()
    sem_by_id = {s.id: s.name for s in semesters}
    sgpa_rows = SgpaResult.query.filter_by(report_id=report_id).all()
    cgpa_rows = {c.reg_no: c for c in CgpaResult.query.filter_by(report_id=report_id).all()}
    students = {s.reg_no: s.name for s in Student.query.all()}

    per_student = {}
    for row in sgpa_rows:
        entry = per_student.setdefault(row.reg_no, {
            "reg_no": row.reg_no, "name": students.get(row.reg_no, ""), "sgpa": {},
        })
        entry["sgpa"][sem_by_id.get(row.semester_id, "?")] = row.sgpa

    table = []
    for reg_no, entry in per_student.items():
        cg = cgpa_rows.get(reg_no)
        entry["cgpa"] = cg.cgpa if cg else None
        entry["rank"] = cg.rank if cg else None
        table.append(entry)

    table.sort(key=lambda r: (r["rank"] is None, r["rank"]))
    return jsonify({"semesters": [s.name for s in semesters], "students": table})


@app.route("/api/reports/<int:report_id>/arrears", methods=["GET"])
def arrears(report_id):
    grades = (Grade.query.filter_by(report_id=report_id, is_arrear=True)
              .join(Student, Grade.reg_no == Student.reg_no).all())
    sem_names = {s.id: s.name for s in Semester.query.filter_by(report_id=report_id).all()}
    students = {s.reg_no: s.name for s in Student.query.all()}
    return jsonify([{
        "reg_no": g.reg_no, "name": students.get(g.reg_no, ""),
        "semester": sem_names.get(g.semester_id, ""),
        "subject": g.subject_code, "grade": g.grade,
    } for g in grades])


@app.route("/api/reports/<int:report_id>/nil-arrear", methods=["GET"])
def nil_arrear(report_id):
    sgpa_rows = SgpaResult.query.filter_by(report_id=report_id).all()
    students = {s.reg_no: s.name for s in Student.query.all()}
    totals = {}
    for row in sgpa_rows:
        totals[row.reg_no] = totals.get(row.reg_no, 0) + (row.arrear_count or 0)
    nil = [{"reg_no": r, "name": students.get(r, "")} for r, total in totals.items() if total == 0]
    return jsonify(nil)


@app.route("/api/reports/<int:report_id>/staff", methods=["GET"])
def staff_performance(report_id):
    rows = StaffPerformance.query.filter_by(report_id=report_id).all()
    sem_names = {s.id: s.name for s in Semester.query.filter_by(report_id=report_id).all()}
    return jsonify([{
        "semester": sem_names.get(r.semester_id, ""), "subject": r.subject_code,
        "faculty_id": r.faculty_id, "faculty_name": r.faculty_name,
        "total_students": r.total_students, "pass": r.pass_count,
        "fail": r.fail_count, "pass_percent": r.pass_percent,
    } for r in rows])


@app.route("/api/reports/<int:report_id>/student/<reg_no>", methods=["GET"])
def student_detail(report_id, reg_no):
    student = db.session.get(Student, reg_no)
    if not student:
        return jsonify({"error": "Student not found."}), 404

    sem_names = {s.id: s.name for s in Semester.query.filter_by(report_id=report_id).all()}
    grades = Grade.query.filter_by(report_id=report_id, reg_no=reg_no).all()
    sgpa_rows = SgpaResult.query.filter_by(report_id=report_id, reg_no=reg_no).all()
    cgpa_row = CgpaResult.query.filter_by(report_id=report_id, reg_no=reg_no).first()

    by_semester = {}
    for g in grades:
        sem = sem_names.get(g.semester_id, "?")
        by_semester.setdefault(sem, []).append({
            "subject": g.subject_code, "grade": g.grade,
            "credit": g.credit, "is_arrear": g.is_arrear,
        })

    return jsonify({
        "reg_no": reg_no, "name": student.name,
        "cgpa": cgpa_row.cgpa if cgpa_row else None,
        "rank": cgpa_row.rank if cgpa_row else None,
        "sgpa": {sem_names.get(r.semester_id, "?"): r.sgpa for r in sgpa_rows},
        "subjects": by_semester,
    })


@app.route("/api/reports/<int:report_id>/download/<fmt>", methods=["GET"])
def download(report_id, fmt):
    report = Report.query.get_or_404(report_id)
    if fmt == "excel":
        return send_file(report.excel_path, as_attachment=True,
                          download_name=f"Academic_Report_{report_id}.xlsx")
    if fmt == "html":
        return send_file(report.html_path, as_attachment=True,
                          download_name=f"Academic_Report_{report_id}.html")
    return jsonify({"error": "Unknown format, use 'excel' or 'html'."}), 400


if __name__ == "__main__":
    PORT = 5000
    # Open the browser shortly after the server starts.
    threading.Timer(1.5, lambda: webbrowser.open(f"http://127.0.0.1:{PORT}")).start()
    # 127.0.0.1 = only this computer can open it (student data stays private).
    app.run(debug=False, host="127.0.0.1", port=PORT)
