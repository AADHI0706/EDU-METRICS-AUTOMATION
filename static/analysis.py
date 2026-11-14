"""
Core result-analysis logic.
Pure pandas/openpyxl, no Flask or DB imports here, so it stays independently
testable/reusable.
"""
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

GRADE_POINTS = {
    "O": 10, "A+": 9, "A": 8, "B+": 7, "B": 6, "C": 5,
    "RA": 0, "U": 0, "F": 0,
}
FAIL_GRADES = {"RA", "U", "F"}
DEFAULT_CREDIT = 3


def normalize_text(value: Any) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.lower() in {"", "nan", "na", "none"}:
        return ""
    return text


def _to_int(raw: Any) -> int:
    try:
        return int(raw)
    except (TypeError, ValueError):
        try:
            return int(float(raw))
        except (TypeError, ValueError):
            return DEFAULT_CREDIT


# --------------------------------------------------------------------------
# Loaders
# --------------------------------------------------------------------------
def load_result_file(path: Path) -> Tuple[Dict[str, pd.DataFrame], Dict[str, Dict[str, int]]]:
    workbook = pd.ExcelFile(path)
    semester_data: Dict[str, pd.DataFrame] = {}
    semester_credit: Dict[str, Dict[str, int]] = {}

    for sheet_name in workbook.sheet_names:
        df = pd.read_excel(workbook, sheet_name=sheet_name)
        df.columns = [normalize_text(c) for c in df.columns]
        df = df.dropna(how="all")
        if df.empty:
            continue

        credit_row = df.iloc[0]
        credits: Dict[str, int] = {}
        for subject in df.columns[2:]:
            credits[subject] = _to_int(credit_row.get(subject))

        semester_data[sheet_name] = df.iloc[1:].reset_index(drop=True)
        semester_credit[sheet_name] = credits

    return semester_data, semester_credit


def load_faculty_file(path: Path) -> Dict[str, pd.DataFrame]:
    workbook = pd.ExcelFile(path)
    faculty_data: Dict[str, pd.DataFrame] = {}

    for sheet_name in workbook.sheet_names:
        df = pd.read_excel(workbook, sheet_name=sheet_name)
        df.columns = [normalize_text(c) for c in df.columns]
        df = df.dropna(how="all")
        if df.empty:
            continue

        if "Subject Code" in df.columns:
            df["Subject Code"] = df["Subject Code"].apply(lambda v: normalize_text(v).upper())
        if "Faculty ID" in df.columns:
            df["Faculty ID"] = df["Faculty ID"].apply(normalize_text)
        if "Faculty Name" in df.columns:
            df["Faculty Name"] = df["Faculty Name"].apply(normalize_text)

        faculty_data[sheet_name] = df

    return faculty_data


def _find_row_with_keyword(raw: pd.DataFrame, keyword: str, start: int = 0) -> Optional[int]:
    keyword = keyword.lower()
    for i in range(start, len(raw)):
        for v in raw.iloc[i]:
            if isinstance(v, str) and keyword in v.lower():
                return i
    return None


def load_consolidated_workbook(
    path: Path,
) -> Tuple[Dict[str, pd.DataFrame], Dict[str, Dict[str, int]], Dict[str, str], str]:
    """
    Parses a workbook already laid out in the department "master sheet"
    template (title rows, "Subject Code" header, "Credit" row, ARREAR/RESULT,
    repeated codes for grade points, SGPA). Only register numbers, names and
    letter grades are read; points / SGPA / status are recomputed.

    Returns (semester_data, semester_credit, sem_labels, department_title).
    Returns empty dicts if the file is not in that layout.
    """
    workbook = pd.ExcelFile(path)
    semester_data: Dict[str, pd.DataFrame] = {}
    semester_credit: Dict[str, Dict[str, int]] = {}
    sem_labels: Dict[str, str] = {}
    title_candidates: List[str] = []

    for sheet_name in workbook.sheet_names:
        raw = pd.read_excel(workbook, sheet_name=sheet_name, header=None)
        if raw.empty:
            continue

        subj_row = _find_row_with_keyword(raw, "subject")
        if subj_row is None or subj_row + 1 >= len(raw):
            continue
        code_row = subj_row + 1

        dep_row = _find_row_with_keyword(raw, "department")
        if dep_row is not None:
            for v in raw.iloc[dep_row]:
                if isinstance(v, str) and "department" in v.lower():
                    title_candidates.append(v.strip())
                    break

        sub_row = _find_row_with_keyword(raw, "result analysis")
        if sub_row is not None:
            for v in raw.iloc[sub_row]:
                if isinstance(v, str) and "result analysis" in v.lower():
                    sem_labels[sheet_name] = v.strip()
                    break

        leading = None
        for j, v in enumerate(raw.iloc[subj_row]):
            if isinstance(v, str) and "subject" in v.lower():
                leading = j
                break
        if leading is None:
            continue

        reg_col = name_col = None
        for j in range(leading):
            v = raw.iloc[subj_row, j]
            if isinstance(v, str):
                vs = v.upper()
                if "REG" in vs:
                    reg_col = j
                elif "NAME" in vs:
                    name_col = j
        if reg_col is None or name_col is None:
            continue

        arrear_col = None
        for j, v in enumerate(raw.iloc[code_row]):
            if isinstance(v, str) and v.strip().upper() == "ARREAR":
                arrear_col = j
                break
        if arrear_col is None:
            continue

        credit_label_col = None
        for j, v in enumerate(raw.iloc[subj_row]):
            if isinstance(v, str) and v.strip().lower() == "credit":
                credit_label_col = j
                break

        grade_cols = list(range(leading, arrear_col))
        subjects = [normalize_text(raw.iloc[code_row, c]).upper() for c in grade_cols]
        pairs = [(s, c) for s, c in zip(subjects, grade_cols) if s]
        if not pairs:
            continue
        subjects = [s for s, _ in pairs]
        grade_cols = [c for _, c in pairs]

        credits: Dict[str, int] = {}
        for i, subject in enumerate(subjects):
            if credit_label_col is not None and credit_label_col + 1 + i < raw.shape[1]:
                credits[subject] = _to_int(raw.iloc[subj_row, credit_label_col + 1 + i])
            else:
                credits[subject] = DEFAULT_CREDIT

        records: List[Dict[str, Any]] = []
        for i in range(code_row + 1, len(raw)):
            reg_no = normalize_text(raw.iloc[i, reg_col])
            if not reg_no:
                continue
            record: Dict[str, Any] = {
                "Reg No": reg_no, "Name": normalize_text(raw.iloc[i, name_col]),
            }
            for subject, col in zip(subjects, grade_cols):
                record[subject] = normalize_text(raw.iloc[i, col]).upper()
            records.append(record)

        if records:
            semester_data[sheet_name] = pd.DataFrame(records)
            semester_credit[sheet_name] = credits

    title = max(set(title_candidates), key=title_candidates.count) if title_candidates else ""
    return semester_data, semester_credit, sem_labels, title


# --------------------------------------------------------------------------
# Calculations
# --------------------------------------------------------------------------
def calculate_sgpa(
    student_df: pd.DataFrame, credits: Dict[str, int]
) -> Tuple[pd.DataFrame, List[Dict[str, Any]]]:
    records: List[Dict[str, Any]] = []
    arrear_list: List[Dict[str, Any]] = []
    subjects = list(student_df.columns[2:])

    for _, row in student_df.iterrows():
        reg_no = normalize_text(row.get(student_df.columns[0]))
        name = normalize_text(row.get(student_df.columns[1]))
        total_points = 0
        total_credits = 0
        arrears = 0

        for subject in subjects:
            grade = normalize_text(row.get(subject)).upper()
            if not grade:
                continue

            credit = credits.get(subject, DEFAULT_CREDIT)
            total_points += GRADE_POINTS.get(grade, 0) * credit
            total_credits += credit

            if grade in FAIL_GRADES:
                arrears += 1
                arrear_list.append(
                    {"Reg No": reg_no, "Name": name, "Subject": subject, "Grade": grade}
                )

        sgpa = round(total_points / total_credits, 2) if total_credits else 0.0
        records.append({"Reg No": reg_no, "Name": name, "SGPA": sgpa, "Arrear": arrears})

    return pd.DataFrame(records), arrear_list


def build_subject_history(semester_data: Dict[str, pd.DataFrame]) -> Dict[str, Dict[str, Any]]:
    history: Dict[str, Dict[str, Any]] = {}

    for df in semester_data.values():
        subjects = list(df.columns[2:])
        for _, row in df.iterrows():
            reg_no = normalize_text(row.get(df.columns[0]))
            name = normalize_text(row.get(df.columns[1]))
            if not reg_no:
                continue
            stored = history.setdefault(reg_no, {"Name": name, "subjects": {}})
            for subject in subjects:
                grade = normalize_text(row.get(subject)).upper()
                if grade:
                    stored["subjects"].setdefault(subject, []).append(grade)

    return history


def generate_staff_data(
    semester_data: Dict[str, pd.DataFrame],
    faculty_data: Optional[Dict[str, pd.DataFrame]] = None,
) -> Tuple[Dict[str, pd.DataFrame], pd.DataFrame]:
    semester_staff_data: Dict[str, pd.DataFrame] = {}

    for sem_name, df in semester_data.items():
        rows: List[Dict[str, Any]] = []
        subjects = list(df.columns[2:])
        faculty_sheet = faculty_data.get(sem_name) if faculty_data else None

        for subject in subjects:
            total = passed = failed = 0
            for raw_grade in df[subject]:
                grade = normalize_text(raw_grade).upper()
                if not grade:
                    continue
                total += 1
                if grade in FAIL_GRADES:
                    failed += 1
                else:
                    passed += 1

            percent = round((passed / total) * 100, 2) if total else 0.0
            faculty_id = faculty_name = "N/A"

            if faculty_sheet is not None and "Subject Code" in faculty_sheet.columns:
                match = faculty_sheet[faculty_sheet["Subject Code"] == subject.upper()]
                if not match.empty:
                    faculty_id = normalize_text(match.iloc[0].get("Faculty ID"))
                    faculty_name = normalize_text(match.iloc[0].get("Faculty Name"))

            rows.append({
                "Subject": subject, "Faculty ID": faculty_id, "Faculty Name": faculty_name,
                "Total Students": total, "Pass": passed, "Fail": failed, "Pass %": percent,
            })

        semester_staff_data[sem_name] = pd.DataFrame(rows)

    staff_df = (
        pd.concat(
            [frame.assign(Semester=name) for name, frame in semester_staff_data.items()],
            ignore_index=True,
        ) if semester_staff_data else pd.DataFrame()
    )
    return semester_staff_data, staff_df


def analyze_results(
    semester_data: Dict[str, pd.DataFrame],
    semester_credit: Dict[str, Dict[str, int]],
    faculty_data: Optional[Dict[str, pd.DataFrame]] = None,
) -> Dict[str, pd.DataFrame]:
    sgpa_records: List[pd.DataFrame] = []
    arrear_rows: List[Dict[str, Any]] = []

    for sem_name, df in semester_data.items():
        credits = semester_credit.get(sem_name, {})
        sgpa_df, arrears = calculate_sgpa(df, credits)
        sgpa_df["Semester"] = sem_name
        sgpa_records.append(sgpa_df)
        arrear_rows.extend(arrears)

    sgpa_all = (
        pd.concat(sgpa_records, ignore_index=True) if sgpa_records
        else pd.DataFrame(columns=["Reg No", "Name", "SGPA", "Arrear", "Semester"])
    )

    sgpa_table = (
        sgpa_all.pivot_table(index=["Reg No", "Name"], columns="Semester", values="SGPA")
        .reset_index().rename_axis(None, axis=1)
    )

    score_columns = [c for c in sgpa_table.columns if c not in {"Reg No", "Name"}]
    sgpa_table["CGPA"] = (
        sgpa_table[score_columns].mean(axis=1).round(2) if score_columns else 0.0
    )

    if not sgpa_table.empty:
        sgpa_table = sgpa_table.sort_values(by="CGPA", ascending=False)
        sgpa_table["Rank"] = sgpa_table["CGPA"].rank(method="dense", ascending=False).astype(int)
    else:
        sgpa_table["Rank"] = pd.Series(dtype=int)

    arrear_df = pd.DataFrame(arrear_rows)
    if arrear_df.empty:
        arrear_df = pd.DataFrame(columns=["Reg No", "Name", "Subject", "Grade"])

    arrear_count = (
        sgpa_all.groupby(["Reg No", "Name"])["Arrear"].sum().reset_index()
        if not sgpa_all.empty else pd.DataFrame(columns=["Reg No", "Name", "Arrear"])
    )
    nil_arrear = arrear_count[arrear_count["Arrear"] == 0].reset_index(drop=True)

    history = build_subject_history(semester_data)
    history_rows: List[Dict[str, Any]] = []
    clearance_rows: List[Dict[str, Any]] = []

    for reg_no, record in history.items():
        subjects = record["subjects"]
        if not any(g in FAIL_GRADES for grades in subjects.values() for g in grades):
            history_rows.append({
                "Reg No": reg_no, "Name": record["Name"],
                "Status": "No Arrear in Entire History",
            })

        for subject, grades in subjects.items():
            saw_failure = False
            for grade in grades:
                if grade in FAIL_GRADES:
                    saw_failure = True
                elif saw_failure:
                    clearance_rows.append({
                        "Reg No": reg_no, "Name": record["Name"],
                        "Subject": subject, "Status": "Cleared",
                    })
                    break

    semester_staff_data, staff_df = generate_staff_data(semester_data, faculty_data)

    results = {
        "SGPA_CGPA": sgpa_table,
        "Arrear_List": arrear_df,
        "Arrear_Count": arrear_count,
        "Nil_Arrear": nil_arrear,
        "History_of_Arrear": pd.DataFrame(history_rows),
        "Arrear_Clearance": pd.DataFrame(clearance_rows),
        "Staff_Percentage": staff_df,
    }
    results.update({f"Faculty_{name}": frame for name, frame in semester_staff_data.items()})
    return results


# --------------------------------------------------------------------------
# Excel output: department "master sheet" template per semester
# --------------------------------------------------------------------------
def _sanitize_sheet_name(name: str) -> str:
    return "".join(c if c not in r"[]:*?/\\" else "_" for c in name)[:31] or "Sheet"


def _unique_sheet_name(name: str, used: set) -> str:
    base = _sanitize_sheet_name(name)
    candidate, suffix = base, 1
    while candidate in used:
        candidate = f"{base[: max(0, 31 - len(str(suffix)) - 1)]}_{suffix}"
        suffix += 1
    used.add(candidate)
    return candidate


def write_master_sheet(
    ws: Worksheet,
    sem_name: str,
    df: pd.DataFrame,
    credits: Dict[str, int],
    department_title: str = "DEPARTMENT OF CSE",
    batch_range: str = "",
    subtitle_override: Optional[str] = None,
) -> None:
    """
    Layout (matches the department template):
      row1  department / batch title (merged)
      row2  "<SEM> RESULT ANALYSIS(<years>)" (merged)
      row3  S.NO | REG.NO | NAME | Subject Code | Credit | per-subject credits | total credit
      row4  subject codes | ARREAR | RESULT | subject codes | SGPA
      row5+ letter grades | arrear count | PASS/FAIL | grade points | SGPA
    """
    subjects = list(df.columns[2:])
    n = len(subjects)
    if n == 0:
        ws.cell(row=1, column=1, value=f"No subject columns found for {sem_name}")
        return

    total_cols = 3 + n + 2 + n + 1

    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    bold = Font(bold=True)
    blue_bold = Font(bold=True, color="0000CC")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    c = ws.cell(row=1, column=1, value=department_title)
    c.font, c.alignment = Font(bold=True, size=13), center

    if subtitle_override:
        subtitle = subtitle_override
    else:
        subtitle = f"{sem_name} RESULT ANALYSIS" + (f"({batch_range})" if batch_range else "")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=total_cols)
    c = ws.cell(row=2, column=1, value=subtitle)
    c.font, c.alignment = Font(bold=True, size=11), center

    h1, h2 = 3, 4

    for col, label in enumerate(["S.NO", "REG.NO", "NAME"], start=1):
        ws.merge_cells(start_row=h1, start_column=col, end_row=h2, end_column=col)
        c = ws.cell(row=h1, column=col, value=label)
        c.font, c.alignment = bold, center

    grade_start, grade_end = 4, 3 + n
    ws.merge_cells(start_row=h1, start_column=grade_start, end_row=h1, end_column=grade_end)
    c = ws.cell(row=h1, column=grade_start, value="Subject Code")
    c.font, c.alignment = bold, center
    for i, subject in enumerate(subjects):
        c = ws.cell(row=h2, column=grade_start + i, value=subject)
        c.font, c.alignment = bold, center

    arrear_col, result_col = grade_end + 1, grade_end + 2
    ws.merge_cells(start_row=h1, start_column=arrear_col, end_row=h1, end_column=result_col)
    c = ws.cell(row=h1, column=arrear_col, value="Credit")
    c.font, c.alignment = blue_bold, center
    for col, label in ((arrear_col, "ARREAR"), (result_col, "RESULT")):
        c = ws.cell(row=h2, column=col, value=label)
        c.font, c.alignment = bold, center

    points_start = result_col + 1
    total_credit = 0
    for i, subject in enumerate(subjects):
        credit_val = credits.get(subject, DEFAULT_CREDIT)
        total_credit += credit_val
        c = ws.cell(row=h1, column=points_start + i, value=credit_val)
        c.font, c.alignment = blue_bold, center
        c = ws.cell(row=h2, column=points_start + i, value=subject)
        c.font, c.alignment = bold, center

    sgpa_col = points_start + n
    c = ws.cell(row=h1, column=sgpa_col, value=total_credit)
    c.font, c.alignment = blue_bold, center
    c = ws.cell(row=h2, column=sgpa_col, value="SGPA")
    c.font, c.alignment = bold, center

    for r in (h1, h2):
        for col in range(1, total_cols + 1):
            ws.cell(row=r, column=col).border = border

    row, serial = h2 + 1, 1
    for _, record in df.iterrows():
        reg_no = normalize_text(record.get(df.columns[0]))
        name = normalize_text(record.get(df.columns[1]))
        if not reg_no:
            continue

        ws.cell(row=row, column=1, value=serial)
        ws.cell(row=row, column=2, value=reg_no)
        ws.cell(row=row, column=3, value=name)

        total_points = total_earned = arrear_count = 0
        for i, subject in enumerate(subjects):
            grade = normalize_text(record.get(subject)).upper()
            gp = GRADE_POINTS.get(grade, 0)
            ws.cell(row=row, column=grade_start + i, value=grade)
            ws.cell(row=row, column=points_start + i, value=gp)
            if grade:
                credit = credits.get(subject, DEFAULT_CREDIT)
                total_points += gp * credit
                total_earned += credit
                if grade in FAIL_GRADES:
                    arrear_count += 1

        ws.cell(row=row, column=arrear_col, value=arrear_count)
        ws.cell(row=row, column=result_col, value="PASS" if arrear_count == 0 else "FAIL")
        ws.cell(row=row, column=sgpa_col,
                value=round(total_points / total_earned, 2) if total_earned else 0.0)

        for col in range(1, total_cols + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = border
            cell.alignment = Alignment(
                horizontal="left" if col == 3 else "center", vertical="center")
        row += 1
        serial += 1

    ws.column_dimensions[get_column_letter(1)].width = 6
    ws.column_dimensions[get_column_letter(2)].width = 16
    ws.column_dimensions[get_column_letter(3)].width = 26
    for col in range(4, total_cols + 1):
        ws.column_dimensions[get_column_letter(col)].width = 9
    ws.freeze_panes = ws.cell(row=h2 + 1, column=4).coordinate


def write_excel_report(
    results: Dict[str, pd.DataFrame],
    path: Path,
    semester_data: Optional[Dict[str, pd.DataFrame]] = None,
    semester_credit: Optional[Dict[str, Dict[str, int]]] = None,
    department_title: str = "DEPARTMENT OF CSE",
    batch_range: str = "",
    sem_labels: Optional[Dict[str, str]] = None,
) -> None:
    """One '<Semester>_Master' sheet per semester in the department template,
    followed by the plain summary sheets."""
    used: set = set()

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        book = writer.book
        if "Sheet" in book.sheetnames:
            default = book["Sheet"]
            if default.max_row == 1 and default.max_column == 1 and default["A1"].value is None:
                book.remove(default)

        if semester_data:
            for sem_name, df in semester_data.items():
                ws = book.create_sheet(title=_unique_sheet_name(f"{sem_name}_Master", used))
                write_master_sheet(
                    ws, sem_name, df, (semester_credit or {}).get(sem_name, {}),
                    department_title, batch_range,
                    subtitle_override=(sem_labels or {}).get(sem_name),
                )

        for sheet_name, frame in results.items():
            frame.to_excel(writer, sheet_name=_unique_sheet_name(sheet_name, used), index=False)


def build_html_report(results: Dict[str, pd.DataFrame]) -> str:
    parts: List[str] = [
        "<html><head><meta charset='utf-8'><title>Academic Result Analysis</title>",
        "<style>body{font-family:Arial,sans-serif;margin:20px;}",
        "table{border-collapse:collapse;width:100%;margin-bottom:24px;}",
        "th,td{border:1px solid #ddd;padding:8px;text-align:left;}",
        "th{background-color:#4CAF50;color:white;}",
        "tr:nth-child(even){background-color:#f2f2f2;}h1,h2{color:#333;}</style></head><body>",
        "<h1>Academic Result Analysis</h1>",
    ]
    for sheet_name, frame in results.items():
        parts.append(f"<h2>{sheet_name.replace('_', ' ')}</h2>")
        parts.append(frame.to_html(index=False, classes="dataframe"))
    parts.append("</body></html>")
    return "\n".join(parts)
