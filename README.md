# Academic Result Analysis Portal

Upload a semester result workbook and get SGPA / CGPA / rank, arrears, nil-arrear
students, faculty pass-rates and a formatted Excel report (one master sheet per
semester in the department template).

## Project structure

```
AcademicPortal/
├── app.py              Flask server + REST API (run this)
├── analysis.py         Result calculations + Excel/HTML report builders
├── database.py         SQLAlchemy models (SQLite)
├── requirements.txt    Python packages
├── run_app.bat         Windows launcher (no build needed)
├── run_app.sh          Mac/Linux launcher
├── build_exe.bat       Builds dist\AcademicResultPortal.exe (Windows)
├── build_exe.sh        Builds a single executable (Mac/Linux)
├── templates/
│   ├── index.html      Main page  "Registrar - Result Ledger"  (/)
│   └── dashboard.html  Alternative sidebar layout               (/dashboard)
├── sample_data/
│   └── CONSOLIDATED_TAURUS.xlsx   example input
├── uploads/            (created automatically) uploaded files
└── generated/          (created automatically) Excel / HTML reports
```

Both pages are **self-contained** (styles and scripts are inside the HTML), so
there is no `static/` folder to lose.

## Run it

**Windows:** double-click `run_app.bat`.
**Mac/Linux:** `./run_app.sh`
**Manual:**
```
pip install -r requirements.txt
python app.py
```
The browser opens at **http://127.0.0.1:5000**. Always use that address -
double-clicking `index.html` shows an unstyled/empty page because it needs the server.

## Make an .exe (Windows)

Install Python 3.9+, then double-click `build_exe.bat`.
Result: `dist\AcademicResultPortal.exe`. Uploads, reports and the database are saved next to it.

## Input formats accepted

1. **Department master-sheet workbook** (title rows, "Subject Code", "Credit", ARREAR/RESULT, SGPA) - like `sample_data/CONSOLIDATED_TAURUS.xlsx`.
2. **Plain workbook** - one sheet per semester: row 1 = subject credits,
   remaining rows = Reg No, Name, then one grade per subject.

Optional faculty workbook: columns `Subject Code`, `Faculty ID`, `Faculty Name` (one sheet per semester).

## Troubleshooting

| Problem | Fix |
|---|---|
| `TemplateNotFound` | `templates/` must sit next to `app.py` and contain `index.html` |
| Page has no styling | Open http://127.0.0.1:5000, not the .html file; keep internet on for fonts/charts |
| `ModuleNotFoundError` | Run `pip install -r requirements.txt` |
| Port 5000 busy | Change `PORT = 5000` at the bottom of `app.py` |
