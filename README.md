# Employee Audit Manager (Single File)

One Python file — easy to upload to GitHub and build a Windows `.exe`.

## Run from source

```bash
pip install -r requirements.txt
python EmployeeAuditManager.py
```

## Build Windows EXE (on a Windows PC)

```bat
pip install -r requirements.txt
pyinstaller --noconfirm --windowed --onefile --name EmployeeAuditManager --collect-all PySide6 --collect-all reportlab EmployeeAuditManager.py
```

Result: `dist\EmployeeAuditManager.exe`

## Build free with GitHub Actions

1. Create a new **public** GitHub repository.
2. Upload **only** these files:
   - `EmployeeAuditManager.py`
   - `requirements.txt`
   - `.github/workflows/build.yml` (optional — included)
3. Actions → **Build Windows EXE** → Run workflow.
4. Download the artifact when finished.

## Features

- HR/Admin PIN login  
- Employees (add/edit/activate/delete)  
- **Daily Audit (EOD)** — submitted, late, links, missing proof, completed, pending  
- Add / edit / remove **daily tasks per employee**  
- Daily Tasks (multi-assign, copy previous day)  
- Attendance  
- Dashboard  
- Print Center (EOD PDF, CSV)  
- Backup / Restore  

Data folder (Windows): `%LOCALAPPDATA%\EmployeeAuditManager\`
