#!/usr/bin/env python3
"""
Employee Audit Manager v1.2.0 — Single-file Windows desktop application
=======================================================================
HR-focused daily audit tool (EOD format) with employees, tasks, attendance,
reports, and backup.

Dependencies:
  pip install PySide6 reportlab python-docx

Build Windows EXE (on Windows or GitHub Actions windows-latest):
  pip install pyinstaller
  pyinstaller --noconfirm --windowed --onefile --name EmployeeAuditManager ^
    --collect-all PySide6 --collect-all reportlab EmployeeAuditManager.py

Data is stored under %%LOCALAPPDATA%%\\EmployeeAuditManager\\ (Windows)
or ~/.employeeauditmanager/ (Linux/macOS).
"""

from __future__ import annotations

import csv
import hashlib
import os
import secrets
import shutil
import sqlite3
import subprocess
import sys
import traceback
import webbrowser
from contextlib import contextmanager
from datetime import datetime, date
from pathlib import Path
from typing import Any, Generator, Optional
from urllib.parse import urlparse

# ---------------------------------------------------------------------------
# Third-party
# ---------------------------------------------------------------------------
from PySide6.QtCore import Qt, QDate, QTime
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QDialog, QVBoxLayout, QHBoxLayout,
    QGridLayout, QLabel, QLineEdit, QPushButton, QComboBox, QDateEdit,
    QTimeEdit, QSpinBox, QTextEdit, QTableWidget, QTableWidgetItem,
    QHeaderView, QAbstractItemView, QCheckBox, QFormLayout, QGroupBox,
    QFrame, QStackedWidget, QListWidget, QListWidgetItem, QMessageBox,
    QSizePolicy, QFileDialog, QSplitter,
)

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False

try:
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# ===========================================================================
# CONFIG
# ===========================================================================
APP_NAME = "EmployeeAuditManager"
APP_DISPLAY_NAME = "Employee Audit Manager"
APP_VERSION = "1.2.0"
SCHEMA_VERSION = 2

TASK_STATUSES = ["Not Started", "In Progress", "Blocked", "Completed", "Cancelled"]
TASK_PRIORITIES = ["Low", "Medium", "High", "Critical"]


def get_data_dir() -> Path:
    if sys.platform == "win32":
        base = os.environ.get("LOCALAPPDATA")
        path = Path(base) / APP_NAME if base else Path.home() / "AppData" / "Local" / APP_NAME
    else:
        path = Path.home() / f".{APP_NAME.lower()}"
    path.mkdir(parents=True, exist_ok=True)
    return path


DATA_DIR = get_data_dir()
DB_PATH = DATA_DIR / "employee_audit.db"
BACKUP_DIR = DATA_DIR / "backups"
REPORTS_DIR = DATA_DIR / "reports"
for _d in (BACKUP_DIR, REPORTS_DIR):
    _d.mkdir(parents=True, exist_ok=True)

# ===========================================================================
# STYLES
# ===========================================================================
APP_STYLESHEET = """
QMainWindow, QDialog { background-color: #f0f2f5; }
QWidget { font-family: "Segoe UI", Arial, sans-serif; font-size: 13px; }
QLabel#pageTitle { font-size: 20px; font-weight: 600; color: #1e3a5f; }
QLabel#cardValue { font-size: 22px; font-weight: 700; color: #1e3a5f; }
QLabel#cardLabel { font-size: 12px; color: #5a6a7a; }
QFrame#card { background: white; border: 1px solid #d0d7de; border-radius: 8px; }
QFrame#sidebar { background-color: #1e3a5f; border: none; }
QPushButton#navBtn {
    background: transparent; color: #c8d6e5; text-align: left;
    padding: 10px 16px; border: none; border-radius: 4px;
}
QPushButton#navBtn:hover { background: #2a4a6f; color: white; }
QPushButton#navBtn[active="true"] { background: #0d6efd; color: white; font-weight: 600; }
QPushButton#primaryBtn {
    background: #0d6efd; color: white; border: none; border-radius: 4px;
    padding: 6px 14px; font-weight: 500;
}
QPushButton#primaryBtn:hover { background: #0b5ed7; }
QPushButton#dangerBtn {
    background: #dc3545; color: white; border: none; border-radius: 4px; padding: 6px 14px;
}
QPushButton#dangerBtn:hover { background: #bb2d3b; }
QPushButton#secondaryBtn {
    background: #e9ecef; color: #212529; border: 1px solid #ced4da;
    border-radius: 4px; padding: 6px 14px;
}
QPushButton#secondaryBtn:hover { background: #dee2e6; }
QLineEdit, QComboBox, QDateEdit, QTimeEdit, QSpinBox, QTextEdit, QPlainTextEdit {
    border: 1px solid #ced4da; border-radius: 4px; padding: 5px 8px; background: white;
}
QTableWidget {
    background: white; border: 1px solid #d0d7de; gridline-color: #e9ecef;
    selection-background-color: #cfe2ff;
}
QHeaderView::section {
    background: #e9ecef; padding: 6px; border: none;
    border-right: 1px solid #dee2e6; border-bottom: 1px solid #dee2e6; font-weight: 600;
}
QGroupBox {
    font-weight: 600; border: 1px solid #d0d7de; border-radius: 6px;
    margin-top: 10px; padding-top: 12px; background: white;
}
QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; }
"""


def btn_primary(text: str) -> QPushButton:
    b = QPushButton(text)
    b.setObjectName("primaryBtn")
    b.setCursor(Qt.CursorShape.PointingHandCursor)
    return b


def btn_secondary(text: str) -> QPushButton:
    b = QPushButton(text)
    b.setObjectName("secondaryBtn")
    b.setCursor(Qt.CursorShape.PointingHandCursor)
    return b


def btn_danger(text: str) -> QPushButton:
    b = QPushButton(text)
    b.setObjectName("dangerBtn")
    b.setCursor(Qt.CursorShape.PointingHandCursor)
    return b


class StatCard(QFrame):
    def __init__(self, label: str, value: str = "0", parent=None):
        super().__init__(parent)
        self.setObjectName("card")
        self.setMinimumHeight(80)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(14, 10, 14, 10)
        self.val = QLabel(value)
        self.val.setObjectName("cardValue")
        self.cap = QLabel(label)
        self.cap.setObjectName("cardLabel")
        lay.addWidget(self.val)
        lay.addWidget(self.cap)

    def set_value(self, v: str) -> None:
        self.val.setText(v)


# ===========================================================================
# UTILS
# ===========================================================================
def today_str() -> str:
    return date.today().isoformat()


def qdate_to_str(qd: QDate) -> str:
    return qd.toString("yyyy-MM-dd")


def str_to_qdate(s: str) -> QDate:
    d = QDate.fromString(s or "", "yyyy-MM-dd")
    return d if d.isValid() else QDate.currentDate()


def qtime_to_str(qt: QTime) -> str:
    return qt.toString("HH:mm") if qt.isValid() else ""


def str_to_qtime(s: Optional[str]) -> QTime:
    if not s:
        return QTime()
    t = QTime.fromString(s[:5], "HH:mm")
    return t if t.isValid() else QTime()


def open_url(url: str) -> bool:
    url = (url or "").strip()
    if not url:
        return False
    if not urlparse(url).scheme:
        url = "https://" + url
    try:
        webbrowser.open(url)
        return True
    except Exception:
        return False


def confirm(parent, title, text) -> bool:
    return QMessageBox.question(
        parent, title, text,
        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        QMessageBox.StandardButton.No,
    ) == QMessageBox.StandardButton.Yes


def info(parent, title, text):
    QMessageBox.information(parent, title, text)


def warn(parent, title, text):
    QMessageBox.warning(parent, title, text)


def error(parent, title, text):
    QMessageBox.critical(parent, title, text)


def open_path(path: Path):
    try:
        if sys.platform == "win32":
            os.startfile(str(path))  # type: ignore
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
    except Exception:
        pass


# ===========================================================================
# DATABASE
# ===========================================================================
class DatabaseError(Exception):
    pass


def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(str(DB_PATH), timeout=30, isolation_level=None)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    return conn


@contextmanager
def db_transaction() -> Generator[sqlite3.Connection, None, None]:
    conn = get_connection()
    try:
        conn.execute("BEGIN")
        yield conn
        conn.execute("COMMIT")
    except Exception:
        conn.execute("ROLLBACK")
        raise
    finally:
        conn.close()


def _table_exists(conn, name) -> bool:
    return conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (name,)
    ).fetchone() is not None


def _get_schema_version(conn) -> int:
    if not _table_exists(conn, "schema_version"):
        return 0
    row = conn.execute("SELECT version FROM schema_version LIMIT 1").fetchone()
    return int(row["version"]) if row else 0


def _set_schema_version(conn, version: int):
    if not _table_exists(conn, "schema_version"):
        conn.execute("CREATE TABLE schema_version (version INTEGER NOT NULL)")
        conn.execute("INSERT INTO schema_version (version) VALUES (?)", (version,))
    else:
        conn.execute("UPDATE schema_version SET version = ?", (version,))


def initialize_database():
    conn = get_connection()
    try:
        current = _get_schema_version(conn)
        if current < SCHEMA_VERSION:
            if DB_PATH.exists() and current > 0:
                stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                shutil.copy2(DB_PATH, BACKUP_DIR / f"pre_migration_{stamp}.db")
            if current < 1:
                _create_v1(conn)
            if current < 2:
                _create_v2(conn)
            _set_schema_version(conn, SCHEMA_VERSION)
    finally:
        conn.close()


def _create_v1(conn):
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS schema_version (version INTEGER NOT NULL);
        CREATE TABLE IF NOT EXISTS manager_auth (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            salt BLOB NOT NULL, hash BLOB NOT NULL,
            created_at TEXT NOT NULL, updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id TEXT NOT NULL UNIQUE, full_name TEXT NOT NULL,
            department TEXT, role TEXT, email TEXT, phone TEXT,
            is_active INTEGER NOT NULL DEFAULT 1, notes TEXT,
            created_at TEXT NOT NULL, updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS tasks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL, description TEXT, task_date TEXT NOT NULL,
            due_time TEXT, priority TEXT NOT NULL DEFAULT 'Medium',
            status TEXT NOT NULL DEFAULT 'Not Started', progress INTEGER NOT NULL DEFAULT 0,
            proof_url TEXT, manager_remarks TEXT,
            created_at TEXT NOT NULL, updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS task_assignments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            task_id INTEGER NOT NULL, employee_id INTEGER NOT NULL,
            FOREIGN KEY (task_id) REFERENCES tasks(id) ON DELETE CASCADE,
            FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE CASCADE,
            UNIQUE(task_id, employee_id)
        );
        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id INTEGER NOT NULL, attendance_date TEXT NOT NULL,
            check_in TEXT, check_out TEXT, remarks TEXT,
            created_at TEXT NOT NULL, updated_at TEXT NOT NULL,
            FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE CASCADE,
            UNIQUE(employee_id, attendance_date)
        );
        CREATE TABLE IF NOT EXISTS attendance_audit (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            attendance_id INTEGER NOT NULL, field_name TEXT NOT NULL,
            old_value TEXT, new_value TEXT, changed_at TEXT NOT NULL,
            FOREIGN KEY (attendance_id) REFERENCES attendance(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT);
        CREATE INDEX IF NOT EXISTS idx_tasks_date ON tasks(task_date);
        CREATE INDEX IF NOT EXISTS idx_employees_active ON employees(is_active);
    """)


def _create_v2(conn):
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS daily_audits (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id INTEGER NOT NULL, audit_date TEXT NOT NULL,
            work_submitted INTEGER NOT NULL DEFAULT 0,
            late INTEGER NOT NULL DEFAULT 0,
            links_open INTEGER NOT NULL DEFAULT 1,
            missing_proof INTEGER NOT NULL DEFAULT 0,
            completed_work TEXT, proof_links TEXT, pending TEXT, manager_remarks TEXT,
            created_at TEXT NOT NULL, updated_at TEXT NOT NULL,
            FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE CASCADE,
            UNIQUE(employee_id, audit_date)
        );
        CREATE INDEX IF NOT EXISTS idx_daily_audits_date ON daily_audits(audit_date);
    """)


# ---- Auth ----
_ITER = 200_000
_SALT_LEN = 32
_HASH_LEN = 32


def _hash_pw(password: str, salt: bytes) -> bytes:
    return hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, _ITER, dklen=_HASH_LEN)


def is_setup_complete() -> bool:
    conn = get_connection()
    try:
        return conn.execute("SELECT 1 FROM manager_auth WHERE id=1").fetchone() is not None
    finally:
        conn.close()


def setup_manager(password: str):
    if is_setup_complete():
        raise RuntimeError("Credentials already set.")
    if len(password) < 4:
        raise ValueError("PIN must be at least 4 characters.")
    salt = secrets.token_bytes(_SALT_LEN)
    digest = _hash_pw(password, salt)
    now = datetime.now().isoformat(timespec="seconds")
    with db_transaction() as conn:
        conn.execute(
            "INSERT INTO manager_auth (id, salt, hash, created_at, updated_at) VALUES (1,?,?,?,?)",
            (salt, digest, now, now),
        )


def verify_manager(password: str) -> bool:
    conn = get_connection()
    try:
        row = conn.execute("SELECT salt, hash FROM manager_auth WHERE id=1").fetchone()
        if not row:
            return False
        return secrets.compare_digest(_hash_pw(password, row["salt"]), row["hash"])
    finally:
        conn.close()


def change_password(old: str, new: str):
    if not verify_manager(old):
        raise ValueError("Current PIN is incorrect.")
    if len(new) < 4:
        raise ValueError("New PIN must be at least 4 characters.")
    salt = secrets.token_bytes(_SALT_LEN)
    digest = _hash_pw(new, salt)
    now = datetime.now().isoformat(timespec="seconds")
    with db_transaction() as conn:
        conn.execute(
            "UPDATE manager_auth SET salt=?, hash=?, updated_at=? WHERE id=1",
            (salt, digest, now),
        )


# ---- Employees ----
def employee_exists(emp_id: str, exclude_pk: Optional[int] = None) -> bool:
    conn = get_connection()
    try:
        if exclude_pk is None:
            return conn.execute(
                "SELECT 1 FROM employees WHERE employee_id=?", (emp_id,)
            ).fetchone() is not None
        return conn.execute(
            "SELECT 1 FROM employees WHERE employee_id=? AND id!=?", (emp_id, exclude_pk)
        ).fetchone() is not None
    finally:
        conn.close()


def get_employees(active_only=False, search="") -> list:
    conn = get_connection()
    try:
        sql = "SELECT * FROM employees WHERE 1=1"
        params: list = []
        if active_only:
            sql += " AND is_active=1"
        if search:
            sql += " AND (full_name LIKE ? OR employee_id LIKE ? OR department LIKE ? OR role LIKE ?)"
            q = f"%{search}%"
            params.extend([q, q, q, q])
        sql += " ORDER BY full_name COLLATE NOCASE"
        return list(conn.execute(sql, params).fetchall())
    finally:
        conn.close()


def get_employee(pk: int):
    conn = get_connection()
    try:
        return conn.execute("SELECT * FROM employees WHERE id=?", (pk,)).fetchone()
    finally:
        conn.close()


def add_employee(employee_id, full_name, department="", role="", email="", phone="",
                 notes="", is_active=True) -> int:
    if employee_exists(employee_id):
        raise DatabaseError(f"Employee ID '{employee_id}' already exists.")
    now = datetime.now().isoformat(timespec="seconds")
    with db_transaction() as conn:
        cur = conn.execute(
            """INSERT INTO employees
            (employee_id, full_name, department, role, email, phone, is_active, notes, created_at, updated_at)
            VALUES (?,?,?,?,?,?,?,?,?,?)""",
            (employee_id.strip(), full_name.strip(), department.strip(), role.strip(),
             email.strip(), phone.strip(), 1 if is_active else 0, notes.strip(), now, now),
        )
        return int(cur.lastrowid)


def update_employee(pk, employee_id, full_name, department="", role="", email="", phone="",
                    notes="", is_active=True):
    if employee_exists(employee_id, exclude_pk=pk):
        raise DatabaseError(f"Employee ID '{employee_id}' already exists.")
    now = datetime.now().isoformat(timespec="seconds")
    with db_transaction() as conn:
        conn.execute(
            """UPDATE employees SET employee_id=?, full_name=?, department=?, role=?,
            email=?, phone=?, is_active=?, notes=?, updated_at=? WHERE id=?""",
            (employee_id.strip(), full_name.strip(), department.strip(), role.strip(),
             email.strip(), phone.strip(), 1 if is_active else 0, notes.strip(), now, pk),
        )


def set_employee_active(pk: int, active: bool):
    now = datetime.now().isoformat(timespec="seconds")
    with db_transaction() as conn:
        conn.execute(
            "UPDATE employees SET is_active=?, updated_at=? WHERE id=?",
            (1 if active else 0, now, pk),
        )


def employee_has_records(pk: int) -> bool:
    conn = get_connection()
    try:
        t = conn.execute(
            "SELECT 1 FROM task_assignments WHERE employee_id=? LIMIT 1", (pk,)
        ).fetchone()
        a = conn.execute(
            "SELECT 1 FROM attendance WHERE employee_id=? LIMIT 1", (pk,)
        ).fetchone()
        d = conn.execute(
            "SELECT 1 FROM daily_audits WHERE employee_id=? LIMIT 1", (pk,)
        ).fetchone()
        return t is not None or a is not None or d is not None
    finally:
        conn.close()


def delete_employee(pk: int):
    with db_transaction() as conn:
        conn.execute("DELETE FROM employees WHERE id=?", (pk,))


# ---- Tasks ----
def get_tasks(task_date=None, employee_pk=None, status=None, priority=None, search="") -> list:
    conn = get_connection()
    try:
        sql = """
            SELECT t.*, GROUP_CONCAT(e.full_name, ', ') AS assignee_names,
                   GROUP_CONCAT(e.id, ',') AS assignee_ids
            FROM tasks t
            LEFT JOIN task_assignments ta ON ta.task_id = t.id
            LEFT JOIN employees e ON e.id = ta.employee_id
            WHERE 1=1
        """
        params: list = []
        if task_date:
            sql += " AND t.task_date=?"
            params.append(task_date)
        if employee_pk is not None:
            sql += " AND ta.employee_id=?"
            params.append(employee_pk)
        if status:
            sql += " AND t.status=?"
            params.append(status)
        if priority:
            sql += " AND t.priority=?"
            params.append(priority)
        if search:
            sql += " AND (t.title LIKE ? OR t.description LIKE ?)"
            q = f"%{search}%"
            params.extend([q, q])
        sql += " GROUP BY t.id ORDER BY t.task_date DESC, t.title"
        return list(conn.execute(sql, params).fetchall())
    finally:
        conn.close()


def get_task(pk: int):
    conn = get_connection()
    try:
        return conn.execute("SELECT * FROM tasks WHERE id=?", (pk,)).fetchone()
    finally:
        conn.close()


def get_task_assignees(task_id: int) -> list:
    conn = get_connection()
    try:
        return [int(r["employee_id"]) for r in conn.execute(
            "SELECT employee_id FROM task_assignments WHERE task_id=?", (task_id,)
        ).fetchall()]
    finally:
        conn.close()


def add_task(title, description, task_date, due_time, priority, status, progress,
             proof_url, manager_remarks, assignee_ids) -> int:
    now = datetime.now().isoformat(timespec="seconds")
    with db_transaction() as conn:
        cur = conn.execute(
            """INSERT INTO tasks
            (title, description, task_date, due_time, priority, status, progress,
             proof_url, manager_remarks, created_at, updated_at)
            VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
            (title.strip(), description.strip(), task_date, due_time.strip() or None,
             priority, status, max(0, min(100, progress)),
             proof_url.strip() or None, manager_remarks.strip() or None, now, now),
        )
        tid = int(cur.lastrowid)
        for eid in assignee_ids:
            conn.execute(
                "INSERT OR IGNORE INTO task_assignments (task_id, employee_id) VALUES (?,?)",
                (tid, eid),
            )
        return tid


def update_task(pk, title, description, task_date, due_time, priority, status, progress,
                proof_url, manager_remarks, assignee_ids):
    now = datetime.now().isoformat(timespec="seconds")
    with db_transaction() as conn:
        conn.execute(
            """UPDATE tasks SET title=?, description=?, task_date=?, due_time=?, priority=?,
            status=?, progress=?, proof_url=?, manager_remarks=?, updated_at=? WHERE id=?""",
            (title.strip(), description.strip(), task_date, due_time.strip() or None,
             priority, status, max(0, min(100, progress)),
             proof_url.strip() or None, manager_remarks.strip() or None, now, pk),
        )
        conn.execute("DELETE FROM task_assignments WHERE task_id=?", (pk,))
        for eid in assignee_ids:
            conn.execute(
                "INSERT OR IGNORE INTO task_assignments (task_id, employee_id) VALUES (?,?)",
                (pk, eid),
            )


def delete_task(pk: int):
    with db_transaction() as conn:
        conn.execute("DELETE FROM tasks WHERE id=?", (pk,))


def copy_tasks_from_date(source_date: str, target_date: str) -> int:
    tasks = get_tasks(task_date=source_date)
    count = 0
    for t in tasks:
        add_task(
            t["title"], t["description"] or "", target_date, t["due_time"] or "",
            t["priority"], "Not Started", 0, "", "", get_task_assignees(t["id"]),
        )
        count += 1
    return count


def update_task_progress(pk, status, progress, proof_url, manager_remarks):
    now = datetime.now().isoformat(timespec="seconds")
    with db_transaction() as conn:
        conn.execute(
            """UPDATE tasks SET status=?, progress=?, proof_url=?, manager_remarks=?, updated_at=?
            WHERE id=?""",
            (status, max(0, min(100, progress)), proof_url.strip() or None,
             manager_remarks.strip() or None, now, pk),
        )


# ---- Attendance ----
def get_attendance(employee_pk=None, date_from=None, date_to=None, attendance_date=None) -> list:
    conn = get_connection()
    try:
        sql = """SELECT a.*, e.full_name, e.employee_id AS emp_code
                 FROM attendance a JOIN employees e ON e.id=a.employee_id WHERE 1=1"""
        params: list = []
        if employee_pk is not None:
            sql += " AND a.employee_id=?"
            params.append(employee_pk)
        if attendance_date:
            sql += " AND a.attendance_date=?"
            params.append(attendance_date)
        if date_from:
            sql += " AND a.attendance_date>=?"
            params.append(date_from)
        if date_to:
            sql += " AND a.attendance_date<=?"
            params.append(date_to)
        sql += " ORDER BY a.attendance_date DESC, e.full_name"
        return list(conn.execute(sql, params).fetchall())
    finally:
        conn.close()


def get_attendance_record(pk: int):
    conn = get_connection()
    try:
        return conn.execute("SELECT * FROM attendance WHERE id=?", (pk,)).fetchone()
    finally:
        conn.close()


def attendance_exists(employee_pk: int, day: str) -> bool:
    conn = get_connection()
    try:
        return conn.execute(
            "SELECT 1 FROM attendance WHERE employee_id=? AND attendance_date=?",
            (employee_pk, day),
        ).fetchone() is not None
    finally:
        conn.close()


def add_attendance(employee_pk, attendance_date, check_in, check_out, remarks) -> int:
    if attendance_exists(employee_pk, attendance_date):
        raise DatabaseError(f"Attendance already exists for this employee on {attendance_date}.")
    now = datetime.now().isoformat(timespec="seconds")
    with db_transaction() as conn:
        cur = conn.execute(
            """INSERT INTO attendance
            (employee_id, attendance_date, check_in, check_out, remarks, created_at, updated_at)
            VALUES (?,?,?,?,?,?,?)""",
            (employee_pk, attendance_date, check_in.strip() or None,
             check_out.strip() or None, remarks.strip() or None, now, now),
        )
        return int(cur.lastrowid)


def update_attendance(pk, check_in, check_out, remarks):
    existing = get_attendance_record(pk)
    if not existing:
        raise DatabaseError("Record not found.")
    now = datetime.now().isoformat(timespec="seconds")
    changes = []
    for field, new_val in (
        ("check_in", check_in.strip() or None),
        ("check_out", check_out.strip() or None),
        ("remarks", remarks.strip() or None),
    ):
        old = existing[field]
        if (old or "") != (new_val or ""):
            changes.append((field, old, new_val))
    with db_transaction() as conn:
        conn.execute(
            "UPDATE attendance SET check_in=?, check_out=?, remarks=?, updated_at=? WHERE id=?",
            (check_in.strip() or None, check_out.strip() or None,
             remarks.strip() or None, now, pk),
        )
        for field, old, new in changes:
            conn.execute(
                """INSERT INTO attendance_audit
                (attendance_id, field_name, old_value, new_value, changed_at)
                VALUES (?,?,?,?,?)""",
                (pk, field, old, new, now),
            )


def delete_attendance(pk: int):
    with db_transaction() as conn:
        conn.execute("DELETE FROM attendance WHERE id=?", (pk,))


def calc_duration(check_in, check_out) -> str:
    if not check_in:
        return ""
    if not check_out:
        return "Missing check-out"
    try:
        t1 = datetime.strptime(check_in[:5], "%H:%M")
        t2 = datetime.strptime(check_out[:5], "%H:%M")
        delta = t2 - t1
        if delta.total_seconds() < 0:
            return "Invalid times"
        h, rem = divmod(int(delta.total_seconds()), 3600)
        return f"{h}h {rem // 60}m"
    except ValueError:
        return "Invalid times"


# ---- Daily audits ----
def get_daily_audits(audit_date=None, employee_pk=None) -> list:
    conn = get_connection()
    try:
        sql = """SELECT d.*, e.full_name, e.employee_id AS emp_code
                 FROM daily_audits d JOIN employees e ON e.id=d.employee_id WHERE 1=1"""
        params: list = []
        if audit_date:
            sql += " AND d.audit_date=?"
            params.append(audit_date)
        if employee_pk is not None:
            sql += " AND d.employee_id=?"
            params.append(employee_pk)
        sql += " ORDER BY e.full_name COLLATE NOCASE"
        return list(conn.execute(sql, params).fetchall())
    finally:
        conn.close()


def get_daily_audit(pk: int):
    conn = get_connection()
    try:
        return conn.execute("SELECT * FROM daily_audits WHERE id=?", (pk,)).fetchone()
    finally:
        conn.close()


def get_daily_audit_for(employee_pk: int, audit_date: str):
    conn = get_connection()
    try:
        return conn.execute(
            "SELECT * FROM daily_audits WHERE employee_id=? AND audit_date=?",
            (employee_pk, audit_date),
        ).fetchone()
    finally:
        conn.close()


def upsert_daily_audit(employee_pk, audit_date, work_submitted, late, links_open,
                       missing_proof, completed_work, proof_links, pending,
                       manager_remarks="") -> int:
    now = datetime.now().isoformat(timespec="seconds")
    existing = get_daily_audit_for(employee_pk, audit_date)
    with db_transaction() as conn:
        if existing:
            conn.execute(
                """UPDATE daily_audits SET work_submitted=?, late=?, links_open=?, missing_proof=?,
                completed_work=?, proof_links=?, pending=?, manager_remarks=?, updated_at=?
                WHERE id=?""",
                (1 if work_submitted else 0, 1 if late else 0, 1 if links_open else 0,
                 1 if missing_proof else 0, completed_work.strip() or None,
                 proof_links.strip() or None, pending.strip() or None,
                 manager_remarks.strip() or None, now, existing["id"]),
            )
            return int(existing["id"])
        cur = conn.execute(
            """INSERT INTO daily_audits
            (employee_id, audit_date, work_submitted, late, links_open, missing_proof,
             completed_work, proof_links, pending, manager_remarks, created_at, updated_at)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
            (employee_pk, audit_date, 1 if work_submitted else 0, 1 if late else 0,
             1 if links_open else 0, 1 if missing_proof else 0,
             completed_work.strip() or None, proof_links.strip() or None,
             pending.strip() or None, manager_remarks.strip() or None, now, now),
        )
        return int(cur.lastrowid)


def delete_daily_audit(pk: int):
    with db_transaction() as conn:
        conn.execute("DELETE FROM daily_audits WHERE id=?", (pk,))


def get_daily_audit_summary(audit_date: str) -> dict:
    rows = get_daily_audits(audit_date=audit_date)
    active = get_employees(active_only=True)
    return {
        "total_active": len(active),
        "recorded": len(rows),
        "not_yet_recorded": max(0, len(active) - len(rows)),
        "submitted": sum(1 for r in rows if r["work_submitted"]),
        "late": sum(1 for r in rows if r["late"]),
        "missing_proof": sum(1 for r in rows if r["missing_proof"]),
        "has_pending": sum(1 for r in rows if (r["pending"] or "").strip()),
    }


# ---- Dashboard ----
def get_dashboard_stats(today: Optional[str] = None) -> dict:
    if today is None:
        today = today_str()
    conn = get_connection()
    try:
        active = conn.execute(
            "SELECT COUNT(*) AS c FROM employees WHERE is_active=1"
        ).fetchone()["c"]
        present = conn.execute(
            """SELECT COUNT(*) AS c FROM attendance
               WHERE attendance_date=? AND check_in IS NOT NULL AND check_in!=''""",
            (today,),
        ).fetchone()["c"]
        tasks_today = conn.execute(
            "SELECT COUNT(*) AS c FROM tasks WHERE task_date=?", (today,)
        ).fetchone()["c"]
        completed = conn.execute(
            "SELECT COUNT(*) AS c FROM tasks WHERE task_date=? AND status='Completed'",
            (today,),
        ).fetchone()["c"]
        pending = conn.execute(
            "SELECT COUNT(*) AS c FROM tasks WHERE task_date=? AND status='Not Started'",
            (today,),
        ).fetchone()["c"]
        in_progress = conn.execute(
            "SELECT COUNT(*) AS c FROM tasks WHERE task_date=? AND status='In Progress'",
            (today,),
        ).fetchone()["c"]
        blocked = conn.execute(
            "SELECT COUNT(*) AS c FROM tasks WHERE task_date=? AND status='Blocked'",
            (today,),
        ).fetchone()["c"]
        emp_progress = conn.execute(
            """SELECT e.id, e.full_name, e.employee_id,
                      COUNT(t.id) AS total,
                      SUM(CASE WHEN t.status='Completed' THEN 1 ELSE 0 END) AS done
               FROM employees e
               LEFT JOIN task_assignments ta ON ta.employee_id=e.id
               LEFT JOIN tasks t ON t.id=ta.task_id AND t.task_date=?
               WHERE e.is_active=1 GROUP BY e.id ORDER BY e.full_name""",
            (today,),
        ).fetchall()
        return {
            "active_employees": active,
            "present": present,
            "missing_attendance": max(0, active - present),
            "tasks_today": tasks_today,
            "completed": completed,
            "pending": pending,
            "in_progress": in_progress,
            "blocked": blocked,
            "completion_pct": round((completed / tasks_today * 100) if tasks_today else 0, 1),
            "employee_progress": emp_progress,
        }
    finally:
        conn.close()


# ---- Backup ----
def create_backup(label: str = "") -> Path:
    if not DB_PATH.exists():
        raise DatabaseError("Database does not exist yet.")
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    name = f"backup_{stamp}"
    if label:
        safe = "".join(c if c.isalnum() or c in "-_" else "_" for c in label)[:40]
        name = f"backup_{stamp}_{safe}"
    dest = BACKUP_DIR / f"{name}.db"
    shutil.copy2(DB_PATH, dest)
    return dest


def list_backups() -> list:
    return sorted(BACKUP_DIR.glob("*.db"), key=lambda p: p.stat().st_mtime, reverse=True)


def restore_backup(backup_path: Path):
    if not backup_path.exists():
        raise DatabaseError("Backup file not found.")
    try:
        test = sqlite3.connect(str(backup_path))
        test.execute("SELECT name FROM sqlite_master LIMIT 1")
        test.close()
    except sqlite3.Error as e:
        raise DatabaseError(f"Invalid backup: {e}") from e
    if DB_PATH.exists():
        create_backup("before_restore")
    shutil.copy2(backup_path, DB_PATH)


# ===========================================================================
# REPORTS (PDF / DOCX / CSV) — compact
# ===========================================================================
def generate_eod_pdf(report_date: str) -> Path:
    if not HAS_REPORTLAB:
        raise RuntimeError("reportlab is not installed.")
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = REPORTS_DIR / f"eod_audit_{report_date}_{stamp}.pdf"
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="RT", parent=styles["Heading1"], fontSize=14, spaceAfter=6))
    styles.add(ParagraphStyle(name="SH", parent=styles["Heading2"], fontSize=11, spaceBefore=10))
    doc = SimpleDocTemplate(str(out), pagesize=A4, leftMargin=15 * mm, rightMargin=15 * mm,
                            topMargin=15 * mm, bottomMargin=15 * mm)
    story = [Paragraph(f"Employee Daily Audit — {report_date}", styles["RT"]),
             Paragraph(f"{APP_DISPLAY_NAME} v{APP_VERSION}", styles["Normal"]), Spacer(1, 8)]
    summary = get_daily_audit_summary(report_date)
    data = [["Metric", "Value"],
            ["Active", str(summary["total_active"])],
            ["Recorded", str(summary["recorded"])],
            ["Submitted", str(summary["submitted"])],
            ["Late", str(summary["late"])],
            ["Missing proof", str(summary["missing_proof"])],
            ["Has pending", str(summary["has_pending"])]]
    t = Table(data, colWidths=[120, 80])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1e3a5f")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
    ]))
    story.append(t)
    story.append(Spacer(1, 10))
    for r in get_daily_audits(audit_date=report_date):
        story.append(Paragraph(f"{r['full_name']} ({r['emp_code']})", styles["SH"]))
        flags = (f"Submitted: {'Yes' if r['work_submitted'] else 'No'} | "
                 f"Late: {'Yes' if r['late'] else 'No'} | "
                 f"Missing proof: {'Yes' if r['missing_proof'] else 'No'}")
        story.append(Paragraph(flags, styles["Normal"]))
        if r["completed_work"]:
            story.append(Paragraph(f"<b>Completed:</b> {(r['completed_work'] or '')[:300]}", styles["Normal"]))
        if r["proof_links"]:
            story.append(Paragraph(f"<b>Links:</b> {(r['proof_links'] or '')[:200]}", styles["Normal"]))
        if r["pending"]:
            story.append(Paragraph(f"<b>Pending:</b> {r['pending']}", styles["Normal"]))
        story.append(Spacer(1, 6))
    doc.build(story)
    return out


def export_employees_csv() -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = REPORTS_DIR / f"employees_{stamp}.csv"
    with open(out, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Employee ID", "Name", "Department", "Role", "Active"])
        for e in get_employees():
            w.writerow([e["employee_id"], e["full_name"], e["department"] or "",
                        e["role"] or "", "Yes" if e["is_active"] else "No"])
    return out


def export_tasks_csv(date_from: str, date_to: str) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = REPORTS_DIR / f"tasks_{date_from}_{date_to}_{stamp}.csv"
    tasks = [t for t in get_tasks() if date_from <= (t["task_date"] or "") <= date_to]
    with open(out, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Date", "Title", "Assignees", "Status", "Priority", "Progress", "Proof"])
        for t in tasks:
            w.writerow([t["task_date"], t["title"], t["assignee_names"] or "",
                        t["status"], t["priority"], t["progress"], t["proof_url"] or ""])
    return out


# ===========================================================================
# DIALOGS: Login / First-run
# ===========================================================================
class FirstRunDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"{APP_DISPLAY_NAME} — HR Setup")
        self.setFixedSize(400, 300)
        self.setModal(True)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 24, 24, 24)
        title = QLabel("Welcome — HR / Admin Setup")
        title.setStyleSheet("font-size: 18px; font-weight: 600; color: #1e3a5f;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(title)
        lay.addWidget(QLabel("Create a PIN/password (min 4 characters) to protect the app."))
        self.p1 = QLineEdit()
        self.p1.setEchoMode(QLineEdit.EchoMode.Password)
        self.p1.setPlaceholderText("PIN / Password")
        self.p2 = QLineEdit()
        self.p2.setEchoMode(QLineEdit.EchoMode.Password)
        self.p2.setPlaceholderText("Confirm")
        self.p2.returnPressed.connect(self._save)
        lay.addWidget(self.p1)
        lay.addWidget(self.p2)
        self.err = QLabel("")
        self.err.setStyleSheet("color: #dc3545;")
        lay.addWidget(self.err)
        row = QHBoxLayout()
        row.addStretch()
        b = btn_primary("Create & Continue")
        b.clicked.connect(self._save)
        row.addWidget(b)
        lay.addLayout(row)

    def _save(self):
        if len(self.p1.text()) < 4:
            self.err.setText("PIN must be at least 4 characters.")
            return
        if self.p1.text() != self.p2.text():
            self.err.setText("Entries do not match.")
            return
        try:
            setup_manager(self.p1.text())
            self.accept()
        except Exception as e:
            error(self, "Setup Failed", str(e))


class LoginDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"{APP_DISPLAY_NAME} — Login")
        self.setFixedSize(360, 200)
        self.setModal(True)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 24, 24, 24)
        title = QLabel("HR / Admin Login")
        title.setStyleSheet("font-size: 18px; font-weight: 600; color: #1e3a5f;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(title)
        self.pwd = QLineEdit()
        self.pwd.setEchoMode(QLineEdit.EchoMode.Password)
        self.pwd.setPlaceholderText("PIN / Password")
        self.pwd.returnPressed.connect(self._try)
        lay.addWidget(self.pwd)
        self.err = QLabel("")
        self.err.setStyleSheet("color: #dc3545;")
        self.err.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(self.err)
        row = QHBoxLayout()
        row.addStretch()
        b = btn_primary("Login")
        b.clicked.connect(self._try)
        row.addWidget(b)
        lay.addLayout(row)

    def _try(self):
        if not self.pwd.text():
            self.err.setText("Enter your PIN.")
            return
        if verify_manager(self.pwd.text()):
            self.accept()
        else:
            self.err.setText("Incorrect PIN.")
            self.pwd.clear()
            self.pwd.setFocus()


# ===========================================================================
# PAGES (simplified but complete)
# ===========================================================================
class DashboardPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 16, 20, 16)
        header = QHBoxLayout()
        t = QLabel("Dashboard")
        t.setObjectName("pageTitle")
        header.addWidget(t)
        header.addStretch()
        header.addWidget(QLabel("Date:"))
        self.date = QDateEdit()
        self.date.setCalendarPopup(True)
        self.date.setDate(QDate.currentDate())
        self.date.setDisplayFormat("yyyy-MM-dd")
        self.date.dateChanged.connect(self.refresh)
        header.addWidget(self.date)
        r = btn_secondary("Refresh")
        r.clicked.connect(self.refresh)
        header.addWidget(r)
        lay.addLayout(header)
        grid = QGridLayout()
        self.cards = {}
        labels = [
            ("active", "Active Employees"), ("present", "Present"),
            ("missing", "Missing Attendance"), ("tasks", "Tasks Today"),
            ("done", "Completed"), ("pending", "Pending"),
            ("prog", "In Progress"), ("blocked", "Blocked"), ("pct", "Completion %"),
        ]
        for i, (k, lab) in enumerate(labels):
            c = StatCard(lab)
            self.cards[k] = c
            grid.addWidget(c, i // 3, i % 3)
        lay.addLayout(grid)
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["Employee", "ID", "Tasks", "Completed"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        lay.addWidget(self.table)
        self.refresh()

    def refresh(self):
        s = get_dashboard_stats(qdate_to_str(self.date.date()))
        self.cards["active"].set_value(str(s["active_employees"]))
        self.cards["present"].set_value(str(s["present"]))
        self.cards["missing"].set_value(str(s["missing_attendance"]))
        self.cards["tasks"].set_value(str(s["tasks_today"]))
        self.cards["done"].set_value(str(s["completed"]))
        self.cards["pending"].set_value(str(s["pending"]))
        self.cards["prog"].set_value(str(s["in_progress"]))
        self.cards["blocked"].set_value(str(s["blocked"]))
        self.cards["pct"].set_value(f"{s['completion_pct']}%")
        rows = s["employee_progress"]
        self.table.setRowCount(len(rows))
        for i, r in enumerate(rows):
            self.table.setItem(i, 0, QTableWidgetItem(r["full_name"] or ""))
            self.table.setItem(i, 1, QTableWidgetItem(r["employee_id"] or ""))
            self.table.setItem(i, 2, QTableWidgetItem(str(r["total"] or 0)))
            self.table.setItem(i, 3, QTableWidgetItem(str(r["done"] or 0)))


class EmployeesPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 16, 20, 16)
        t = QLabel("Employees")
        t.setObjectName("pageTitle")
        lay.addWidget(t)
        bar = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search…")
        self.search.textChanged.connect(self.refresh)
        bar.addWidget(self.search, stretch=1)
        self.filt = QComboBox()
        self.filt.addItems(["All", "Active only", "Inactive only"])
        self.filt.currentIndexChanged.connect(self.refresh)
        bar.addWidget(self.filt)
        add = btn_primary("Add Employee")
        add.clicked.connect(self._add)
        bar.addWidget(add)
        lay.addLayout(bar)
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["ID", "Name", "Department", "Role", "Status"])
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.doubleClicked.connect(self._edit)
        lay.addWidget(self.table)
        row = QHBoxLayout()
        for text, slot in [("Edit", self._edit), ("Activate/Deactivate", self._toggle),
                           ("Delete", self._delete)]:
            b = btn_danger(text) if text == "Delete" else btn_secondary(text)
            b.clicked.connect(slot)
            row.addWidget(b)
        row.addStretch()
        lay.addLayout(row)
        self._pk: dict = {}
        self.refresh()

    def refresh(self):
        mode = self.filt.currentIndex()
        emps = get_employees(active_only=(mode == 1), search=self.search.text().strip())
        if mode == 2:
            emps = [e for e in emps if not e["is_active"]]
        self.table.setRowCount(len(emps))
        self._pk.clear()
        for i, e in enumerate(emps):
            self._pk[i] = e["id"]
            self.table.setItem(i, 0, QTableWidgetItem(e["employee_id"]))
            self.table.setItem(i, 1, QTableWidgetItem(e["full_name"]))
            self.table.setItem(i, 2, QTableWidgetItem(e["department"] or ""))
            self.table.setItem(i, 3, QTableWidgetItem(e["role"] or ""))
            self.table.setItem(i, 4, QTableWidgetItem("Active" if e["is_active"] else "Inactive"))

    def _sel(self):
        rows = self.table.selectionModel().selectedRows()
        return self._pk.get(rows[0].row()) if rows else None

    def _form(self, data=None):
        dlg = QDialog(self)
        dlg.setWindowTitle("Edit Employee" if data else "Add Employee")
        dlg.setMinimumWidth(400)
        form = QFormLayout(dlg)
        fields = {}
        for key, label in [("employee_id", "Employee ID *"), ("full_name", "Full Name *"),
                           ("department", "Department"), ("role", "Role"),
                           ("email", "Email"), ("phone", "Phone")]:
            fields[key] = QLineEdit()
            form.addRow(label, fields[key])
        fields["active"] = QCheckBox("Active")
        fields["active"].setChecked(True)
        form.addRow("", fields["active"])
        fields["notes"] = QTextEdit()
        fields["notes"].setMaximumHeight(60)
        form.addRow("Notes", fields["notes"])
        if data:
            for k in ("employee_id", "full_name", "department", "role", "email", "phone"):
                fields[k].setText(data.get(k) or "")
            fields["active"].setChecked(bool(data.get("is_active", 1)))
            fields["notes"].setPlainText(data.get("notes") or "")
        br = QHBoxLayout()
        br.addStretch()
        c = btn_secondary("Cancel")
        c.clicked.connect(dlg.reject)
        s = btn_primary("Save")
        s.clicked.connect(dlg.accept)
        br.addWidget(c)
        br.addWidget(s)
        form.addRow(br)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return None
        return {
            "employee_id": fields["employee_id"].text().strip(),
            "full_name": fields["full_name"].text().strip(),
            "department": fields["department"].text().strip(),
            "role": fields["role"].text().strip(),
            "email": fields["email"].text().strip(),
            "phone": fields["phone"].text().strip(),
            "is_active": fields["active"].isChecked(),
            "notes": fields["notes"].toPlainText().strip(),
        }

    def _add(self):
        data = self._form()
        if not data:
            return
        if not data["employee_id"] or not data["full_name"]:
            warn(self, "Validation", "ID and Name required.")
            return
        try:
            add_employee(**data)
            info(self, "OK", "Employee added.")
            self.refresh()
        except Exception as e:
            error(self, "Error", str(e))

    def _edit(self):
        pk = self._sel()
        if pk is None:
            warn(self, "Select", "Select an employee.")
            return
        emp = get_employee(pk)
        data = self._form(dict(emp) if emp else None)
        if not data:
            return
        try:
            update_employee(pk, **data)
            info(self, "OK", "Updated.")
            self.refresh()
        except Exception as e:
            error(self, "Error", str(e))

    def _toggle(self):
        pk = self._sel()
        if pk is None:
            return
        emp = get_employee(pk)
        if emp and confirm(self, "Confirm", "Toggle active status?"):
            set_employee_active(pk, not bool(emp["is_active"]))
            self.refresh()

    def _delete(self):
        pk = self._sel()
        if pk is None:
            return
        msg = "Delete permanently?"
        if employee_has_records(pk):
            msg += "\nThis employee has related records. Prefer deactivation."
        if confirm(self, "Delete", msg):
            try:
                delete_employee(pk)
                self.refresh()
            except Exception as e:
                error(self, "Error", str(e))


# ---- Daily Audit (core HR page) ----
class QuickTaskDialog(QDialog):
    def __init__(self, parent=None, employee_pk=None, task_date=None, task_data=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Task" if task_data else "Add Daily Task")
        self.setMinimumWidth(420)
        self.employee_pk = employee_pk
        self._date = task_date or today_str()
        lay = QVBoxLayout(self)
        form = QFormLayout()
        self.title = QLineEdit()
        self.desc = QTextEdit()
        self.desc.setMaximumHeight(60)
        self.priority = QComboBox()
        self.priority.addItems(TASK_PRIORITIES)
        self.priority.setCurrentText("Medium")
        self.status = QComboBox()
        self.status.addItems(TASK_STATUSES)
        self.proof = QLineEdit()
        form.addRow("Title *", self.title)
        form.addRow("Description", self.desc)
        form.addRow("Priority", self.priority)
        form.addRow("Status", self.status)
        form.addRow("Proof URL", self.proof)
        lay.addLayout(form)
        if task_data:
            self.title.setText(task_data.get("title", ""))
            self.desc.setPlainText(task_data.get("description") or "")
            self.priority.setCurrentText(task_data.get("priority", "Medium"))
            self.status.setCurrentText(task_data.get("status", "Not Started"))
            self.proof.setText(task_data.get("proof_url") or "")
            self._date = task_data.get("task_date") or self._date
        br = QHBoxLayout()
        br.addStretch()
        c = btn_secondary("Cancel")
        c.clicked.connect(self.reject)
        s = btn_primary("Save Task")
        s.clicked.connect(self.accept)
        br.addWidget(c)
        br.addWidget(s)
        lay.addLayout(br)

    def get_data(self):
        st = self.status.currentText()
        return {
            "title": self.title.text().strip(),
            "description": self.desc.toPlainText().strip(),
            "task_date": self._date,
            "due_time": "",
            "priority": self.priority.currentText(),
            "status": st,
            "progress": 100 if st == "Completed" else 0,
            "proof_url": self.proof.text().strip(),
            "manager_remarks": "",
            "assignee_ids": [self.employee_pk] if self.employee_pk else [],
        }


class DailyAuditDialog(QDialog):
    def __init__(self, parent=None, employee_pk=None, audit_date=None, data=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Daily Audit" if data else "Record Daily Audit")
        self.setMinimumSize(600, 520)
        self.data = data
        lay = QVBoxLayout(self)
        form = QFormLayout()
        self.emp = QComboBox()
        for e in get_employees(active_only=True):
            self.emp.addItem(f"{e['full_name']} ({e['employee_id']})", e["id"])
        self.date = QDateEdit()
        self.date.setCalendarPopup(True)
        self.date.setDate(QDate.currentDate())
        self.date.setDisplayFormat("yyyy-MM-dd")
        self.submitted = QCheckBox("Work submitted")
        self.late = QCheckBox("Late")
        self.links = QCheckBox("Links / proofs open")
        self.links.setChecked(True)
        self.missing = QCheckBox("Missing proof")
        self.completed = QTextEdit()
        self.completed.setMaximumHeight(70)
        self.proofs = QTextEdit()
        self.proofs.setMaximumHeight(55)
        self.pending = QTextEdit()
        self.pending.setMaximumHeight(45)
        self.remarks = QTextEdit()
        self.remarks.setMaximumHeight(40)
        form.addRow("Employee *", self.emp)
        form.addRow("Date *", self.date)
        form.addRow("", self.submitted)
        form.addRow("", self.late)
        form.addRow("", self.links)
        form.addRow("", self.missing)
        form.addRow("Completed work", self.completed)
        form.addRow("Proof links", self.proofs)
        form.addRow("Pending", self.pending)
        form.addRow("HR remarks", self.remarks)
        lay.addLayout(form)
        # Tasks panel
        grp = QGroupBox("Daily tasks for this employee")
        gl = QVBoxLayout(grp)
        self.task_list = QListWidget()
        self.task_list.setMaximumHeight(100)
        gl.addWidget(self.task_list)
        tr = QHBoxLayout()
        for text, slot in [("Add Task", self._add_task), ("Edit Task", self._edit_task),
                           ("Remove Task", self._rem_task)]:
            b = btn_danger(text) if "Remove" in text else (
                btn_primary(text) if "Add" in text else btn_secondary(text)
            )
            b.clicked.connect(slot)
            tr.addWidget(b)
        tr.addStretch()
        gl.addLayout(tr)
        lay.addWidget(grp)
        if data:
            self.emp.setEnabled(False)
            self.date.setEnabled(False)
            idx = self.emp.findData(data.get("employee_id"))
            if idx >= 0:
                self.emp.setCurrentIndex(idx)
            self.date.setDate(str_to_qdate(data.get("audit_date", "")))
            self.submitted.setChecked(bool(data.get("work_submitted")))
            self.late.setChecked(bool(data.get("late")))
            self.links.setChecked(bool(data.get("links_open", 1)))
            self.missing.setChecked(bool(data.get("missing_proof")))
            self.completed.setPlainText(data.get("completed_work") or "")
            self.proofs.setPlainText(data.get("proof_links") or "")
            self.pending.setPlainText(data.get("pending") or "")
            self.remarks.setPlainText(data.get("manager_remarks") or "")
        else:
            if employee_pk is not None:
                idx = self.emp.findData(employee_pk)
                if idx >= 0:
                    self.emp.setCurrentIndex(idx)
            if audit_date:
                self.date.setDate(str_to_qdate(audit_date))
        self.emp.currentIndexChanged.connect(self._reload_tasks)
        self.date.dateChanged.connect(self._reload_tasks)
        self._reload_tasks()
        br = QHBoxLayout()
        br.addStretch()
        c = btn_secondary("Cancel")
        c.clicked.connect(self.reject)
        s = btn_primary("Save Audit")
        s.clicked.connect(self.accept)
        br.addWidget(c)
        br.addWidget(s)
        lay.addLayout(br)

    def _ed(self):
        return self.emp.currentData(), qdate_to_str(self.date.date())

    def _reload_tasks(self):
        self.task_list.clear()
        emp, day = self._ed()
        if emp is None:
            return
        for t in get_tasks(task_date=day, employee_pk=emp):
            item = QListWidgetItem(f"[{t['status']}] {t['title']}")
            item.setData(Qt.ItemDataRole.UserRole, t["id"])
            self.task_list.addItem(item)

    def _add_task(self):
        emp, day = self._ed()
        if emp is None:
            warn(self, "Select", "Choose employee first.")
            return
        dlg = QuickTaskDialog(self, emp, day)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        data = dlg.get_data()
        if not data["title"]:
            warn(self, "Validation", "Title required.")
            return
        try:
            add_task(**data)
            self._reload_tasks()
        except Exception as e:
            error(self, "Error", str(e))

    def _edit_task(self):
        item = self.task_list.currentItem()
        if not item:
            return
        tid = item.data(Qt.ItemDataRole.UserRole)
        t = get_task(tid)
        if not t:
            return
        emp, day = self._ed()
        dlg = QuickTaskDialog(self, emp, day, dict(t))
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        data = dlg.get_data()
        if not data["title"]:
            return
        try:
            update_task(tid, **data)
            self._reload_tasks()
        except Exception as e:
            error(self, "Error", str(e))

    def _rem_task(self):
        item = self.task_list.currentItem()
        if not item:
            return
        if confirm(self, "Remove", "Remove this task?"):
            delete_task(item.data(Qt.ItemDataRole.UserRole))
            self._reload_tasks()

    def get_data(self):
        return {
            "employee_pk": self.emp.currentData(),
            "audit_date": qdate_to_str(self.date.date()),
            "work_submitted": self.submitted.isChecked(),
            "late": self.late.isChecked(),
            "links_open": self.links.isChecked(),
            "missing_proof": self.missing.isChecked(),
            "completed_work": self.completed.toPlainText(),
            "proof_links": self.proofs.toPlainText(),
            "pending": self.pending.toPlainText(),
            "manager_remarks": self.remarks.toPlainText(),
        }


class DailyAuditPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 16, 20, 16)
        t = QLabel("Daily Audit (EOD)")
        t.setObjectName("pageTitle")
        lay.addWidget(t)
        bar = QHBoxLayout()
        bar.addWidget(QLabel("Date:"))
        self.date = QDateEdit()
        self.date.setCalendarPopup(True)
        self.date.setDate(QDate.currentDate())
        self.date.setDisplayFormat("yyyy-MM-dd")
        self.date.dateChanged.connect(self.refresh)
        bar.addWidget(self.date)
        for text, slot in [
            ("Record / Update", self._add),
            ("Pre-fill all", self._prefill),
            ("Add Task for Selected", self._quick_task),
        ]:
            b = btn_primary(text) if "Record" in text else btn_secondary(text)
            b.clicked.connect(slot)
            bar.addWidget(b)
        bar.addStretch()
        lay.addLayout(bar)
        cards = QHBoxLayout()
        self.sc = {}
        for k, lab in [("active", "Active"), ("rec", "Recorded"), ("miss", "Not recorded"),
                       ("sub", "Submitted"), ("late", "Late"), ("mp", "Missing proof"),
                       ("pend", "Has pending")]:
            c = StatCard(lab)
            self.sc[k] = c
            cards.addWidget(c)
        lay.addLayout(cards)
        self.table = QTableWidget(0, 9)
        self.table.setHorizontalHeaderLabels([
            "Employee", "ID", "Submitted", "Late", "Links", "Missing proof",
            "Pending?", "Tasks", "Completed (preview)",
        ])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(8, QHeaderView.ResizeMode.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.doubleClicked.connect(self._edit)
        lay.addWidget(self.table)
        row = QHBoxLayout()
        for text, slot in [("Edit", self._edit), ("Delete", self._delete),
                           ("Manage Tasks", self._manage)]:
            b = btn_danger(text) if text == "Delete" else btn_secondary(text)
            b.clicked.connect(slot)
            row.addWidget(b)
        row.addStretch()
        lay.addLayout(row)
        note = QLabel(
            "HR workflow: Pre-fill → fill each EOD → add/remove daily tasks per person. "
            "Export from Print Center → EOD Daily Audit."
        )
        note.setStyleSheet("color:#5a6a7a;font-size:11px;")
        note.setWordWrap(True)
        lay.addWidget(note)
        self._pk = {}
        self._emp = {}
        self.refresh()

    def refresh(self):
        day = qdate_to_str(self.date.date())
        rows = get_daily_audits(audit_date=day)
        s = get_daily_audit_summary(day)
        self.sc["active"].set_value(str(s["total_active"]))
        self.sc["rec"].set_value(str(s["recorded"]))
        self.sc["miss"].set_value(str(s["not_yet_recorded"]))
        self.sc["sub"].set_value(str(s["submitted"]))
        self.sc["late"].set_value(str(s["late"]))
        self.sc["mp"].set_value(str(s["missing_proof"]))
        self.sc["pend"].set_value(str(s["has_pending"]))
        self.table.setRowCount(len(rows))
        self._pk.clear()
        self._emp.clear()
        for i, r in enumerate(rows):
            self._pk[i] = r["id"]
            self._emp[i] = r["employee_id"]
            self.table.setItem(i, 0, QTableWidgetItem(r["full_name"] or ""))
            self.table.setItem(i, 1, QTableWidgetItem(r["emp_code"] or ""))
            self.table.setItem(i, 2, QTableWidgetItem("Yes" if r["work_submitted"] else "No"))
            self.table.setItem(i, 3, QTableWidgetItem("Yes" if r["late"] else "No"))
            self.table.setItem(i, 4, QTableWidgetItem("Yes" if r["links_open"] else "No"))
            self.table.setItem(i, 5, QTableWidgetItem("Yes" if r["missing_proof"] else "No"))
            self.table.setItem(i, 6, QTableWidgetItem("Yes" if (r["pending"] or "").strip() else "No"))
            n = len(get_tasks(task_date=day, employee_pk=r["employee_id"]))
            self.table.setItem(i, 7, QTableWidgetItem(str(n)))
            prev = (r["completed_work"] or "").replace("\n", " ")[:70]
            self.table.setItem(i, 8, QTableWidgetItem(prev))

    def _row(self):
        rows = self.table.selectionModel().selectedRows()
        return rows[0].row() if rows else None

    def _add(self):
        dlg = DailyAuditDialog(self, audit_date=qdate_to_str(self.date.date()))
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        data = dlg.get_data()
        if data["employee_pk"] is None:
            return
        try:
            upsert_daily_audit(**data)
            info(self, "Saved", "Audit saved.")
            self.refresh()
        except Exception as e:
            error(self, "Error", str(e))

    def _edit(self):
        r = self._row()
        if r is None:
            self._add()
            return
        rec = get_daily_audit(self._pk[r])
        if not rec:
            return
        dlg = DailyAuditDialog(self, data=dict(rec))
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        try:
            upsert_daily_audit(**dlg.get_data())
            self.refresh()
        except Exception as e:
            error(self, "Error", str(e))

    def _delete(self):
        r = self._row()
        if r is None:
            return
        if confirm(self, "Delete", "Delete this audit entry?"):
            delete_daily_audit(self._pk[r])
            self.refresh()

    def _quick_task(self):
        r = self._row()
        if r is None:
            warn(self, "Select", "Select an employee row first.")
            return
        day = qdate_to_str(self.date.date())
        dlg = QuickTaskDialog(self, self._emp[r], day)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        data = dlg.get_data()
        if not data["title"]:
            return
        try:
            add_task(**data)
            info(self, "OK", "Task added.")
            self.refresh()
        except Exception as e:
            error(self, "Error", str(e))

    def _manage(self):
        self._edit()

    def _prefill(self):
        day = qdate_to_str(self.date.date())
        emps = get_employees(active_only=True)
        if not emps:
            warn(self, "No employees", "Add employees first.")
            return
        if not confirm(self, "Pre-fill", f"Create empty rows for {len(emps)} employees on {day}?"):
            return
        n = 0
        for e in emps:
            if get_daily_audit_for(e["id"], day):
                continue
            upsert_daily_audit(e["id"], day, False, False, True, False, "", "", "")
            n += 1
        info(self, "Done", f"Created {n} entries.")
        self.refresh()


# ---- Tasks page (compact) ----
class TasksPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 16, 20, 16)
        t = QLabel("Daily Tasks")
        t.setObjectName("pageTitle")
        lay.addWidget(t)
        bar = QHBoxLayout()
        bar.addWidget(QLabel("Date:"))
        self.date = QDateEdit()
        self.date.setCalendarPopup(True)
        self.date.setDate(QDate.currentDate())
        self.date.setDisplayFormat("yyyy-MM-dd")
        self.date.dateChanged.connect(self.refresh)
        bar.addWidget(self.date)
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search…")
        self.search.textChanged.connect(self.refresh)
        bar.addWidget(self.search, stretch=1)
        add = btn_primary("Create Task")
        add.clicked.connect(self._add)
        bar.addWidget(add)
        copy = btn_secondary("Copy from previous day")
        copy.clicked.connect(self._copy)
        bar.addWidget(copy)
        lay.addLayout(bar)
        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(
            ["Title", "Assignees", "Status", "Priority", "Progress", "Proof"]
        )
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.doubleClicked.connect(self._edit)
        lay.addWidget(self.table)
        row = QHBoxLayout()
        e = btn_secondary("Edit")
        e.clicked.connect(self._edit)
        d = btn_danger("Delete")
        d.clicked.connect(self._delete)
        row.addWidget(e)
        row.addWidget(d)
        row.addStretch()
        lay.addLayout(row)
        self._pk = {}
        self.refresh()

    def refresh(self):
        day = qdate_to_str(self.date.date())
        tasks = get_tasks(task_date=day, search=self.search.text().strip())
        self.table.setRowCount(len(tasks))
        self._pk.clear()
        for i, t in enumerate(tasks):
            self._pk[i] = t["id"]
            self.table.setItem(i, 0, QTableWidgetItem(t["title"] or ""))
            self.table.setItem(i, 1, QTableWidgetItem(t["assignee_names"] or ""))
            self.table.setItem(i, 2, QTableWidgetItem(t["status"] or ""))
            self.table.setItem(i, 3, QTableWidgetItem(t["priority"] or ""))
            self.table.setItem(i, 4, QTableWidgetItem(f"{t['progress']}%"))
            self.table.setItem(i, 5, QTableWidgetItem(t["proof_url"] or ""))

    def _sel(self):
        rows = self.table.selectionModel().selectedRows()
        return self._pk.get(rows[0].row()) if rows else None

    def _form(self, data=None, assignees=None):
        dlg = QDialog(self)
        dlg.setWindowTitle("Edit Task" if data else "Create Task")
        dlg.setMinimumWidth(480)
        form = QFormLayout(dlg)
        title = QLineEdit()
        desc = QTextEdit()
        desc.setMaximumHeight(60)
        tdate = QDateEdit()
        tdate.setCalendarPopup(True)
        tdate.setDate(self.date.date())
        tdate.setDisplayFormat("yyyy-MM-dd")
        pri = QComboBox()
        pri.addItems(TASK_PRIORITIES)
        pri.setCurrentText("Medium")
        st = QComboBox()
        st.addItems(TASK_STATUSES)
        proof = QLineEdit()
        form.addRow("Title *", title)
        form.addRow("Description", desc)
        form.addRow("Date", tdate)
        form.addRow("Priority", pri)
        form.addRow("Status", st)
        form.addRow("Proof URL", proof)
        alist = QListWidget()
        alist.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        alist.setMaximumHeight(100)
        for e in get_employees(active_only=True):
            item = QListWidgetItem(f"{e['full_name']} ({e['employee_id']})")
            item.setData(Qt.ItemDataRole.UserRole, e["id"])
            alist.addItem(item)
        form.addRow("Assignees", alist)
        if data:
            title.setText(data.get("title", ""))
            desc.setPlainText(data.get("description") or "")
            tdate.setDate(str_to_qdate(data.get("task_date", "")))
            pri.setCurrentText(data.get("priority", "Medium"))
            st.setCurrentText(data.get("status", "Not Started"))
            proof.setText(data.get("proof_url") or "")
            for i in range(alist.count()):
                if alist.item(i).data(Qt.ItemDataRole.UserRole) in (assignees or []):
                    alist.item(i).setSelected(True)
        br = QHBoxLayout()
        br.addStretch()
        c = btn_secondary("Cancel")
        c.clicked.connect(dlg.reject)
        s = btn_primary("Save")
        s.clicked.connect(dlg.accept)
        br.addWidget(c)
        br.addWidget(s)
        form.addRow(br)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return None
        return {
            "title": title.text().strip(),
            "description": desc.toPlainText().strip(),
            "task_date": qdate_to_str(tdate.date()),
            "due_time": "",
            "priority": pri.currentText(),
            "status": st.currentText(),
            "progress": 100 if st.currentText() == "Completed" else 0,
            "proof_url": proof.text().strip(),
            "manager_remarks": "",
            "assignee_ids": [it.data(Qt.ItemDataRole.UserRole) for it in alist.selectedItems()],
        }

    def _add(self):
        data = self._form()
        if not data or not data["title"]:
            return
        try:
            add_task(**data)
            self.refresh()
        except Exception as e:
            error(self, "Error", str(e))

    def _edit(self):
        pk = self._sel()
        if pk is None:
            return
        t = get_task(pk)
        data = self._form(dict(t) if t else None, get_task_assignees(pk))
        if not data or not data["title"]:
            return
        try:
            update_task(pk, **data)
            self.refresh()
        except Exception as e:
            error(self, "Error", str(e))

    def _delete(self):
        pk = self._sel()
        if pk and confirm(self, "Delete", "Delete this task?"):
            delete_task(pk)
            self.refresh()

    def _copy(self):
        target = qdate_to_str(self.date.date())
        source = qdate_to_str(self.date.date().addDays(-1))
        if confirm(self, "Copy", f"Copy tasks from {source} to {target}?"):
            n = copy_tasks_from_date(source, target)
            info(self, "Copied", f"{n} task(s) copied.")
            self.refresh()


# ---- Attendance (compact) ----
class AttendancePage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 16, 20, 16)
        t = QLabel("Attendance")
        t.setObjectName("pageTitle")
        lay.addWidget(t)
        bar = QHBoxLayout()
        self.emp = QComboBox()
        self.emp.addItem("All", None)
        for e in get_employees():
            self.emp.addItem(f"{e['full_name']} ({e['employee_id']})", e["id"])
        self.emp.currentIndexChanged.connect(self.refresh)
        bar.addWidget(self.emp)
        self.dfrom = QDateEdit()
        self.dfrom.setCalendarPopup(True)
        self.dfrom.setDate(QDate.currentDate().addDays(-7))
        self.dfrom.setDisplayFormat("yyyy-MM-dd")
        self.dfrom.dateChanged.connect(self.refresh)
        bar.addWidget(QLabel("From:"))
        bar.addWidget(self.dfrom)
        self.dto = QDateEdit()
        self.dto.setCalendarPopup(True)
        self.dto.setDate(QDate.currentDate())
        self.dto.setDisplayFormat("yyyy-MM-dd")
        self.dto.dateChanged.connect(self.refresh)
        bar.addWidget(QLabel("To:"))
        bar.addWidget(self.dto)
        add = btn_primary("Record")
        add.clicked.connect(self._add)
        bar.addWidget(add)
        lay.addLayout(bar)
        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(
            ["Date", "Employee", "In", "Out", "Duration", "Remarks"]
        )
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        lay.addWidget(self.table)
        row = QHBoxLayout()
        e = btn_secondary("Edit")
        e.clicked.connect(self._edit)
        d = btn_danger("Delete")
        d.clicked.connect(self._delete)
        row.addWidget(e)
        row.addWidget(d)
        row.addStretch()
        lay.addLayout(row)
        self._pk = {}
        self.refresh()

    def refresh(self):
        recs = get_attendance(
            employee_pk=self.emp.currentData(),
            date_from=qdate_to_str(self.dfrom.date()),
            date_to=qdate_to_str(self.dto.date()),
        )
        self.table.setRowCount(len(recs))
        self._pk.clear()
        for i, a in enumerate(recs):
            self._pk[i] = a["id"]
            self.table.setItem(i, 0, QTableWidgetItem(a["attendance_date"] or ""))
            self.table.setItem(i, 1, QTableWidgetItem(a["full_name"] or ""))
            self.table.setItem(i, 2, QTableWidgetItem(a["check_in"] or ""))
            self.table.setItem(i, 3, QTableWidgetItem(a["check_out"] or ""))
            self.table.setItem(i, 4, QTableWidgetItem(calc_duration(a["check_in"], a["check_out"])))
            self.table.setItem(i, 5, QTableWidgetItem(a["remarks"] or ""))

    def _sel(self):
        rows = self.table.selectionModel().selectedRows()
        return self._pk.get(rows[0].row()) if rows else None

    def _form(self, data=None):
        dlg = QDialog(self)
        dlg.setWindowTitle("Edit Attendance" if data else "Record Attendance")
        form = QFormLayout(dlg)
        emp = QComboBox()
        for e in get_employees(active_only=True):
            emp.addItem(f"{e['full_name']} ({e['employee_id']})", e["id"])
        day = QDateEdit()
        day.setCalendarPopup(True)
        day.setDate(QDate.currentDate())
        day.setDisplayFormat("yyyy-MM-dd")
        cin = QTimeEdit()
        cin.setDisplayFormat("HH:mm")
        cin.setTime(QTime(9, 0))
        cout = QTimeEdit()
        cout.setDisplayFormat("HH:mm")
        cout.setTime(QTime(18, 0))
        miss = QCheckBox("Missing check-out")
        rem = QTextEdit()
        rem.setMaximumHeight(50)
        form.addRow("Employee", emp)
        form.addRow("Date", day)
        form.addRow("Check-in", cin)
        form.addRow("Check-out", cout)
        form.addRow("", miss)
        form.addRow("Remarks", rem)
        if data:
            idx = emp.findData(data.get("employee_id"))
            if idx >= 0:
                emp.setCurrentIndex(idx)
            emp.setEnabled(False)
            day.setDate(str_to_qdate(data.get("attendance_date", "")))
            day.setEnabled(False)
            cin.setTime(str_to_qtime(data.get("check_in")))
            if data.get("check_out"):
                cout.setTime(str_to_qtime(data["check_out"]))
            else:
                miss.setChecked(True)
            rem.setPlainText(data.get("remarks") or "")
        br = QHBoxLayout()
        br.addStretch()
        c = btn_secondary("Cancel")
        c.clicked.connect(dlg.reject)
        s = btn_primary("Save")
        s.clicked.connect(dlg.accept)
        br.addWidget(c)
        br.addWidget(s)
        form.addRow(br)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return None
        return {
            "employee_pk": emp.currentData(),
            "attendance_date": qdate_to_str(day.date()),
            "check_in": qtime_to_str(cin.time()),
            "check_out": "" if miss.isChecked() else qtime_to_str(cout.time()),
            "remarks": rem.toPlainText().strip(),
        }

    def _add(self):
        data = self._form()
        if not data:
            return
        try:
            add_attendance(**data)
            self.refresh()
        except Exception as e:
            error(self, "Error", str(e))

    def _edit(self):
        pk = self._sel()
        if pk is None:
            return
        rec = get_attendance_record(pk)
        data = self._form(dict(rec) if rec else None)
        if not data:
            return
        try:
            update_attendance(pk, data["check_in"], data["check_out"], data["remarks"])
            self.refresh()
        except Exception as e:
            error(self, "Error", str(e))

    def _delete(self):
        pk = self._sel()
        if pk and confirm(self, "Delete", "Delete this record?"):
            delete_attendance(pk)
            self.refresh()


# ---- Print Center ----
class PrintCenterPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._last: Optional[Path] = None
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 16, 20, 16)
        t = QLabel("Print Center")
        t.setObjectName("pageTitle")
        lay.addWidget(t)
        form = QFormLayout()
        self.rtype = QComboBox()
        self.rtype.addItems([
            "EOD Daily Audit PDF",
            "Employees CSV",
            "Tasks CSV (date range)",
        ])
        form.addRow("Report:", self.rtype)
        self.dfrom = QDateEdit()
        self.dfrom.setCalendarPopup(True)
        self.dfrom.setDate(QDate.currentDate())
        self.dfrom.setDisplayFormat("yyyy-MM-dd")
        form.addRow("Date / From:", self.dfrom)
        self.dto = QDateEdit()
        self.dto.setCalendarPopup(True)
        self.dto.setDate(QDate.currentDate())
        self.dto.setDisplayFormat("yyyy-MM-dd")
        form.addRow("To:", self.dto)
        lay.addLayout(form)
        row = QHBoxLayout()
        exp = btn_primary("Export")
        exp.clicked.connect(self._export)
        op = btn_secondary("Open last / folder")
        op.clicked.connect(self._open)
        row.addWidget(exp)
        row.addWidget(op)
        row.addStretch()
        lay.addLayout(row)
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumHeight(120)
        lay.addWidget(self.log)
        lay.addStretch()

    def _export(self):
        r = self.rtype.currentText()
        d1 = qdate_to_str(self.dfrom.date())
        d2 = qdate_to_str(self.dto.date())
        try:
            if r.startswith("EOD"):
                path = generate_eod_pdf(d1)
            elif r.startswith("Employees"):
                path = export_employees_csv()
            else:
                path = export_tasks_csv(d1, d2)
            self._last = path
            self.log.append(f"Created: {path}")
            info(self, "Done", f"Saved to:\n{path}")
        except Exception as e:
            error(self, "Export failed", str(e))
            self.log.append(f"Error: {e}")

    def _open(self):
        if self._last and self._last.exists():
            open_path(self._last)
        else:
            open_path(REPORTS_DIR)


# ---- Backup ----
class BackupPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 16, 20, 16)
        t = QLabel("Backup & Restore")
        t.setObjectName("pageTitle")
        lay.addWidget(t)
        row = QHBoxLayout()
        self.label = QLineEdit()
        self.label.setPlaceholderText("Optional label")
        row.addWidget(self.label, stretch=1)
        c = btn_primary("Create Backup")
        c.clicked.connect(self._create)
        row.addWidget(c)
        lay.addLayout(row)
        self.list = QListWidget()
        lay.addWidget(self.list)
        row2 = QHBoxLayout()
        for text, slot in [("Refresh", self._refresh), ("Restore selected", self._restore),
                           ("Open folder", lambda: open_path(BACKUP_DIR))]:
            b = btn_secondary(text)
            b.clicked.connect(slot)
            row2.addWidget(b)
        row2.addStretch()
        lay.addLayout(row2)
        lay.addWidget(QLabel(f"DB: {DB_PATH}\nBackups: {BACKUP_DIR}"))
        self._refresh()

    def _refresh(self):
        self.list.clear()
        for p in list_backups():
            self.list.addItem(f"{p.name}  ({p.stat().st_size // 1024} KB)")

    def _create(self):
        try:
            path = create_backup(self.label.text().strip())
            info(self, "Backup", f"Saved:\n{path}")
            self._refresh()
        except Exception as e:
            error(self, "Error", str(e))

    def _restore(self):
        item = self.list.currentItem()
        if not item:
            return
        name = item.text().split("  (")[0]
        path = BACKUP_DIR / name
        if confirm(self, "Restore", f"Restore {name}? Current DB will be backed up first."):
            try:
                restore_backup(path)
                info(self, "Restored", "Please restart the application.")
                self._refresh()
            except Exception as e:
                error(self, "Error", str(e))


# ---- Settings ----
class SettingsPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 16, 20, 16)
        t = QLabel("Settings")
        t.setObjectName("pageTitle")
        lay.addWidget(t)
        grp = QGroupBox("Change HR / Admin PIN")
        form = QFormLayout(grp)
        self.old = QLineEdit()
        self.old.setEchoMode(QLineEdit.EchoMode.Password)
        self.new = QLineEdit()
        self.new.setEchoMode(QLineEdit.EchoMode.Password)
        self.conf = QLineEdit()
        self.conf.setEchoMode(QLineEdit.EchoMode.Password)
        form.addRow("Current:", self.old)
        form.addRow("New:", self.new)
        form.addRow("Confirm:", self.conf)
        b = btn_primary("Change PIN")
        b.clicked.connect(self._change)
        form.addRow("", b)
        lay.addWidget(grp)
        about = QGroupBox("About")
        al = QVBoxLayout(about)
        al.addWidget(QLabel(f"<b>{APP_DISPLAY_NAME}</b> v{APP_VERSION}"))
        al.addWidget(QLabel(f"Data: {DATA_DIR}"))
        al.addWidget(QLabel(f"Database: {DB_PATH}"))
        al.addWidget(QLabel(
            "Login is for HR/Admin only. HR can also be listed under Employees "
            "to record their own EOD. Local PIN protects against casual access only."
        ))
        lay.addWidget(about)
        lay.addStretch()

    def _change(self):
        if self.new.text() != self.conf.text():
            warn(self, "Mismatch", "New PINs do not match.")
            return
        try:
            change_password(self.old.text(), self.new.text())
            info(self, "OK", "PIN updated.")
            self.old.clear()
            self.new.clear()
            self.conf.clear()
        except Exception as e:
            error(self, "Failed", str(e))


# ===========================================================================
# MAIN WINDOW
# ===========================================================================
NAV = [
    ("Dashboard", "dashboard"),
    ("Employees", "employees"),
    ("Daily Audit", "daily_audit"),
    ("Daily Tasks", "tasks"),
    ("Attendance", "attendance"),
    ("Print Center", "print"),
    ("Backup", "backup"),
    ("Settings", "settings"),
]


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_DISPLAY_NAME}  v{APP_VERSION}")
        self.setMinimumSize(1100, 700)
        self.resize(1280, 800)
        self.setStyleSheet(APP_STYLESHEET)
        central = QWidget()
        self.setCentralWidget(central)
        root = QHBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)
        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(190)
        sl = QVBoxLayout(sidebar)
        sl.setContentsMargins(8, 16, 8, 16)
        lab = QLabel(APP_DISPLAY_NAME)
        lab.setStyleSheet("color:white;font-size:13px;font-weight:700;padding:4px 8px 12px 8px;")
        lab.setWordWrap(True)
        sl.addWidget(lab)
        self.nav_btns = {}
        for text, key in NAV:
            b = QPushButton(text)
            b.setObjectName("navBtn")
            b.setCursor(Qt.CursorShape.PointingHandCursor)
            b.setProperty("active", "false")
            b.clicked.connect(lambda checked=False, k=key: self._go(k))
            sl.addWidget(b)
            self.nav_btns[key] = b
        sl.addStretch()
        ver = QLabel(f"v{APP_VERSION}")
        ver.setStyleSheet("color:#8aa0b8;font-size:11px;padding:8px;")
        sl.addWidget(ver)
        root.addWidget(sidebar)
        self.stack = QStackedWidget()
        self.pages = {
            "dashboard": DashboardPage(),
            "employees": EmployeesPage(),
            "daily_audit": DailyAuditPage(),
            "tasks": TasksPage(),
            "attendance": AttendancePage(),
            "print": PrintCenterPage(),
            "backup": BackupPage(),
            "settings": SettingsPage(),
        }
        for _, key in NAV:
            self.stack.addWidget(self.pages[key])
        root.addWidget(self.stack, stretch=1)
        self._go("dashboard")

    def _go(self, key: str):
        keys = [k for _, k in NAV]
        if key not in keys:
            return
        self.stack.setCurrentIndex(keys.index(key))
        for k, b in self.nav_btns.items():
            b.setProperty("active", "true" if k == key else "false")
            b.style().unpolish(b)
            b.style().polish(b)
        page = self.pages.get(key)
        if page and hasattr(page, "refresh"):
            try:
                page.refresh()
            except Exception:
                pass


# ===========================================================================
# ENTRY
# ===========================================================================
def main() -> int:
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )
    app = QApplication(sys.argv)
    app.setApplicationName(APP_DISPLAY_NAME)
    app.setApplicationVersion(APP_VERSION)
    app.setStyle("Fusion")
    try:
        initialize_database()
    except Exception as e:
        QMessageBox.critical(None, "Database Error", f"{e}\n\n{traceback.format_exc()}")
        return 1
    if not is_setup_complete():
        if FirstRunDialog().exec() != QDialog.DialogCode.Accepted:
            return 0
    else:
        if LoginDialog().exec() != QDialog.DialogCode.Accepted:
            return 0
    win = MainWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
