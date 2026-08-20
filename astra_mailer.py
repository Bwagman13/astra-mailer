"""
Astra Mailer — Personalized Mass Email Generator
A desktop application for generating personalized one-on-one meeting emails.

Usage:
    python astra_mailer.py

Requires:
    pip install PySide6 openpyxl anthropic python-dotenv
"""

import sys
import os
import json
import csv
import re
import html
import platform
import datetime
import base64
from pathlib import Path
from typing import Optional

# ─── Friendly error if dependencies are missing ─────────────────────────────

_missing = []
for _pkg, _import in [("PySide6", "PySide6"), ("openpyxl", "openpyxl"),
                        ("anthropic", "anthropic"), ("python-dotenv", "dotenv")]:
    try:
        __import__(_import)
    except ImportError:
        _missing.append(_pkg)

if _missing:
    print("\n" + "="*55)
    print("  Astra Mailer — Missing Required Packages")
    print("="*55)
    print(f"\n  The following packages need to be installed:\n")
    for p in _missing:
        print(f"    - {p}")
    print(f"\n  To fix this, run:\n")
    print(f"    pip install {' '.join(_missing)}")
    print(f"\n  Or run INSTALL.bat to set everything up automatically.\n")
    print("="*55 + "\n")
    try:
        input("  Press Enter to close...")
    except Exception:
        pass
    sys.exit(1)

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QLabel, QPushButton, QTextEdit, QPlainTextEdit,
    QTableWidget, QTableWidgetItem, QFileDialog, QComboBox,
    QProgressBar, QMessageBox, QGroupBox, QFormLayout, QLineEdit,
    QHeaderView, QSplitter, QCheckBox, QStatusBar, QFrame,
    QScrollArea, QSizePolicy, QSpacerItem
)
from PySide6.QtCore import Qt, QThread, Signal, QSize, QSettings
from PySide6.QtGui import QFont, QColor, QPalette, QIcon, QAction, QPixmap

from dotenv import load_dotenv

# Load .env from app directory
load_dotenv(Path(__file__).parent / ".env")


# ─── Platform Detection ──────────────────────────────────────────────────────

IS_WINDOWS = platform.system() == "Windows"
IS_MAC = platform.system() == "Darwin"
IS_LINUX = platform.system() == "Linux"

def get_system_font() -> str:
    if IS_MAC:
        return ".AppleSystemUIFont"
    elif IS_WINDOWS:
        return "Segoe UI"
    else:
        return "Ubuntu"


# ─── Constants ───────────────────────────────────────────────────────────────

APP_NAME = "Astra Mailer"
APP_VERSION = "1.0.0"
SETTINGS_ORG = "AstraMailer"

AVATAR_PATH = Path(__file__).parent / "Astra.png"
ICON_PATH = Path(__file__).parent / "astra_icon.ico"
ICON_PNG_PATH = Path(__file__).parent / "astra_icon.png"

COLUMN_HINTS = {
    "name": ["name", "student name", "student", "full name", "first name", "firstname", "last name", "lastname"],
    "email": ["email", "e-mail", "email address", "student email", "mail"],
    "meeting_time": ["meeting time", "meeting", "time", "date", "scheduled", "appointment",
                     "meeting date", "date/time", "datetime", "slot", "session"],
    "class": ["class", "course", "section", "subject", "period"],
    "topic": ["topic", "agenda", "discussion", "notes", "meeting topic"],
    "advisor": ["advisor", "adviser", "counselor", "teacher", "instructor"],
}

TONE_OPTIONS = ["Professional", "Warm", "Concise"]

DEFAULT_TEMPLATE = """Dear {name},

I hope this message finds you well. I'd like to confirm our upcoming one-on-one meeting scheduled for {meeting_time}.

This will be a great opportunity for us to discuss your progress and any questions you may have.

Please let me know if you need to reschedule.

Best regards,
Astra"""


# ─── Anime-Inspired Stylesheet ──────────────────────────────────────────────

def get_stylesheet() -> str:
    groupbox_margin_top = "20px" if IS_MAC else "16px"
    groupbox_padding_top = "32px" if IS_MAC else "28px"
    combo_padding = "8px 14px" if IS_MAC else "7px 12px"
    mono_font = "Menlo" if IS_MAC else "Consolas" if IS_WINDOWS else "Ubuntu Mono"

    return f"""
/* ── Main Window ── */
QMainWindow {{
    background-color: #fdf2f8;
}}

/* ── Tabs ── */
QTabWidget::pane {{
    border: 2px solid #f9a8d4;
    background: #fffbfe;
    border-radius: 12px;
}}
QTabBar::tab {{
    background: #fce7f3;
    color: #9d174d;
    padding: 10px 22px;
    margin-right: 3px;
    border-top-left-radius: 10px;
    border-top-right-radius: 10px;
    font-size: 13px;
    font-weight: 600;
    border: 1px solid #f9a8d4;
    border-bottom: none;
}}
QTabBar::tab:selected {{
    background: #fffbfe;
    color: #be185d;
    border-bottom: 3px solid #ec4899;
}}
QTabBar::tab:hover:!selected {{
    background: #fbcfe8;
}}

/* ── Group Boxes ── */
QGroupBox {{
    font-weight: 600;
    font-size: 13px;
    color: #831843;
    border: 2px solid #f9a8d4;
    border-radius: 12px;
    margin-top: {groupbox_margin_top};
    padding: 16px;
    padding-top: {groupbox_padding_top};
    background: #fffbfe;
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    left: 14px;
    padding: 0 8px;
    color: #be185d;
}}

/* ── Buttons ── */
QPushButton {{
    background-color: #ec4899;
    color: white;
    border: none;
    padding: 9px 20px;
    border-radius: 10px;
    font-weight: 600;
    font-size: 13px;
}}
QPushButton:hover {{
    background-color: #db2777;
}}
QPushButton:pressed {{
    background-color: #be185d;
}}
QPushButton:disabled {{
    background-color: #d1d5db;
    color: #9ca3af;
}}
QPushButton[secondary="true"] {{
    background-color: #fce7f3;
    color: #9d174d;
    border: 1px solid #f9a8d4;
}}
QPushButton[secondary="true"]:hover {{
    background-color: #fbcfe8;
}}
QPushButton[danger="true"] {{
    background-color: #ef4444;
}}
QPushButton[danger="true"]:hover {{
    background-color: #dc2626;
}}

/* ── Text Inputs — EXPLICIT dark text color ── */
QTextEdit, QPlainTextEdit {{
    border: 2px solid #f9a8d4;
    border-radius: 10px;
    padding: 10px;
    font-size: 13px;
    color: #1f2937;
    background: #ffffff;
    selection-background-color: #ec4899;
    selection-color: white;
}}
QTextEdit:focus, QPlainTextEdit:focus {{
    border-color: #ec4899;
    background: #fff5f9;
}}
QLineEdit {{
    border: 2px solid #f9a8d4;
    border-radius: 10px;
    padding: 8px 12px;
    font-size: 13px;
    color: #1f2937;
    background: #ffffff;
    selection-background-color: #ec4899;
    selection-color: white;
    min-height: 20px;
}}
QLineEdit:focus {{
    border-color: #ec4899;
    background: #fff5f9;
}}

/* ── Combo Boxes ── */
QComboBox {{
    border: 2px solid #f9a8d4;
    border-radius: 10px;
    padding: {combo_padding};
    font-size: 13px;
    color: #1f2937;
    background: white;
}}
QComboBox:focus {{
    border-color: #ec4899;
}}
QComboBox QAbstractItemView {{
    color: #1f2937;
    background: white;
    selection-background-color: #fce7f3;
    selection-color: #9d174d;
}}

/* ── Tables ── */
QTableWidget {{
    border: 2px solid #f9a8d4;
    border-radius: 10px;
    gridline-color: #fce7f3;
    font-size: 12px;
    color: #1f2937;
    background: white;
    alternate-background-color: #fef7fb;
}}
QTableWidget::item {{
    padding: 6px;
    color: #1f2937;
}}
QTableWidget::item:selected {{
    background-color: #fce7f3;
    color: #9d174d;
}}
QHeaderView::section {{
    background-color: #fce7f3;
    color: #9d174d;
    padding: 8px;
    border: none;
    border-right: 1px solid #f9a8d4;
    border-bottom: 1px solid #f9a8d4;
    font-weight: 600;
    font-size: 12px;
}}

/* ── Progress Bar ── */
QProgressBar {{
    border: 2px solid #f9a8d4;
    border-radius: 10px;
    text-align: center;
    font-size: 12px;
    height: 24px;
    color: #9d174d;
    background: #fce7f3;
}}
QProgressBar::chunk {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #ec4899, stop:1 #a855f7);
    border-radius: 8px;
}}

/* ── Labels ── */
QLabel {{
    color: #6b7280;
    font-size: 13px;
}}

/* ── Status Bar ── */
QStatusBar {{
    background: #fce7f3;
    color: #9d174d;
    font-size: 12px;
    border-top: 1px solid #f9a8d4;
}}

/* ── Scroll Area ── */
QScrollArea {{
    border: none;
    background: #fffbfe;
}}
QScrollArea > QWidget > QWidget {{
    background: #fffbfe;
}}

/* ── Form labels ── */
QFormLayout QLabel {{
    color: #9d174d;
    font-weight: 500;
}}
"""


# ─── Spreadsheet Parser ─────────────────────────────────────────────────────

def parse_spreadsheet(filepath: str) -> tuple[list[str], list[dict]]:
    ext = Path(filepath).suffix.lower()
    if ext == ".csv":
        for encoding in ["utf-8-sig", "utf-8", "latin-1"]:
            try:
                with open(filepath, "r", encoding=encoding) as f:
                    reader = csv.DictReader(f)
                    headers = reader.fieldnames or []
                    rows = [row for row in reader if any(v.strip() for v in row.values() if v)]
                return headers, rows
            except UnicodeDecodeError:
                continue
        raise ValueError("Could not decode CSV file. Try saving it as UTF-8.")
    elif ext in (".xlsx", ".xls"):
        import openpyxl
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        ws = wb.active
        all_rows = list(ws.iter_rows(values_only=True))
        wb.close()
        if not all_rows:
            raise ValueError("Spreadsheet is empty.")
        headers = [str(h).strip() if h else f"Column_{i+1}" for i, h in enumerate(all_rows[0])]
        rows = []
        for row in all_rows[1:]:
            if not any(cell is not None and str(cell).strip() for cell in row):
                continue
            row_dict = {}
            for i, header in enumerate(headers):
                val = row[i] if i < len(row) else None
                row_dict[header] = str(val).strip() if val is not None else ""
            rows.append(row_dict)
        return headers, rows
    else:
        raise ValueError(f"Unsupported file type: {ext}. Please use .csv or .xlsx")


def auto_map_columns(headers: list[str]) -> dict[str, str]:
    mapping = {}
    headers_lower = {h: h.lower().strip() for h in headers}
    for role, hints in COLUMN_HINTS.items():
        best_match = None
        for header, header_low in headers_lower.items():
            if header_low in hints:
                best_match = header
                break
            for hint in hints:
                if hint in header_low or header_low in hint:
                    best_match = header
                    break
            if best_match:
                break
        if best_match:
            mapping[role] = best_match
    return mapping


# ─── AI Personalization Engine ───────────────────────────────────────────────

PERSONALIZATION_PROMPT = """You are an email personalization assistant. Your job is to lightly personalize a boilerplate email template for a specific student based on their data.

RULES:
- Preserve the original intent, structure, and tone of the template
- Naturally incorporate the student's specific details (name, meeting time, class, etc.)
- Do NOT invent or hallucinate any information not provided
- Do NOT add extra paragraphs or significantly lengthen the email
- Do NOT use overly enthusiastic or fake-sounding language
- Keep it {tone} in tone
- If a field is missing or empty, simply skip it — do not mention it's missing
- Return ONLY the personalized email body, no subject line, no extra commentary

TEMPLATE:
{template}

STUDENT DATA:
{student_data}

Write the personalized email now:"""


class EmailGeneratorThread(QThread):
    progress = Signal(int, str)
    email_ready = Signal(int, str)
    error = Signal(int, str)
    finished_all = Signal()

    def __init__(self, template, rows, column_map, tone, api_key):
        super().__init__()
        self.template = template
        self.rows = rows
        self.column_map = column_map
        self.tone = tone
        self.api_key = api_key
        self._cancelled = False

    def cancel(self):
        self._cancelled = True

    def run(self):
        import anthropic
        client = anthropic.Anthropic(api_key=self.api_key)
        for i, row in enumerate(self.rows):
            if self._cancelled:
                break
            self.progress.emit(i, f"Generating email for row {i+1}...")
            student_data_parts = []
            for role, header in self.column_map.items():
                value = row.get(header, "").strip()
                if value:
                    student_data_parts.append(f"- {role.replace('_', ' ').title()}: {value}")
            mapped_headers = set(self.column_map.values())
            for header, value in row.items():
                if header not in mapped_headers and value.strip():
                    student_data_parts.append(f"- {header}: {value}")
            if not student_data_parts:
                self.error.emit(i, "No data found for this row")
                continue
            student_data_str = "\n".join(student_data_parts)
            prompt = PERSONALIZATION_PROMPT.format(
                tone=self.tone.lower(),
                template=self.template,
                student_data=student_data_str
            )
            try:
                response = client.messages.create(
                    model="claude-sonnet-4-20250514",
                    max_tokens=1024,
                    messages=[{"role": "user", "content": prompt}]
                )
                email_text = response.content[0].text.strip()
                self.email_ready.emit(i, email_text)
            except Exception as e:
                self.error.emit(i, str(e))
        self.finished_all.emit()


# ─── Outlook Integration (Optional, Windows only) ───────────────────────────

def check_outlook_available() -> bool:
    """Check if Outlook COM is available, with a timeout to prevent hanging."""
    if not IS_WINDOWS:
        return False
    try:
        import threading
        result = [False]

        def _try_outlook():
            try:
                import win32com.client
                outlook = win32com.client.Dispatch("Outlook.Application")
                _ = outlook.GetNamespace("MAPI")
                result[0] = True
            except Exception:
                result[0] = False

        t = threading.Thread(target=_try_outlook, daemon=True)
        t.start()
        t.join(timeout=5)  # Give it 5 seconds max
        return result[0]
    except Exception:
        return False


def create_outlook_drafts(emails):
    if not IS_WINDOWS:
        return 0, ["Outlook drafts are only supported on Windows"]
    import win32com.client
    outlook = win32com.client.Dispatch("Outlook.Application")
    successes = 0
    errors = []
    for em in emails:
        try:
            mail = outlook.CreateItem(0)
            mail.To = em["to"]
            mail.Subject = em["subject"]
            mail.Body = em["body"]
            mail.Save()
            successes += 1
        except Exception as e:
            errors.append(f"{em.get('to', '?')}: {e}")
    return successes, errors


# ─── Export Functions ────────────────────────────────────────────────────────

def export_to_csv(emails, filepath):
    with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["Name", "Email", "Subject", "Body"])
        writer.writeheader()
        for em in emails:
            writer.writerow({
                "Name": em.get("name", ""),
                "Email": em.get("to", ""),
                "Subject": em.get("subject", ""),
                "Body": em.get("body", "")
            })


def export_to_html(emails, filepath):
    parts = ["""<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>Astra Mailer — Email Export</title>
<style>
body{font-family:Segoe UI,sans-serif;max-width:800px;margin:40px auto;background:#fdf2f8;color:#1f2937}
.email-card{background:white;border:2px solid #f9a8d4;border-radius:12px;padding:24px;margin:20px 0}
.email-header{display:flex;justify-content:space-between;border-bottom:1px solid #fce7f3;padding-bottom:12px;margin-bottom:12px}
.email-to{font-weight:600;color:#be185d}.email-subject{color:#9d174d;font-size:14px}
.email-body{white-space:pre-wrap;line-height:1.6;font-size:14px}
h1{color:#be185d;font-size:24px}
.meta{color:#9ca3af;font-size:12px;margin-top:8px}
</style></head><body>
<h1>Astra Mailer — Generated Emails</h1>
<p class="meta">Generated on """ + datetime.datetime.now().strftime("%B %d, %Y at %I:%M %p") + f" — {len(emails)} emails</p>"]
    for i, em in enumerate(emails, 1):
        parts.append(f"""<div class="email-card">
<div class="email-header">
<div><span class="email-to">{html.escape(em.get('name', ''))}</span> &lt;{html.escape(em.get('to', ''))}&gt;</div>
<div class="email-subject">Subject: {html.escape(em.get('subject', ''))}</div>
</div>
<div class="email-body">{html.escape(em.get('body', ''))}</div>
</div>""")
    parts.append("</body></html>")
    with open(filepath, "w", encoding="utf-8") as f:
        f.write("\n".join(parts))


def export_individual_txt(emails, folder):
    os.makedirs(folder, exist_ok=True)
    for i, em in enumerate(emails, 1):
        safe_name = re.sub(r'[^\w\s-]', '', em.get("name", f"student_{i}")).strip().replace(" ", "_")
        filepath = os.path.join(folder, f"{i:03d}_{safe_name}.txt")
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(f"To: {em.get('to', '')}\n")
            f.write(f"Subject: {em.get('subject', '')}\n")
            f.write(f"{'='*50}\n\n")
            f.write(em.get("body", ""))


# ─── Main Window ─────────────────────────────────────────────────────────────

class AstraMailerWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} v{APP_VERSION}")
        self.setMinimumSize(960, 720)
        self.resize(1080, 800)

        self.settings = QSettings(SETTINGS_ORG, APP_NAME)

        # State
        self.spreadsheet_headers = []
        self.spreadsheet_rows = []
        self.column_map = {}
        self.generated_emails = []
        self.generator_thread = None
        self.outlook_available = False

        # Load API key from .env (already loaded by dotenv)
        self._env_api_key = os.environ.get("ANTHROPIC_API_KEY", "").strip()

        print("[Astra] Building UI...")
        self._setup_ui()
        print("[Astra] Restoring state...")
        self._restore_state()
        print("[Astra] Checking Outlook...")
        self._check_outlook()
        print("[Astra] Ready!")

    def _get_api_key(self) -> str:
        """Get API key: prefer .env, fall back to manual input field."""
        if self._env_api_key:
            return self._env_api_key
        return self.api_key_input.text().strip()

    # ── UI Setup ──────────────────────────────────────────────────────────

    def _setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(20, 12, 20, 8)
        main_layout.setSpacing(6)

        # ── Header with Avatar ──
        header_widget = QWidget()
        header_layout = QHBoxLayout(header_widget)
        header_layout.setContentsMargins(8, 0, 8, 0)
        header_layout.setSpacing(14)

        # Avatar image
        avatar_label = QLabel()
        if AVATAR_PATH.exists():
            pixmap = QPixmap(str(AVATAR_PATH))
            scaled = pixmap.scaled(80, 80, Qt.AspectRatioMode.KeepAspectRatio,
                                   Qt.TransformationMode.SmoothTransformation)
            avatar_label.setPixmap(scaled)
        else:
            avatar_label.setText("✉")
            avatar_label.setFont(QFont(get_system_font(), 36))
        avatar_label.setFixedSize(85, 85)
        avatar_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        avatar_label.setStyleSheet("background: transparent;")
        header_layout.addWidget(avatar_label)

        # Title text
        title_block = QWidget()
        title_layout = QVBoxLayout(title_block)
        title_layout.setContentsMargins(0, 8, 0, 8)
        title_layout.setSpacing(2)

        header = QLabel(APP_NAME)
        header.setFont(QFont(get_system_font(), 22, QFont.Weight.Bold))
        header.setStyleSheet("color: #be185d; background: transparent;")
        title_layout.addWidget(header)

        subtitle = QLabel("Personalized emails for your students, powered by AI ")
        subtitle.setStyleSheet("color: #9d174d; font-size: 13px; background: transparent;")
        title_layout.addWidget(subtitle)

        # API status indicator in header
        self.api_status_label = QLabel()
        self.api_status_label.setStyleSheet("font-size: 11px; background: transparent;")
        if self._env_api_key:
            self.api_status_label.setText("♦ API key loaded from config")
            self.api_status_label.setStyleSheet("color: #059669; font-size: 11px; background: transparent;")
        else:
            self.api_status_label.setText("⚠ No API key found — enter one in Setup tab")
            self.api_status_label.setStyleSheet("color: #d97706; font-size: 11px; background: transparent;")
        title_layout.addWidget(self.api_status_label)

        header_layout.addWidget(title_block)
        header_layout.addStretch()

        main_layout.addWidget(header_widget)

        # ── Tabs ──
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        self.tabs.addTab(self._build_setup_tab(), "① Setup")
        self.tabs.addTab(self._build_data_tab(), "② Data")
        self.tabs.addTab(self._build_generate_tab(), "▶ Generate")
        self.tabs.addTab(self._build_review_tab(), "④ Review & Export")

        self.statusBar().showMessage("Ready! Upload a spreadsheet to get started ")

    def _build_setup_tab(self) -> QWidget:
        # Content widget
        content = QWidget()
        layout = QVBoxLayout(content)
        layout.setSpacing(12)
        layout.setContentsMargins(8, 8, 8, 8)

        # ── API Key (only show full input if not loaded from .env) ──
        api_group = QGroupBox("♦ AI Configuration")
        api_layout = QFormLayout(api_group)

        if self._env_api_key:
            # Key is already loaded — show a friendly confirmation
            key_preview = self._env_api_key[:12] + "..." + self._env_api_key[-4:]
            key_label = QLabel(f"✅ API key loaded: {key_preview}")
            key_label.setStyleSheet("color: #059669; font-weight: 600; font-size: 13px;")
            api_layout.addRow(key_label)

            key_note = QLabel("This was saved during installation. You're all set!")
            key_note.setStyleSheet("color: #9ca3af; font-size: 12px;")
            api_layout.addRow(key_note)

            # Hidden input for compatibility — pre-filled
            self.api_key_input = QLineEdit()
            self.api_key_input.setText(self._env_api_key)
            self.api_key_input.setVisible(False)
            api_layout.addRow(self.api_key_input)
        else:
            # No .env key — show the input field
            self.api_key_input = QLineEdit()
            self.api_key_input.setEchoMode(QLineEdit.EchoMode.Password)
            self.api_key_input.setPlaceholderText("Paste your API key here (sk-ant-...)")
            api_layout.addRow("API Key:", self.api_key_input)

            key_note = QLabel("Ask the person who sent you this app for the key!")
            key_note.setStyleSheet("color: #d97706; font-size: 12px;")
            api_layout.addRow(key_note)

        self.tone_combo = QComboBox()
        self.tone_combo.addItems(TONE_OPTIONS)
        api_layout.addRow("Email Tone:", self.tone_combo)

        ai_note = QLabel(
            "ℹ️ Student data (name, email, meeting time) is sent to the AI for "
            "personalization. Nothing is stored."
        )
        ai_note.setWordWrap(True)
        ai_note.setStyleSheet("color: #9ca3af; font-size: 11px; padding: 4px 0;")
        api_layout.addRow(ai_note)

        layout.addWidget(api_group)

        # ── Outlook Status ──
        outlook_group = QGroupBox("✉ Outlook Connection")
        outlook_layout = QVBoxLayout(outlook_group)

        self.outlook_status_label = QLabel("Checking...")
        outlook_layout.addWidget(self.outlook_status_label)

        self.outlook_refresh_btn = QPushButton("Refresh")
        self.outlook_refresh_btn.setProperty("secondary", True)
        self.outlook_refresh_btn.setMaximumWidth(120)
        self.outlook_refresh_btn.clicked.connect(self._check_outlook)
        outlook_layout.addWidget(self.outlook_refresh_btn)

        layout.addWidget(outlook_group)

        # ── Email Template (subject ABOVE body) ──
        template_group = QGroupBox("✎ Email Template")
        template_layout = QVBoxLayout(template_group)

        template_hint = QLabel(
            "Write your boilerplate email below. You can use placeholders like "
            "{name} and {meeting_time}, or just write naturally — the AI handles it!"
        )
        template_hint.setWordWrap(True)
        template_hint.setStyleSheet("color: #9d174d; font-size: 12px; margin-bottom: 6px;")
        template_layout.addWidget(template_hint)

        # Subject line FIRST (like a real email)
        subj_layout = QHBoxLayout()
        subj_label = QLabel("Subject:")
        subj_label.setStyleSheet("color: #9d174d; font-weight: 600; font-size: 13px;")
        subj_label.setFixedWidth(60)
        subj_layout.addWidget(subj_label)
        self.subject_input = QLineEdit()
        self.subject_input.setPlaceholderText("e.g., Your Upcoming One-on-One Meeting")
        self.subject_input.setMinimumHeight(36)
        subj_layout.addWidget(self.subject_input)
        template_layout.addLayout(subj_layout)

        # Body SECOND
        body_label = QLabel("Body:")
        body_label.setStyleSheet("color: #9d174d; font-weight: 600; font-size: 13px; margin-top: 4px;")
        template_layout.addWidget(body_label)

        self.template_edit = QPlainTextEdit()
        self.template_edit.setPlaceholderText("Type or paste your email template here...")
        self.template_edit.setMinimumHeight(180)
        template_layout.addWidget(self.template_edit)

        layout.addWidget(template_group)
        layout.addStretch()

        # Wrap in scroll area
        scroll = QScrollArea()
        scroll.setWidget(content)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        return scroll

    def _build_data_tab(self) -> QWidget:
        content = QWidget()
        layout = QVBoxLayout(content)
        layout.setSpacing(12)
        layout.setContentsMargins(8, 8, 8, 8)

        # Upload
        upload_group = QGroupBox("◈ Spreadsheet Upload")
        upload_layout = QVBoxLayout(upload_group)

        btn_row = QHBoxLayout()
        self.upload_btn = QPushButton("Upload Spreadsheet (.csv / .xlsx)")
        self.upload_btn.clicked.connect(self._upload_spreadsheet)
        btn_row.addWidget(self.upload_btn)

        self.file_label = QLabel("No file loaded yet")
        self.file_label.setStyleSheet("color: #9ca3af;")
        btn_row.addWidget(self.file_label)
        btn_row.addStretch()
        upload_layout.addLayout(btn_row)

        layout.addWidget(upload_group)

        # Column mapping
        mapping_group = QGroupBox("⇔ Column Mapping")
        mapping_layout = QFormLayout(mapping_group)

        self.map_combos = {}
        for role in COLUMN_HINTS:
            combo = QComboBox()
            combo.addItem("— Not Mapped —")
            combo.setMinimumWidth(220)
            self.map_combos[role] = combo
            label = role.replace("_", " ").title()
            mapping_layout.addRow(f"{label}:", combo)

        layout.addWidget(mapping_group)

        # Data preview
        preview_group = QGroupBox("◉ Data Preview")
        preview_layout = QVBoxLayout(preview_group)

        self.data_table = QTableWidget()
        self.data_table.setAlternatingRowColors(True)
        self.data_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.data_table.horizontalHeader().setStretchLastSection(True)
        preview_layout.addWidget(self.data_table)

        self.row_count_label = QLabel("")
        self.row_count_label.setStyleSheet("color: #9d174d; font-size: 12px;")
        preview_layout.addWidget(self.row_count_label)

        layout.addWidget(preview_group)

        scroll = QScrollArea()
        scroll.setWidget(content)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        return scroll

    def _build_generate_tab(self) -> QWidget:
        content = QWidget()
        layout = QVBoxLayout(content)
        layout.setSpacing(12)
        layout.setContentsMargins(8, 8, 8, 8)

        # Summary
        summary_group = QGroupBox("≡ Generation Summary")
        summary_layout = QVBoxLayout(summary_group)

        self.summary_label = QLabel("Upload a spreadsheet and configure your template first!")
        self.summary_label.setWordWrap(True)
        self.summary_label.setStyleSheet("color: #6b7280;")
        summary_layout.addWidget(self.summary_label)

        layout.addWidget(summary_group)

        # Controls
        ctrl_layout = QHBoxLayout()

        self.generate_btn = QPushButton("▶  Generate All Emails!")
        self.generate_btn.clicked.connect(self._start_generation)
        self.generate_btn.setEnabled(False)
        self.generate_btn.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #ec4899, stop:1 #a855f7);
                color: white; border: none; padding: 12px 28px;
                border-radius: 12px; font-weight: 700; font-size: 14px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #db2777, stop:1 #9333ea);
            }
            QPushButton:disabled { background: #d1d5db; color: #9ca3af; }
        """)
        ctrl_layout.addWidget(self.generate_btn)

        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.setProperty("danger", True)
        self.cancel_btn.clicked.connect(self._cancel_generation)
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.setMaximumWidth(120)
        ctrl_layout.addWidget(self.cancel_btn)

        ctrl_layout.addStretch()
        layout.addLayout(ctrl_layout)

        # Progress
        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        layout.addWidget(self.progress_bar)

        self.progress_label = QLabel("")
        self.progress_label.setStyleSheet("color: #9d174d; font-size: 12px;")
        layout.addWidget(self.progress_label)

        # Log
        log_group = QGroupBox("✎ Generation Log")
        log_layout = QVBoxLayout(log_group)

        self.log_display = QPlainTextEdit()
        self.log_display.setReadOnly(True)
        self.log_display.setMaximumHeight(250)
        mono = "Menlo" if IS_MAC else "Consolas" if IS_WINDOWS else "Ubuntu Mono"
        self.log_display.setStyleSheet(
            f"font-family: {mono}, monospace; font-size: 12px; color: #1f2937;"
        )
        log_layout.addWidget(self.log_display)

        layout.addWidget(log_group)
        layout.addStretch()

        scroll = QScrollArea()
        scroll.setWidget(content)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        return scroll

    def _build_review_tab(self) -> QWidget:
        content = QWidget()
        layout = QVBoxLayout(content)
        layout.setSpacing(12)
        layout.setContentsMargins(8, 8, 8, 8)

        # Results table
        results_group = QGroupBox("④ Generated Emails")
        results_layout = QVBoxLayout(results_group)

        self.results_table = QTableWidget()
        self.results_table.setColumnCount(4)
        self.results_table.setHorizontalHeaderLabels(["Status", "Name", "Email", "Preview"])
        self.results_table.setAlternatingRowColors(True)
        self.results_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self.results_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.results_table.cellClicked.connect(self._show_email_detail)
        results_layout.addWidget(self.results_table)

        self.results_count_label = QLabel("")
        self.results_count_label.setStyleSheet("color: #9d174d; font-size: 12px;")
        results_layout.addWidget(self.results_count_label)

        layout.addWidget(results_group)

        # Detail viewer
        detail_group = QGroupBox("◉ Email Detail")
        detail_layout = QVBoxLayout(detail_group)

        self.detail_display = QPlainTextEdit()
        self.detail_display.setReadOnly(True)
        self.detail_display.setMinimumHeight(150)
        detail_layout.addWidget(self.detail_display)

        detail_btns = QHBoxLayout()
        self.copy_btn = QPushButton("≡ Copy to Clipboard")
        self.copy_btn.setProperty("secondary", True)
        self.copy_btn.clicked.connect(self._copy_selected_email)
        self.copy_btn.setEnabled(False)
        detail_btns.addWidget(self.copy_btn)
        detail_btns.addStretch()
        detail_layout.addLayout(detail_btns)

        layout.addWidget(detail_group)

        # Export actions
        export_group = QGroupBox("⬇ Export & Send")
        export_layout = QHBoxLayout(export_group)

        self.export_csv_btn = QPushButton("CSV Export")
        self.export_csv_btn.clicked.connect(self._export_csv)
        self.export_csv_btn.setEnabled(False)
        export_layout.addWidget(self.export_csv_btn)

        self.export_html_btn = QPushButton("HTML Review")
        self.export_html_btn.clicked.connect(self._export_html)
        self.export_html_btn.setEnabled(False)
        export_layout.addWidget(self.export_html_btn)

        self.export_txt_btn = QPushButton("Individual TXT")
        self.export_txt_btn.clicked.connect(self._export_txt)
        self.export_txt_btn.setEnabled(False)
        export_layout.addWidget(self.export_txt_btn)

        self.outlook_draft_btn = QPushButton("✉ Outlook Drafts")
        self.outlook_draft_btn.clicked.connect(self._create_outlook_drafts)
        self.outlook_draft_btn.setEnabled(False)
        export_layout.addWidget(self.outlook_draft_btn)

        layout.addWidget(export_group)

        scroll = QScrollArea()
        scroll.setWidget(content)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        return scroll

    # ── Outlook ───────────────────────────────────────────────────────────

    def _check_outlook(self):
        self.outlook_available = check_outlook_available()
        if self.outlook_available:
            self.outlook_status_label.setText("✅ Outlook is connected and ready for drafts!")
            self.outlook_status_label.setStyleSheet("color: #059669; font-size: 13px; font-weight: 600;")
        elif not IS_WINDOWS:
            self.outlook_status_label.setText(
                "ℹ️ Outlook drafts are Windows-only. Use export options instead!\n"
                "   Astra can use Outlook drafts on her Windows PC."
            )
            self.outlook_status_label.setStyleSheet("color: #6b7280; font-size: 13px;")
        else:
            self.outlook_status_label.setText(
                "⚠️ Outlook not detected. Export options still work great!\n"
                "   (Needs Microsoft Outlook desktop app + pywin32)"
            )
            self.outlook_status_label.setStyleSheet("color: #d97706; font-size: 13px;")

    # ── Spreadsheet ───────────────────────────────────────────────────────

    def _upload_spreadsheet(self):
        filepath, _ = QFileDialog.getOpenFileName(
            self, "Open Spreadsheet", "",
            "Spreadsheets (*.csv *.xlsx *.xls);;All Files (*)"
        )
        if not filepath:
            return
        try:
            headers, rows = parse_spreadsheet(filepath)
        except Exception as e:
            QMessageBox.critical(self, "Error Reading File", str(e))
            return
        if not rows:
            QMessageBox.warning(self, "Empty File", "The file has no data rows.")
            return

        self.spreadsheet_headers = headers
        self.spreadsheet_rows = rows

        self.file_label.setText(f"✅ {Path(filepath).name} — {len(rows)} students")
        self.file_label.setStyleSheet("color: #059669; font-weight: 600;")

        for role, combo in self.map_combos.items():
            combo.clear()
            combo.addItem("— Not Mapped —")
            combo.addItems(headers)

        auto = auto_map_columns(headers)
        for role, header in auto.items():
            if role in self.map_combos:
                idx = self.map_combos[role].findText(header)
                if idx >= 0:
                    self.map_combos[role].setCurrentIndex(idx)

        preview_rows = rows[:50]
        self.data_table.setRowCount(len(preview_rows))
        self.data_table.setColumnCount(len(headers))
        self.data_table.setHorizontalHeaderLabels(headers)
        for r, row in enumerate(preview_rows):
            for c, header in enumerate(headers):
                item = QTableWidgetItem(row.get(header, ""))
                self.data_table.setItem(r, c, item)
        self.data_table.resizeColumnsToContents()
        self.row_count_label.setText(
            f"Showing {len(preview_rows)} of {len(rows)} rows"
            if len(rows) > 50 else f"{len(rows)} rows loaded"
        )

        self._update_generate_readiness()
        self.statusBar().showMessage(f"Loaded {len(rows)} students from {Path(filepath).name} ")

    # ── Generation ────────────────────────────────────────────────────────

    def _get_column_map(self):
        mapping = {}
        for role, combo in self.map_combos.items():
            text = combo.currentText()
            if text and text != "— Not Mapped —":
                mapping[role] = text
        return mapping

    def _update_generate_readiness(self):
        has_data = len(self.spreadsheet_rows) > 0
        has_template = bool(self.template_edit.toPlainText().strip())
        has_key = bool(self._get_api_key())

        ready = has_data and has_template and has_key
        self.generate_btn.setEnabled(ready)

        parts = []
        if has_key:
            parts.append("✅ API key ready")
        else:
            parts.append("❌ API key not set")

        if has_template:
            parts.append("✅ Template ready")
        else:
            parts.append("❌ Email template is empty")

        if has_data:
            mapping = self._get_column_map()
            parts.append(f"✅ {len(self.spreadsheet_rows)} students loaded")
            if mapping:
                parts.append(f"   Mapped: {', '.join(f'{k} → {v}' for k, v in mapping.items())}")
        else:
            parts.append("❌ No spreadsheet loaded")

        self.summary_label.setText("\n".join(parts))

    def _start_generation(self):
        self._update_generate_readiness()
        if not self.generate_btn.isEnabled():
            return

        template = self.template_edit.toPlainText().strip()
        api_key = self._get_api_key()
        tone = self.tone_combo.currentText()
        self.column_map = self._get_column_map()

        if "name" not in self.column_map:
            QMessageBox.warning(self, "Missing Mapping",
                                "Please map at least the 'Name' column before generating.")
            self.tabs.setCurrentIndex(1)
            return

        self.generated_emails = [
            {"name": "", "to": "", "subject": "", "body": "", "status": "pending"}
            for _ in self.spreadsheet_rows
        ]
        self.log_display.clear()
        self.progress_bar.setMaximum(len(self.spreadsheet_rows))
        self.progress_bar.setValue(0)

        subject_template = self.subject_input.text().strip() or "Your Upcoming One-on-One Meeting"
        for i, row in enumerate(self.spreadsheet_rows):
            name = row.get(self.column_map.get("name", ""), f"Student {i+1}")
            email = row.get(self.column_map.get("email", ""), "")
            self.generated_emails[i]["name"] = name
            self.generated_emails[i]["to"] = email
            self.generated_emails[i]["subject"] = subject_template

        self.generate_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)

        self.generator_thread = EmailGeneratorThread(
            template, self.spreadsheet_rows, self.column_map, tone, api_key
        )
        self.generator_thread.progress.connect(self._on_gen_progress)
        self.generator_thread.email_ready.connect(self._on_email_ready)
        self.generator_thread.error.connect(self._on_gen_error)
        self.generator_thread.finished_all.connect(self._on_gen_finished)
        self.generator_thread.start()

        self.statusBar().showMessage("Generating emails... ")

    def _cancel_generation(self):
        if self.generator_thread:
            self.generator_thread.cancel()
            self.log_display.appendPlainText("⚠ Cancellation requested...")

    def _on_gen_progress(self, idx, msg):
        self.progress_bar.setValue(idx)
        self.progress_label.setText(msg)

    def _on_email_ready(self, idx, email_text):
        self.generated_emails[idx]["body"] = email_text
        self.generated_emails[idx]["status"] = "done"
        name = self.generated_emails[idx]["name"]
        self.log_display.appendPlainText(f"✅ [{idx+1}] {name} — done")
        self.progress_bar.setValue(idx + 1)

    def _on_gen_error(self, idx, error_msg):
        self.generated_emails[idx]["status"] = "error"
        self.generated_emails[idx]["body"] = f"[ERROR] {error_msg}"
        name = self.generated_emails[idx]["name"]
        self.log_display.appendPlainText(f"❌ [{idx+1}] {name} — {error_msg}")
        self.progress_bar.setValue(idx + 1)

    def _on_gen_finished(self):
        self.generate_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)

        done = sum(1 for e in self.generated_emails if e["status"] == "done")
        errs = sum(1 for e in self.generated_emails if e["status"] == "error")

        self.progress_label.setText(f"Complete! {done} emails generated, {errs} errors")
        self.statusBar().showMessage(f"Done! {done} emails ready ")
        self.log_display.appendPlainText(f"\n{'='*40}\nDone! {done} emails, {errs} errors.\n")

        self._populate_review_table()

        has_emails = done > 0
        self.export_csv_btn.setEnabled(has_emails)
        self.export_html_btn.setEnabled(has_emails)
        self.export_txt_btn.setEnabled(has_emails)
        self.outlook_draft_btn.setEnabled(has_emails and self.outlook_available)

        self.tabs.setCurrentIndex(3)

    def _populate_review_table(self):
        emails = self.generated_emails
        self.results_table.setRowCount(len(emails))
        for i, em in enumerate(emails):
            status_icon = "✅" if em["status"] == "done" else "❌" if em["status"] == "error" else "⏳"
            self.results_table.setItem(i, 0, QTableWidgetItem(status_icon))
            self.results_table.setItem(i, 1, QTableWidgetItem(em["name"]))
            self.results_table.setItem(i, 2, QTableWidgetItem(em["to"]))
            preview = em["body"][:100].replace("\n", " ") + "..." if len(em["body"]) > 100 else em["body"]
            self.results_table.setItem(i, 3, QTableWidgetItem(preview))
            if em["status"] == "error":
                for c in range(4):
                    item = self.results_table.item(i, c)
                    if item:
                        item.setBackground(QColor("#fecaca"))
        self.results_table.resizeColumnsToContents()
        self.results_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        done = sum(1 for e in emails if e["status"] == "done")
        self.results_count_label.setText(f"{done} of {len(emails)} emails generated successfully ")

    def _show_email_detail(self, row, col):
        if 0 <= row < len(self.generated_emails):
            em = self.generated_emails[row]
            detail = f"To: {em['to']}\nSubject: {em['subject']}\n{'='*50}\n\n{em['body']}"
            self.detail_display.setPlainText(detail)
            self.copy_btn.setEnabled(True)

    def _copy_selected_email(self):
        text = self.detail_display.toPlainText()
        if text:
            QApplication.clipboard().setText(text)
            self.statusBar().showMessage("Copied to clipboard! ≡", 3000)

    # ── Export ────────────────────────────────────────────────────────────

    def _get_exportable_emails(self):
        return [e for e in self.generated_emails if e["status"] == "done"]

    def _export_csv(self):
        filepath, _ = QFileDialog.getSaveFileName(self, "Export CSV", "emails_export.csv",
                                                   "CSV Files (*.csv)")
        if filepath:
            try:
                export_to_csv(self._get_exportable_emails(), filepath)
                QMessageBox.information(self, "Export Complete",
                                        f"Exported {len(self._get_exportable_emails())} emails! \n{filepath}")
            except Exception as e:
                QMessageBox.critical(self, "Export Error", str(e))

    def _export_html(self):
        filepath, _ = QFileDialog.getSaveFileName(self, "Export HTML", "emails_review.html",
                                                   "HTML Files (*.html)")
        if filepath:
            try:
                export_to_html(self._get_exportable_emails(), filepath)
                QMessageBox.information(self, "Export Complete",
                                        f"HTML review exported! \n{filepath}")
            except Exception as e:
                QMessageBox.critical(self, "Export Error", str(e))

    def _export_txt(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Export Folder")
        if folder:
            try:
                export_individual_txt(self._get_exportable_emails(), folder)
                QMessageBox.information(self, "Export Complete",
                                        f"Exported {len(self._get_exportable_emails())} files! \n{folder}")
            except Exception as e:
                QMessageBox.critical(self, "Export Error", str(e))

    def _create_outlook_drafts(self):
        emails = self._get_exportable_emails()
        if not emails:
            return
        confirm = QMessageBox.question(
            self, "Create Outlook Drafts",
            f"This will create {len(emails)} draft emails in Outlook.\n\n"
            "Drafts will NOT be sent automatically — review and send them yourself.\n\nContinue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return
        try:
            successes, errors = create_outlook_drafts(emails)
            msg = f"Created {successes} drafts in Outlook! "
            if errors:
                msg += f"\n\n{len(errors)} errors:\n" + "\n".join(errors[:5])
            QMessageBox.information(self, "Outlook Drafts", msg)
        except Exception as e:
            QMessageBox.critical(self, "Outlook Error", str(e))

    # ── State Persistence ─────────────────────────────────────────────────

    def _restore_state(self):
        template = self.settings.value("template", DEFAULT_TEMPLATE)
        self.template_edit.setPlainText(template)
        subject = self.settings.value("subject", "Your Upcoming One-on-One Meeting")
        self.subject_input.setText(subject)
        tone = self.settings.value("tone", "Professional")
        idx = self.tone_combo.findText(tone)
        if idx >= 0:
            self.tone_combo.setCurrentIndex(idx)

    def closeEvent(self, event):
        self.settings.setValue("template", self.template_edit.toPlainText())
        self.settings.setValue("subject", self.subject_input.text())
        self.settings.setValue("tone", self.tone_combo.currentText())
        event.accept()

    def showEvent(self, event):
        super().showEvent(event)
        self.template_edit.textChanged.connect(self._update_generate_readiness)
        self.api_key_input.textChanged.connect(self._update_generate_readiness)


# ─── Entry Point ─────────────────────────────────────────────────────────────

def main():
    app = QApplication(sys.argv)
    app.setFont(QFont(get_system_font(), 13 if IS_MAC else 10))
    app.setStyleSheet(get_stylesheet())

    # Set app/taskbar icon
    if IS_WINDOWS and ICON_PATH.exists():
        app.setWindowIcon(QIcon(str(ICON_PATH)))
    elif ICON_PNG_PATH.exists():
        app.setWindowIcon(QIcon(str(ICON_PNG_PATH)))
    elif AVATAR_PATH.exists():
        app.setWindowIcon(QIcon(str(AVATAR_PATH)))

    window = AstraMailerWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        # Write crash log so user can report the error
        log_path = Path(__file__).parent / "crash_log.txt"
        import traceback
        with open(log_path, "w") as f:
            f.write(f"Astra Mailer Crash Log\n")
            f.write(f"Time: {datetime.datetime.now()}\n")
            f.write(f"Python: {sys.version}\n")
            f.write(f"Platform: {platform.platform()}\n\n")
            traceback.print_exc(file=f)
        print(f"\n  Astra Mailer crashed. Error log saved to:\n  {log_path}\n")
        try:
            input("  Press Enter to close...")
        except Exception:
            pass
        sys.exit(1)
