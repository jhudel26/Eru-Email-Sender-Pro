import sys
import os
import time
import re
import json
import logging
from datetime import datetime
import pandas as pd
import win32com.client
from openpyxl import Workbook
from openpyxl.styles import Font

from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QTableWidget, QTableWidgetItem,
    QProgressBar, QTextEdit, QMessageBox, QSplitter,
    QLineEdit, QToolBar, QLabel, QFrame, QScrollArea,
    QGroupBox, QSizePolicy, QSpacerItem, QHeaderView, QComboBox,
    QTabWidget
)
from PySide6.QtGui import QFont, QAction, QIcon, QPalette, QColor, QPixmap, QTextCursor, QTextBlockFormat, QKeySequence
from PySide6.QtCore import Qt, QThread, Signal, QSize

# =====================================================
# SETTINGS MANAGER
# =====================================================
class SettingsManager:
    def __init__(self, config_file="settings.json"):
        # Handle both script and executable environments
        if getattr(sys, 'frozen', False):
            # Running as PyInstaller executable
            # Use user's home directory for settings to ensure writability
            app_data_dir = os.path.join(os.path.expanduser("~"), "EruEmailSender")
            os.makedirs(app_data_dir, exist_ok=True)
            self.config_file = os.path.join(app_data_dir, config_file)
        else:
            # Running as script
            script_dir = os.path.dirname(os.path.abspath(__file__))
            self.config_file = os.path.join(script_dir, config_file)
        self.default_settings = {
            "window_geometry": None,
            "paragraph_spacing": 12,
            "email_templates": {},
            "last_excel_path": "",
            "auto_save_interval": 5,
            "retry_failed_emails": True,
            "max_retries": 3,
            "last_selected_template": "default"
        }
        self.settings = self.load_settings()
    
    def load_settings(self):
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    loaded_settings = json.load(f)
                    # Merge with defaults to handle missing keys
                    settings = self.default_settings.copy()
                    settings.update(loaded_settings)
                    return settings
            else:
                return self.default_settings.copy()
        except Exception as e:
            return self.default_settings.copy()
    
    def save_settings(self):
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, indent=2, ensure_ascii=False)
            return True
        except Exception:
            return False
    
    def get(self, key, default=None):
        return self.settings.get(key, default)
    
    def set(self, key, value):
        self.settings[key] = value
        self.save_settings()

# =====================================================
# HELPER FUNCTION TO GET SURNAME
# =====================================================
def get_surname(fullname):
    """
    Returns the surname (part before comma) from a full name.
    If no comma is found, returns the full name.
    """
    fullname = str(fullname).strip()
    if "," in fullname:
        return fullname.split(",")[0].strip()
    return fullname

# =====================================================
# EMAIL VALIDATION
# =====================================================
def validate_email(email):
    """
    Validate email address format using regex
    Returns (is_valid: bool, error_message: str)
    """
    if not email or not str(email).strip():
        return False, "Email address is empty"
    
    email = str(email).strip()
    
    # Basic email regex pattern
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    
    if not re.match(pattern, email):
        return False, f"Invalid email format: {email}"
    
    # Additional checks
    if email.count('@') != 1:
        return False, f"Email must contain exactly one @ symbol: {email}"
    
    local, domain = email.split('@')
    
    if len(local) == 0:
        return False, "Local part of email is empty"
    
    if len(domain) == 0:
        return False, "Domain part of email is empty"
    
    if domain.startswith('.') or domain.endswith('.'):
        return False, f"Domain cannot start or end with dot: {domain}"
    
    return True, ""

def validate_emails_in_dataframe(df, email_column="Email"):
    """
    Validate all emails in a dataframe column
    Returns (valid_df, invalid_emails: list)
    """
    if df is None or email_column not in df.columns:
        return df, []
    
    invalid_emails = []
    valid_indices = []
    
    for idx, row in df.iterrows():
        email = str(row[email_column]).strip()
        is_valid, error = validate_email(email)
        
        if is_valid or email == "":  # Allow empty emails (will be filtered later)
            valid_indices.append(idx)
        else:
            invalid_emails.append({
                'row': idx + 2,  # +2 for Excel row number (header + 1-based)
                'email': email,
                'error': error
            })
    
    valid_df = df.iloc[valid_indices].copy() if valid_indices else df.iloc[0:0].copy()
    return valid_df, invalid_emails

# =====================================================
# OUTLOOK-SAFE HTML BUILDER
# =====================================================
def build_outlook_safe_html(editor_html: str, para_spacing_px: int = 12) -> str:
    """
    Take rich HTML from QTextEdit.toHtml() and wrap it in an Outlook/Word-safe
    HTML shell with CSS resets to avoid extra spacing and reflow.
    """
    html = editor_html or ""
    PARA_SPACE_PX = max(0, int(para_spacing_px))

    # Extract inner <body> when present
    body_match = re.search(r"<body[^>]*>([\s\S]*?)</body>", html, re.IGNORECASE)
    inner = body_match.group(1) if body_match else html
    # Normalize divs to paragraphs
    inner = re.sub(r"<div\b([^>]*)>", r"<p\1>", inner, flags=re.IGNORECASE)
    inner = re.sub(r"</div>", r"</p>", inner, flags=re.IGNORECASE)
    had_paragraphs = bool(re.search(r"<p\b", inner, flags=re.IGNORECASE))
    # Convert multiple <br> into spacer blocks (Outlook-safe)
    inner = re.sub(
        r"(?:<br\s*/?>\s*){2,}",
        rf'''<table role="presentation" border="0" cellspacing="0" cellpadding="0" width="100%"><tr><td style="padding:0 0 {PARA_SPACE_PX}px 0;"><span style="font-size:1px; line-height:1px;">&nbsp;</span></td></tr></table>''',
        inner,
        flags=re.IGNORECASE
    )

    # Outlook spacing via tables
    # 1) Blank paragraphs -> spacer table
    inner = re.sub(
        r"<p\b[^>]*>\s*</p>",
        rf'''<table role="presentation" border="0" cellspacing="0" cellpadding="0" width="100%"><tr><td height="{PARA_SPACE_PX}" style="font-size:0; line-height:0;">&nbsp;</td></tr></table>''',
        inner,
        flags=re.IGNORECASE
    )
    # 2) Normal paragraphs -> table with content row + spacer row
    def _wrap_para(m):
        content = m.group(1)
        return (
            f'<table role="presentation" border="0" cellspacing="0" cellpadding="0" width="100%">'
            f'<tr><td style="line-height:1.35; mso-line-height-rule:exactly; font-family: Segoe UI, Arial, sans-serif;">{content}</td></tr>'
            f'<tr><td height="{PARA_SPACE_PX}" style="font-size:0; line-height:0;">&nbsp;</td></tr></table>'
        )
    inner = re.sub(
        r"<p\b[^>]*>([\s\S]*?)</p>",
        _wrap_para,
        inner,
        flags=re.IGNORECASE
    )
    # Fallback: if no paragraphs were present and no tables inserted, add spacing after single <br>
    if ('role="presentation"' not in inner) and (not had_paragraphs):
        inner = re.sub(
            r"<br\s*/?>",
            rf'''<br/><table role="presentation" border="0" cellspacing="0" cellpadding="0" width="100%"><tr><td height="{PARA_SPACE_PX}" style="font-size:0; line-height:0;">&nbsp;</td></tr></table>''',
            inner,
            flags=re.IGNORECASE
        )

    # Build final skeleton with resets
    wrapped = f"""<!DOCTYPE html>
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="x-ua-compatible" content="IE=edge">
  <meta name="format-detection" content="telephone=no, date=no, address=no, email=no">
  <meta name="x-apple-disable-message-reformatting">
  <!--[if mso]>
  <xml>
   <o:OfficeDocumentSettings>
    <o:AllowPNG/>
    <o:PixelsPerInch>96</o:PixelsPerInch>
   </o:OfficeDocumentSettings>
  </xml>
  <style type="text/css">
    body, table, td, div, p, a {{ font-family: Segoe UI, Arial, sans-serif !important; }}
    p {{ margin:0 !important; }}
  </style>
  <![endif]-->
  <style>
    body, table, td, div, p, a {{ font-family: Segoe UI, Arial, sans-serif; -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%; }}
    p {{ margin:0 !important; }} /* kept for safety but spacing handled by tables */
    .content {{ font-size: 11pt; line-height: 1.35; color:#2b2b2b; }}
    img {{ border:0; outline:0; text-decoration:none; -ms-interpolation-mode:bicubic; }}
  </style>
</head>
<body style="Margin:0; padding:0; background:#ffffff;">
  <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
    <tr>
      <td align="left" style="padding:0;">
        <div class="content" style="font-family: Segoe UI, Arial, sans-serif; font-size:11pt; line-height:1.35; mso-line-height-rule:exactly;">
          {inner}
        </div>
      </td>
    </tr>
  </table>
</body>
</html>"""
    return wrapped

# =====================================================
# EMAIL WORKER THREAD WITH COM RETRY LOGIC
# =====================================================
class EmailWorker(QThread):
    progress_updated = Signal(int)
    log_updated = Signal(str)
    status_updated = Signal(int, str)
    finished_sending = Signal()

    def __init__(self, dataframe, subject, body_template, para_spacing_px=12, max_retries=3):
        super().__init__()
        self.df = dataframe
        self.subject = subject
        self.body_template = body_template
        self.para_spacing_px = int(para_spacing_px) if para_spacing_px is not None else 12
        self.max_retries = max_retries
        self.running = True

    def stop(self):
        self.running = False

    def run(self):
        try:
            import pythoncom
            pythoncom.CoInitialize()

            # ===== TRY CONNECTING TO OUTLOOK =====
            outlook = None
            for attempt in range(5):
                try:
                    outlook = win32com.client.Dispatch("Outlook.Application")
                    self.log_updated.emit("✅ Connected to Outlook COM.")
                    break
                except Exception as e:
                    self.log_updated.emit(f"⚠ Attempt {attempt +1} failed to connect Outlook: {e}")
                    time.sleep(3)

            if not outlook:
                self.log_updated.emit("❌ Could not connect to Outlook. Make sure it is open and fully ready.")
                self.finished_sending.emit()
                return

            namespace = outlook.GetNamespace("MAPI")
            try:
                outbox = namespace.GetDefaultFolder(4)  # olFolderOutbox
                sent = namespace.GetDefaultFolder(5)    # olFolderSentMail
            except Exception as e:
                self.log_updated.emit(f"❌ Outlook is busy or not ready: {e}")
                self.finished_sending.emit()
                return

            total = len(self.df)
            processed_count = 0
            failed_emails = []  # Track failed emails for retry

            # First pass: try to send all emails
            for index, row in self.df.iterrows():
                if not self.running:
                    self.log_updated.emit("⛔ Sending stopped by user.")
                    break

                success = self._send_single_email(outlook, outbox, sent, row, index)
                if not success:
                    failed_emails.append((index, row))
                
                processed_count += 1
                percent = int((processed_count / total) * 100)
                self.progress_updated.emit(percent)

            # Retry failed emails if any
            if failed_emails and self.max_retries > 0:
                self.log_updated.emit(f"🔄 Retrying {len(failed_emails)} failed emails...")
                
                for retry_attempt in range(self.max_retries):
                    if not self.running or not failed_emails:
                        break
                    
                    self.log_updated.emit(f"🔄 Retry attempt {retry_attempt + 1}/{self.max_retries}")
                    still_failed = []
                    
                    for index, row in failed_emails:
                        if not self.running:
                            break
                        
                        success = self._send_single_email(outlook, outbox, sent, row, index)
                        if not success:
                            still_failed.append((index, row))
                    
                    failed_emails = still_failed
                    
                    if failed_emails:
                        time.sleep(2)  # Wait before next retry attempt

            if failed_emails:
                self.log_updated.emit(f"⚠️ {len(failed_emails)} emails still failed after all retries")

            self.finished_sending.emit()

        except Exception as e:
            self.log_updated.emit(f"FATAL ERROR in worker: {str(e)}")
            self.finished_sending.emit()

    def _send_single_email(self, outlook, outbox, sent, row, index):
        """Send a single email with error handling"""
        try:
            email = str(row["Email"]).strip()
            cc_value = str(row["CC"]).strip()
            attachment = str(row["Attachment Path"]).strip()

            if not email:
                self.status_updated.emit(index, "Failed")
                self.log_updated.emit(f"❌ No email for row {index + 2}")
                return False
                
            if not os.path.exists(attachment):
                self.status_updated.emit(index, "Failed")
                self.log_updated.emit(f"❌ Attachment not found: {attachment}")
                return False

            start_sent_count = sent.Items.Count

            # --- ✅ NEW LOGIC HERE ---
            full_name = str(row["Full Name"]).strip()   # e.g. "Dela Cruz, Juan"
            surname = get_surname(full_name)            # e.g. "Dela Cruz"

            # Replace placeholders accordingly
            subject = self.subject.replace("{{fullname}}", full_name)
            body_raw = self.body_template.replace("{{fullname}}", surname)
            body = build_outlook_safe_html(body_raw, self.para_spacing_px)

            mail = outlook.CreateItem(0)  # olMailItem
            mail.To = email
            if cc_value:  # Only set CC if not empty
                mail.CC = cc_value
            mail.Subject = subject        # ✅ Full name in subject
            try:
                mail.BodyFormat = 2  # 2 = olFormatHTML
            except Exception:
                pass
            mail.HTMLBody = body          # ✅ Surname in body
            mail.Importance = 2  # High
            mail.ReadReceiptRequested = True
            mail.Attachments.Add(attachment)
            mail.Send()

            # Wait until sent or max 30 sec
            for _ in range(30):
                if outbox.Items.Count == 0 or sent.Items.Count > start_sent_count:
                    break
                time.sleep(1)

            self.status_updated.emit(index, "Sent")
            self.log_updated.emit(f"✅ Sent to {email}")
            return True

        except Exception as e:
            self.status_updated.emit(index, "Failed")
            self.log_updated.emit(f"❌ Error sending to {email}: {str(e)}")
            return False


# =====================================================
# MAIN APPLICATION
# =====================================================
class EmailApp(QWidget):
    def __init__(self):
        super().__init__()

        # Initialize settings manager
        self.settings = SettingsManager()
        
        # Setup file-based logging
        self.setup_logging()
        
        self.setWindowTitle("📧 Eru Email Sender Pro")
        self.setMinimumSize(1400, 1000)
        self.setStyleSheet(self.modern_styles())
        
        # Set application icon and style
        window_icon = self.create_app_icon()
        self.setWindowIcon(window_icon)
        
        # Set application-wide icon for all windows
        app_icon = QIcon()
        icon_sizes = [16, 32, 48, 64, 128, 256]
        
        # Determine icon path based on environment
        icon_base_path = "EMAIL.ico"
        if getattr(sys, 'frozen', False):
            # First try PyInstaller's temporary directory where bundled files are extracted
            if hasattr(sys, '_MEIPASS'):
                icon_base_path = os.path.join(sys._MEIPASS, "EMAIL.ico")
            else:
                # Fallback to executable directory
                app_dir = os.path.dirname(sys.executable)
                icon_base_path = os.path.join(app_dir, "EMAIL.ico")
        
        for size in icon_sizes:
            app_icon.addFile(icon_base_path, QSize(size, size))
        QApplication.setWindowIcon(app_icon)
        
        # Start maximized
        self.showMaximized()

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)
        
        # HEADER SECTION
        header_widget = self.create_header()
        main_layout.addWidget(header_widget)
        
        # CONTROL BUTTONS SECTION
        controls_widget = self.create_controls_section()
        main_layout.addWidget(controls_widget)

        # MAIN CONTENT AREA WITH TABS
        self.tab_widget = QTabWidget()
        self.tab_widget.setObjectName("mainTabWidget")
        
        # TAB 1: MAIN DASHBOARD
        dashboard_tab = QWidget()
        dashboard_layout = QVBoxLayout(dashboard_tab)
        dashboard_layout.setContentsMargins(0, 0, 0, 0)
        dashboard_layout.setSpacing(15)
        
        # Dashboard content splitter
        content_splitter = QSplitter(Qt.Horizontal)
        content_splitter.setHandleWidth(2)
        content_splitter.setStyleSheet("""
            QSplitter::handle {
                background-color: #3a3f5a;
                border-radius: 1px;
            }
        """)
        
        # LEFT PANEL - DATA TABLE (wider)
        left_panel = self.create_table_panel()
        content_splitter.addWidget(left_panel)
        
        # RIGHT PANEL - SENDING PROGRESS & ACTIVITY LOGS
        right_panel = self.create_status_section()
        content_splitter.addWidget(right_panel)
        
        # Set splitter proportions (70% table, 30% status)
        content_splitter.setSizes([980, 420])
        content_splitter.setStretchFactor(0, 7)
        content_splitter.setStretchFactor(1, 3)
        
        dashboard_layout.addWidget(content_splitter)
        self.tab_widget.addTab(dashboard_tab, "📊 Main Dashboard")
        
        # TAB 2: EMAIL COMPOSER (WIDE VIEW)
        composer_tab = QWidget()
        composer_layout = QVBoxLayout(composer_tab)
        composer_layout.setContentsMargins(0, 0, 0, 0)
        composer_layout.setSpacing(15)
        
        # Email composer with wide view
        email_composer_widget = self.create_email_panel()
        composer_layout.addWidget(email_composer_widget)
        self.tab_widget.addTab(composer_tab, "✉️ Email Composer")
        
        main_layout.addWidget(self.tab_widget)

        # CONNECTIONS
        self.export_button.clicked.connect(self.export_template)
        self.load_button.clicked.connect(self.load_excel)
        self.start_button.clicked.connect(self.start_sending)
        self.stop_button.clicked.connect(self.stop_sending)
        
        self.df = None
        self.worker = None
        
        # Initialize UI state
        self.load_templates()  # Load saved templates
        # Unblock signals after initialization is complete
        self.template_combo.blockSignals(False)
        self.setup_keyboard_shortcuts()  # Setup keyboard shortcuts
        self.update_ui_state()

    # =================================================
    # UI COMPONENT CREATION METHODS
    # =================================================
    def create_app_icon(self):
        """Create app icon using the EMAIL.ico file"""
        # Handle both script and executable environments
        icon_path = "EMAIL.ico"
        
        # If running as executable, check PyInstaller's temporary directory first
        if getattr(sys, 'frozen', False):
            # First try PyInstaller's temporary directory where bundled files are extracted
            if hasattr(sys, '_MEIPASS'):
                icon_path = os.path.join(sys._MEIPASS, "EMAIL.ico")
            else:
                # Fallback to executable directory
                app_dir = os.path.dirname(sys.executable)
                icon_path = os.path.join(app_dir, "EMAIL.ico")
        
        if os.path.exists(icon_path):
            return QIcon(icon_path)
        else:
            # Fallback to a simple colored icon if EMAIL.ico is not found
            pixmap = QPixmap(32, 32)
            pixmap.fill(QColor("#4a90e2"))
            return QIcon(pixmap)
    
    def create_header(self):
        """Create the header section with title and description"""
        header_frame = QFrame()
        header_frame.setFrameStyle(QFrame.NoFrame)
        header_layout = QVBoxLayout(header_frame)
        header_layout.setContentsMargins(0, 0, 0, 0)
        
        title_label = QLabel("📧 Eru Email Sender Pro")
        title_label.setObjectName("headerTitle")
        title_label.setAlignment(Qt.AlignCenter)
        
        subtitle_label = QLabel("Professional Email Automation System")
        subtitle_label.setObjectName("headerSubtitle")
        subtitle_label.setAlignment(Qt.AlignCenter)
        
        header_layout.addWidget(title_label)
        header_layout.addWidget(subtitle_label)
        
        return header_frame
    
    def create_controls_section(self):
        """Create the control buttons section"""
        controls_frame = QFrame()
        controls_frame.setObjectName("controlsFrame")
        controls_layout = QHBoxLayout(controls_frame)
        controls_layout.setContentsMargins(20, 15, 20, 15)
        controls_layout.setSpacing(15)
        
        # Create buttons with icons
        self.export_button = QPushButton("📄 Export Template")
        self.export_button.setObjectName("primaryButton")
        
        self.load_button = QPushButton("📁 Load Excel")
        self.load_button.setObjectName("primaryButton")
        
        self.start_button = QPushButton("▶️ Start Sending")
        self.start_button.setObjectName("successButton")
        
        self.stop_button = QPushButton("⏹️ Stop")
        self.stop_button.setObjectName("dangerButton")
        
        # Add buttons to layout
        controls_layout.addWidget(self.export_button)
        controls_layout.addWidget(self.load_button)
        controls_layout.addStretch()
        controls_layout.addWidget(self.start_button)
        controls_layout.addWidget(self.stop_button)
        
        return controls_frame
    
    def create_table_panel(self):
        """Create the left panel with data table"""
        table_frame = QFrame()
        table_frame.setObjectName("tableFrame")
        table_layout = QVBoxLayout(table_frame)
        table_layout.setContentsMargins(0, 0, 0, 0)
        
        # Table header with counter
        header_row = QHBoxLayout()
        table_header = QLabel("📋 Recipient Data")
        table_header.setObjectName("sectionTitle")
        
        self.recipient_counter = QLabel("📊 0 recipients loaded")
        self.recipient_counter.setObjectName("recipientCounter")
        self.recipient_counter.setStyleSheet("""
            QLabel#recipientCounter {
                color: #64748b;
                font-size: 10pt;
                font-weight: 500;
                padding: 4px 8px;
                background: #f1f5f9;
                border-radius: 12px;
                border: 1px solid #e2e8f0;
            }
        """)
        
        header_row.addWidget(table_header)
        header_row.addStretch()
        header_row.addWidget(self.recipient_counter)
        table_layout.addLayout(header_row)
        
        # Create table
        self.table = QTableWidget()
        self.table.setObjectName("dataTable")
        self.table.setAlternatingRowColors(True)
        self.table.setShowGrid(True)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setStretchLastSection(True)
        
        table_layout.addWidget(self.table)
        return table_frame
    
    def create_email_panel(self):
        """Create the right panel with email composer"""
        email_frame = QFrame()
        email_frame.setObjectName("emailFrame")
        email_layout = QVBoxLayout(email_frame)
        email_layout.setContentsMargins(0, 0, 0, 0)
        
        # Email composer header
        email_header = QLabel("✉️ Email Composer")
        email_header.setObjectName("sectionTitle")
        email_layout.addWidget(email_header)
        
        # Subject input
        subject_group = QGroupBox("Subject")
        subject_group.setObjectName("inputGroup")
        subject_layout = QVBoxLayout(subject_group)
        
        # Template management row
        template_row = QHBoxLayout()
        template_label = QLabel("Template:")
        self.template_combo = QComboBox()
        self.template_combo.setObjectName("templateCombo")
        self.template_combo.addItem("Default HR Notice", "default")
        # Block signals during initialization to prevent overwriting saved settings
        self.template_combo.blockSignals(True)
        self.template_combo.currentIndexChanged.connect(self.load_template)
        
        self.save_template_btn = QPushButton("💾 Save Template")
        self.save_template_btn.setObjectName("secondaryButton")
        self.save_template_btn.clicked.connect(self.save_template)
        
        self.delete_template_btn = QPushButton("🗑️ Delete")
        self.delete_template_btn.setObjectName("dangerButton")
        self.delete_template_btn.clicked.connect(self.delete_template)
        
        template_row.addWidget(template_label)
        template_row.addWidget(self.template_combo)
        template_row.addWidget(self.save_template_btn)
        template_row.addWidget(self.delete_template_btn)
        template_row.addStretch()
        
        subject_layout.addLayout(template_row)
        
        self.subject_input = QLineEdit()
        self.subject_input.setObjectName("subjectInput")
        self.subject_input.setPlaceholderText("Enter email subject here...")
        self.subject_input.setText("NOTICE TO SUBMIT LACKING EMPLOYMENT REQUIREMENTS - {{fullname}}")
        subject_layout.addWidget(self.subject_input)
        
        email_layout.addWidget(subject_group)
        
        # Formatting toolbar
        toolbar = QToolBar()
        toolbar.setObjectName("formatToolbar")
        toolbar.setMovable(False)
        
        bold_action = QAction("🔤 Bold", self)
        bold_action.triggered.connect(self.make_bold)
        toolbar.addAction(bold_action)
        
        italic_action = QAction("𝐈 Italic", self)
        italic_action.triggered.connect(self.make_italic)
        toolbar.addAction(italic_action)
        
        underline_action = QAction("U̲ Underline", self)
        underline_action.triggered.connect(self.make_underline)
        toolbar.addAction(underline_action)
        
        toolbar.addSeparator()
        
        preview_action = QAction("👁️ Preview", self)
        preview_action.triggered.connect(self.preview_email)
        toolbar.addAction(preview_action)
        
        email_layout.addWidget(toolbar)
        
        # Spacing control
        spacing_row = QHBoxLayout()
        spacing_label = QLabel("Paragraph spacing:")
        spacing_label.setObjectName("spacingLabel")
        self.spacing_select = QComboBox()
        self.spacing_select.setObjectName("spacingSelect")
        self.spacing_select.addItem("Tight", 8)
        self.spacing_select.addItem("Normal", 12)
        self.spacing_select.addItem("Relaxed", 16)
        
        # Load saved spacing setting
        saved_spacing = self.settings.get("paragraph_spacing", 12)
        for i in range(self.spacing_select.count()):
            if self.spacing_select.itemData(i) == saved_spacing:
                self.spacing_select.setCurrentIndex(i)
                break
        
        self.spacing_select.currentIndexChanged.connect(self.on_spacing_changed)
        spacing_row.addWidget(spacing_label)
        spacing_row.addWidget(self.spacing_select)
        spacing_row.addStretch()
        email_layout.addLayout(spacing_row)
        
        # Email body
        body_group = QGroupBox("Message Body")
        body_group.setObjectName("inputGroup")
        body_layout = QVBoxLayout(body_group)
        
        self.email_editor = QTextEdit()
        self.email_editor.setObjectName("emailEditor")
        self.email_editor.setFont(QFont("Segoe UI", 11))
        self.email_editor.setMinimumHeight(500)  # Increased height for wide view
        self.email_editor.setMinimumWidth(800)   # Set minimum width for wide view
        self.email_editor.setAcceptRichText(True)
        # Ensure scrollbars are always visible when needed
        self.email_editor.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.email_editor.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        
        default_body = """
<p>Dear {{fullname}},</p>

<p>This is to formally inform you that you still have outstanding mandatory employment requirements as of this date, despite prior reminders and your signed Affidavit of Undertaking upon commencement of employment.</p>

<p>As stated in your Affidavit of Undertaking, you committed to submit all required documents within prescribed period. <b>You are hereby given five (5) days from receipt of this email notice </b> to complete and submit pending requirements. Please see the attached <b>Notice of Incomplete Employment Requirements</b> for full details. Failure to comply within the given timeframe, may result in appropriate administrative action in accordance with Company policy.</p>

<p>Please submit the required documents through this same email thread. For any clarification, please coordinate with <b>HR-DMRC or your assigned account supervisor.</b></p>

<p>Thanks,<br>Jhudel S. Orola<br>HR Staff - Data Management & Records Control<br>Acabar Marketing International Inc.<br>(02) 8887-8170 Local 153</p>

<p><img class="x_CToWUd" height="77" width="250" src="https://ci3.googleusercontent.com/mail-sig/AIorK4x0oCXqeBBsjR9hQB3HLxhAJPc1msod_2dqrIiATYz-sDfATgJdOa_R6eWlr16--ykbMmeApG_G3we-" data-imagetype="External"></p>
"""
        self.email_editor.setHtml(default_body)
        # Apply spacing to the entire document so the composer preview matches
        try:
            self.apply_editor_paragraph_spacing(int(self.spacing_select.currentData()))
        except Exception:
            self.apply_editor_paragraph_spacing(12)
        
        body_layout.addWidget(self.email_editor)
        email_layout.addWidget(body_group)
        
        return email_frame
    
    def create_status_section(self):
        """Create the status section with progress and logs"""
        status_frame = QFrame()
        status_frame.setObjectName("statusFrame")
        status_layout = QVBoxLayout(status_frame)
        status_layout.setContentsMargins(0, 0, 0, 0)
        
        # Progress section
        progress_group = QGroupBox("📊 Sending Progress")
        progress_group.setObjectName("progressGroup")
        progress_layout = QVBoxLayout(progress_group)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setObjectName("progressBar")
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("%p%")
        
        progress_layout.addWidget(self.progress_bar)
        status_layout.addWidget(progress_group)
        
        # Logs section
        logs_group = QGroupBox("📝 Activity Logs")
        logs_group.setObjectName("logsGroup")
        logs_layout = QVBoxLayout(logs_group)
        
        self.log_box = QTextEdit()
        self.log_box.setObjectName("logBox")
        self.log_box.setReadOnly(True)
        # Removed maximum height to allow expansion in right panel
        
        logs_layout.addWidget(self.log_box)
        status_layout.addWidget(logs_group)
        
        return status_frame
    
    def update_ui_state(self):
        """Update UI state based on data availability"""
        has_data = self.df is not None and len(self.df) > 0
        self.start_button.setEnabled(has_data)
        self.stop_button.setEnabled(False)
        
        # Update recipient counter
        if has_data:
            recipient_count = len(self.df)
            self.recipient_counter.setText(f"📊 {recipient_count} recipient{'s' if recipient_count != 1 else ''} loaded")
        else:
            self.recipient_counter.setText("📊 0 recipients loaded")

    # =================================================
    # EXPORT EXCEL TEMPLATE
    # =================================================
    def export_template(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Excel Template", "Email_Template.xlsx", "Excel Files (*.xlsx)")
        if not file_path:
            return
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Email Sending Setup"
            headers = ["Full Name", "Email", "CC", "Attachment Path"]
            for col, header in enumerate(headers, start=1):
                ws.cell(row=1, column=col, value=header).font = Font(bold=True)
            ws.cell(row=2, column=1).value = "Dela Cruz, Juan"
            ws.cell(row=2, column=2).value = "juan@email.com"
            ws.cell(row=2, column=3).value = ""
            ws.cell(row=2, column=4).value = "C:\\Path\\To\\Attachment.pdf"
            for col_letter, width in zip(["A", "B", "C", "D"], [25, 30, 30, 40]):
                ws.column_dimensions[col_letter].width = width
            wb.save(file_path)
            QMessageBox.information(self, "Success", "Excel template exported successfully.")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    # =================================================
    # LOAD EXCEL
    # =================================================
    def load_excel(self):
        # Use last path from settings if available
        last_path = self.settings.get("last_excel_path", "")
        file_path, _ = QFileDialog.getOpenFileName(self, "Open Excel", last_path, "Excel Files (*.xlsx)")
        if not file_path:
            return
        try:
            df = pd.read_excel(file_path, sheet_name="Email Sending Setup")
            df = df.rename(columns={
                df.columns[0]: "Full Name",
                df.columns[1]: "Email",
                df.columns[2]: "CC",
                df.columns[3]: "Attachment Path"
            })
            df = df.fillna("")
            
            # Validate emails
            valid_df, invalid_emails = validate_emails_in_dataframe(df, "Email")
            
            # Show validation results
            if invalid_emails:
                error_msg = "The following emails have invalid format:\n\n"
                for item in invalid_emails[:10]:  # Show max 10 errors
                    error_msg += f"Row {item['row']}: {item['email']} - {item['error']}\n"
                
                if len(invalid_emails) > 10:
                    error_msg += f"... and {len(invalid_emails) - 10} more errors\n"
                
                error_msg += "\nThese rows will be excluded from sending."
                QMessageBox.warning(self, "Email Validation Warning", error_msg)
            
            # Use validated dataframe
            df = valid_df
            df["Status"] = "Pending"
            self.df = df[["Full Name", "Email", "CC", "Attachment Path", "Status"]]
            self.populate_table()
            
            # Update recipient counter
            recipient_count = len(self.df)
            self.recipient_counter.setText(f"📊 {recipient_count} recipient{'s' if recipient_count != 1 else ''} loaded")
            
            # Save the path for next time
            self.settings.set("last_excel_path", file_path)
            
            if len(invalid_emails) > 0:
                self.log(f"⚠️ Excel loaded with {len(invalid_emails)} invalid emails excluded.")
            else:
                self.log("✅ Excel loaded and validated successfully.")
            self.update_ui_state()
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    # =================================================
    # POPULATE TABLE
    # =================================================
    def populate_table(self):
        self.table.setRowCount(len(self.df))
        self.table.setColumnCount(len(self.df.columns))
        self.table.setHorizontalHeaderLabels(self.df.columns)
        
        # Set column widths
        header = self.table.horizontalHeader()
        header.setStretchLastSection(False)  # Don't stretch last section (Status)
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)  # Full Name
        header.setSectionResizeMode(1, QHeaderView.Stretch)  # Email
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)  # CC
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)  # Attachment Path
        header.setSectionResizeMode(4, QHeaderView.Fixed)  # Status - fixed width
        header.setDefaultSectionSize(100)  # Default width for stretch columns
        
        # Set specific width for Status column
        header.resizeSection(4, 80)  # Status column width
        
        for i in range(len(self.df)):
            for j in range(len(self.df.columns)):
                item = QTableWidgetItem(str(self.df.iloc[i, j]))
                
                # Color code status
                if j == 4:  # Status column
                    status = str(self.df.iloc[i, j]).lower()
                    if status == "sent":
                        item.setBackground(QColor("#d4edda"))
                    elif status == "failed":
                        item.setBackground(QColor("#f8d7da"))
                    elif status == "pending":
                        item.setBackground(QColor("#fff3cd"))
                
                self.table.setItem(i, j, item)

    # =================================================
    # START / STOP SENDING
    # =================================================
    def start_sending(self):
        if self.df is None:
            QMessageBox.warning(self, "⚠️ Warning", "Please load an Excel file first.")
            return
        if len(self.df) == 0:
            QMessageBox.warning(self, "⚠️ Warning", "No rows to send.")
            return

        subject = self.subject_input.text()
        body = self.email_editor.toHtml()

        # Get settings
        try:
            spacing_px = int(self.spacing_select.currentData())
            max_retries = self.settings.get("max_retries", 3)
        except Exception:
            spacing_px = 12
            max_retries = 3

        # Save current settings
        self.settings.set("paragraph_spacing", spacing_px)

        self.worker = EmailWorker(self.df, subject, body, spacing_px, max_retries)
        self.worker.progress_updated.connect(self.progress_bar.setValue)
        self.worker.log_updated.connect(self.log)
        self.worker.status_updated.connect(self.update_status)
        self.worker.finished_sending.connect(self.finish_message)
        self.worker.start()
        
        # Update UI state
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.log("ℹ️ Email sending started. Opening Outlook")

    def stop_sending(self):
        if self.worker:
            self.worker.stop()
            self.log("⛔ Sending stopped by user.")
            self.update_ui_state()

    # =================================================
    # UPDATE STATUS & LOGS
    # =================================================
    def update_status(self, row, status):
        self.table.setItem(row, 4, QTableWidgetItem(status))
        # Update the dataframe as well
        if self.df is not None and row < len(self.df):
            self.df.iloc[row, 4] = status

    def log(self, message):
        """Log message to both UI and file"""
        self.log_box.append(message)
        
        # Also log to file
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"[{timestamp}] {message}"
        
        try:
            with open("email_sender.log", "a", encoding="utf-8") as f:
                f.write(log_message + "\n")
        except Exception:
            pass  # Silently fail if logging doesn't work

    # =================================================
    # LOGGING SYSTEM
    # =================================================
    def setup_logging(self):
        """Setup file-based logging system"""
        try:
            # Create logs directory if it doesn't exist
            if not os.path.exists("logs"):
                os.makedirs("logs")
            
            # Setup logging configuration
            log_file = os.path.join("logs", f"email_sender_{datetime.now().strftime('%Y%m%d')}.log")
            
            logging.basicConfig(
                level=logging.INFO,
                format='%(asctime)s - %(levelname)s - %(message)s',
                handlers=[
                    logging.FileHandler(log_file, encoding='utf-8'),
                    logging.StreamHandler()  # Also log to console
                ]
            )
            
            # Log application start
            logging.info("Eru Email Sender Pro started")
            
        except Exception as e:
            print(f"Failed to setup logging: {e}")

    def finish_message(self):
        QMessageBox.information(self, "✅ Complete", "All emails have been processed.")
        self.update_ui_state()

    # =================================================
    # TEXT FORMATTING
    # =================================================
    def make_bold(self):
        fmt = self.email_editor.currentCharFormat()
        fmt.setFontWeight(QFont.Bold if fmt.fontWeight() != QFont.Bold else QFont.Normal)
        self.email_editor.setCurrentCharFormat(fmt)

    def make_italic(self):
        fmt = self.email_editor.currentCharFormat()
        fmt.setFontItalic(not fmt.fontItalic())
        self.email_editor.setCurrentCharFormat(fmt)

    def make_underline(self):
        fmt = self.email_editor.currentCharFormat()
        fmt.setFontUnderline(not fmt.fontUnderline())
        self.email_editor.setCurrentCharFormat(fmt)

    # =================================================
    # EMAIL PREVIEW
    # =================================================
    def preview_email(self):
        """Preview email with sample data"""
        from PySide6.QtWidgets import QDialog, QVBoxLayout, QTextBrowser, QHBoxLayout, QLabel, QComboBox
        
        # Create preview dialog
        dialog = QDialog(self)
        dialog.setWindowTitle("📧 Email Preview")
        dialog.setMinimumSize(800, 600)
        dialog.setStyleSheet(self.modern_styles())
        
        layout = QVBoxLayout(dialog)
        
        # Sample recipient selection
        sample_row = QHBoxLayout()
        sample_label = QLabel("Sample Recipient:")
        sample_combo = QComboBox()
        sample_combo.addItem("Sample: Dela Cruz, Juan", "Dela Cruz, Juan")
        sample_combo.addItem("Sample: Smith, John", "Smith, John")
        sample_combo.addItem("Sample: Garcia, Maria", "Garcia, Maria")
        sample_row.addWidget(sample_label)
        sample_row.addWidget(sample_combo)
        sample_row.addStretch()
        layout.addLayout(sample_row)
        
        # Preview browser
        preview_browser = QTextBrowser()
        preview_browser.setObjectName("previewBrowser")
        layout.addWidget(preview_browser)
        
        # Update preview when sample changes
        def update_preview():
            sample_name = sample_combo.currentData()
            subject = self.subject_input.text().replace("{{fullname}}", sample_name)
            body_html = self.email_editor.toHtml().replace("{{fullname}}", get_surname(sample_name))
            
            # Apply Outlook-safe formatting
            try:
                spacing_px = int(self.spacing_select.currentData())
                final_html = build_outlook_safe_html(body_html, spacing_px)
            except Exception:
                final_html = body_html
            
            preview_html = f"""
            <div style="font-family: Segoe UI, Arial, sans-serif; padding: 20px;">
                <h3 style="color: #1e40af; margin-bottom: 15px;">Subject: {subject}</h3>
                <div style="border-top: 1px solid #e2e8f0; padding-top: 15px;">
                    {final_html}
                </div>
            </div>
            """
            preview_browser.setHtml(preview_html)
        
        sample_combo.currentIndexChanged.connect(lambda: update_preview())
        update_preview()  # Initial preview
        
        # Dialog buttons
        from PySide6.QtWidgets import QDialogButtonBox
        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(dialog.close)
        layout.addWidget(buttons)
        
        dialog.exec()

    # =================================================
    # COMPOSER SPACING PREVIEW
    # =================================================
    def apply_editor_paragraph_spacing(self, px: int):
        doc = self.email_editor.document()
        cursor = QTextCursor(doc)
        cursor.beginEditBlock()
        block = doc.begin()
        while block.isValid():
            bfmt = block.blockFormat()
            bfmt.setTopMargin(0)
            bfmt.setBottomMargin(max(0, int(px)))
            # Use proportional line height ~135% for readability
            try:
                bfmt.setLineHeight(135, QTextBlockFormat.ProportionalHeight)
            except Exception:
                pass
            c = QTextCursor(block)
            c.setBlockFormat(bfmt)
            block = block.next()
        cursor.endEditBlock()

    def on_spacing_changed(self, _index: int):
        try:
            px = int(self.spacing_select.currentData())
        except Exception:
            px = 12
        self.apply_editor_paragraph_spacing(px)
        self.settings.set("paragraph_spacing", px)

    # =================================================
    # TEMPLATE MANAGEMENT
    # =================================================
    def load_templates(self):
        """Load templates from settings into combo box"""
        templates = self.settings.get("email_templates", {})
        
        # Clear existing items except default
        self.template_combo.clear()
        self.template_combo.addItem("Default HR Notice", "default")
        
        # Add saved templates
        for name in templates.keys():
            self.template_combo.addItem(name, name)
        
        # Restore last selected template
        last_template = self.settings.get("last_selected_template", "default")
        
        template_found = False
        for i in range(self.template_combo.count()):
            item_data = self.template_combo.itemData(i)
            if item_data == last_template:
                # Temporarily block signals to avoid triggering load_template twice
                self.template_combo.blockSignals(True)
                self.template_combo.setCurrentIndex(i)
                self.template_combo.blockSignals(False)
                # Load the template content directly without signal
                self._load_template_content(i)
                template_found = True
                break
        
        if not template_found:
            # Default template is already selected at index 0
            self._load_template_content(0)
    
    def _load_template_content(self, index):
        """Load template content without saving the last selected template"""
        if index == 0:  # Default template
            self.subject_input.setText("NOTICE TO SUBMIT LACKING EMPLOYMENT REQUIREMENTS - {{fullname}}")
            default_body = """
<p>Dear {{fullname}},</p>

<p>This is to formally inform you that you still have outstanding mandatory employment requirements as of this date, despite prior reminders and your signed Affidavit of Undertaking upon commencement of employment.</p>

<p>As stated in your Affidavit of Undertaking, you committed to submit all required documents within prescribed period. <b>You are hereby given five (5) days from receipt of this email notice </b> to complete and submit pending requirements. Please see the attached <b>Notice of Incomplete Employment Requirements</b> for full details. Failure to comply within the given timeframe, may result in appropriate administrative action in accordance with Company policy.</p>

<p>Please submit the required documents through this same email thread. For any clarification, please coordinate with <b>HR-DMRC or your assigned account supervisor.</b></p>

<p>Thanks,<br>Jhudel S. Orola<br>HR Staff - Data Management & Records Control<br>Acabar Marketing International Inc.<br>(02) 8887-8170 Local 153</p>

<p><img class="x_CToWUd" height="77" width="250" src="https://ci3.googleusercontent.com/mail-sig/AIorK4x0oCXqeBBsjR9hQB3HLxhAJPc1msod_2dqrIiATYz-sDfATgJdOa_R6eWlr16--ykbMmeApG_G3we-" data-imagetype="External"></p>
"""
            self.email_editor.setHtml(default_body)
        else:
            template_name = self.template_combo.itemData(index)  # Use index instead of currentData
            templates = self.settings.get("email_templates", {})
            if template_name in templates:
                template = templates[template_name]
                self.subject_input.setText(template.get("subject", ""))
                self.email_editor.setHtml(template.get("body", ""))
        
        # Apply current spacing
        try:
            px = int(self.spacing_select.currentData())
            self.apply_editor_paragraph_spacing(px)
        except Exception:
            pass

    def load_template(self, index):
        """Load selected template into editor"""
        # Load the content
        self._load_template_content(index)
        
        # Save the last selected template
        template_name = self.template_combo.currentData()
        self.settings.set("last_selected_template", template_name)
    
    def save_template(self):
        """Save current email as template"""
        from PySide6.QtWidgets import QInputDialog
        
        name, ok = QInputDialog.getText(self, "Save Template", "Enter template name:")
        if not ok or not name.strip():
            return
        
        name = name.strip()
        templates = self.settings.get("email_templates", {})
        templates[name] = {
            "subject": self.subject_input.text(),
            "body": self.email_editor.toHtml()
        }
        
        self.settings.set("email_templates", templates)
        self.load_templates()  # Refresh combo box
        
        # Select the newly saved template
        for i in range(self.template_combo.count()):
            if self.template_combo.itemData(i) == name:
                self.template_combo.setCurrentIndex(i)
                break
        
        QMessageBox.information(self, "Success", f"Template '{name}' saved successfully.")
    
    def delete_template(self):
        """Delete selected template"""
        if self.template_combo.currentIndex() == 0:
            QMessageBox.warning(self, "Warning", "Cannot delete the default template.")
            return
        
        template_name = self.template_combo.currentData()
        reply = QMessageBox.question(self, "Confirm Delete", 
                                   f"Are you sure you want to delete template '{template_name}'?",
                                   QMessageBox.Yes | QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            templates = self.settings.get("email_templates", {})
            if template_name in templates:
                del templates[template_name]
                self.settings.set("email_templates", templates)
                self.load_templates()  # Refresh combo box
                self.template_combo.setCurrentIndex(0)  # Select default
                QMessageBox.information(self, "Success", f"Template '{template_name}' deleted.")

    # =================================================
    # KEYBOARD SHORTCUTS
    # =================================================
    def setup_keyboard_shortcuts(self):
        """Setup keyboard shortcuts for common actions"""
        
        # Export Template: Ctrl+E
        export_shortcut = QAction(self)
        export_shortcut.setShortcut(QKeySequence("Ctrl+E"))
        export_shortcut.triggered.connect(self.export_template)
        self.addAction(export_shortcut)
        
        # Load Excel: Ctrl+O
        load_shortcut = QAction(self)
        load_shortcut.setShortcut(QKeySequence("Ctrl+O"))
        load_shortcut.triggered.connect(self.load_excel)
        self.addAction(load_shortcut)
        
        # Start Sending: Ctrl+S
        start_shortcut = QAction(self)
        start_shortcut.setShortcut(QKeySequence("Ctrl+S"))
        start_shortcut.triggered.connect(self.start_sending)
        self.addAction(start_shortcut)
        
        # Stop Sending: Ctrl+Shift+S
        stop_shortcut = QAction(self)
        stop_shortcut.setShortcut(QKeySequence("Ctrl+Shift+S"))
        stop_shortcut.triggered.connect(self.stop_sending)
        self.addAction(stop_shortcut)
        
        # Preview Email: Ctrl+P
        preview_shortcut = QAction(self)
        preview_shortcut.setShortcut(QKeySequence("Ctrl+P"))
        preview_shortcut.triggered.connect(self.preview_email)
        self.addAction(preview_shortcut)
        
        # Save Template: Ctrl+T
        save_template_shortcut = QAction(self)
        save_template_shortcut.setShortcut(QKeySequence("Ctrl+T"))
        save_template_shortcut.triggered.connect(self.save_template)
        self.addAction(save_template_shortcut)
        
        # Bold: Ctrl+B
        bold_shortcut = QAction(self)
        bold_shortcut.setShortcut(QKeySequence("Ctrl+B"))
        bold_shortcut.triggered.connect(self.make_bold)
        self.addAction(bold_shortcut)
        
        # Italic: Ctrl+I
        italic_shortcut = QAction(self)
        italic_shortcut.setShortcut(QKeySequence("Ctrl+I"))
        italic_shortcut.triggered.connect(self.make_italic)
        self.addAction(italic_shortcut)
        
        # Underline: Ctrl+U
        underline_shortcut = QAction(self)
        underline_shortcut.setShortcut(QKeySequence("Ctrl+U"))
        underline_shortcut.triggered.connect(self.make_underline)
        self.addAction(underline_shortcut)

    # =================================================
    # MODERN STYLES
    # =================================================
    def modern_styles(self):
        return """
        /* =============================================
           GLOBAL STYLES
           ============================================= */
        QWidget {
            background-color: #f8fafc;
            color: #1e293b;
            font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, 'Helvetica Neue', Arial, sans-serif;
            font-size: 11pt;
        }
        
        /* =============================================
           HEADER STYLES
           ============================================= */
        #headerTitle {
            font-size: 24pt;
            font-weight: 600;
            color: #1e40af;
            margin: 10px 0;
        }
        
        #headerSubtitle {
            font-size: 12pt;
            color: #64748b;
            margin-bottom: 10px;
        }
        
        /* =============================================
           CONTROLS FRAME
           ============================================= */
        #controlsFrame {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #ffffff, stop:1 #f1f5f9);
            border: 1px solid #e2e8f0;
            border-radius: 12px;
            margin: 5px 0;
        }
        
        /* =============================================
           BUTTON STYLES
           ============================================= */
        QPushButton {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #3b82f6, stop:1 #2563eb);
            color: white;
            border: none;
            border-radius: 8px;
            padding: 12px 24px;
            font-weight: 500;
            font-size: 11pt;
            min-width: 120px;
        }
        
        QPushButton:hover {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #60a5fa, stop:1 #3b82f6);
        }
        
        QPushButton:pressed {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #2563eb, stop:1 #1d4ed8);
        }
        
        QPushButton:disabled {
            background: #e2e8f0;
            color: #94a3b8;
        }
        
        #primaryButton {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #6366f1, stop:1 #4f46e5);
        }
        
        #primaryButton:hover {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #818cf8, stop:1 #6366f1);
        }
        
        #successButton {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #10b981, stop:1 #059669);
        }
        
        #successButton:hover {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #34d399, stop:1 #10b981);
        }
        
        #dangerButton {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #ef4444, stop:1 #dc2626);
        }
        
        #dangerButton:hover {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #f87171, stop:1 #ef4444);
        }
        
        #secondaryButton {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #64748b, stop:1 #475569);
        }
        
        #secondaryButton:hover {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #94a3b8, stop:1 #64748b);
        }
        
        /* =============================================
           SECTION TITLES
           ============================================= */
        #sectionTitle {
            font-size: 14pt;
            font-weight: 600;
            color: #1e40af;
            margin: 5px 0;
            padding: 8px 0;
            border-bottom: 2px solid #e2e8f0;
        }
        
        /* =============================================
           FRAME STYLES
           ============================================= */
        #tableFrame, #emailFrame, #statusFrame {
            background: white;
            border: 1px solid #e2e8f0;
            border-radius: 12px;
            padding: 15px;
        }
        
        /* =============================================
           INPUT STYLES
           ============================================= */
        QLineEdit, QTextEdit, QComboBox {
            background: white;
            border: 2px solid #e2e8f0;
            border-radius: 8px;
            padding: 10px;
            font-size: 11pt;
        }
        
        QLineEdit:focus, QTextEdit:focus, QComboBox:focus {
            border: 2px solid #3b82f6;
        }
        
        /* =============================================
           TABLE STYLES
           ============================================= */
        QTableWidget {
            background: white;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            gridline-color: #f1f5f9;
            selection-background-color: #3b82f6;
            alternate-background-color: #f8fafc;
        }
        
        QTableWidget::item {
            padding: 10px;
            border-bottom: 1px solid #f1f5f9;
        }
        
        QTableWidget::item:selected {
            background: #3b82f6;
            color: white;
        }
        
        QHeaderView::section {
            background: #f8fafc;
            color: #374151;
            padding: 12px;
            border: none;
            border-right: 1px solid #e2e8f0;
            border-bottom: 2px solid #e2e8f0;
            font-weight: 600;
        }
        
        /* =============================================
           PROGRESS BAR
           ============================================= */
        QProgressBar {
            background: #f1f5f9;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            text-align: center;
            color: #374151;
            font-weight: 600;
            height: 24px;
        }
        
        QProgressBar::chunk {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                stop:0 #10b981, stop:1 #059669);
            border-radius: 6px;
            margin: 2px;
        }
        
        /* =============================================
           GROUP BOX
           ============================================= */
        QGroupBox {
            background: #f8fafc;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            margin-top: 10px;
            padding-top: 20px;
            font-weight: 600;
            color: #374151;
        }
        
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 10px;
            padding: 0 5px;
        }
        
        /* =============================================
           COMBO BOX
           ============================================= */
        QComboBox::drop-down {
            border: none;
            width: 20px;
        }
        
        QComboBox::down-arrow {
            image: none;
            border-left: 5px solid transparent;
            border-right: 5px solid transparent;
            border-top: 5px solid #64748b;
        }
        
        QComboBox QAbstractItemView {
            background: white;
            border: 1px solid #e2e8f0;
            selection-background-color: #3b82f6;
            color: #374151;
        }
        
        /* =============================================
           SCROLLBAR
           ============================================= */
        QScrollBar:vertical {
            background: #f1f5f9;
            border: none;
            border-radius: 6px;
            width: 12px;
        }
        
        QScrollBar::handle:vertical {
            background: #cbd5e1;
            border-radius: 6px;
            min-height: 20px;
        }
        
        QScrollBar::handle:vertical:hover {
            background: #94a3b8;
        }
        
        /* =============================================
           SPECIAL ELEMENTS
           ============================================= */
        #recipientCounter {
            background: #eff6ff;
            color: #1d4ed8;
            border: 1px solid #3b82f6;
            border-radius: 20px;
            padding: 6px 12px;
            font-weight: 600;
            font-size: 10pt;
        }
        
        #templateCombo, #spacingSelect {
            background: #f8fafc;
            border: 2px solid #3b82f6;
        }
        
        #subjectInput {
            background: #eff6ff;
            border: 2px solid #3b82f6;
            font-weight: 500;
        }
        
        #emailEditor {
            background: white;
            border: 2px solid #e2e8f0;
            font-family: 'Segoe UI', system-ui, sans-serif;
            line-height: 1.5;
        }
        
        #logBox {
            background: #1e293b;
            color: #e2e8f0;
            border: 2px solid #334155;
            font-family: 'Consolas', 'Monaco', monospace;
            font-size: 10pt;
        }
        
        /* =============================================
           TAB WIDGET STYLING
           ============================================= */
        QTabWidget::pane {
            border: 1px solid #e2e8f0;
            background: white;
            border-radius: 8px;
            top: -1px;
        }
        
        QTabBar::tab {
            background: #f8fafc;
            border: 1px solid #e2e8f0;
            padding: 12px 24px;
            margin-right: 2px;
            border-top-left-radius: 8px;
            border-top-right-radius: 8px;
            font-weight: 500;
            font-size: 11pt;
            color: #64748b;
        }
        
        QTabBar::tab:selected {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #3b82f6, stop:1 #2563eb);
            color: white;
            border-bottom: 2px solid #2563eb;
        }
        
        QTabBar::tab:hover:!selected {
            background: #e2e8f0;
            color: #1e40af;
        }
        
        #mainTabWidget QTabBar::tab {
            min-width: 150px;
        }
        """

# =====================================================
# RUN APPLICATION
# =====================================================
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = EmailApp()
    window.show()
    sys.exit(app.exec())
