import sys
import os
import time
import re
import json
import logging
import sqlite3
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
from PySide6.QtCore import Qt, QThread, Signal, QSize, QTimer

# =====================================================
# DATABASE MANAGER
# =====================================================
class DatabaseManager:
    """Manages SQLite database for persistent storage"""
    
    def __init__(self, db_name="eru_email_sender.db"):
        # Handle both script and executable environments
        if getattr(sys, 'frozen', False):
            # Running as PyInstaller executable
            app_data_dir = os.path.join(os.path.expanduser("~"), "EruEmailSender")
            os.makedirs(app_data_dir, exist_ok=True)
            self.db_path = os.path.join(app_data_dir, db_name)
        else:
            # Running as script
            script_dir = os.path.dirname(os.path.abspath(__file__))
            self.db_path = os.path.join(script_dir, db_name)
        
        self.initialize_database()
    
    def initialize_database(self):
        """Create database tables if they don't exist"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Settings table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            # Templates table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS templates (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT UNIQUE NOT NULL,
                    subject TEXT,
                    body TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            # Campaigns table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS campaigns (
                    id TEXT PRIMARY KEY,
                    name TEXT NOT NULL,
                    total_recipients INTEGER DEFAULT 0,
                    confirmed_count INTEGER DEFAULT 0,
                    failed_count INTEGER DEFAULT 0,
                    unknown_count INTEGER DEFAULT 0,
                    cancelled_count INTEGER DEFAULT 0,
                    status TEXT DEFAULT 'pending',
                    started_at TIMESTAMP,
                    completed_at TIMESTAMP,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            # Recipients table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS recipients (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    campaign_id TEXT NOT NULL,
                    account TEXT,
                    full_name TEXT,
                    email TEXT,
                    cc TEXT,
                    attachment_path TEXT,
                    status TEXT DEFAULT 'pending',
                    attempt_number INTEGER DEFAULT 0,
                    last_error TEXT,
                    last_attempt_time TIMESTAMP,
                    row_index INTEGER,
                    FOREIGN KEY (campaign_id) REFERENCES campaigns(id)
                )
            ''')
            
            # Send attempts table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS send_attempts (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    recipient_id INTEGER NOT NULL,
                    campaign_id TEXT NOT NULL,
                    attempt_number INTEGER NOT NULL,
                    status TEXT NOT NULL,
                    error_message TEXT,
                    timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (recipient_id) REFERENCES recipients(id),
                    FOREIGN KEY (campaign_id) REFERENCES campaigns(id)
                )
            ''')
            
            # Send logs table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS send_logs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    campaign_id TEXT,
                    recipient_id INTEGER,
                    log_level TEXT DEFAULT 'info',
                    message TEXT,
                    timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (campaign_id) REFERENCES campaigns(id),
                    FOREIGN KEY (recipient_id) REFERENCES recipients(id)
                )
            ''')
            
            conn.commit()
            conn.close()
            
        except Exception as e:
            print(f"Error initializing database: {e}")
    
    def get_connection(self):
        """Get a database connection"""
        return sqlite3.connect(self.db_path)
    
    def save_setting(self, key, value):
        """Save a setting to the database"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                INSERT OR REPLACE INTO settings (key, value, updated_at)
                VALUES (?, ?, CURRENT_TIMESTAMP)
            ''', (key, json.dumps(value) if isinstance(value, (dict, list)) else str(value)))
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error saving setting: {e}")
            return False
    
    def get_setting(self, key, default=None):
        """Get a setting from the database"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute('SELECT value FROM settings WHERE key = ?', (key,))
            result = cursor.fetchone()
            conn.close()
            
            if result:
                value = result[0]
                # Try to parse as JSON
                try:
                    return json.loads(value)
                except:
                    return value
            return default
        except Exception as e:
            print(f"Error getting setting: {e}")
            return default
    
    def create_campaign(self, campaign_id, name, total_recipients):
        """Create a new campaign"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO campaigns (id, name, total_recipients, status, started_at)
                VALUES (?, ?, ?, 'in_progress', CURRENT_TIMESTAMP)
            ''', (campaign_id, name, total_recipients))
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error creating campaign: {e}")
            return False
    
    def update_campaign_status(self, campaign_id, status):
        """Update campaign status"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            if status == 'completed':
                cursor.execute('''
                    UPDATE campaigns 
                    SET status = ?, completed_at = CURRENT_TIMESTAMP
                    WHERE id = ?
                ''', (status, campaign_id))
            else:
                cursor.execute('''
                    UPDATE campaigns SET status = ? WHERE id = ?
                ''', (status, campaign_id))
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error updating campaign status: {e}")
            return False
    
    def update_campaign_counts(self, campaign_id):
        """Update campaign count statistics"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            
            cursor.execute('''
                UPDATE campaigns SET
                    confirmed_count = (SELECT COUNT(*) FROM recipients WHERE campaign_id = ? AND status = 'confirmed'),
                    failed_count = (SELECT COUNT(*) FROM recipients WHERE campaign_id = ? AND status = 'failed'),
                    unknown_count = (SELECT COUNT(*) FROM recipients WHERE campaign_id = ? AND status = 'unknown'),
                    cancelled_count = (SELECT COUNT(*) FROM recipients WHERE campaign_id = ? AND status = 'cancelled')
                WHERE id = ?
            ''', (campaign_id, campaign_id, campaign_id, campaign_id, campaign_id))
            
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error updating campaign counts: {e}")
            return False
    
    def add_recipient(self, campaign_id, account, full_name, email, cc, attachment_path, row_index):
        """Add a recipient to a campaign"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO recipients (campaign_id, account, full_name, email, cc, attachment_path, row_index)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (campaign_id, account, full_name, email, cc, attachment_path, row_index))
            conn.commit()
            recipient_id = cursor.lastrowid
            conn.close()
            return recipient_id
        except Exception as e:
            print(f"Error adding recipient: {e}")
            return None
    
    def update_recipient_status(self, recipient_id, status, error_message=None):
        """Update recipient status"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            
            # Increment attempt number
            cursor.execute('''
                UPDATE recipients 
                SET status = ?, 
                    attempt_number = attempt_number + 1,
                    last_error = ?,
                    last_attempt_time = CURRENT_TIMESTAMP
                WHERE id = ?
            ''', (status, error_message, recipient_id))
            
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error updating recipient status: {e}")
            return False
    
    def log_send_attempt(self, recipient_id, campaign_id, attempt_number, status, error_message=None):
        """Log a send attempt"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO send_attempts (recipient_id, campaign_id, attempt_number, status, error_message)
                VALUES (?, ?, ?, ?, ?)
            ''', (recipient_id, campaign_id, attempt_number, status, error_message))
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error logging send attempt: {e}")
            return False
    
    def log_message(self, campaign_id, recipient_id, log_level, message):
        """Log a message"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO send_logs (campaign_id, recipient_id, log_level, message)
                VALUES (?, ?, ?, ?)
            ''', (campaign_id, recipient_id, log_level, message))
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error logging message: {e}")
            return False
    
    def get_interrupted_campaigns(self):
        """Get campaigns that were interrupted (in_progress status)"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                SELECT id, name, total_recipients, confirmed_count, failed_count, 
                       unknown_count, cancelled_count, started_at
                FROM campaigns 
                WHERE status = 'in_progress'
                ORDER BY started_at DESC
            ''')
            campaigns = cursor.fetchall()
            conn.close()
            return campaigns
        except Exception as e:
            print(f"Error getting interrupted campaigns: {e}")
            return []
    
    def get_campaign_recipients(self, campaign_id):
        """Get all recipients for a campaign"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                SELECT id, account, full_name, email, cc, attachment_path, status, 
                       attempt_number, last_error, row_index
                FROM recipients 
                WHERE campaign_id = ?
                ORDER BY row_index
            ''', (campaign_id,))
            recipients = cursor.fetchall()
            conn.close()
            return recipients
        except Exception as e:
            print(f"Error getting campaign recipients: {e}")
            return []
    
    def get_pending_recipients(self, campaign_id):
        """Get pending recipients for a campaign"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                SELECT id, account, full_name, email, cc, attachment_path, row_index
                FROM recipients 
                WHERE campaign_id = ? AND status = 'pending'
                ORDER BY row_index
            ''', (campaign_id,))
            recipients = cursor.fetchall()
            conn.close()
            return recipients
        except Exception as e:
            print(f"Error getting pending recipients: {e}")
            return []

# =====================================================
# SETTINGS MANAGER
# =====================================================
class SettingsManager:
    def __init__(self, config_file="settings.json"):
        # Initialize database manager
        self.db_manager = DatabaseManager()
        
        # Handle both script and executable environments for JSON fallback
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
        
        # Migrate from JSON to database if JSON exists
        self.migrate_from_json()
        
        # Load settings from database
        self.settings = self.load_settings()
    
    def migrate_from_json(self):
        """Migrate settings from JSON file to database"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    json_settings = json.load(f)
                    # Save each setting to database
                    for key, value in json_settings.items():
                        self.db_manager.save_setting(key, value)
                # Backup and remove old JSON file
                backup_file = self.config_file + ".backup"
                os.rename(self.config_file, backup_file)
                print(f"Migrated settings from JSON to database. Backup saved to {backup_file}")
            except Exception as e:
                print(f"Error migrating settings: {e}")
    
    def load_settings(self):
        """Load settings from database"""
        try:
            settings = self.default_settings.copy()
            # Load known settings from database
            for key in self.default_settings.keys():
                value = self.db_manager.get_setting(key)
                if value is not None:
                    settings[key] = value
            return settings
        except Exception as e:
            print(f"Error loading settings: {e}")
            return self.default_settings.copy()
    
    def save_settings(self):
        """Save settings to database"""
        try:
            for key, value in self.settings.items():
                self.db_manager.save_setting(key, value)
            return True
        except Exception:
            return False
    
    def get(self, key, default=None):
        return self.settings.get(key, default)
    
    def set(self, key, value):
        self.settings[key] = value
        self.db_manager.save_setting(key, value)

# =====================================================
# EMAIL STATE MACHINE
# =====================================================
class EmailState:
    """Email sending states for proper state tracking"""
    PENDING = "Pending"
    VALIDATING = "Validating"
    SENDING = "Sending"
    SUBMITTED = "Submitted"
    CONFIRMED = "Confirmed"
    FAILED = "Failed"
    UNKNOWN = "Unknown"
    SKIPPED = "Skipped"
    CANCELLED = "Cancelled"
    
    @staticmethod
    def is_terminal(state):
        """States that are final and should not be automatically retried"""
        return state in [EmailState.CONFIRMED, EmailState.SKIPPED, EmailState.CANCELLED]
    
    @staticmethod
    def can_retry(state):
        """States that can be safely retried"""
        return state in [EmailState.FAILED]
    
    @staticmethod
    def is_uncertain(state):
        """States where we're unsure if the email was sent"""
        return state in [EmailState.UNKNOWN, EmailState.SUBMITTED]

# =====================================================
# OUTLOOK CLIENT MANAGER
# =====================================================
class OutlookClient:
    """Manages Outlook COM connection lifecycle"""
    
    def __init__(self):
        self.outlook = None
        self.namespace = None
        self.outbox = None
        self.sent = None
        self.is_connected = False
    
    def connect(self, max_attempts=5, retry_delay=3):
        """Connect to Outlook with retry logic"""
        import pythoncom
        
        for attempt in range(max_attempts):
            try:
                pythoncom.CoInitialize()
                self.outlook = win32com.client.Dispatch("Outlook.Application")
                self.namespace = self.outlook.GetNamespace("MAPI")
                self.outbox = self.namespace.GetDefaultFolder(4)  # olFolderOutbox
                self.sent = self.namespace.GetDefaultFolder(5)    # olFolderSentMail
                self.is_connected = True
                return True, "Connected to Outlook successfully"
            except Exception as e:
                if attempt < max_attempts - 1:
                    time.sleep(retry_delay)
                else:
                    return False, f"Failed to connect to Outlook after {max_attempts} attempts: {str(e)}"
        
        return False, "Unknown connection error"
    
    def disconnect(self):
        """Safely disconnect from Outlook"""
        try:
            import pythoncom
            self.is_connected = False
            self.outlook = None
            self.namespace = None
            self.outbox = None
            self.sent = None
            pythoncom.CoUninitialize()
        except Exception:
            pass  # Best effort cleanup
    
    def is_available(self):
        """Check if Outlook connection is available"""
        if not self.is_connected or self.outlook is None:
            return False
        try:
            # Try to access a simple property to test connection
            _ = self.outlook.Version
            return True
        except Exception:
            self.is_connected = False
            return False
    
    def create_email(self):
        """Create a new email item"""
        if not self.is_available():
            raise Exception("Outlook is not available")
        return self.outlook.CreateItem(0)  # olMailItem

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

def parse_cc_addresses(cc_string):
    """Parse multiple CC addresses from semicolon or comma separated string"""
    if not cc_string or not str(cc_string).strip():
        return []
    
    cc_string = str(cc_string).strip()
    # Try semicolon separator first, then comma
    if ';' in cc_string:
        addresses = [addr.strip() for addr in cc_string.split(';')]
    elif ',' in cc_string:
        addresses = [addr.strip() for addr in cc_string.split(',')]
    else:
        addresses = [cc_string]
    
    # Filter out empty addresses
    return [addr for addr in addresses if addr]

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
# EXCEL VALIDATION
# =====================================================
def validate_excel_data(df):
    """
    Comprehensive Excel data validation before sending
    Returns (is_valid: bool, validation_results: dict)
    """
    validation_results = {
        'total_rows': len(df),
        'valid_rows': 0,
        'invalid_emails': [],
        'missing_attachments': [],
        'empty_required_fields': [],
        'duplicate_recipients': [],
        'invalid_cc': [],
        'warnings': []
    }
    
    if df is None or len(df) == 0:
        validation_results['warnings'].append("Excel file is empty")
        return False, validation_results
    
    # Check required columns
    required_columns = ['Account', 'Full Name', 'Email', 'CC', 'Attachment Path']
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        validation_results['warnings'].append(f"Missing columns: {', '.join(missing_columns)}")
        return False, validation_results
    
    # Track seen emails for duplicate detection
    seen_emails = {}
    
    for idx, row in df.iterrows():
        row_num = idx + 2  # Excel row number
        has_issues = False
        
        # Validate email
        email = str(row['Email']).strip()
        if email:
            is_valid, error = validate_email(email)
            if not is_valid:
                validation_results['invalid_emails'].append({
                    'row': row_num,
                    'email': email,
                    'error': error
                })
                has_issues = True
            
            # Check for duplicates
            if email.lower() in seen_emails:
                validation_results['duplicate_recipients'].append({
                    'row': row_num,
                    'email': email,
                    'duplicate_row': seen_emails[email.lower()]
                })
                has_issues = True
            else:
                seen_emails[email.lower()] = row_num
        else:
            validation_results['empty_required_fields'].append({
                'row': row_num,
                'field': 'Email'
            })
            has_issues = True
        
        # Validate CC addresses
        cc_value = str(row['CC']).strip()
        if cc_value:
            cc_addresses = parse_cc_addresses(cc_value)
            for cc_addr in cc_addresses:
                is_valid, error = validate_email(cc_addr)
                if not is_valid:
                    validation_results['invalid_cc'].append({
                        'row': row_num,
                        'cc': cc_addr,
                        'error': error
                    })
                    has_issues = True
        
        # Validate attachment path
        attachment = str(row['Attachment Path']).strip()
        if attachment and not os.path.exists(attachment):
            validation_results['missing_attachments'].append({
                'row': row_num,
                'path': attachment
            })
            has_issues = True
        
        # Check other required fields
        if not str(row['Account']).strip():
            validation_results['empty_required_fields'].append({
                'row': row_num,
                'field': 'Account'
            })
            has_issues = True
        
        if not str(row['Full Name']).strip():
            validation_results['empty_required_fields'].append({
                'row': row_num,
                'field': 'Full Name'
            })
            has_issues = True
        
        if not has_issues:
            validation_results['valid_rows'] += 1
    
    # Determine overall validity
    is_valid = (
        len(validation_results['invalid_emails']) == 0 and
        len(validation_results['missing_attachments']) == 0 and
        len(validation_results['empty_required_fields']) == 0 and
        len(validation_results['invalid_cc']) == 0
    )
    
    return is_valid, validation_results

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
# EMAIL WORKER THREAD WITH ENHANCED SAFETY
# =====================================================
class EmailWorker(QThread):
    progress_updated = Signal(int)
    log_updated = Signal(str)
    status_updated = Signal(int, str)
    finished_sending = Signal()
    validation_complete = Signal(bool, object)  # is_valid, validation_results
    campaign_created = Signal(str)  # campaign_id

    def __init__(self, dataframe, subject, body_template, para_spacing_px=12, max_retries=3, 
                 send_delay=0, importance=2, request_read_receipt=True, campaign_name=None, db_manager=None, is_test_send=False):
        super().__init__()
        self.df = dataframe
        self.subject = subject
        self.body_template = body_template
        self.para_spacing_px = int(para_spacing_px) if para_spacing_px is not None else 12
        self.max_retries = max_retries
        self.send_delay = send_delay
        self.importance = importance  # 0=Low, 1=Normal, 2=High
        self.request_read_receipt = request_read_receipt
        self.running = True
        self.paused = False
        
        # Database manager for persistence
        self.db_manager = db_manager
        self.is_test_send = is_test_send
        
        # Campaign management
        self.campaign_id = None
        self.campaign_name = campaign_name or f"Campaign {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        
        # Track recipient states and attempts
        self.recipient_states = {}
        self.recipient_attempts = {}
        self.recipient_db_ids = {}  # Map row index to database ID
        
        # Initialize states
        for idx in range(len(dataframe)):
            self.recipient_states[idx] = EmailState.PENDING
            self.recipient_attempts[idx] = 0

    def stop(self):
        self.running = False

    def pause(self):
        self.paused = True

    def resume(self):
        self.paused = False

    def run(self):
        import pythoncom
        
        try:
            pythoncom.CoInitialize()
            
            try:
                # Phase 0: Create campaign in database (skip for test sends)
                if self.db_manager and not self.is_test_send:
                    # Generate unique campaign ID
                    date_str = datetime.now().strftime('%Y%m%d')
                    # Simple ID generation - in production you'd want something more robust
                    self.campaign_id = f"ESM-{date_str}-{int(time.time())}"
                    
                    # Create campaign
                    self.db_manager.create_campaign(self.campaign_id, self.campaign_name, len(self.df))
                    self.campaign_created.emit(self.campaign_id)
                    self.log_updated.emit(f"📋 Campaign created: {self.campaign_id}")
                    
                    # Add recipients to database
                    for index, row in self.df.iterrows():
                        recipient_id = self.db_manager.add_recipient(
                            self.campaign_id,
                            str(row.get("Account", "")),
                            str(row.get("Full Name", "")),
                            str(row.get("Email", "")),
                            str(row.get("CC", "")),
                            str(row.get("Attachment Path", "")),
                            index
                        )
                        if recipient_id:
                            self.recipient_db_ids[index] = recipient_id
                
                # Phase 1: Validation
                self.log_updated.emit("🔍 Starting Excel data validation...")
                is_valid, validation_results = validate_excel_data(self.df)
                self.validation_complete.emit(is_valid, validation_results)
                
                if not is_valid and not self.is_test_send:
                    self.log_updated.emit("❌ Validation failed. Please fix the errors before sending.")
                    if self.db_manager:
                        self.db_manager.update_campaign_status(self.campaign_id, 'validation_failed')
                    self.finished_sending.emit()
                    return
                
                self.log_updated.emit(f"✅ Validation complete. {validation_results['valid_rows']} valid rows ready to send.")
                
                # Phase 2: Connect to Outlook
                outlook_client = OutlookClient()
                success, message = outlook_client.connect()
                
                if not success:
                    self.log_updated.emit(f"❌ {message}")
                    self.log_updated.emit("Outlook is unavailable. Please make sure Microsoft Outlook is installed and configured, then try again.")
                    if self.db_manager:
                        self.db_manager.update_campaign_status(self.campaign_id, 'outlook_unavailable')
                    self.finished_sending.emit()
                    return
                
                self.log_updated.emit("✅ Connected to Outlook successfully.")
                
                # Phase 3: Send emails
                total = len(self.df)
                processed_count = 0
                
                for index, row in self.df.iterrows():
                    if not self.running:
                        self.log_updated.emit("⛔ Sending stopped by user.")
                        # Mark remaining as cancelled
                        for idx in range(index, len(self.df)):
                            if self.recipient_states[idx] == EmailState.PENDING:
                                self.recipient_states[idx] = EmailState.CANCELLED
                                self.status_updated.emit(idx, EmailState.CANCELLED)
                        break
                    
                    # Wait if paused
                    while self.paused and self.running:
                        time.sleep(0.1)
                    
                    if not self.running:
                        break
                    
                    # Skip if already processed
                    if EmailState.is_terminal(self.recipient_states[index]):
                        processed_count += 1
                        percent = int((processed_count / total) * 100)
                        self.progress_updated.emit(percent)
                        continue
                    
                    # Process this recipient
                    self.recipient_states[index] = EmailState.SENDING
                    self.status_updated.emit(index, EmailState.SENDING)
                    
                    result = self._send_single_email(outlook_client, row, index)
                    
                    if result['success']:
                        self.recipient_states[index] = EmailState.CONFIRMED
                        self.status_updated.emit(index, EmailState.CONFIRMED)
                        self.log_updated.emit(f"✅ Sent to {result['email']}")
                        
                        # Update database
                        if self.db_manager and index in self.recipient_db_ids:
                            recipient_id = self.recipient_db_ids[index]
                            self.db_manager.update_recipient_status(recipient_id, EmailState.CONFIRMED)
                            self.db_manager.log_send_attempt(recipient_id, self.campaign_id, 
                                                           self.recipient_attempts[index] + 1, 
                                                           EmailState.CONFIRMED)
                            
                    elif result['state'] == EmailState.UNKNOWN:
                        self.recipient_states[index] = EmailState.UNKNOWN
                        self.status_updated.emit(index, EmailState.UNKNOWN)
                        self.log_updated.emit(f"⚠️ Uncertain state for {result['email']}: {result['error']}")
                        
                        # Update database
                        if self.db_manager and index in self.recipient_db_ids:
                            recipient_id = self.recipient_db_ids[index]
                            self.db_manager.update_recipient_status(recipient_id, EmailState.UNKNOWN, result['error'])
                            self.db_manager.log_send_attempt(recipient_id, self.campaign_id,
                                                           self.recipient_attempts[index] + 1,
                                                           EmailState.UNKNOWN, result['error'])
                            
                    else:
                        self.recipient_states[index] = EmailState.FAILED
                        self.status_updated.emit(index, EmailState.FAILED)
                        self.log_updated.emit(f"❌ Failed to send to {result['email']}: {result['error']}")
                        
                        # Update database
                        if self.db_manager and index in self.recipient_db_ids:
                            recipient_id = self.recipient_db_ids[index]
                            self.db_manager.update_recipient_status(recipient_id, EmailState.FAILED, result['error'])
                            self.db_manager.log_send_attempt(recipient_id, self.campaign_id,
                                                           self.recipient_attempts[index] + 1,
                                                           EmailState.FAILED, result['error'])
                    
                    # Apply delay between emails
                    if self.send_delay > 0 and index < len(self.df) - 1:
                        time.sleep(self.send_delay)
                    
                    processed_count += 1
                    percent = int((processed_count / total) * 100)
                    self.progress_updated.emit(percent)
                
                # Phase 4: Retry failed emails (only true failures, not unknown)
                failed_indices = [idx for idx, state in self.recipient_states.items() 
                                if state == EmailState.FAILED and self.recipient_attempts[idx] < self.max_retries]
                
                if failed_indices and self.max_retries > 0:
                    self.log_updated.emit(f"🔄 Retrying {len(failed_indices)} failed emails...")
                    
                    for retry_attempt in range(self.max_retries):
                        if not self.running or not failed_indices:
                            break
                        
                        self.log_updated.emit(f"🔄 Retry attempt {retry_attempt + 1}/{self.max_retries}")
                        still_failed = []
                        
                        for index in failed_indices:
                            if not self.running:
                                break
                            
                            while self.paused and self.running:
                                time.sleep(0.1)
                            
                            if not self.running:
                                break
                            
                            row = self.df.iloc[index]
                            self.recipient_states[index] = EmailState.SENDING
                            self.status_updated.emit(index, EmailState.SENDING)
                            self.recipient_attempts[index] += 1
                            
                            result = self._send_single_email(outlook_client, row, index)
                            
                            if result['success']:
                                self.recipient_states[index] = EmailState.CONFIRMED
                                self.status_updated.emit(index, EmailState.CONFIRMED)
                                self.log_updated.emit(f"✅ Sent to {result['email']}")
                            elif result['state'] == EmailState.UNKNOWN:
                                self.recipient_states[index] = EmailState.UNKNOWN
                                self.status_updated.emit(index, EmailState.UNKNOWN)
                                still_failed.append(index)
                            else:
                                self.recipient_states[index] = EmailState.FAILED
                                self.status_updated.emit(index, EmailState.FAILED)
                                still_failed.append(index)
                            
                            if self.send_delay > 0:
                                time.sleep(self.send_delay)
                        
                        failed_indices = still_failed
                        
                        if failed_indices:
                            time.sleep(2)  # Wait before next retry attempt
                
                # Summary
                confirmed = sum(1 for state in self.recipient_states.values() if state == EmailState.CONFIRMED)
                failed = sum(1 for state in self.recipient_states.values() if state == EmailState.FAILED)
                unknown = sum(1 for state in self.recipient_states.values() if state == EmailState.UNKNOWN)
                cancelled = sum(1 for state in self.recipient_states.values() if state == EmailState.CANCELLED)
                
                self.log_updated.emit(f"📊 Sending complete: {confirmed} sent, {failed} failed, {unknown} unknown, {cancelled} cancelled")
                
                if unknown > 0:
                    self.log_updated.emit("⚠️ Some emails have unknown status. These were not automatically retried to prevent duplicate sends.")
                
                # Update campaign in database
                if self.db_manager and not self.is_test_send:
                    self.db_manager.update_campaign_counts(self.campaign_id)
                    if cancelled > 0:
                        self.db_manager.update_campaign_status(self.campaign_id, 'cancelled')
                    else:
                        self.db_manager.update_campaign_status(self.campaign_id, 'completed')
                
                self.finished_sending.emit()
                
            finally:
                outlook_client.disconnect()
                pythoncom.CoUninitialize()
                
        except Exception as e:
            self.log_updated.emit(f"FATAL ERROR in worker: {str(e)}")
            self.finished_sending.emit()

    def _send_single_email(self, outlook_client, row, index):
        """Send a single email with enhanced error handling and state tracking"""
        try:
            email = str(row["Email"]).strip()
            cc_value = str(row["CC"]).strip()
            attachment = str(row["Attachment Path"]).strip()
            
            if not email:
                return {'success': False, 'state': EmailState.FAILED, 'email': email, 'error': 'No email address'}
            
            # Check attachment (should have been validated, but double-check)
            if attachment and not os.path.exists(attachment):
                return {'success': False, 'state': EmailState.FAILED, 'email': email, 'error': f'Attachment not found: {attachment}'}
            
            # Get starting sent count
            start_sent_count = outlook_client.sent.Items.Count
            
            # Prepare email content
            account = str(row["Account"]).strip()
            full_name = str(row["Full Name"]).strip()
            surname = get_surname(full_name)
            
            # Replace placeholders
            subject = self.subject.replace("{{account}}", account).replace("{{Account}}", account).replace("{{fullname}}", full_name)
            body_raw = self.body_template.replace("{{fullname}}", surname)
            body = build_outlook_safe_html(body_raw, self.para_spacing_px)
            
            # Create email
            mail = outlook_client.create_email()
            mail.To = email
            
            # Handle CC addresses
            if cc_value:
                cc_addresses = parse_cc_addresses(cc_value)
                valid_cc = []
                for cc_addr in cc_addresses:
                    is_valid, _ = validate_email(cc_addr)
                    if is_valid:
                        valid_cc.append(cc_addr)
                if valid_cc:
                    mail.CC = ";".join(valid_cc)
            
            mail.Subject = subject
            
            try:
                mail.BodyFormat = 2  # olFormatHTML
            except Exception:
                pass
            
            mail.HTMLBody = body
            mail.Importance = self.importance  # 0=Low, 1=Normal, 2=High
            mail.ReadReceiptRequested = self.request_read_receipt
            
            # Add attachment if path provided
            if attachment:
                mail.Attachments.Add(attachment)
            
            # Send the email
            mail.Send()
            
            # Wait for confirmation (max 30 seconds)
            confirmation_received = False
            for wait_attempt in range(30):
                try:
                    if outlook_client.outbox.Items.Count == 0 or outlook_client.sent.Items.Count > start_sent_count:
                        confirmation_received = True
                        break
                    time.sleep(1)
                except Exception:
                    # Outlook became unavailable during wait
                    return {'success': False, 'state': EmailState.UNKNOWN, 'email': email, 'error': 'Outlook unavailable during confirmation wait'}
            
            if confirmation_received:
                return {'success': True, 'state': EmailState.CONFIRMED, 'email': email, 'error': None}
            else:
                # Timeout - we don't know if it was sent
                return {'success': False, 'state': EmailState.UNKNOWN, 'email': email, 'error': 'Confirmation timeout - email may have been sent'}
            
        except Exception as e:
            error_msg = str(e)
            # Check if this might be a transient error
            if "unavailable" in error_msg.lower() or "not ready" in error_msg.lower():
                return {'success': False, 'state': EmailState.UNKNOWN, 'email': email, 'error': error_msg}
            else:
                return {'success': False, 'state': EmailState.FAILED, 'email': email, 'error': error_msg}


# =====================================================
# MAIN APPLICATION
# =====================================================
class EmailApp(QWidget):
    def __init__(self):
        super().__init__()

        # Initialize database manager
        self.db_manager = DatabaseManager()
        
        # Initialize settings manager (now uses database)
        self.settings = SettingsManager()
        
        # Setup file-based logging
        self.setup_logging()
        
        # Check for interrupted campaigns
        self.check_interrupted_campaigns()
        
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
        self.current_campaign_id = None
        
        # Statistics update timer (initially stopped)
        self.stats_timer = QTimer()
        self.stats_timer.timeout.connect(self.update_statistics)
        
        # Initialize UI state
        self.load_templates()  # Load saved templates
        # Unblock signals after initialization is complete
        self.template_combo.blockSignals(False)
        self.setup_keyboard_shortcuts()  # Setup keyboard shortcuts
        self.update_ui_state()
    
    def check_interrupted_campaigns(self):
        """Check for interrupted campaigns and show recovery dialog"""
        try:
            interrupted = self.db_manager.get_interrupted_campaigns()
            if interrupted:
                self.show_recovery_dialog(interrupted)
        except Exception as e:
            print(f"Error checking interrupted campaigns: {e}")
    
    def show_recovery_dialog(self, campaigns):
        """Show dialog for interrupted campaign recovery"""
        from PySide6.QtWidgets import QDialog, QVBoxLayout, QTextBrowser, QHBoxLayout, QLabel, QPushButton, QListWidget, QListWidgetItem
        
        dialog = QDialog(self)
        dialog.setWindowTitle("⚠️ Interrupted Campaigns Found")
        dialog.setMinimumSize(600, 400)
        dialog.setStyleSheet(self.modern_styles())
        
        layout = QVBoxLayout(dialog)
        
        # Information label
        info_label = QLabel("The following campaigns were interrupted and can be resumed:")
        layout.addWidget(info_label)
        
        # Campaign list
        campaign_list = QListWidget()
        for campaign in campaigns:
            (campaign_id, name, total, confirmed, failed, unknown, cancelled, started_at) = campaign
            pending = total - confirmed - failed - unknown - cancelled
            item_text = f"{name} ({campaign_id})\n"
            item_text += f"  Total: {total} | Sent: {confirmed} | Failed: {failed} | Unknown: {unknown} | Pending: {pending}"
            item = QListWidgetItem(item_text)
            item.setData(Qt.UserRole, campaign_id)
            campaign_list.addItem(item)
        
        layout.addWidget(campaign_list)
        
        # Buttons
        button_layout = QHBoxLayout()
        
        resume_button = QPushButton("📥 Resume Selected")
        resume_button.setObjectName("successButton")
        resume_button.clicked.connect(lambda: self.resume_campaign(dialog, campaign_list))
        
        review_button = QPushButton("👁️ Review Details")
        review_button.setObjectName("primaryButton")
        review_button.clicked.connect(lambda: self.review_campaign(dialog, campaign_list))
        
        cancel_button = QPushButton("❌ Cancel")
        cancel_button.setObjectName("dangerButton")
        cancel_button.clicked.connect(dialog.close)
        
        button_layout.addWidget(resume_button)
        button_layout.addWidget(review_button)
        button_layout.addStretch()
        button_layout.addWidget(cancel_button)
        
        layout.addLayout(button_layout)
        
        dialog.exec()
    
    def resume_campaign(self, dialog, campaign_list):
        """Resume selected campaign"""
        selected_items = campaign_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "No Selection", "Please select a campaign to resume.")
            return
        
        campaign_id = selected_items[0].data(Qt.UserRole)
        dialog.close()
        
        # Load campaign data
        try:
            recipients = self.db_manager.get_campaign_recipients(campaign_id)
            if recipients:
                # Convert to dataframe
                data = []
                for rec in recipients:
                    (rec_id, account, full_name, email, cc, attachment_path, status, 
                     attempt_number, last_error, row_index) = rec
                    data.append({
                        'Account': account,
                        'Full Name': full_name,
                        'Email': email,
                        'CC': cc,
                        'Attachment Path': attachment_path,
                        'Status': status
                    })
                
                self.df = pd.DataFrame(data)
                self.populate_table()
                self.update_ui_state()
                
                QMessageBox.information(self, "Campaign Loaded", 
                                      f"Campaign {campaign_id} has been loaded. "
                                      f"You can review the status and send pending emails.")
                self.log(f"📥 Resumed campaign: {campaign_id}")
            else:
                QMessageBox.warning(self, "Error", "No recipients found for this campaign.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load campaign: {str(e)}")
    
    def review_campaign(self, dialog, campaign_list):
        """Review selected campaign details"""
        selected_items = campaign_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "No Selection", "Please select a campaign to review.")
            return
        
        campaign_id = selected_items[0].data(Qt.UserRole)
        
        # Show detailed information
        try:
            recipients = self.db_manager.get_campaign_recipients(campaign_id)
            if recipients:
                # Create a simple details dialog
                from PySide6.QtWidgets import QDialog, QVBoxLayout, QTextEdit
                
                details_dialog = QDialog(self)
                details_dialog.setWindowTitle(f"Campaign Details: {campaign_id}")
                details_dialog.setMinimumSize(800, 600)
                details_dialog.setStyleSheet(self.modern_styles())
                
                layout = QVBoxLayout(details_dialog)
                
                details_text = QTextEdit()
                details_text.setReadOnly(True)
                
                # Build details text
                details = f"Campaign ID: {campaign_id}\n\n"
                details += "RECIPIENTS:\n"
                details += "-" * 80 + "\n"
                
                for rec in recipients:
                    (rec_id, account, full_name, email, cc, attachment_path, status, 
                     attempt_number, last_error, row_index) = rec
                    details += f"Row {row_index + 2}: {full_name} ({email})\n"
                    details += f"  Status: {status} | Attempts: {attempt_number}\n"
                    if last_error:
                        details += f"  Last Error: {last_error}\n"
                    details += "\n"
                
                details_text.setPlainText(details)
                layout.addWidget(details_text)
                
                close_button = QPushButton("Close")
                close_button.clicked.connect(details_dialog.close)
                layout.addWidget(close_button)
                
                details_dialog.exec()
            else:
                QMessageBox.warning(self, "Error", "No recipients found for this campaign.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load campaign details: {str(e)}")

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
        
        self.settings_button = QPushButton("⚙️ Settings")
        self.settings_button.setObjectName("secondaryButton")
        self.settings_button.clicked.connect(self.show_settings_dialog)
        
        # Enhanced sending controls
        self.pause_button = QPushButton("⏸️ Pause")
        self.pause_button.setObjectName("secondaryButton")
        self.pause_button.setEnabled(False)
        
        self.resume_button = QPushButton("▶️ Resume")
        self.resume_button.setObjectName("successButton")
        self.resume_button.setEnabled(False)
        
        self.retry_failed_button = QPushButton("🔄 Retry Failed")
        self.retry_failed_button.setObjectName("primaryButton")
        self.retry_failed_button.setEnabled(False)
        
        self.test_send_button = QPushButton("🧪 Test Send")
        self.test_send_button.setObjectName("secondaryButton")
        
        self.view_history_button = QPushButton("📜 History")
        self.view_history_button.setObjectName("secondaryButton")
        self.view_history_button.clicked.connect(self.view_campaign_history)
        
        # Add buttons to layout
        controls_layout.addWidget(self.export_button)
        controls_layout.addWidget(self.load_button)
        controls_layout.addWidget(self.settings_button)
        controls_layout.addWidget(self.test_send_button)
        controls_layout.addWidget(self.view_history_button)
        controls_layout.addStretch()
        controls_layout.addWidget(self.start_button)
        controls_layout.addWidget(self.pause_button)
        controls_layout.addWidget(self.resume_button)
        controls_layout.addWidget(self.retry_failed_button)
        controls_layout.addWidget(self.stop_button)
        
        return controls_frame
    
    def create_table_panel(self):
        """Create the left panel with data table"""
        table_frame = QFrame()
        table_frame.setObjectName("tableFrame")
        table_layout = QVBoxLayout(table_frame)
        table_layout.setContentsMargins(0, 0, 0, 0)
        
        # Statistics cards
        stats_row = QHBoxLayout()
        
        # Total card
        self.stat_total, self.stat_total_value = self.create_stat_card("📊 Total", "0", "#3b82f6")
        stats_row.addWidget(self.stat_total)
        
        # Sent card
        self.stat_sent, self.stat_sent_value = self.create_stat_card("✅ Sent", "0", "#10b981")
        stats_row.addWidget(self.stat_sent)
        
        # Failed card
        self.stat_failed, self.stat_failed_value = self.create_stat_card("❌ Failed", "0", "#ef4444")
        stats_row.addWidget(self.stat_failed)
        
        # Pending card
        self.stat_pending, self.stat_pending_value = self.create_stat_card("⏳ Pending", "0", "#f59e0b")
        stats_row.addWidget(self.stat_pending)
        
        table_layout.addLayout(stats_row)
        
        # Table header with counter
        header_row = QHBoxLayout()
        table_header = QLabel("📋 Recipient Data")
        table_header.setObjectName("sectionTitle")
        
        self.recipient_counter = QLabel("📊 0 recipients loaded")
        self.recipient_counter.setObjectName("recipientCounter")
        
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
    
    def create_stat_card(self, title, value, color):
        """Create a statistics card"""
        card = QFrame()
        card.setObjectName("statCard")
        card.setStyleSheet(f"""
            QFrame#statCard {{
                background: white;
                border: 1px solid #e2e8f0;
                border-radius: 12px;
                padding: 16px;
            }}
        """)
        
        layout = QVBoxLayout(card)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(6)
        
        title_label = QLabel(title)
        title_label.setStyleSheet(f"""
            QLabel {{
                color: #64748b;
                font-size: 9pt;
                font-weight: 600;
                letter-spacing: 0.2px;
            }}
        """)
        
        value_label = QLabel(value)
        value_label.setStyleSheet(f"""
            QLabel {{
                color: {color};
                font-size: 20pt;
                font-weight: 700;
                letter-spacing: -0.5px;
            }}
        """)
        value_label.setAlignment(Qt.AlignCenter)
        
        layout.addWidget(title_label)
        layout.addWidget(value_label)
        
        return card, value_label
    
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
        self.subject_input.setText("NOTICE TO SUBMIT LACKING EMPLOYMENT REQUIREMENTS - {{fullname}} - {{account}}")
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
        self.pause_button.setEnabled(False)
        self.resume_button.setEnabled(False)
        
        # Check if there are failed emails to retry
        has_failed = False
        if has_data:
            for idx in range(len(self.df)):
                status = str(self.df.iloc[idx, 5]).lower()
                if status == "failed":
                    has_failed = True
                    break
        self.retry_failed_button.setEnabled(has_failed)
        
        # Update statistics cards
        if has_data:
            self.update_statistics()
        else:
            self.stat_total_value.setText("0")
            self.stat_sent_value.setText("0")
            self.stat_failed_value.setText("0")
            self.stat_pending_value.setText("0")
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
            headers = ["Account", "Full Name", "Email", "CC", "Attachment Path"]
            for col, header in enumerate(headers, start=1):
                ws.cell(row=1, column=col, value=header).font = Font(bold=True)
            ws.cell(row=2, column=1).value = "ACC-001"
            ws.cell(row=2, column=2).value = "Dela Cruz, Juan"
            ws.cell(row=2, column=3).value = "juan@email.com"
            ws.cell(row=2, column=4).value = ""
            ws.cell(row=2, column=5).value = "C:\\Path\\To\\Attachment.pdf"
            for col_letter, width in zip(["A", "B", "C", "D", "E"], [15, 25, 30, 30, 40]):
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
                df.columns[0]: "Account",
                df.columns[1]: "Full Name",
                df.columns[2]: "Email",
                df.columns[3]: "CC",
                df.columns[4]: "Attachment Path"
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
            df["Status"] = EmailState.PENDING
            self.df = df[["Account", "Full Name", "Email", "CC", "Attachment Path", "Status"]]
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
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)  # Account
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)  # Full Name
        header.setSectionResizeMode(2, QHeaderView.Stretch)  # Email
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)  # CC
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)  # Attachment Path
        header.setSectionResizeMode(5, QHeaderView.Fixed)  # Status - fixed width
        header.setDefaultSectionSize(100)  # Default width for stretch columns
        
        # Set specific width for Status column
        header.resizeSection(5, 100)  # Status column width
        
        for i in range(len(self.df)):
            for j in range(len(self.df.columns)):
                item = QTableWidgetItem(str(self.df.iloc[i, j]))
                
                # Color code status
                if j == 5:  # Status column
                    status = str(self.df.iloc[i, j]).lower()
                    if status == "confirmed" or status == "sent":
                        item.setBackground(QColor("#d4edda"))
                    elif status == "failed":
                        item.setBackground(QColor("#f8d7da"))
                    elif status == "unknown":
                        item.setBackground(QColor("#fff3cd"))
                    elif status == "pending":
                        item.setBackground(QColor("#e2e8f0"))
                    elif status == "sending":
                        item.setBackground(QColor("#cce5ff"))
                
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
            send_delay = self.settings.get("send_delay", 0)
            importance = self.settings.get("importance", 2)
            request_read_receipt = self.settings.get("request_read_receipt", True)
        except Exception:
            spacing_px = 12
            max_retries = 3
            send_delay = 0
            importance = 2
            request_read_receipt = True

        # Save current settings
        self.settings.set("paragraph_spacing", spacing_px)

        # Create worker with enhanced safety features
        self.worker = EmailWorker(
            self.df, subject, body, spacing_px, max_retries,
            send_delay=send_delay,
            importance=importance,
            request_read_receipt=request_read_receipt,
            campaign_name=f"Campaign {datetime.now().strftime('%Y-%m-%d %H:%M')}",
            db_manager=self.db_manager
        )
        
        self.worker.progress_updated.connect(self.progress_bar.setValue)
        self.worker.log_updated.connect(self.log)
        self.worker.status_updated.connect(self.update_status)
        self.worker.validation_complete.connect(self.on_validation_complete)
        self.worker.campaign_created.connect(self.on_campaign_created)
        self.worker.finished_sending.connect(self.finish_message)
        self.worker.start()
        
        # Update UI state
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.pause_button.setEnabled(True)
        self.resume_button.setEnabled(False)
        self.stats_timer.start(1000)  # Start statistics update timer
        self.log("ℹ️ Email sending process started.")

    def stop_sending(self):
        if self.worker:
            self.worker.stop()
            self.stats_timer.stop()  # Stop statistics update timer
            self.log("⛔ Sending stopped by user.")
            self.update_ui_state()
    
    def pause_sending(self):
        if self.worker:
            self.worker.pause()
            self.log("⏸️ Sending paused.")
            self.pause_button.setEnabled(False)
            self.resume_button.setEnabled(True)
    
    def resume_sending(self):
        if self.worker:
            self.worker.resume()
            self.log("▶️ Sending resumed.")
            self.pause_button.setEnabled(True)
            self.resume_button.setEnabled(False)
    
    def retry_failed_emails(self):
        """Retry only failed emails"""
        if self.df is None:
            QMessageBox.warning(self, "⚠️ Warning", "Please load an Excel file first.")
            return
        
        # Find failed emails
        failed_indices = []
        for idx in range(len(self.df)):
            status = str(self.df.iloc[idx, 5]).lower()
            if status == "failed":
                failed_indices.append(idx)
        
        if not failed_indices:
            QMessageBox.information(self, "No Failed Emails", "There are no failed emails to retry.")
            return
        
        reply = QMessageBox.question(self, "Confirm Retry", 
                                   f"Retry {len(failed_indices)} failed emails?",
                                   QMessageBox.Yes | QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            # Reset failed status to pending
            for idx in failed_indices:
                self.df.iloc[idx, 5] = EmailState.PENDING
                self.table.setItem(idx, 5, QTableWidgetItem(EmailState.PENDING))
            
            self.log(f"🔄 {len(failed_indices)} failed emails reset to pending. Click Start Sending to retry.")
            self.update_ui_state()
    
    def test_send(self):
        """Send a test email to verify configuration"""
        if self.df is None or len(self.df) == 0:
            QMessageBox.warning(self, "⚠️ Warning", "Please load an Excel file first.")
            return
        
        from PySide6.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, QDialogButtonBox
        
        dialog = QDialog(self)
        dialog.setWindowTitle("🧪 Test Send")
        dialog.setMinimumSize(400, 200)
        dialog.setStyleSheet(self.modern_styles())
        
        layout = QVBoxLayout(dialog)
        
        # Recipient selection
        layout.addWidget(QLabel("Select recipient for test email:"))
        recipient_combo = QComboBox()
        
        for idx in range(len(self.df)):
            name = str(self.df.iloc[idx, 1])
            email = str(self.df.iloc[idx, 2])
            recipient_combo.addItem(f"{name} ({email})", idx)
        
        layout.addWidget(recipient_combo)
        
        # Dialog buttons
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        
        if dialog.exec() == QDialog.Accepted:
            idx = recipient_combo.currentData()
            row = self.df.iloc[idx]
            
            subject = self.subject_input.text()
            body = self.email_editor.toHtml()
            spacing_px = int(self.spacing_select.currentData())
            
            # Create a single-item worker for test send
            test_df = self.df.iloc[[idx]].copy()
            
            self.worker = EmailWorker(
                test_df, subject, body, spacing_px, 0,
                send_delay=0,
                importance=self.settings.get("importance", 2),
                request_read_receipt=self.settings.get("request_read_receipt", True),
                campaign_name=f"Test Send {datetime.now().strftime('%Y-%m-%d %H:%M')}",
                db_manager=None,  # Don't save test sends to database
                is_test_send=True
            )
            
            self.worker.progress_updated.connect(self.progress_bar.setValue)
            self.worker.log_updated.connect(self.log)
            self.worker.status_updated.connect(self.update_status)
            self.worker.validation_complete.connect(self.on_validation_complete)
            self.worker.finished_sending.connect(lambda: self.finish_test_send(dialog))
            self.worker.start()
            
            self.log("🧪 Test send initiated...")
    
    def finish_test_send(self, dialog):
        """Handle test send completion"""
        if dialog:
            QMessageBox.information(dialog, "Test Send Complete", "Test email has been sent.")
        self.update_ui_state()
    
    def view_campaign_history(self):
        """View campaign history from database"""
        try:
            from PySide6.QtWidgets import QDialog, QVBoxLayout, QTextEdit, QHBoxLayout, QLabel, QPushButton, QListWidget, QListWidgetItem
            
            # Get all campaigns from database
            conn = self.db_manager.get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                SELECT id, name, total_recipients, confirmed_count, failed_count, 
                       unknown_count, cancelled_count, status, started_at, completed_at
                FROM campaigns 
                ORDER BY started_at DESC
                LIMIT 50
            ''')
            campaigns = cursor.fetchall()
            conn.close()
            
            if not campaigns:
                QMessageBox.information(self, "No History", "No campaign history found.")
                return
            
            # Create history dialog
            history_dialog = QDialog(self)
            history_dialog.setWindowTitle("📜 Campaign History")
            history_dialog.setMinimumSize(800, 600)
            history_dialog.setStyleSheet(self.modern_styles())
            
            layout = QVBoxLayout(history_dialog)
            
            # Campaign list
            layout.addWidget(QLabel("Select a campaign to view details:"))
            campaign_list = QListWidget()
            
            for campaign in campaigns:
                (campaign_id, name, total, confirmed, failed, unknown, cancelled, status, started_at, completed_at) = campaign
                pending = total - confirmed - failed - unknown - cancelled
                
                item_text = f"{name} ({campaign_id})\n"
                item_text += f"  Status: {status} | Total: {total} | Sent: {confirmed} | Failed: {failed} | Unknown: {unknown} | Pending: {pending}"
                item_text += f"\n  Started: {started_at}"
                if completed_at:
                    item_text += f" | Completed: {completed_at}"
                
                item = QListWidgetItem(item_text)
                item.setData(Qt.UserRole, campaign_id)
                campaign_list.addItem(item)
            
            layout.addWidget(campaign_list)
            
            # Buttons
            button_layout = QHBoxLayout()
            
            view_details_button = QPushButton("👁️ View Details")
            view_details_button.setObjectName("primaryButton")
            view_details_button.clicked.connect(lambda: self.view_campaign_details(history_dialog, campaign_list))
            
            delete_button = QPushButton("🗑️ Delete Campaign")
            delete_button.setObjectName("dangerButton")
            delete_button.clicked.connect(lambda: self.delete_campaign(history_dialog, campaign_list))
            
            close_button = QPushButton("❌ Close")
            close_button.setObjectName("secondaryButton")
            close_button.clicked.connect(history_dialog.close)
            
            button_layout.addWidget(view_details_button)
            button_layout.addWidget(delete_button)
            button_layout.addStretch()
            button_layout.addWidget(close_button)
            
            layout.addLayout(button_layout)
            
            history_dialog.exec()
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load campaign history: {str(e)}")
    
    def view_campaign_details(self, parent_dialog, campaign_list):
        """View detailed information about a selected campaign"""
        selected_items = campaign_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(parent_dialog, "No Selection", "Please select a campaign to view.")
            return
        
        campaign_id = selected_items[0].data(Qt.UserRole)
        
        try:
            recipients = self.db_manager.get_campaign_recipients(campaign_id)
            if recipients:
                from PySide6.QtWidgets import QDialog, QVBoxLayout, QTextEdit
                
                details_dialog = QDialog(parent_dialog)
                details_dialog.setWindowTitle(f"Campaign Details: {campaign_id}")
                details_dialog.setMinimumSize(900, 700)
                details_dialog.setStyleSheet(self.modern_styles())
                
                layout = QVBoxLayout(details_dialog)
                
                details_text = QTextEdit()
                details_text.setReadOnly(True)
                
                # Build details text
                details = f"Campaign ID: {campaign_id}\n\n"
                details += "RECIPIENTS:\n"
                details += "=" * 80 + "\n"
                
                for rec in recipients:
                    (rec_id, account, full_name, email, cc, attachment_path, status, 
                     attempt_number, last_error, row_index) = rec
                    details += f"Row {row_index + 2}: {full_name} ({email})\n"
                    details += f"  Status: {status} | Attempts: {attempt_number}\n"
                    if last_error:
                        details += f"  Last Error: {last_error}\n"
                    details += "\n"
                
                details_text.setPlainText(details)
                layout.addWidget(details_text)
                
                close_button = QPushButton("Close")
                close_button.clicked.connect(details_dialog.close)
                layout.addWidget(close_button)
                
                details_dialog.exec()
            else:
                QMessageBox.warning(parent_dialog, "Error", "No recipients found for this campaign.")
        except Exception as e:
            QMessageBox.critical(parent_dialog, "Error", f"Failed to load campaign details: {str(e)}")
    
    def delete_campaign(self, parent_dialog, campaign_list):
        """Delete a selected campaign"""
        selected_items = campaign_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(parent_dialog, "No Selection", "Please select a campaign to delete.")
            return
        
        campaign_id = selected_items[0].data(Qt.UserRole)
        
        reply = QMessageBox.question(parent_dialog, "Confirm Delete", 
                                   f"Are you sure you want to delete campaign {campaign_id}?\n\nThis action cannot be undone.",
                                   QMessageBox.Yes | QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            try:
                conn = self.db_manager.get_connection()
                cursor = conn.cursor()
                
                # Delete related records first
                cursor.execute('DELETE FROM send_logs WHERE campaign_id = ?', (campaign_id,))
                cursor.execute('DELETE FROM send_attempts WHERE campaign_id = ?', (campaign_id,))
                cursor.execute('DELETE FROM recipients WHERE campaign_id = ?', (campaign_id,))
                cursor.execute('DELETE FROM campaigns WHERE id = ?', (campaign_id,))
                
                conn.commit()
                conn.close()
                
                # Refresh the list
                parent_dialog.close()
                self.view_campaign_history()  # Reopen with updated data
                
                QMessageBox.information(self, "Success", f"Campaign {campaign_id} has been deleted.")
            except Exception as e:
                QMessageBox.critical(parent_dialog, "Error", f"Failed to delete campaign: {str(e)}")

    def on_campaign_created(self, campaign_id):
        """Handle campaign creation"""
        self.current_campaign_id = campaign_id
        self.log(f"📋 Campaign ID: {campaign_id}")
    
    def on_validation_complete(self, is_valid, validation_results):
        """Handle validation completion"""
        if not is_valid and not (self.worker and self.worker.is_test_send):
            # Show validation error dialog (skip for test sends)
            error_msg = "Excel Validation Failed\n\n"
            
            if validation_results['invalid_emails']:
                error_msg += f"Invalid Emails ({len(validation_results['invalid_emails'])}):\n"
                for item in validation_results['invalid_emails'][:5]:
                    error_msg += f"  Row {item['row']}: {item['email']} - {item['error']}\n"
                if len(validation_results['invalid_emails']) > 5:
                    error_msg += f"  ... and {len(validation_results['invalid_emails']) - 5} more\n"
                error_msg += "\n"
            
            if validation_results['missing_attachments']:
                error_msg += f"Missing Attachments ({len(validation_results['missing_attachments'])}):\n"
                for item in validation_results['missing_attachments'][:5]:
                    error_msg += f"  Row {item['row']}: {item['path']}\n"
                if len(validation_results['missing_attachments']) > 5:
                    error_msg += f"  ... and {len(validation_results['missing_attachments']) - 5} more\n"
                error_msg += "\n"
            
            if validation_results['empty_required_fields']:
                error_msg += f"Empty Required Fields ({len(validation_results['empty_required_fields'])}):\n"
                for item in validation_results['empty_required_fields'][:5]:
                    error_msg += f"  Row {item['row']}: {item['field']}\n"
                if len(validation_results['empty_required_fields']) > 5:
                    error_msg += f"  ... and {len(validation_results['empty_required_fields']) - 5} more\n"
                error_msg += "\n"
            
            if validation_results['invalid_cc']:
                error_msg += f"Invalid CC Addresses ({len(validation_results['invalid_cc'])}):\n"
                for item in validation_results['invalid_cc'][:5]:
                    error_msg += f"  Row {item['row']}: {item['cc']} - {item['error']}\n"
                if len(validation_results['invalid_cc']) > 5:
                    error_msg += f"  ... and {len(validation_results['invalid_cc']) - 5} more\n"
                error_msg += "\n"
            
            if validation_results['duplicate_recipients']:
                error_msg += f"Duplicate Recipients ({len(validation_results['duplicate_recipients'])}):\n"
                for item in validation_results['duplicate_recipients'][:5]:
                    error_msg += f"  Row {item['row']}: {item['email']} (duplicate of row {item['duplicate_row']})\n"
                if len(validation_results['duplicate_recipients']) > 5:
                    error_msg += f"  ... and {len(validation_results['duplicate_recipients']) - 5} more\n"
                error_msg += "\n"
            
            error_msg += "Please fix these issues before sending."
            QMessageBox.critical(self, "Validation Failed", error_msg)

    # =================================================
    # UPDATE STATUS & LOGS
    # =================================================
    def update_status(self, row, status):
        self.table.setItem(row, 5, QTableWidgetItem(status))
        # Update the dataframe as well
        if self.df is not None and row < len(self.df):
            self.df.iloc[row, 5] = status
    
    def update_statistics(self):
        """Update statistics cards"""
        if self.df is not None and len(self.df) > 0:
            recipient_count = len(self.df)
            confirmed = sum(1 for idx in range(len(self.df)) if str(self.df.iloc[idx, 5]).lower() == "confirmed")
            failed = sum(1 for idx in range(len(self.df)) if str(self.df.iloc[idx, 5]).lower() == "failed")
            pending = sum(1 for idx in range(len(self.df)) if str(self.df.iloc[idx, 5]).lower() == "pending")
            unknown = sum(1 for idx in range(len(self.df)) if str(self.df.iloc[idx, 5]).lower() == "unknown")
            
            # Update stat cards
            self.stat_total_value.setText(str(recipient_count))
            self.stat_sent_value.setText(str(confirmed))
            self.stat_failed_value.setText(str(failed))
            self.stat_pending_value.setText(str(pending))
            
            # Update recipient counter with status breakdown
            counter_text = f"📊 {recipient_count} recipient{'s' if recipient_count != 1 else ''} loaded"
            if confirmed > 0 or failed > 0 or pending > 0 or unknown > 0:
                counter_text += f" | ✅ {confirmed} | ❌ {failed} | ⏳ {pending} | ❓ {unknown}"
            self.recipient_counter.setText(counter_text)

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
        self.stats_timer.stop()  # Stop statistics update timer
        QMessageBox.information(self, "✅ Complete", "Email sending process has completed.")
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
        sample_combo.addItem("Sample: ACC-001 - Dela Cruz, Juan", {"account": "ACC-001", "name": "Dela Cruz, Juan"})
        sample_combo.addItem("Sample: ACC-002 - Smith, John", {"account": "ACC-002", "name": "Smith, John"})
        sample_combo.addItem("Sample: ACC-003 - Garcia, Maria", {"account": "ACC-003", "name": "Garcia, Maria"})
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
            sample_data = sample_combo.currentData()
            sample_account = sample_data["account"]
            sample_name = sample_data["name"]
            subject = self.subject_input.text().replace("{{account}}", sample_account).replace("{{Account}}", sample_account).replace("{{fullname}}", sample_name)
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
            self.subject_input.setText("NOTICE TO SUBMIT LACKING EMPLOYMENT REQUIREMENTS - {{fullname}} - {{account}}")
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
    # SETTINGS DIALOG
    # =================================================
    def show_settings_dialog(self):
        """Show settings dialog"""
        from PySide6.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QLabel, QSpinBox, QCheckBox, QComboBox, QDialogButtonBox, QTabWidget, QWidget, QFormLayout, QGroupBox
        
        dialog = QDialog(self)
        dialog.setWindowTitle("⚙️ Settings")
        dialog.setMinimumSize(600, 450)
        dialog.setStyleSheet(self.modern_styles())
        
        layout = QVBoxLayout(dialog)
        
        # Tab widget for settings categories
        tab_widget = QTabWidget()
        
        # General Settings Tab
        general_tab = QWidget()
        general_layout = QVBoxLayout(general_tab)
        
        general_group = QGroupBox("General Settings")
        general_group.setObjectName("inputGroup")
        general_form = QFormLayout()
        
        # Max retries
        max_retries_spin = QSpinBox()
        max_retries_spin.setRange(0, 10)
        max_retries_spin.setValue(self.settings.get("max_retries", 3))
        general_form.addRow("Max Retries:", max_retries_spin)
        
        # Send delay
        send_delay_spin = QSpinBox()
        send_delay_spin.setRange(0, 60)
        send_delay_spin.setValue(self.settings.get("send_delay", 0))
        send_delay_spin.setSuffix(" seconds")
        general_form.addRow("Send Delay:", send_delay_spin)
        
        # Auto-save interval
        auto_save_spin = QSpinBox()
        auto_save_spin.setRange(1, 60)
        auto_save_spin.setValue(self.settings.get("auto_save_interval", 5))
        auto_save_spin.setSuffix(" minutes")
        general_form.addRow("Auto-save Interval:", auto_save_spin)
        
        general_group.setLayout(general_form)
        general_layout.addWidget(general_group)
        general_layout.addStretch()
        
        # Email Settings Tab
        email_tab = QWidget()
        email_layout = QVBoxLayout(email_tab)
        
        email_group = QGroupBox("Email Settings")
        email_group.setObjectName("inputGroup")
        email_form = QFormLayout()
        
        # Email importance
        importance_combo = QComboBox()
        importance_combo.addItem("Low", 0)
        importance_combo.addItem("Normal", 1)
        importance_combo.addItem("High", 2)
        current_importance = self.settings.get("importance", 2)
        for i in range(importance_combo.count()):
            if importance_combo.itemData(i) == current_importance:
                importance_combo.setCurrentIndex(i)
                break
        email_form.addRow("Email Importance:", importance_combo)
        
        # Read receipt
        read_receipt_check = QCheckBox()
        read_receipt_check.setChecked(self.settings.get("request_read_receipt", True))
        email_form.addRow("Request Read Receipt:", read_receipt_check)
        
        # Retry failed emails
        retry_check = QCheckBox()
        retry_check.setChecked(self.settings.get("retry_failed_emails", True))
        email_form.addRow("Retry Failed Emails:", retry_check)
        
        email_group.setLayout(email_form)
        email_layout.addWidget(email_group)
        email_layout.addStretch()
        
        # UI Settings Tab
        ui_tab = QWidget()
        ui_layout = QVBoxLayout(ui_tab)
        
        ui_group = QGroupBox("Interface Settings")
        ui_group.setObjectName("inputGroup")
        ui_form = QFormLayout()
        
        # Paragraph spacing
        spacing_combo = QComboBox()
        spacing_combo.addItem("Tight", 8)
        spacing_combo.addItem("Normal", 12)
        spacing_combo.addItem("Relaxed", 16)
        current_spacing = self.settings.get("paragraph_spacing", 12)
        for i in range(spacing_combo.count()):
            if spacing_combo.itemData(i) == current_spacing:
                spacing_combo.setCurrentIndex(i)
                break
        ui_form.addRow("Default Paragraph Spacing:", spacing_combo)
        
        ui_group.setLayout(ui_form)
        ui_layout.addWidget(ui_group)
        ui_layout.addStretch()
        
        # Add tabs
        tab_widget.addTab(general_tab, "📋 General")
        tab_widget.addTab(email_tab, "✉️ Email")
        tab_widget.addTab(ui_tab, "🎨 Interface")
        
        layout.addWidget(tab_widget)
        
        # Dialog buttons
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        
        # Show dialog
        if dialog.exec() == QDialog.Accepted:
            # Save settings
            self.settings.set("max_retries", max_retries_spin.value())
            self.settings.set("send_delay", send_delay_spin.value())
            self.settings.set("auto_save_interval", auto_save_spin.value())
            self.settings.set("importance", importance_combo.currentData())
            self.settings.set("request_read_receipt", read_receipt_check.isChecked())
            self.settings.set("retry_failed_emails", retry_check.isChecked())
            self.settings.set("paragraph_spacing", spacing_combo.currentData())
            
            # Apply paragraph spacing change
            self.apply_editor_paragraph_spacing(spacing_combo.currentData())
            
            QMessageBox.information(self, "Settings Saved", "Settings have been saved successfully.")

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
            font-size: 28pt;
            font-weight: 700;
            color: #0f172a;
            margin: 5px 0;
            letter-spacing: -0.5px;
        }
        
        #headerSubtitle {
            font-size: 11pt;
            color: #64748b;
            margin-bottom: 15px;
            font-weight: 400;
            letter-spacing: 0.2px;
        }
        
        /* =============================================
           CONTROLS FRAME
           ============================================= */
        #controlsFrame {
            background: white;
            border: 1px solid #e2e8f0;
            border-radius: 16px;
            margin: 8px 0;
        }
        
        /* =============================================
           BUTTON STYLES
           ============================================= */
        QPushButton {
            background: #2563eb;
            color: white;
            border: none;
            border-radius: 10px;
            padding: 10px 20px;
            font-weight: 600;
            font-size: 10pt;
            min-width: 100px;
            letter-spacing: 0.3px;
        }
        
        QPushButton:hover {
            background: #1d4ed8;
        }
        
        QPushButton:pressed {
            background: #1e40af;
        }
        
        QPushButton:disabled {
            background: #e2e8f0;
            color: #94a3b8;
        }
        
        #primaryButton {
            background: #4f46e5;
        }
        
        #primaryButton:hover {
            background: #4338ca;
        }
        
        #successButton {
            background: #059669;
        }
        
        #successButton:hover {
            background: #047857;
        }
        
        #dangerButton {
            background: #dc2626;
        }
        
        #dangerButton:hover {
            background: #b91c1c;
        }
        
        #secondaryButton {
            background: #475569;
        }
        
        #secondaryButton:hover {
            background: #334155;
        }
        
        /* =============================================
           SECTION TITLES
           ============================================= */
        #sectionTitle {
            font-size: 15pt;
            font-weight: 700;
            color: #0f172a;
            margin: 8px 0;
            padding: 10px 0;
            border-bottom: 2px solid #e2e8f0;
            letter-spacing: -0.3px;
        }
        
        /* =============================================
           FRAME STYLES
           ============================================= */
        #tableFrame, #emailFrame, #statusFrame {
            background: white;
            border: 1px solid #e2e8f0;
            border-radius: 16px;
            padding: 20px;
        }
        
        /* =============================================
           INPUT STYLES
           ============================================= */
        QLineEdit, QTextEdit, QComboBox {
            background: white;
            border: 1px solid #cbd5e1;
            border-radius: 8px;
            padding: 10px 12px;
            font-size: 11pt;
            selection-background-color: #dbeafe;
        }
        
        QLineEdit:focus, QTextEdit:focus, QComboBox:focus {
            border: 2px solid #3b82f6;
            outline: none;
        }
        
        QLineEdit:hover, QTextEdit:hover, QComboBox:hover {
            border: 1px solid #94a3b8;
        }
        
        /* =============================================
           TABLE STYLES
           ============================================= */
        QTableWidget {
            background: white;
            border: 1px solid #e2e8f0;
            border-radius: 12px;
            gridline-color: #f1f5f9;
            selection-background-color: #dbeafe;
            selection-color: #0f172a;
            alternate-background-color: #f8fafc;
        }
        
        QTableWidget::item {
            padding: 12px 10px;
            border-bottom: 1px solid #f1f5f9;
        }
        
        QTableWidget::item:selected {
            background: #dbeafe;
            color: #0f172a;
        }
        
        QHeaderView::section {
            background: #f8fafc;
            color: #475569;
            padding: 14px 12px;
            border: none;
            border-right: 1px solid #e2e8f0;
            border-bottom: 2px solid #e2e8f0;
            font-weight: 600;
            font-size: 10pt;
            letter-spacing: 0.3px;
        }
        
        /* =============================================
           PROGRESS BAR
           ============================================= */
        QProgressBar {
            background: #f1f5f9;
            border: 1px solid #e2e8f0;
            border-radius: 10px;
            text-align: center;
            color: #475569;
            font-weight: 600;
            height: 28px;
            font-size: 10pt;
        }
        
        QProgressBar::chunk {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                stop:0 #3b82f6, stop:1 #2563eb);
            border-radius: 8px;
            margin: 3px;
        }
        
        /* =============================================
           GROUP BOX
           ============================================= */
        QGroupBox {
            background: #f8fafc;
            border: 1px solid #e2e8f0;
            border-radius: 12px;
            margin-top: 12px;
            padding-top: 24px;
            font-weight: 600;
            color: #475569;
            font-size: 11pt;
        }
        
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 12px;
            padding: 0 8px;
            letter-spacing: 0.2px;
        }
        
        /* =============================================
           COMBO BOX
           ============================================= */
        QComboBox::drop-down {
            border: none;
            width: 24px;
        }
        
        QComboBox::down-arrow {
            image: none;
            border-left: 4px solid transparent;
            border-right: 4px solid transparent;
            border-top: 5px solid #64748b;
        }
        
        QComboBox QAbstractItemView {
            background: white;
            border: 1px solid #e2e8f0;
            selection-background-color: #dbeafe;
            selection-color: #0f172a;
            color: #1e293b;
            border-radius: 8px;
        }
        
        /* =============================================
           SCROLLBAR
           ============================================= */
        QScrollBar:vertical {
            background: #f1f5f9;
            border: none;
            border-radius: 8px;
            width: 10px;
        }
        
        QScrollBar::handle:vertical {
            background: #cbd5e1;
            border-radius: 8px;
            min-height: 24px;
        }
        
        QScrollBar::handle:vertical:hover {
            background: #94a3b8;
        }
        
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
            height: 0px;
        }
        
        /* =============================================
           SPECIAL ELEMENTS
           ============================================= */
        #recipientCounter {
            background: #eff6ff;
            color: #1d4ed8;
            border: 1px solid #bfdbfe;
            border-radius: 24px;
            padding: 8px 16px;
            font-weight: 600;
            font-size: 10pt;
            letter-spacing: 0.2px;
        }
        
        #templateCombo, #spacingSelect {
            background: #f8fafc;
            border: 1px solid #cbd5e1;
        }
        
        #subjectInput {
            background: #eff6ff;
            border: 1px solid #bfdbfe;
            font-weight: 500;
        }
        
        #emailEditor {
            background: white;
            border: 1px solid #e2e8f0;
            font-family: 'Segoe UI', system-ui, sans-serif;
            line-height: 1.6;
        }
        
        #logBox {
            background: #0f172a;
            color: #e2e8f0;
            border: 1px solid #1e293b;
            font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
            font-size: 9pt;
            line-height: 1.4;
        }
        
        /* =============================================
           STAT CARDS
           ============================================= */
        #statCard {
            background: white;
            border: 1px solid #e2e8f0;
            border-radius: 12px;
            padding: 16px;
        }
        
        /* =============================================
           TAB WIDGET STYLING
           ============================================= */
        QTabWidget::pane {
            border: 1px solid #e2e8f0;
            background: white;
            border-radius: 12px;
            top: -1px;
        }
        
        QTabBar::tab {
            background: #f8fafc;
            border: 1px solid #e2e8f0;
            padding: 14px 28px;
            margin-right: 4px;
            border-top-left-radius: 12px;
            border-top-right-radius: 12px;
            font-weight: 600;
            font-size: 10pt;
            color: #64748b;
            letter-spacing: 0.2px;
        }
        
        QTabBar::tab:selected {
            background: white;
            color: #0f172a;
            border-bottom: 2px solid #2563eb;
        }
        
        QTabBar::tab:hover:!selected {
            background: #f1f5f9;
            color: #1e40af;
        }
        
        #mainTabWidget QTabBar::tab {
            min-width: 160px;
        }
        
        /* =============================================
           DIALOG STYLES
           ============================================= */
        QDialog {
            background: #f8fafc;
        }
        
        QMessageBox {
            background: white;
        }
        
        QMessageBox QPushButton {
            min-width: 80px;
            padding: 8px 16px;
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