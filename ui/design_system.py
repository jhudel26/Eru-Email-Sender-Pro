"""Centralized design system for Eru Email Sender Pro."""

from dataclasses import dataclass


@dataclass(frozen=True)
class ThemeColors:
    bg_app: str
    bg_sidebar: str
    bg_surface: str
    bg_elevated: str
    bg_input: str
    bg_hover: str
    bg_active: str
    border: str
    border_subtle: str
    text_primary: str
    text_secondary: str
    text_muted: str
    accent: str
    accent_hover: str
    accent_pressed: str
    success: str
    success_bg: str
    warning: str
    warning_bg: str
    error: str
    error_bg: str
    info: str
    info_bg: str
    log_bg: str
    log_text: str
    scrollbar: str
    scrollbar_hover: str


LIGHT = ThemeColors(
    bg_app="#F3F2F1",
    bg_sidebar="#FAFAFA",
    bg_surface="#FFFFFF",
    bg_elevated="#FFFFFF",
    bg_input="#FFFFFF",
    bg_hover="#F3F2F1",
    bg_active="#EDEBE9",
    border="#EDEBE9",
    border_subtle="#F3F2F1",
    text_primary="#201F1E",
    text_secondary="#605E5C",
    text_muted="#A19F9D",
    accent="#0078D4",
    accent_hover="#106EBE",
    accent_pressed="#005A9E",
    success="#107C10",
    success_bg="#DFF6DD",
    warning="#CA5010",
    warning_bg="#FFF4CE",
    error="#D13438",
    error_bg="#FDE7E9",
    info="#0078D4",
    info_bg="#DEECF9",
    log_bg="#1B1A19",
    log_text="#F3F2F1",
    scrollbar="#C8C6C4",
    scrollbar_hover="#A19F9D",
)

DARK = ThemeColors(
    bg_app="#1B1A19",
    bg_sidebar="#252423",
    bg_surface="#2D2C2C",
    bg_elevated="#323130",
    bg_input="#252423",
    bg_hover="#3B3A39",
    bg_active="#484644",
    border="#484644",
    border_subtle="#3B3A39",
    text_primary="#FAFAFA",
    text_secondary="#C8C6C4",
    text_muted="#797775",
    accent="#2899F5",
    accent_hover="#47AAFF",
    accent_pressed="#0078D4",
    success="#6BB700",
    success_bg="#1E3A1E",
    warning="#FFB900",
    warning_bg="#3D3500",
    error="#F1707B",
    error_bg="#442726",
    info="#2899F5",
    info_bg="#1A3449",
    log_bg="#141414",
    log_text="#E1DFDD",
    scrollbar="#605E5C",
    scrollbar_hover="#797775",
)

SPACING = {
    "xs": 4,
    "sm": 8,
    "md": 12,
    "lg": 16,
    "xl": 24,
    "xxl": 32,
}

FONT_FAMILY = "'Segoe UI', 'Segoe UI Variable', system-ui, sans-serif"


def build_stylesheet(theme: str = "light") -> str:
    c = DARK if theme == "dark" else LIGHT
    return f"""
    /* ===== GLOBAL ===== */
    QWidget {{
        background-color: {c.bg_app};
        color: {c.text_primary};
        font-family: {FONT_FAMILY};
        font-size: 10pt;
    }}

    QWidget#sidebar {{
        background-color: {c.bg_sidebar};
        border-right: 1px solid {c.border};
    }}

    QWidget#mainContent {{
        background-color: {c.bg_app};
    }}

    QWidget#contentArea {{
        background-color: {c.bg_app};
    }}

    QWidget#pageContainer {{
        background-color: transparent;
    }}

    /* ===== TYPOGRAPHY ===== */
    QLabel#pageTitle {{
        font-size: 20pt;
        font-weight: 600;
        color: {c.text_primary};
        background: transparent;
    }}

    QLabel#pageSubtitle {{
        font-size: 10pt;
        color: {c.text_secondary};
        background: transparent;
    }}

    QLabel#sectionTitle {{
        font-size: 11pt;
        font-weight: 600;
        color: {c.text_primary};
        background: transparent;
        padding: 0;
        border: none;
    }}

    QLabel#sectionDesc {{
        font-size: 9pt;
        color: {c.text_secondary};
        background: transparent;
    }}

    QLabel#brandTitle {{
        font-size: 11pt;
        font-weight: 600;
        color: {c.text_primary};
        background: transparent;
    }}

    QLabel#brandSubtitle {{
        font-size: 8pt;
        color: {c.text_muted};
        background: transparent;
    }}

    QLabel#statLabel {{
        font-size: 9pt;
        font-weight: 500;
        color: {c.text_secondary};
        background: transparent;
    }}

    QLabel#statValue {{
        font-size: 22pt;
        font-weight: 600;
        background: transparent;
    }}

    QLabel#statusDot {{
        font-size: 8pt;
        background: transparent;
    }}

    QLabel#statusText {{
        font-size: 9pt;
        color: {c.text_secondary};
        background: transparent;
    }}

    /* ===== SIDEBAR NAV ===== */
    QPushButton#navButton {{
        background: transparent;
        color: {c.text_secondary};
        border: none;
        border-radius: 4px;
        padding: 10px 12px;
        text-align: left;
        font-size: 10pt;
        font-weight: 400;
        min-height: 36px;
    }}

    QPushButton#navButton:hover {{
        background: {c.bg_hover};
        color: {c.text_primary};
    }}

    QPushButton#navButton:checked {{
        background: {c.bg_active};
        color: {c.text_primary};
        font-weight: 600;
        border-left: 3px solid {c.accent};
        padding-left: 9px;
    }}

    QPushButton#navButtonSettings {{
        background: transparent;
        color: {c.text_secondary};
        border: none;
        border-radius: 4px;
        padding: 10px 12px;
        text-align: left;
        font-size: 10pt;
        min-height: 36px;
    }}

    QPushButton#navButtonSettings:hover {{
        background: {c.bg_hover};
        color: {c.text_primary};
    }}

    /* ===== BUTTONS ===== */
    QPushButton {{
        background: {c.accent};
        color: #FFFFFF;
        border: none;
        border-radius: 4px;
        padding: 8px 16px;
        font-weight: 600;
        font-size: 10pt;
        min-height: 32px;
    }}

    QPushButton:hover {{
        background: {c.accent_hover};
    }}

    QPushButton:pressed {{
        background: {c.accent_pressed};
    }}

    QPushButton:disabled {{
        background: {c.border_subtle};
        color: {c.text_muted};
    }}

    QPushButton#primaryButton {{
        background: {c.accent};
        color: #FFFFFF;
    }}

    QPushButton#primaryButton:hover {{
        background: {c.accent_hover};
    }}

    QPushButton#successButton {{
        background: {c.success};
        color: #FFFFFF;
    }}

    QPushButton#successButton:hover {{
        background: #0B6A0B;
    }}

    QPushButton#dangerButton {{
        background: {c.error};
        color: #FFFFFF;
    }}

    QPushButton#dangerButton:hover {{
        background: #A4262C;
    }}

    QPushButton#secondaryButton {{
        background: transparent;
        color: {c.text_primary};
        border: 1px solid {c.border};
    }}

    QPushButton#secondaryButton:hover {{
        background: {c.bg_hover};
        border-color: {c.text_muted};
    }}

    QPushButton#ghostButton {{
        background: transparent;
        color: {c.accent};
        border: none;
        padding: 6px 12px;
        font-weight: 500;
    }}

    QPushButton#ghostButton:hover {{
        background: {c.info_bg};
    }}

    QPushButton#toolbarButton {{
        background: transparent;
        color: {c.text_primary};
        border: 1px solid transparent;
        border-radius: 4px;
        padding: 6px 10px;
        min-height: 28px;
        min-width: 28px;
        font-weight: 600;
        font-size: 9pt;
    }}

    QPushButton#toolbarButton:hover {{
        background: {c.bg_hover};
        border-color: {c.border};
    }}

    QPushButton#toolbarButton:checked {{
        background: {c.bg_active};
        border-color: {c.border};
    }}

    /* ===== INPUTS ===== */
    QLineEdit, QTextEdit, QPlainTextEdit, QSpinBox {{
        background: {c.bg_input};
        color: {c.text_primary};
        border: 1px solid {c.border};
        border-radius: 4px;
        padding: 8px 12px;
        font-size: 10pt;
        selection-background-color: {c.info_bg};
        selection-color: {c.text_primary};
    }}

    QLineEdit:focus, QTextEdit:focus, QPlainTextEdit:focus, QSpinBox:focus {{
        border: 1px solid {c.accent};
    }}

    QLineEdit:hover, QTextEdit:hover, QPlainTextEdit:hover {{
        border-color: {c.text_muted};
    }}

    QLineEdit:disabled {{
        background: {c.bg_hover};
        color: {c.text_muted};
    }}

    QComboBox {{
        background: {c.bg_input};
        color: {c.text_primary};
        border: 1px solid {c.border};
        border-radius: 4px;
        padding: 8px 12px;
        min-height: 20px;
        font-size: 10pt;
    }}

    QComboBox:hover {{
        border-color: {c.text_muted};
    }}

    QComboBox:focus {{
        border: 1px solid {c.accent};
    }}

    QComboBox::drop-down {{
        border: none;
        width: 24px;
    }}

    QComboBox::down-arrow {{
        image: none;
        border-left: 4px solid transparent;
        border-right: 4px solid transparent;
        border-top: 5px solid {c.text_secondary};
    }}

    QComboBox QAbstractItemView {{
        background: {c.bg_surface};
        color: {c.text_primary};
        border: 1px solid {c.border};
        selection-background-color: {c.bg_active};
        outline: none;
    }}

    /* ===== CARDS & SURFACES ===== */
    QFrame#card, QFrame#statCard {{
        background: {c.bg_surface};
        border: 1px solid {c.border};
        border-radius: 6px;
    }}

    QFrame#surfacePanel {{
        background: {c.bg_surface};
        border: 1px solid {c.border};
        border-radius: 6px;
    }}

    QFrame#toolbarFrame {{
        background: {c.bg_surface};
        border: 1px solid {c.border};
        border-bottom: none;
        border-top-left-radius: 6px;
        border-top-right-radius: 6px;
    }}

    QFrame#editorFrame {{
        background: {c.bg_surface};
        border: 1px solid {c.border};
        border-top: none;
        border-bottom-left-radius: 6px;
        border-bottom-right-radius: 6px;
    }}

    /* ===== TABLES ===== */
    QTableWidget {{
        background: {c.bg_surface};
        alternate-background-color: {c.bg_app};
        color: {c.text_primary};
        border: 1px solid {c.border};
        border-radius: 6px;
        gridline-color: {c.border_subtle};
        selection-background-color: {c.info_bg};
        selection-color: {c.text_primary};
    }}

    QTableWidget::item {{
        padding: 8px 12px;
        border-bottom: 1px solid {c.border_subtle};
    }}

    QTableWidget::item:hover {{
        background: {c.bg_hover};
    }}

    QHeaderView::section {{
        background: {c.bg_app};
        color: {c.text_secondary};
        padding: 10px 12px;
        border: none;
        border-bottom: 1px solid {c.border};
        border-right: 1px solid {c.border_subtle};
        font-weight: 600;
        font-size: 9pt;
    }}

    /* ===== PROGRESS ===== */
    QProgressBar {{
        background: {c.bg_hover};
        border: none;
        border-radius: 4px;
        text-align: center;
        color: {c.text_secondary};
        font-weight: 600;
        font-size: 9pt;
        min-height: 8px;
        max-height: 8px;
    }}

    QProgressBar::chunk {{
        background: {c.accent};
        border-radius: 4px;
    }}

    QProgressBar#progressBarLarge {{
        min-height: 24px;
        max-height: 24px;
        font-size: 9pt;
    }}

    QProgressBar#progressBarLarge::chunk {{
        border-radius: 4px;
    }}

    /* ===== GROUP BOX ===== */
    QGroupBox {{
        background: {c.bg_surface};
        border: 1px solid {c.border};
        border-radius: 6px;
        margin-top: 16px;
        padding: 16px 12px 12px 12px;
        font-weight: 600;
        color: {c.text_secondary};
    }}

    QGroupBox::title {{
        subcontrol-origin: margin;
        left: 12px;
        padding: 0 6px;
        color: {c.text_primary};
    }}

    /* ===== SCROLLBARS ===== */
    QScrollBar:vertical {{
        background: transparent;
        width: 8px;
        margin: 4px 2px;
    }}

    QScrollBar::handle:vertical {{
        background: {c.scrollbar};
        border-radius: 4px;
        min-height: 24px;
    }}

    QScrollBar::handle:vertical:hover {{
        background: {c.scrollbar_hover};
    }}

    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
        height: 0;
    }}

    QScrollBar:horizontal {{
        background: transparent;
        height: 8px;
        margin: 2px 4px;
    }}

    QScrollBar::handle:horizontal {{
        background: {c.scrollbar};
        border-radius: 4px;
        min-width: 24px;
    }}

    /* ===== LOG BOX ===== */
    QTextEdit#logBox {{
        background: {c.log_bg};
        color: {c.log_text};
        border: 1px solid {c.border};
        border-radius: 6px;
        font-family: 'Cascadia Code', 'Consolas', monospace;
        font-size: 9pt;
        padding: 8px;
    }}

    /* ===== EMAIL EDITOR ===== */
    QTextEdit#emailEditor {{
        background: {c.bg_surface};
        border: none;
        padding: 16px;
        font-family: {FONT_FAMILY};
        font-size: 11pt;
    }}

    /* ===== LIST WIDGET ===== */
    QListWidget {{
        background: {c.bg_surface};
        border: 1px solid {c.border};
        border-radius: 6px;
        outline: none;
    }}

    QListWidget::item {{
        padding: 12px 16px;
        border-bottom: 1px solid {c.border_subtle};
        color: {c.text_primary};
    }}

    QListWidget::item:hover {{
        background: {c.bg_hover};
    }}

    QListWidget::item:selected {{
        background: {c.info_bg};
        color: {c.text_primary};
    }}

    /* ===== TOOLBAR ===== */
    QToolBar {{
        background: transparent;
        border: none;
        spacing: 4px;
        padding: 4px;
    }}

    QToolBar QToolButton {{
        background: transparent;
        border: 1px solid transparent;
        border-radius: 4px;
        padding: 6px 10px;
        color: {c.text_primary};
        font-size: 9pt;
    }}

    QToolBar QToolButton:hover {{
        background: {c.bg_hover};
        border-color: {c.border};
    }}

    /* ===== SPLITTER ===== */
    QSplitter::handle {{
        background: {c.border};
        width: 1px;
    }}

    /* ===== CHECKBOX ===== */
    QCheckBox {{
        color: {c.text_primary};
        spacing: 8px;
    }}

    QCheckBox::indicator {{
        width: 18px;
        height: 18px;
        border: 1px solid {c.border};
        border-radius: 3px;
        background: {c.bg_input};
    }}

    QCheckBox::indicator:checked {{
        background: {c.accent};
        border-color: {c.accent};
    }}

    /* ===== DIALOGS ===== */
    QDialog {{
        background: {c.bg_app};
    }}

    /* ===== TOAST ===== */
    QFrame#toast {{
        background: {c.bg_elevated};
        border: 1px solid {c.border};
        border-radius: 6px;
        border-left: 4px solid {c.accent};
    }}

    QFrame#toastSuccess {{
        border-left-color: {c.success};
    }}

    QFrame#toastWarning {{
        border-left-color: {c.warning};
    }}

    QFrame#toastError {{
        border-left-color: {c.error};
    }}

    QLabel#toastMessage {{
        color: {c.text_primary};
        font-size: 10pt;
        background: transparent;
    }}

    /* ===== SEARCH ===== */
    QLineEdit#searchInput {{
        background: {c.bg_input};
        border: 1px solid {c.border};
        border-radius: 4px;
        padding: 8px 12px 8px 32px;
        min-height: 20px;
    }}

    /* ===== FORM LABELS ===== */
    QLabel#fieldLabel {{
        font-size: 9pt;
        font-weight: 600;
        color: {c.text_secondary};
        background: transparent;
        padding-bottom: 4px;
    }}

    /* ===== OUTLOOK STATUS BAR ===== */
    QFrame#outlookStatus {{
        background: {c.bg_hover};
        border-radius: 4px;
        padding: 4px;
    }}
    """
