"""Reusable UI components for Eru Email Sender Pro."""

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QFrame,
    QScrollArea, QSizePolicy, QStackedWidget, QGraphicsOpacityEffect,
)
from PySide6.QtCore import Qt, Signal, QPropertyAnimation, QEasingCurve, QTimer, QSize
from PySide6.QtGui import QFont

from ui.design_system import SPACING, LIGHT, DARK


class NavButton(QPushButton):
    def __init__(self, text: str, page_id: str, parent=None):
        super().__init__(text, parent)
        self.page_id = page_id
        self.setObjectName("navButton")
        self.setCheckable(True)
        self.setAutoExclusive(True)
        self.setCursor(Qt.PointingHandCursor)


class AppSidebar(QFrame):
    navigate = Signal(str)

    NAV_ITEMS = [
        ("Dashboard", "dashboard"),
        ("Recipients", "recipients"),
        ("Compose", "compose"),
        ("Templates", "templates"),
        ("History", "history"),
        ("Logs", "logs"),
    ]

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("sidebar")
        self.setFixedWidth(220)
        self._buttons: dict[str, NavButton] = {}

        layout = QVBoxLayout(self)
        layout.setContentsMargins(SPACING["md"], SPACING["lg"], SPACING["md"], SPACING["lg"])
        layout.setSpacing(SPACING["xs"])

        brand = QVBoxLayout()
        brand.setSpacing(2)
        title = QLabel("Eru Email Sender")
        title.setObjectName("brandTitle")
        subtitle = QLabel("Pro")
        subtitle.setObjectName("brandSubtitle")
        brand.addWidget(title)
        brand.addWidget(subtitle)
        layout.addLayout(brand)
        layout.addSpacing(SPACING["xl"])

        for label, page_id in self.NAV_ITEMS:
            btn = NavButton(label, page_id)
            btn.clicked.connect(lambda checked, pid=page_id: self.navigate.emit(pid))
            self._buttons[page_id] = btn
            layout.addWidget(btn)

        layout.addStretch()

        self.outlook_frame = QFrame()
        self.outlook_frame.setObjectName("outlookStatus")
        outlook_layout = QHBoxLayout(self.outlook_frame)
        outlook_layout.setContentsMargins(SPACING["sm"], SPACING["sm"], SPACING["sm"], SPACING["sm"])
        self.outlook_dot = QLabel("●")
        self.outlook_dot.setObjectName("statusDot")
        self.outlook_dot.setStyleSheet(f"color: {LIGHT.text_muted};")
        self.outlook_label = QLabel("Outlook")
        self.outlook_label.setObjectName("statusText")
        self.outlook_status = QLabel("Checking...")
        self.outlook_status.setObjectName("statusText")
        outlook_layout.addWidget(self.outlook_dot)
        col = QVBoxLayout()
        col.setSpacing(0)
        col.addWidget(self.outlook_label)
        col.addWidget(self.outlook_status)
        outlook_layout.addLayout(col)
        outlook_layout.addStretch()
        layout.addWidget(self.outlook_frame)

        self.settings_btn = QPushButton("Settings")
        self.settings_btn.setObjectName("navButtonSettings")
        self.settings_btn.setCursor(Qt.PointingHandCursor)
        layout.addWidget(self.settings_btn)

    def set_active(self, page_id: str):
        if page_id in self._buttons:
            self._buttons[page_id].setChecked(True)

    def set_outlook_status(self, connected: bool, detail: str = ""):
        theme = DARK  # dot colors work in both
        if connected:
            self.outlook_dot.setStyleSheet(f"color: {LIGHT.success};")
            self.outlook_status.setText(detail or "Connected")
        else:
            self.outlook_dot.setStyleSheet(f"color: {LIGHT.error};")
            self.outlook_status.setText(detail or "Disconnected")


class PageHeader(QWidget):
    def __init__(self, title: str, subtitle: str = "", parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, SPACING["lg"])

        text_col = QVBoxLayout()
        text_col.setSpacing(4)
        self.title_label = QLabel(title)
        self.title_label.setObjectName("pageTitle")
        self.subtitle_label = QLabel(subtitle)
        self.subtitle_label.setObjectName("pageSubtitle")
        self.subtitle_label.setVisible(bool(subtitle))
        text_col.addWidget(self.title_label)
        text_col.addWidget(self.subtitle_label)
        layout.addLayout(text_col)
        layout.addStretch()

        self.actions_layout = QHBoxLayout()
        self.actions_layout.setSpacing(SPACING["sm"])
        layout.addLayout(self.actions_layout)

    def add_action(self, widget):
        self.actions_layout.addWidget(widget)

    def set_title(self, title: str, subtitle: str = ""):
        self.title_label.setText(title)
        if subtitle:
            self.subtitle_label.setText(subtitle)
            self.subtitle_label.setVisible(True)
        else:
            self.subtitle_label.setVisible(False)


class StatCard(QFrame):
    def __init__(self, label: str, value: str = "0", accent: str = "#0078D4", parent=None):
        super().__init__(parent)
        self.setObjectName("statCard")
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setMinimumHeight(88)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(SPACING["lg"], SPACING["md"], SPACING["lg"], SPACING["md"])
        layout.setSpacing(SPACING["xs"])

        self.label = QLabel(label)
        self.label.setObjectName("statLabel")
        self.value = QLabel(value)
        self.value.setObjectName("statValue")
        self.value.setStyleSheet(f"color: {accent};")
        layout.addWidget(self.label)
        layout.addWidget(self.value)

    def set_value(self, val: str):
        self.value.setText(val)


class ContentScrollArea(QScrollArea):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWidgetResizable(True)
        self.setFrameShape(QFrame.NoFrame)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        self.content = QWidget()
        self.content.setObjectName("pageContainer")
        self.layout = QVBoxLayout(self.content)
        self.layout.setContentsMargins(SPACING["xl"], SPACING["lg"], SPACING["xl"], SPACING["xl"])
        self.layout.setSpacing(SPACING["lg"])
        self.setWidget(self.content)


class ToastWidget(QFrame):
    def __init__(self, message: str, toast_type: str = "info", parent=None):
        super().__init__(parent)
        obj_names = {"success": "toastSuccess", "warning": "toastWarning", "error": "toastError"}
        self.setObjectName(obj_names.get(toast_type, "toast"))
        self.setMinimumWidth(320)
        self.setMaximumWidth(420)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(SPACING["md"], SPACING["md"], SPACING["md"], SPACING["md"])
        lbl = QLabel(message)
        lbl.setObjectName("toastMessage")
        lbl.setWordWrap(True)
        layout.addWidget(lbl)


class ToastManager:
    def __init__(self, parent: QWidget):
        self._parent = parent
        self._toasts: list[ToastWidget] = []

    def show(self, message: str, toast_type: str = "info", duration_ms: int = 3500):
        toast = ToastWidget(message, toast_type, self._parent)
        toast.setParent(self._parent)
        toast.raise_()

        margin = SPACING["xl"]
        toast.adjustSize()
        pw = self._parent.width()
        ph = self._parent.height()
        tw = toast.width()
        x = pw - tw - margin
        y = margin + len(self._toasts) * (toast.height() + SPACING["sm"])
        toast.move(x, y)
        toast.show()

        effect = QGraphicsOpacityEffect(toast)
        toast.setGraphicsEffect(effect)
        anim = QPropertyAnimation(effect, b"opacity")
        anim.setDuration(200)
        anim.setStartValue(0.0)
        anim.setEndValue(1.0)
        anim.setEasingCurve(QEasingCurve.OutCubic)
        anim.start()

        self._toasts.append(toast)

        def dismiss():
            if toast in self._toasts:
                self._toasts.remove(toast)
            toast.deleteLater()
            self._reposition()

        QTimer.singleShot(duration_ms, dismiss)

    def _reposition(self):
        margin = SPACING["xl"]
        pw = self._parent.width()
        for i, toast in enumerate(self._toasts):
            x = pw - toast.width() - margin
            y = margin + i * (toast.height() + SPACING["sm"])
            toast.move(x, y)


class StatusBadge(QLabel):
    COLORS = {
        "confirmed": ("#107C10", "#DFF6DD"),
        "sent": ("#107C10", "#DFF6DD"),
        "failed": ("#D13438", "#FDE7E9"),
        "unknown": ("#CA5010", "#FFF4CE"),
        "pending": ("#605E5C", "#F3F2F1"),
        "sending": ("#0078D4", "#DEECF9"),
        "cancelled": ("#797775", "#F3F2F1"),
        "valid": ("#107C10", "#DFF6DD"),
        "invalid": ("#D13438", "#FDE7E9"),
    }

    def __init__(self, status: str, parent=None):
        super().__init__(parent)
        self.set_status(status)

    def set_status(self, status: str):
        key = str(status).lower()
        fg, bg = self.COLORS.get(key, ("#605E5C", "#F3F2F1"))
        display = str(status).capitalize() if status else "Unknown"
        self.setText(display)
        self.setStyleSheet(f"""
            QLabel {{
                background: {bg};
                color: {fg};
                border-radius: 4px;
                padding: 2px 8px;
                font-size: 8pt;
                font-weight: 600;
            }}
        """)


def make_field_label(text: str) -> QLabel:
    lbl = QLabel(text)
    lbl.setObjectName("fieldLabel")
    return lbl


def make_section(title: str, description: str = "") -> QVBoxLayout:
    section = QVBoxLayout()
    section.setSpacing(SPACING["sm"])
    t = QLabel(title)
    t.setObjectName("sectionTitle")
    section.addWidget(t)
    if description:
        d = QLabel(description)
        d.setObjectName("sectionDesc")
        d.setWordWrap(True)
        section.addWidget(d)
    return section
