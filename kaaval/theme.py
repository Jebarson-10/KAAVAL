"""KAAVAL — Premium light theme with smooth animations, hover effects, and production-grade polish."""

LIGHT_STYLESHEET = """
/* ── Global ── */
* { outline: none; }
QWidget {
    background-color: #f8fafc;
    color: #1e293b;
    font-family: "Segoe UI", "Nirmala UI", sans-serif;
    font-size: 13px;
}

QMainWindow { background-color: #f1f5f9; }

/* ── Tab Widget ── */
QTabWidget::pane {
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    background: #ffffff;
    top: -1px;
}
QTabBar::tab {
    background: #f1f5f9;
    color: #64748b;
    border: 1px solid #e2e8f0;
    border-bottom: none;
    padding: 9px 22px;
    margin-right: 2px;
    border-top-left-radius: 8px;
    border-top-right-radius: 8px;
    font-weight: 600;
    font-size: 13px;
}
QTabBar::tab:selected {
    background: #ffffff;
    color: #1e40af;
    border-bottom: 3px solid #3b82f6;
    font-weight: 700;
}
QTabBar::tab:hover:!selected {
    background: #e8edf5;
    color: #1e40af;
}

/* ── Table ── */
QTableWidget {
    background-color: #ffffff;
    gridline-color: #e8edf3;
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    color: #1e293b;
    selection-background-color: #dbeafe;
    selection-color: #0f172a;
}
QTableWidget QLineEdit {
    background-color: #ffffff;
    color: #1e293b;
    border: 1.5px solid #3b82f6;
    border-radius: 0;
    padding: 2px 4px;
    selection-background-color: #3b82f6;
    selection-color: #ffffff;
}
QHeaderView::section {
    background: qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #f8fafc, stop:1 #f1f5f9);
    color: #475569;
    border: 1px solid #e2e8f0;
    padding: 7px 8px;
    font-weight: 700;
    font-size: 12px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}
QHeaderView::section:hover {
    background: #e8edf5;
    color: #1e40af;
}
QTableCornerButton::section {
    background: #f8fafc;
    border: 1px solid #e2e8f0;
}

/* ── Input Fields ── */
QLineEdit, QTextEdit, QPlainTextEdit {
    background-color: #ffffff;
    border: 1.5px solid #cbd5e1;
    border-radius: 8px;
    padding: 7px 12px;
    color: #1e293b;
    selection-background-color: #3b82f6;
    selection-color: white;
}
QLineEdit:focus, QTextEdit:focus {
    border: 2px solid #3b82f6;
    background-color: #ffffff;
}
QLineEdit:hover, QTextEdit:hover {
    border: 1.5px solid #93c5fd;
}

/* ── Combo Box ── */
QComboBox {
    background-color: #ffffff;
    border: 1.5px solid #cbd5e1;
    border-radius: 8px;
    padding: 7px 28px 7px 12px;
    color: #1e293b;
    min-height: 22px;
    min-width: 100px;
}
QComboBox:hover { border: 1.5px solid #93c5fd; }
QComboBox:focus { border: 2px solid #3b82f6; }
QComboBox::drop-down {
    border: none;
    width: 28px;
    subcontrol-position: right center;
}
QComboBox QAbstractItemView {
    background-color: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    selection-background-color: #dbeafe;
    selection-color: #1e40af;
    padding: 4px;
    outline: none;
}
QComboBox QAbstractItemView::item {
    padding: 6px 12px;
    min-height: 24px;
}

/* ── Buttons (base — overridden per-button) ── */
QPushButton {
    background-color: #f1f5f9;
    color: #334155;
    border: 1.5px solid #cbd5e1;
    border-radius: 8px;
    padding: 8px 16px;
    font-weight: 600;
    min-height: 20px;
}
QPushButton:hover {
    background-color: #e2e8f0;
    border: 1.5px solid #93c5fd;
    color: #1e40af;
}
QPushButton:pressed {
    background-color: #dbeafe;
    padding-top: 9px; padding-bottom: 7px;
}
QPushButton:disabled {
    background-color: #f8fafc;
    color: #94a3b8;
    border: 1px solid #e2e8f0;
}

/* ── Group Box ── */
QGroupBox {
    border: 1.5px solid #e2e8f0;
    border-radius: 10px;
    margin-top: 14px;
    padding-top: 18px;
    background-color: #ffffff;
    font-weight: 700;
    color: #1e40af;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    padding: 2px 14px;
    color: #1e40af;
    font-size: 13px;
}

/* ── Check Box ── */
QCheckBox { spacing: 8px; color: #334155; }
QCheckBox::indicator {
    width: 18px; height: 18px; border-radius: 4px;
    border: 2px solid #94a3b8; background-color: #ffffff;
}
QCheckBox::indicator:checked {
    background-color: #3b82f6; border: 2px solid #3b82f6;
}
QCheckBox::indicator:hover { border: 2px solid #3b82f6; }

/* ── List Widget ── */
QListWidget {
    background-color: #ffffff;
    border: 1.5px solid #e2e8f0;
    border-radius: 8px;
    color: #1e293b;
    outline: none;
}
QListWidget::item {
    padding: 7px 12px;
    border-bottom: 1px solid #f1f5f9;
    border-radius: 4px;
}
QListWidget::item:selected {
    background-color: #dbeafe;
    color: #1e40af;
}
QListWidget::item:hover {
    background-color: #f0f7ff;
}

/* ── Splitter ── */
QSplitter::handle {
    background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
        stop:0 transparent, stop:0.4 #cbd5e1, stop:0.5 #94a3b8, stop:0.6 #cbd5e1, stop:1 transparent);
    border-radius: 3px;
}
QSplitter::handle:vertical { height: 6px; }
QSplitter::handle:horizontal { width: 6px; }
QSplitter::handle:hover {
    background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
        stop:0 transparent, stop:0.3 #93c5fd, stop:0.5 #3b82f6, stop:0.7 #93c5fd, stop:1 transparent);
}

/* ── Scroll Bars ── */
QScrollBar:vertical {
    background: transparent; width: 10px; border: none; margin: 2px;
}
QScrollBar::handle:vertical {
    background: #cbd5e1; border-radius: 5px; min-height: 30px;
}
QScrollBar::handle:vertical:hover { background: #3b82f6; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar:horizontal {
    background: transparent; height: 10px; border: none; margin: 2px;
}
QScrollBar::handle:horizontal {
    background: #cbd5e1; border-radius: 5px; min-width: 30px;
}
QScrollBar::handle:horizontal:hover { background: #3b82f6; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0; }

/* ── Progress Bar ── */
QProgressBar {
    background-color: #e2e8f0;
    border: 1px solid #cbd5e1;
    border-radius: 8px;
    text-align: center;
    color: #1e293b;
    font-weight: 600;
    min-height: 24px;
}
QProgressBar::chunk {
    background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
        stop:0 #3b82f6, stop:0.5 #2563eb, stop:1 #1d4ed8);
    border-radius: 7px;
}

/* ── Text Browser (Chat) ── */
QTextBrowser {
    background-color: #ffffff;
    border: 1.5px solid #e2e8f0;
    border-radius: 10px;
    color: #1e293b;
    padding: 10px;
}

/* ── Labels ── */
QLabel { color: #334155; background: transparent; }

/* ── Menu ── */
QMenu {
    background-color: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 6px;
}
QMenu::item {
    padding: 8px 28px;
    border-radius: 6px;
}
QMenu::item:selected {
    background-color: #dbeafe;
    color: #1e40af;
}
QMenu::separator { height: 1px; background: #e2e8f0; margin: 4px 10px; }

/* ── Status Bar ── */
QStatusBar {
    background: qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #f8fafc, stop:1 #f1f5f9);
    border-top: 1px solid #e2e8f0;
    color: #64748b;
    font-size: 12px;
    padding: 2px 10px;
}

/* ── Tool Tips ── */
QToolTip {
    background-color: #1e293b;
    color: #f8fafc;
    border: none;
    border-radius: 6px;
    padding: 6px 10px;
    font-size: 12px;
}

/* ── Splash / Welcome ── */
QDialog { background-color: #ffffff; border-radius: 12px; }

/* ── Spin Box ── */
QSpinBox {
    background-color: #ffffff;
    border: 1.5px solid #cbd5e1;
    border-radius: 8px;
    padding: 5px 10px;
    color: #1e293b;
    min-height: 22px;
}
QSpinBox:focus { border: 2px solid #3b82f6; }
QSpinBox:hover { border: 1.5px solid #93c5fd; }
"""


# ── Button tone styling ──
BUTTON_TONES = {
    "primary":   {"bg": "qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #2563eb, stop:1 #3b82f6)", "fg": "#ffffff", "hover": "#1d4ed8", "hover_border": "#60a5fa"},
    "success":   {"bg": "qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #059669, stop:1 #10b981)", "fg": "#ffffff", "hover": "#047857", "hover_border": "#34d399"},
    "danger":    {"bg": "qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #dc2626, stop:1 #ef4444)", "fg": "#ffffff", "hover": "#b91c1c", "hover_border": "#f87171"},
    "export":    {"bg": "qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #7c3aed, stop:1 #8b5cf6)", "fg": "#ffffff", "hover": "#6d28d9", "hover_border": "#a78bfa"},
    "format":    {"bg": "#ffffff", "fg": "#0f766e", "hover": "#f0fdfa", "hover_border": "#14b8a6"},
    "align":     {"bg": "#ffffff", "fg": "#2563eb", "hover": "#eff6ff", "hover_border": "#60a5fa"},
    "merge":     {"bg": "#ffffff", "fg": "#c2410c", "hover": "#fff7ed", "hover_border": "#fb923c"},
    "neutral":   {"bg": "#ffffff", "fg": "#334155", "hover": "#f1f5f9", "hover_border": "#3b82f6"},
    "structure": {"bg": "#ffffff", "fg": "#059669", "hover": "#ecfdf5", "hover_border": "#34d399"},
    "training":  {"bg": "qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #d97706, stop:1 #f59e0b)", "fg": "#ffffff", "hover": "#b45309", "hover_border": "#fbbf24"},
}


def style_tool_button(button, tone: str) -> None:
    """Apply premium themed styling to toolbar-style buttons with proper sizing."""
    from PySide6.QtCore import Qt
    colors = BUTTON_TONES.get(tone, BUTTON_TONES["neutral"])
    button.setMinimumHeight(30)
    button.setCursor(Qt.CursorShape.PointingHandCursor)
    button.setStyleSheet(
        "QPushButton {"
        f"  background: {colors['bg']};"
        f"  color: {colors['fg']};"
        "  border: 1.5px solid rgba(0,0,0,0.08);"
        "  border-radius: 8px;"
        "  padding: 6px 12px;"
        "  font-weight: 600;"
        "  font-size: 12px;"
        "}"
        "QPushButton:hover {"
        f"  background: {colors['hover']};"
        f"  border: 1.5px solid {colors['hover_border']};"
        f"  color: {colors['fg']};"
        "}"
        "QPushButton:pressed {"
        "  padding-top: 7px; padding-bottom: 5px;"
        "}"
        "QPushButton:disabled {"
        "  background: #f1f5f9;"
        "  color: #94a3b8;"
        "  border: 1px solid #e2e8f0;"
        "}"
    )


def style_sidebar_button(button, tone: str) -> None:
    """Style buttons used in side panels with proper min-width and text handling."""
    from PySide6.QtCore import Qt
    colors = BUTTON_TONES.get(tone, BUTTON_TONES["neutral"])
    button.setMinimumHeight(34)
    button.setMinimumWidth(90)
    button.setCursor(Qt.CursorShape.PointingHandCursor)
    button.setStyleSheet(
        "QPushButton {"
        f"  background: {colors['bg']};"
        f"  color: {colors['fg']};"
        "  border: 1.5px solid rgba(0,0,0,0.08);"
        "  border-radius: 8px;"
        "  padding: 7px 14px;"
        "  font-weight: 600;"
        "  font-size: 12px;"
        "}"
        "QPushButton:hover {"
        f"  background: {colors['hover']};"
        f"  border: 1.5px solid {colors['hover_border']};"
        f"  color: {colors['fg']};"
        "}"
        "QPushButton:pressed {"
        "  padding-top: 8px; padding-bottom: 6px;"
        "}"
        "QPushButton:disabled {"
        "  background: #f1f5f9;"
        "  color: #94a3b8;"
        "  border: 1px solid #e2e8f0;"
        "}"
    )
CHAT_CSS = """
<style>
body { font-family: 'Segoe UI', 'Nirmala UI', sans-serif; background: #ffffff; color: #1e293b; }
.user-msg {
    background: linear-gradient(135deg, #2563eb, #3b82f6);
    color: white; border-radius: 18px 18px 4px 18px;
    padding: 11px 16px; margin: 6px 0 6px 80px;
    max-width: 75%; float: right; clear: both;
    box-shadow: 0 2px 8px rgba(37,99,235,0.15);
    font-size: 13px; line-height: 1.5;
}
.ai-msg {
    background: #f1f5f9; color: #1e293b;
    border-radius: 18px 18px 18px 4px;
    padding: 11px 16px; margin: 6px 80px 6px 0;
    max-width: 75%; float: left; clear: both;
    border: 1px solid #e2e8f0;
    font-size: 13px; line-height: 1.5;
}
.sources {
    font-size: 11px; color: #64748b; margin: 2px 0 10px 0; clear: both;
    padding: 6px 10px; background: #f8fafc; border-radius: 8px;
    border: 1px solid #e2e8f0;
}
.clear { clear: both; height: 1px; }
</style>
"""


# ── Sidebar panel styles ──
SIDEBAR_HEADER_STYLE = (
    "font-size: 15px; font-weight: 700; color: #1e40af; padding: 8px 12px;"
    "background: qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #ffffff, stop:1 #f8fafc);"
    "border: 1px solid #e2e8f0; border-radius: 8px;"
)

PETITION_STATUS_STYLE = "color: #1e40af; font-weight: 600; font-size: 13px; padding: 6px 4px;"
PETITION_STATUS_DONE_STYLE = "color: #059669; font-weight: 600; font-size: 13px; padding: 6px 4px;"

PETITION_STATS_STYLE = (
    "background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #eef6ff, stop:1 #f0f7ff);"
    "color: #0f172a; border: 1px solid #bfdbfe; border-radius: 8px; "
    "padding: 8px 12px; font-weight: 600; font-size: 12px;"
)

PETITION_LOG_STYLE = (
    "QTextEdit { background: #0f172a; color: #dbeafe; border: 1px solid #1e3a5f; "
    "border-radius: 10px; padding: 10px; "
    "font-family: 'Cascadia Mono', 'Consolas', 'Courier New', monospace; font-size: 12px; }"
)

SECTION_LABEL_STYLE = "color: #64748b; font-size: 11px; font-weight: 700; letter-spacing: 0.5px; padding: 2px 0;"
