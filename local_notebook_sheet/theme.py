"""Premium dark theme stylesheet for the application."""

DARK_STYLESHEET = """
/* ── Global ── */
QWidget {
    background-color: #0f1117;
    color: #e2e8f0;
    font-family: "Segoe UI", "Inter", system-ui, sans-serif;
    font-size: 13px;
}

QMainWindow {
    background-color: #0f1117;
}

/* ── Tab Widget ── */
QTabWidget::pane {
    border: 1px solid #1e293b;
    border-radius: 8px;
    background: #0f1117;
    top: -1px;
}
QTabBar::tab {
    background: #1a1d27;
    color: #94a3b8;
    border: 1px solid #1e293b;
    border-bottom: none;
    padding: 8px 20px;
    margin-right: 2px;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
    font-weight: 600;
}
QTabBar::tab:selected {
    background: #0f1117;
    color: #a78bfa;
    border-bottom: 2px solid #a78bfa;
}
QTabBar::tab:hover:!selected {
    background: #1e293b;
    color: #c4b5fd;
}

/* ── Table ── */
QTableWidget {
    background-color: #12141d;
    alternate-background-color: #151822;
    gridline-color: #1e293b;
    border: 1px solid #1e293b;
    border-radius: 6px;
    selection-background-color: rgba(99, 102, 241, 0.3);
    selection-color: #e2e8f0;
    color: #e2e8f0;
}
QTableWidget::item {
    padding: 4px 6px;
    border-bottom: 1px solid #1a1d27;
}
QTableWidget::item:selected {
    background-color: rgba(99, 102, 241, 0.25);
}
QHeaderView::section {
    background-color: #1a1d27;
    color: #94a3b8;
    border: 1px solid #1e293b;
    padding: 6px 8px;
    font-weight: 700;
    font-size: 12px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}
QHeaderView::section:hover {
    background-color: #1e293b;
    color: #c4b5fd;
}
QTableCornerButton::section {
    background-color: #1a1d27;
    border: 1px solid #1e293b;
}

/* ── Input Fields ── */
QLineEdit, QTextEdit, QPlainTextEdit {
    background-color: #1a1d27;
    border: 1px solid #2d3348;
    border-radius: 6px;
    padding: 6px 10px;
    color: #e2e8f0;
    selection-background-color: #6366f1;
}
QLineEdit:focus, QTextEdit:focus {
    border: 1px solid #6366f1;
    background-color: #1e2233;
}

/* ── Combo Box ── */
QComboBox {
    background-color: #1a1d27;
    border: 1px solid #2d3348;
    border-radius: 6px;
    padding: 6px 10px;
    color: #e2e8f0;
    min-height: 22px;
}
QComboBox:hover {
    border: 1px solid #6366f1;
}
QComboBox::drop-down {
    border: none;
    width: 24px;
}
QComboBox QAbstractItemView {
    background-color: #1a1d27;
    border: 1px solid #2d3348;
    selection-background-color: #6366f1;
    color: #e2e8f0;
}

/* ── Buttons ── */
QPushButton {
    background-color: #1e293b;
    color: #e2e8f0;
    border: 1px solid #334155;
    border-radius: 6px;
    padding: 7px 14px;
    font-weight: 600;
    min-height: 20px;
}
QPushButton:hover {
    background-color: #334155;
    border: 1px solid #6366f1;
}
QPushButton:pressed {
    background-color: #1e293b;
    padding-top: 8px;
    padding-bottom: 6px;
}
QPushButton:disabled {
    background-color: #0f1117;
    color: #475569;
    border: 1px solid #1e293b;
}

/* ── Group Box ── */
QGroupBox {
    border: 1px solid #1e293b;
    border-radius: 8px;
    margin-top: 12px;
    padding-top: 16px;
    background-color: #12141d;
    font-weight: 700;
    color: #a78bfa;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    padding: 2px 12px;
    color: #a78bfa;
}

/* ── Check Box ── */
QCheckBox {
    spacing: 8px;
    color: #e2e8f0;
}
QCheckBox::indicator {
    width: 18px;
    height: 18px;
    border-radius: 4px;
    border: 2px solid #475569;
    background-color: #1a1d27;
}
QCheckBox::indicator:checked {
    background-color: #6366f1;
    border: 2px solid #6366f1;
}
QCheckBox::indicator:hover {
    border: 2px solid #6366f1;
}

/* ── List Widget ── */
QListWidget {
    background-color: #12141d;
    border: 1px solid #1e293b;
    border-radius: 6px;
    color: #e2e8f0;
    outline: none;
}
QListWidget::item {
    padding: 6px 10px;
    border-bottom: 1px solid #1a1d27;
}
QListWidget::item:selected {
    background-color: rgba(99, 102, 241, 0.25);
    color: #c4b5fd;
}
QListWidget::item:hover {
    background-color: #1a1d27;
}

/* ── Splitter ── */
QSplitter::handle {
    background-color: #1e293b;
    width: 3px;
}
QSplitter::handle:hover {
    background-color: #6366f1;
}

/* ── Scroll Bars ── */
QScrollBar:vertical {
    background: #0f1117;
    width: 10px;
    border: none;
    margin: 0;
}
QScrollBar::handle:vertical {
    background: #334155;
    border-radius: 5px;
    min-height: 30px;
}
QScrollBar::handle:vertical:hover {
    background: #6366f1;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0;
}
QScrollBar:horizontal {
    background: #0f1117;
    height: 10px;
    border: none;
}
QScrollBar::handle:horizontal {
    background: #334155;
    border-radius: 5px;
    min-width: 30px;
}
QScrollBar::handle:horizontal:hover {
    background: #6366f1;
}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
    width: 0;
}

/* ── Progress Bar ── */
QProgressBar {
    background-color: #1a1d27;
    border: 1px solid #2d3348;
    border-radius: 6px;
    text-align: center;
    color: #e2e8f0;
    font-weight: 600;
    min-height: 22px;
}
QProgressBar::chunk {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #6366f1, stop:1 #a78bfa);
    border-radius: 5px;
}

/* ── Text Browser (Chat) ── */
QTextBrowser {
    background-color: #12141d;
    border: 1px solid #1e293b;
    border-radius: 8px;
    color: #e2e8f0;
    padding: 8px;
}

/* ── Labels ── */
QLabel {
    color: #cbd5e1;
    background: transparent;
}

/* ── Form Layout Labels ── */
QFormLayout QLabel {
    font-weight: 600;
    color: #94a3b8;
}

/* ── Menu ── */
QMenu {
    background-color: #1a1d27;
    border: 1px solid #2d3348;
    border-radius: 8px;
    padding: 4px;
}
QMenu::item {
    padding: 6px 24px;
    border-radius: 4px;
}
QMenu::item:selected {
    background-color: #6366f1;
    color: white;
}
QMenu::separator {
    height: 1px;
    background: #2d3348;
    margin: 4px 8px;
}

/* ── Status Bar ── */
QStatusBar {
    background-color: #0d0f15;
    border-top: 1px solid #1e293b;
    color: #64748b;
    font-size: 12px;
    padding: 2px 8px;
}

/* ── Tool Tips ── */
QToolTip {
    background-color: #1e293b;
    color: #e2e8f0;
    border: 1px solid #334155;
    border-radius: 4px;
    padding: 4px 8px;
    font-size: 12px;
}
"""


# ── Button styling with themed tones ──

BUTTON_TONES = {
    "primary": {"bg": "qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #6366f1, stop:1 #818cf8)", "fg": "#ffffff", "hover_border": "#a78bfa"},
    "success": {"bg": "qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #059669, stop:1 #10b981)", "fg": "#ffffff", "hover_border": "#34d399"},
    "danger":  {"bg": "qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #dc2626, stop:1 #ef4444)", "fg": "#ffffff", "hover_border": "#f87171"},
    "export":  {"bg": "qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #7c3aed, stop:1 #a855f7)", "fg": "#ffffff", "hover_border": "#c084fc"},
    "format":  {"bg": "#1e293b", "fg": "#a5f3fc", "hover_border": "#22d3ee"},
    "align":   {"bg": "#1e293b", "fg": "#93c5fd", "hover_border": "#60a5fa"},
    "merge":   {"bg": "#1e293b", "fg": "#fdba74", "hover_border": "#fb923c"},
    "neutral": {"bg": "#1e293b", "fg": "#e2e8f0", "hover_border": "#6366f1"},
    "structure": {"bg": "#1e293b", "fg": "#86efac", "hover_border": "#4ade80"},
}


def style_tool_button(button, tone: str) -> None:
    """Apply themed styling to a toolbar button."""
    from PySide6.QtCore import Qt
    colors = BUTTON_TONES.get(tone, BUTTON_TONES["neutral"])
    button.setMinimumHeight(30)
    button.setCursor(Qt.CursorShape.PointingHandCursor)
    button.setStyleSheet(
        "QPushButton {"
        f"  background: {colors['bg']};"
        f"  color: {colors['fg']};"
        "  border: 1px solid rgba(255,255,255,0.08);"
        "  border-radius: 6px;"
        "  padding: 6px 12px;"
        "  font-weight: 600;"
        "  font-size: 12px;"
        "}"
        "QPushButton:hover {"
        f"  border: 1px solid {colors['hover_border']};"
        "  padding: 5px 11px;"
        "}"
        "QPushButton:pressed {"
        "  padding-top: 7px; padding-bottom: 5px;"
        "}"
    )


# ── Toolbar button labels with icons ──

TOOLBAR_ACTIONS = [
    # (label, icon, tone, group)
    ("Open",       "📂", "neutral",   "file"),
    ("Save",       "💾", "success",   "file"),
    ("Save As",    "📄", "success",   "file"),
    ("Export",     "📊", "export",    "file"),
    (None, None, None, "sep"),  # separator
    ("Bold",       "𝐁",  "format",   "format"),
    ("Italic",     "𝐼",  "format",   "format"),
    ("Underline",  "U̲",  "format",   "format"),
    ("Txt Color",  "🎨", "format",   "format"),
    ("Fill",       "🪣", "format",   "format"),
    (None, None, None, "sep"),
    ("Left",       "◀", "align",    "align"),
    ("Center",     "◆", "align",    "align"),
    ("Right",      "▶", "align",    "align"),
    (None, None, None, "sep"),
    ("Merge",      "⊞",  "merge",    "struct"),
    ("Unmerge",    "⊟",  "danger",   "struct"),
    ("+ Row",      "↕",  "structure","struct"),
    ("+ Col",      "↔",  "structure","struct"),
    (None, None, None, "sep"),
    ("Undo",       "↩",  "neutral",  "edit"),
    ("Redo",       "↪",  "neutral",  "edit"),
    ("Clear",      "🗑", "danger",   "edit"),
]


# ── Chat HTML styling ──

CHAT_CSS = """
<style>
body { font-family: 'Segoe UI', sans-serif; background: #12141d; color: #e2e8f0; }
.user-msg {
    background: linear-gradient(135deg, #4f46e5, #6366f1);
    color: white; border-radius: 16px 16px 4px 16px;
    padding: 10px 14px; margin: 6px 0 6px 60px;
    max-width: 80%; float: right; clear: both;
}
.ai-msg {
    background: #1e293b; color: #e2e8f0;
    border-radius: 16px 16px 16px 4px;
    padding: 10px 14px; margin: 6px 60px 6px 0;
    max-width: 80%; float: left; clear: both;
    border: 1px solid #2d3348;
}
.sources { font-size: 11px; color: #64748b; margin: 2px 0 8px 0; clear: both; }
.clear { clear: both; }
</style>
"""
