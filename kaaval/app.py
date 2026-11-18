from __future__ import annotations

import json
import os
import re
import sys
import uuid
from copy import copy
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from PySide6.QtCore import QSignalBlocker, QTimer, Qt, Signal, QThread
from PySide6.QtGui import (
    QColor, QFont, QFontDatabase, QShortcut, QKeySequence, QDragEnterEvent,
    QDropEvent, QPen, QPainter, QTextCharFormat, QTextCursor, QTextDocument,
)
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QColorDialog,
    QComboBox,
    QFileDialog,
    QFontComboBox,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMenu,
    QMessageBox,
    QProgressBar,
    QProgressDialog,
    QScrollArea,
    QPushButton,
    QSplitter,
    QStackedWidget,
    QStatusBar,
    QStyle,
    QStyledItemDelegate,
    QStyleOptionViewItem,
    QTableWidget,
    QTableWidgetItem,
    QTextBrowser,
    QTextEdit,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from ai import DEFAULT_GEMINI_MODEL, GeminiClient, OllamaClient, extraction_summary, ocr_crop, tesseract_status
from document import crop_normalized, existing_supported, page_count, render_page, supported_file_filter
from formula import FormulaEngine, column_name
from templates import (
    APP_DIR, ExtractionTemplate, TemplateField, TemplateStore,
    FewShotStore, FewShotExample, FewShotTemplate,
    load_settings, save_settings,
)
from bamini import bamini_display_runs, has_tamil_unicode
from widgets import ImageSelectionCanvas
from theme import (
    LIGHT_STYLESHEET, style_tool_button, CHAT_CSS,
    style_sidebar_button, SECTION_LABEL_STYLE, SIDEBAR_HEADER_STYLE,
    PETITION_STATUS_STYLE, PETITION_STATUS_DONE_STYLE,
    PETITION_STATS_STYLE, PETITION_LOG_STYLE,
)

try:
    import pytesseract  # type: ignore
    _tess_path = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    if os.path.exists(_tess_path):
        pytesseract.pytesseract.tesseract_cmd = _tess_path
except ImportError:
    pass

NATIVE_FILE_FILTER = "Local AI Sheet (*.laisheet)"
AUTOSAVE_FILE = APP_DIR / "autosave.laisheet"
MONEY_CHECK_COLUMN = "Money Check"
CLOUD_VERSION_COLUMN = "Cloud Version"
DEFAULT_CELL_FONT_SIZE = 10

VISION_MODEL_PRESETS = [
    ("Qwen 2.5 VL (Best for Tamil OCR)", "qwen2.5-vl"),
    ("MiniCPM-V 2.6 (Fast & Accurate)", "minicpm-v"),
    ("Llama 3.2 Vision (Strong Reasoning)", "llama3.2-vision"),
    ("Gemma 3 12B (Tamil Reasoning Only)", "gemma3:12b"),
    ("Gemma 4 E4B (Balanced)", "gemma4:e4b"),
    ("Gemma 4 26B (Advanced)", "gemma4:26b"),
    ("Gemma 4 31B (Strongest Reasoning)", "gemma4:31b"),
]


class SpreadsheetDelegate(QStyledItemDelegate):
    def __init__(self, parent=None):
        super().__init__(parent)
        families = {name.casefold(): name for name in QFontDatabase.families()}
        self.bamini_family = next(
            (families[name] for name in ("bamini", "bamini tamil", "bamini-tamil01") if name in families),
            next((name for key, name in sorted(families.items()) if key.startswith("bamini")), ""),
        )

    def initStyleOption(self, option, index) -> None:
        super().initStyleOption(option, index)
        if self.bamini_family and option.font.family().casefold().startswith("bamini"):
            option.font.setFamily(self.bamini_family)

    def _paint_unicode_tamil(self, painter, option, index) -> None:
        self.initStyleOption(option, index)
        widget = option.widget
        style = widget.style() if widget else QApplication.style()
        rect = style.subElementRect(QStyle.SubElement.SE_ItemViewItemText, option, widget)
        text = option.text
        option.text = ""
        style.drawControl(QStyle.ControlElement.CE_ItemViewItem, option, painter, widget)

        document = QTextDocument()
        document.setDocumentMargin(0)
        text_option = document.defaultTextOption()
        text_option.setAlignment(option.displayAlignment & Qt.AlignmentFlag.AlignHorizontal_Mask)
        if not option.features & QStyleOptionViewItem.ViewItemFeature.WrapText:
            text_option.setWrapMode(text_option.WrapMode.NoWrap)
        document.setDefaultTextOption(text_option)
        document.setTextWidth(max(0, rect.width() - 6))
        cursor = QTextCursor(document)
        for part, tamil in bamini_display_runs(text):
            font = QFont(option.font)
            if tamil:
                font.setFamily(self.bamini_family)
            elif font.family().casefold().startswith("bamini"):
                font.setFamily("Segoe UI")
            char_format = QTextCharFormat()
            char_format.setFont(font)
            char_format.setForeground(option.palette.text())
            cursor.insertText(part, char_format)

        offset = max(0, rect.height() - document.size().height())
        if option.displayAlignment & Qt.AlignmentFlag.AlignVCenter:
            offset /= 2
        elif not option.displayAlignment & Qt.AlignmentFlag.AlignBottom:
            offset = 0
        painter.save()
        painter.setClipRect(rect, Qt.ClipOperation.IntersectClip)
        painter.translate(rect.left() + 3, rect.top() + offset)
        document.drawContents(painter)
        painter.restore()

    def paint(self, painter: QPainter, option: 'QStyleOptionViewItem', index) -> None:
        painter.save()
        
        # 1. Let Qt draw the standard background, fill color, and text
        opt = QStyleOptionViewItem(option)
        is_selected = opt.state & QStyle.StateFlag.State_Selected
        
        # Prevent default blue selection styling by masking out the selected state
        if is_selected:
            opt.state &= ~QStyle.StateFlag.State_Selected

        if self.bamini_family and has_tamil_unicode(str(index.data() or "")):
            self._paint_unicode_tamil(painter, opt, index)
        else:
            super().paint(painter, opt, index)

        # 2. Draw our custom light-gray selection highlight ON TOP using Multiply
        # This brilliantly tints the cell and any custom fill colors while keeping text 100% visible.
        if is_selected:
            painter.setCompositionMode(QPainter.CompositionMode.CompositionMode_Multiply)
            painter.fillRect(option.rect, QColor("#e2e8f0"))
            painter.setCompositionMode(QPainter.CompositionMode.CompositionMode_SourceOver)

        # 3. Draw custom borders
        border_data = index.data(Qt.ItemDataRole.UserRole)
        if border_data and isinstance(border_data, dict):
            rect = option.rect
            for side, (thickness, color_hex) in border_data.items():
                pen = QPen(QColor(color_hex), thickness)
                pen.setJoinStyle(Qt.PenJoinStyle.MiterJoin)
                painter.setPen(pen)
                
                # Offset by half thickness to stay inside the cell
                offset = thickness / 2.0
                
                if side == "top":
                    painter.drawLine(rect.left(), rect.top() + offset, rect.right(), rect.top() + offset)
                elif side == "bottom":
                    painter.drawLine(rect.left(), rect.bottom() - offset, rect.right(), rect.bottom() - offset)
                elif side == "left":
                    painter.drawLine(rect.left() + offset, rect.top(), rect.left() + offset, rect.bottom())
                elif side == "right":
                    painter.drawLine(rect.right() - offset, rect.top(), rect.right() - offset, rect.bottom())

        painter.restore()


from PySide6.QtWidgets import QDialog, QDialogButtonBox, QSpinBox

class BorderDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Border Settings")
        
        self.thickness = QSpinBox()
        self.thickness.setRange(1, 10)
        self.thickness.setValue(1)
        
        self.color_btn = QPushButton("Pick Color")
        self.color = "#000000"
        self.color_btn.clicked.connect(self.pick_color)
        self.color_btn.setStyleSheet("background-color: #000000; color: white;")
        
        self.type_combo = QComboBox()
        self.type_combo.addItems(["All Borders", "Outer Borders", "Inner Borders", "Top", "Bottom", "Left", "Right", "Clear Borders"])
        
        layout = QFormLayout(self)
        layout.addRow("Border Type:", self.type_combo)
        layout.addRow("Thickness (px):", self.thickness)
        layout.addRow("Color:", self.color_btn)
        
        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addRow(btns)
        
    def pick_color(self):
        c = QColorDialog.getColor(QColor(self.color), self)
        if c.isValid():
            self.color = c.name()
            # Set background and contrasting text color
            fg = "white" if (c.red() * 0.299 + c.green() * 0.587 + c.blue() * 0.114) < 186 else "black"
            self.color_btn.setStyleSheet(f"background-color: {self.color}; color: {fg};")


class SpreadsheetWidget(QWidget):
    data_changed = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.raw_values: dict[tuple[int, int], str] = {}
        self.merged_ranges: list[tuple[int, int, int, int]] = []
        self.current_path: Path | None = None
        self._dirty = False
        self._updating = False

        self.history: list[dict[str, object]] = []
        self.history_index: int = -1

        self.table = QTableWidget(200, 26)
        self.table.setItemDelegate(SpreadsheetDelegate(self.table))
        self.table.setAlternatingRowColors(False)
        self.table.setWordWrap(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.verticalHeader().setDefaultSectionSize(28)
        self.table.setHorizontalHeaderLabels([column_name(i) for i in range(26)])
        self.table.cellChanged.connect(self._cell_changed)
        self.table.currentCellChanged.connect(self._current_cell_changed)
        self.table.horizontalHeader().sectionResized.connect(lambda: self.table.resizeRowsToContents())
        self.table.setFont(QFont("Segoe UI", DEFAULT_CELL_FONT_SIZE))

        # Right-click context menu
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)

        QShortcut(QKeySequence("Delete"), self.table, self.clear_selected_cells)
        QShortcut(QKeySequence("Backspace"), self.table, self.clear_selected_cells)
        QShortcut(QKeySequence.StandardKey.Undo, self, self.undo)
        QShortcut(QKeySequence.StandardKey.Redo, self, self.redo)
        QShortcut(QKeySequence("Ctrl+B"), self, self.apply_bold)
        QShortcut(QKeySequence("Ctrl+I"), self, self.apply_italic)
        QShortcut(QKeySequence("Ctrl+U"), self, self.apply_underline)
        QShortcut(QKeySequence("Ctrl+S"), self, self.save_native)
        QShortcut(QKeySequence("Ctrl+Shift+S"), self, self.save_native_as)
        QShortcut(QKeySequence("Ctrl+E"), self, self.export_xlsx)
        QShortcut(QKeySequence("Ctrl+F"), self, self._toggle_search_bar)

        self.formula_input = QLineEdit()
        self.formula_input.setPlaceholderText("Enter value or formula (e.g. =SUM(A1:A10))")
        self.formula_input.returnPressed.connect(self._commit_formula_input)
        self.cell_label = QLabel("A1")
        self.cell_label.setMinimumWidth(50)
        self.cell_label.setStyleSheet("font-weight: 700; color: #1e40af; padding: 0 6px;")
        self.status_label = QLabel("✦ Ready")
        self.status_label.setStyleSheet("color: #64748b; font-size: 12px;")

        # Search bar (hidden by default)
        self.search_bar = QWidget()
        search_layout = QHBoxLayout(self.search_bar)
        search_layout.setContentsMargins(0, 0, 0, 0)
        search_layout.setSpacing(4)
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("🔍 Search cells... (Ctrl+F)")
        self.search_input.textChanged.connect(self._on_search)
        self.search_input.returnPressed.connect(self._search_next)
        btn_prev = QPushButton("◀")
        btn_prev.setFixedWidth(32)
        btn_prev.clicked.connect(self._search_prev)
        btn_next = QPushButton("▶")
        btn_next.setFixedWidth(32)
        btn_next.clicked.connect(self._search_next)
        btn_close = QPushButton("✕")
        btn_close.setFixedWidth(32)
        btn_close.clicked.connect(self._hide_search_bar)
        self.search_count_label = QLabel("")
        self.search_count_label.setStyleSheet("color: #64748b; font-size: 11px;")
        search_layout.addWidget(self.search_input, 1)
        search_layout.addWidget(self.search_count_label)
        search_layout.addWidget(btn_prev)
        search_layout.addWidget(btn_next)
        search_layout.addWidget(btn_close)
        self.search_bar.hide()
        self._search_matches: list[tuple[int, int]] = []
        self._search_index = -1

        toolbar = self._build_toolbar()
        formula_bar = QHBoxLayout()
        formula_bar.addWidget(self.cell_label)
        formula_bar.addWidget(self.formula_input, 1)
        formula_bar.addWidget(self.status_label)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(5)
        layout.addWidget(toolbar)
        layout.addLayout(formula_bar)
        layout.addWidget(self.search_bar)
        layout.addWidget(self.table)

        self.engine = FormulaEngine(self._raw_value)
        
        self.history = [self.to_state()]
        self.history_index = 0

    def _build_toolbar(self) -> QWidget:
        ribbon = QWidget()
        ribbon.setObjectName("sheetRibbon")
        ribbon.setStyleSheet("""
            QWidget#sheetRibbon { background: #eef3f7; border: 1px solid #d5dfe8; border-radius: 6px; }
            QWidget#sheetRibbon QToolButton {
                background: transparent; border: 1px solid transparent;
                border-radius: 4px; padding: 3px 6px; min-height: 22px;
            }
            QWidget#sheetRibbon QToolButton:hover { background: #dbeaf5; border-color: #bccfe0; }
            QWidget#sheetRibbon QToolButton:pressed { background: #caddeb; }
            QWidget#sheetRibbon QComboBox, QWidget#sheetRibbon QSpinBox {
                background: white; border: 1px solid #bccbd8;
                border-radius: 4px; padding: 2px 5px; min-height: 22px;
            }
            QWidget#sheetRibbon QComboBox:focus, QWidget#sheetRibbon QSpinBox:focus {
                border-color: #47719a;
            }
        """)
        layout = QVBoxLayout(ribbon)
        layout.setContentsMargins(6, 5, 6, 5)
        layout.setSpacing(3)
        commands = QHBoxLayout()
        formatting = QHBoxLayout()
        for row in (commands, formatting):
            row.setSpacing(3)
            layout.addLayout(row)

        def button(row, label, title, handler=None, *, icon=None, actions=()):
            control = QToolButton()
            control.setText(label)
            control.setToolTip(title)
            control.setAccessibleName(title)
            if icon is not None:
                control.setIcon(self.style().standardIcon(icon))
                control.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonIconOnly)
            if actions:
                menu = QMenu(control)
                for name, action in actions:
                    menu.addAction(name, action)
                control.setMenu(menu)
                control.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
            else:
                control.clicked.connect(handler)
            row.addWidget(control)
            return control

        button(commands, "", "Open workbook", self.open_native, icon=QStyle.StandardPixmap.SP_DialogOpenButton)
        button(commands, "", "Save workbook", self.save_native, icon=QStyle.StandardPixmap.SP_DialogSaveButton)
        button(commands, "", "Export to Excel (Ctrl+E)", self.export_xlsx, icon=QStyle.StandardPixmap.SP_DialogApplyButton)
        button(commands, "File", "File commands", actions=[
            ("Save As", self.save_native_as),
            ("Clear Sheet", self.clear_sheet),
        ])
        button(commands, "\u21b6", "Undo (Ctrl+Z)", self.undo)
        button(commands, "\u21b7", "Redo (Ctrl+Y)", self.redo)
        button(commands, "Find", "Find cells (Ctrl+F)", self._toggle_search_bar)
        commands.addStretch()
        button(commands, "Tools", "Open tools", actions=[
            ("Petition Extractor", lambda: self._show_tool_panel("petition")),
            ("AI Chat", lambda: self._show_tool_panel("chat")),
            ("AI Training", lambda: self._show_tool_panel("training")),
        ])
        button(commands, "\u25eb", "Toggle tools panel", self._toggle_tool_panel)

        self.font_combo = QFontComboBox()
        self.font_combo.setCurrentFont(self.table.font())
        self.font_combo.setMinimumWidth(74)
        self.font_combo.setMaximumWidth(150)
        self.font_combo.setToolTip("Font family")
        self.font_combo.setAccessibleName("Font family")
        self.font_combo.currentFontChanged.connect(self.apply_font)
        formatting.addWidget(self.font_combo, 1)

        self.font_size = QSpinBox()
        self.font_size.setRange(6, 144)
        self.font_size.setValue(DEFAULT_CELL_FONT_SIZE)
        self.font_size.setFixedWidth(56)
        self.font_size.setKeyboardTracking(False)
        self.font_size.setToolTip("Font size in points")
        self.font_size.setAccessibleName("Font size")
        self.font_size.valueChanged.connect(self.apply_font_size)
        formatting.addWidget(self.font_size)

        button(formatting, "B", "Bold (Ctrl+B)", self.apply_bold)
        button(formatting, "I", "Italic (Ctrl+I)", self.apply_italic)
        button(formatting, "U", "Underline (Ctrl+U)", self.apply_underline)
        button(formatting, "Format", "Cell formatting", actions=[
            ("Text Color", self.apply_text_color),
            ("Fill Color", self.apply_fill_color),
            ("Borders", self.apply_borders),
            ("Align Left", lambda: self.apply_alignment(Qt.AlignmentFlag.AlignLeft)),
            ("Align Center", lambda: self.apply_alignment(Qt.AlignmentFlag.AlignCenter)),
            ("Align Right", lambda: self.apply_alignment(Qt.AlignmentFlag.AlignRight)),
            ("Wrap Text", self.toggle_word_wrap),
            ("Convert to Bamini", self.encode_selected_to_bamini),
            ("Convert to Unicode", self.decode_selected_to_unicode),
        ])
        button(formatting, "Cells", "Row, column, and merge commands", actions=[
            ("Merge Cells", self.merge_selection),
            ("Unmerge Cells", self.unmerge_selection),
            ("Fit Rows", self.auto_fit_rows),
            ("Fit Columns", self.auto_fit_columns),
            ("Insert Row Above", self.insert_row_above),
            ("Insert Column Left", self.insert_column_left),
            ("Delete Row", self.delete_row),
            ("Delete Column", self.delete_column),
        ])
        formatting.addStretch()
        return ribbon

    def _toggle_tool_panel(self) -> None:
        window = self.window()
        if hasattr(window, "toggle_tool_panel"):
            window.toggle_tool_panel()


    def _raw_value(self, row: int, col: int) -> str:
        return self.raw_values.get((row, col), "")

    def _ensure_item(self, row: int, col: int) -> QTableWidgetItem:
        item = self.table.item(row, col)
        if item is None:
            item = QTableWidgetItem("")
            font = self.table.font()
            if font.pointSize() <= 0:
                font.setPointSize(DEFAULT_CELL_FONT_SIZE)
            item.setFont(font)
            self.table.setItem(row, col, item)
        return item

    def _cell_changed(self, row: int, col: int) -> None:
        if self._updating:
            return
        item = self.table.item(row, col)
        text = item.text() if item else ""
        raw = self._raw_value(row, col)
        if text == (self.engine.evaluate(raw) if raw.startswith("=") else raw):
            return
        if text:
            self.raw_values[(row, col)] = text
        else:
            self.raw_values.pop((row, col), None)
        self.recalculate()
        self.table.resizeRowToContents(row)
        self._mark_changed()

    def _current_cell_changed(self, row: int, col: int, _old_row: int, _old_col: int) -> None:
        if row < 0 or col < 0:
            return
        self.cell_label.setText(f"{column_name(col)}{row + 1}")
        self.formula_input.setText(self.raw_values.get((row, col), self._display_value(row, col)))
        self._sync_font_size()

    def _sync_font_size(self) -> None:
        item = self.table.currentItem()
        font = item.font() if item else self.table.font()
        with QSignalBlocker(self.font_size):
            size = font.pointSize()
            self.font_size.setValue(size if size > 0 else DEFAULT_CELL_FONT_SIZE)

    def _commit_formula_input(self) -> None:
        row = self.table.currentRow()
        col = self.table.currentColumn()
        if row < 0 or col < 0:
            return
        text = self.formula_input.text()
        if text:
            self.raw_values[(row, col)] = text
        else:
            self.raw_values.pop((row, col), None)
            self._updating = True
            try:
                item = self.table.item(row, col)
                if item:
                    item.setText("")
            finally:
                self._updating = False
        self.recalculate()
        self._mark_changed()

    def _mark_changed(self) -> None:
        if self._updating:
            return
        self._dirty = True
        self.status_label.setText("Unsaved changes")

        self.history = self.history[:self.history_index + 1]
        self.history.append(self.to_state())
        if len(self.history) > 50:
            self.history.pop(0)
        else:
            self.history_index += 1

        self.data_changed.emit()

    def _display_value(self, row: int, col: int) -> str:
        item = self.table.item(row, col)
        return item.text() if item else ""

    def recalculate(self) -> None:
        self._updating = True
        try:
            for (row, col), raw in list(self.raw_values.items()):
                item = self._ensure_item(row, col)
                item.setText(self.engine.evaluate(raw) if raw.startswith("=") else raw)
        finally:
            self._updating = False

    def selected_items_or_cells(self) -> list[QTableWidgetItem]:
        items: list[QTableWidgetItem] = []
        for selection in self.table.selectedRanges():
            for row in range(selection.topRow(), selection.bottomRow() + 1):
                for col in range(selection.leftColumn(), selection.rightColumn() + 1):
                    items.append(self._ensure_item(row, col))
        if not items and self.table.currentRow() >= 0 and self.table.currentColumn() >= 0:
            items.append(self._ensure_item(self.table.currentRow(), self.table.currentColumn()))
        return items

    def apply_bold(self) -> None:
        self._toggle_font_style(QFont.bold, QFont.setBold)

    def apply_italic(self) -> None:
        self._toggle_font_style(QFont.italic, QFont.setItalic)

    def apply_underline(self) -> None:
        self._toggle_font_style(QFont.underline, QFont.setUnderline)

    def _toggle_font_style(self, getter, setter) -> None:
        items = self.selected_items_or_cells()
        enabled = not all(getter(item.font()) for item in items)
        self._format_fonts(items, lambda font: setter(font, enabled))

    def apply_font(self, font: QFont) -> None:
        self._format_fonts(self.selected_items_or_cells(), lambda current: current.setFamily(font.family()))

    def apply_font_size(self, size: int) -> None:
        self._format_fonts(self.selected_items_or_cells(), lambda font: font.setPointSize(size))

    def _format_fonts(self, items, change) -> None:
        rows = set()
        self._updating = True
        try:
            for item in items:
                font = item.font()
                change(font)
                if font != item.font():
                    item.setFont(font)
                    rows.add(item.row())
            for row in rows:
                self.table.resizeRowToContents(row)
        finally:
            self._updating = False
        if rows:
            self._mark_changed()

    def encode_selected_to_bamini(self) -> None:
        from bamini import unicode_to_bamini
        self._updating = True
        try:
            for item in self.selected_items_or_cells():
                row, col = item.row(), item.column()
                old_text = self.raw_values.get((row, col), "")
                if old_text:
                    new_text = unicode_to_bamini(old_text)
                    item.setText(new_text)
                    f = item.font()
                    f.setFamily("Bamini TAMIL")
                    item.setFont(f)
                    self.raw_values[(row, col)] = new_text
        finally:
            self._updating = False
        self.recalculate()
        self._mark_changed()

    def decode_selected_to_unicode(self) -> None:
        from bamini import bamini_to_unicode
        self._updating = True
        try:
            for item in self.selected_items_or_cells():
                row, col = item.row(), item.column()
                old_text = self.raw_values.get((row, col), "")
                if old_text:
                    new_text = bamini_to_unicode(old_text)
                    item.setText(new_text)
                    f = item.font()
                    f.setFamily("Nirmala UI")
                    item.setFont(f)
                    self.raw_values[(row, col)] = new_text
        finally:
            self._updating = False
        self.recalculate()
        self._mark_changed()

    def apply_text_color(self) -> None:
        color = QColorDialog.getColor(parent=self)
        if not color.isValid():
            return
        for item in self.selected_items_or_cells():
            item.setForeground(color)
        self._mark_changed()

    def apply_fill_color(self) -> None:
        color = QColorDialog.getColor(parent=self)
        if not color.isValid():
            return
        for item in self.selected_items_or_cells():
            item.setBackground(color)
        self._mark_changed()

    def apply_alignment(self, alignment: Qt.AlignmentFlag) -> None:
        for item in self.selected_items_or_cells():
            item.setTextAlignment(alignment | Qt.AlignmentFlag.AlignVCenter)
        self._mark_changed()

    def merge_selection(self) -> None:
        ranges = self.table.selectedRanges()
        if not ranges:
            return
        selection = ranges[0]
        row_span = selection.rowCount()
        col_span = selection.columnCount()
        if row_span <= 1 and col_span <= 1:
            return
        self.table.setSpan(selection.topRow(), selection.leftColumn(), row_span, col_span)
        self.merged_ranges.append((selection.topRow(), selection.leftColumn(), row_span, col_span))
        self._mark_changed()

    def unmerge_selection(self) -> None:
        row = self.table.currentRow()
        col = self.table.currentColumn()
        if row < 0 or col < 0:
            return
        top_left = None
        for merged in self.merged_ranges:
            start_row, start_col, row_span, col_span = merged
            if start_row <= row < start_row + row_span and start_col <= col < start_col + col_span:
                top_left = merged
                break
        if top_left:
            start_row, start_col, _, _ = top_left
            self.table.setSpan(start_row, start_col, 1, 1)
            self.merged_ranges.remove(top_left)
            self._mark_changed()

    def add_row(self) -> None:
        self.table.insertRow(self.table.rowCount())
        self._mark_changed()

    def add_column(self) -> None:
        col = self.table.columnCount()
        self.table.insertColumn(col)
        self.table.setHorizontalHeaderItem(col, QTableWidgetItem(column_name(col)))
        self._mark_changed()

    def clear_selected_cells(self) -> None:
        self._updating = True
        try:
            for item in self.selected_items_or_cells():
                item.setText("")
                row, col = item.row(), item.column()
                self.raw_values.pop((row, col), None)
        finally:
            self._updating = False
        self.recalculate()
        self._mark_changed()

    # ── Right-click context menu ──

    def _show_context_menu(self, pos) -> None:
        menu = QMenu(self)
        menu.setStyleSheet(
            "QMenu { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; padding: 4px; }"
            "QMenu::item { padding: 7px 24px; border-radius: 4px; }"
            "QMenu::item:selected { background: #dbeafe; color: #1e40af; }"
            "QMenu::separator { height: 1px; background: #e2e8f0; margin: 4px 8px; }"
        )
        menu.addAction("✂ Cut", self._cut_cells)
        menu.addAction("📋 Copy", self._copy_cells)
        menu.addAction("📋 Paste", self._paste_cells)
        menu.addSeparator()
        menu.addAction("⬆ Insert Row Above", self.insert_row_above)
        menu.addAction("⬇ Insert Row Below", self.insert_row_below)
        menu.addAction("⬅ Insert Column Left", self.insert_column_left)
        menu.addAction("➡ Insert Column Right", self.insert_column_right)
        menu.addSeparator()
        menu.addAction("❌ Delete Row", self.delete_row)
        menu.addAction("❌ Delete Column", self.delete_column)
        menu.addAction("🗑 Clear Cells", self.clear_selected_cells)
        menu.addSeparator()
        menu.addAction("⊞ Merge Cells", self.merge_selection)
        menu.addAction("⊟ Unmerge Cells", self.unmerge_selection)
        menu.addSeparator()
        menu.addAction("↩ Toggle Word Wrap", self.toggle_word_wrap)
        menu.addAction("↕ Auto-Fit Rows", self.auto_fit_rows)
        menu.addAction("↔ Auto-Fit Columns", self.auto_fit_columns)
        menu.addSeparator()
        menu.addAction("𝐁 Bold", self.apply_bold)
        menu.addAction("𝐼 Italic", self.apply_italic)
        menu.addAction("U̲ Underline", self.apply_underline)
        menu.addAction("🎨 Text Color", self.apply_text_color)
        menu.addAction("🪣 Fill Color", self.apply_fill_color)
        menu.addAction("🔲 Borders...", self.apply_borders)
        menu.exec(self.table.viewport().mapToGlobal(pos))

    def apply_borders(self) -> None:
        dlg = BorderDialog(self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            b_type = dlg.type_combo.currentText()
            thickness = dlg.thickness.value()
            color = dlg.color
            
            ranges = self.table.selectedRanges()
            if not ranges:
                return
            
            for sel in ranges:
                for r in range(sel.topRow(), sel.bottomRow() + 1):
                    for c in range(sel.leftColumn(), sel.rightColumn() + 1):
                        item = self._ensure_item(r, c)
                        border_data = item.data(Qt.ItemDataRole.UserRole)
                        if not isinstance(border_data, dict):
                            border_data = {}
                        else:
                            border_data = border_data.copy()
                        
                        if b_type == "Clear Borders":
                            border_data = {}
                        else:
                            val = (thickness, color)
                            if b_type == "All Borders":
                                border_data["top"] = val
                                border_data["bottom"] = val
                                border_data["left"] = val
                                border_data["right"] = val
                            elif b_type == "Top" and r == sel.topRow():
                                border_data["top"] = val
                            elif b_type == "Bottom" and r == sel.bottomRow():
                                border_data["bottom"] = val
                            elif b_type == "Left" and c == sel.leftColumn():
                                border_data["left"] = val
                            elif b_type == "Right" and c == sel.rightColumn():
                                border_data["right"] = val
                            elif b_type == "Outer Borders":
                                if r == sel.topRow(): border_data["top"] = val
                                if r == sel.bottomRow(): border_data["bottom"] = val
                                if c == sel.leftColumn(): border_data["left"] = val
                                if c == sel.rightColumn(): border_data["right"] = val
                            elif b_type == "Inner Borders":
                                if r > sel.topRow(): border_data["top"] = val
                                if r < sel.bottomRow(): border_data["bottom"] = val
                                if c > sel.leftColumn(): border_data["left"] = val
                                if c < sel.rightColumn(): border_data["right"] = val
                        
                        item.setData(Qt.ItemDataRole.UserRole, border_data)
            
            self._mark_changed()
            self.table.viewport().update()

    def toggle_word_wrap(self) -> None:
        self.table.setWordWrap(not self.table.wordWrap())
        if self.table.wordWrap():
            self.table.resizeRowsToContents()

    def auto_fit_rows(self) -> None:
        self.table.resizeRowsToContents()

    def auto_fit_columns(self) -> None:
        self.table.resizeColumnsToContents()

    def _cut_cells(self) -> None:
        self._copy_cells()
        self.clear_selected_cells()

    def _copy_cells(self) -> None:
        ranges = self.table.selectedRanges()
        if not ranges:
            return
        sel = ranges[0]
        rows = []
        for r in range(sel.topRow(), sel.bottomRow() + 1):
            cols = []
            for c in range(sel.leftColumn(), sel.rightColumn() + 1):
                cols.append(self._display_value(r, c))
            rows.append("\t".join(cols))
        QApplication.clipboard().setText("\n".join(rows))

    def _paste_cells(self) -> None:
        text = QApplication.clipboard().text()
        if not text:
            return
        row = self.table.currentRow()
        col = self.table.currentColumn()
        if row < 0 or col < 0:
            return
        self._updating = True
        try:
            for r_offset, line in enumerate(text.split("\n")):
                for c_offset, val in enumerate(line.split("\t")):
                    target_r = row + r_offset
                    target_c = col + c_offset
                    if target_r < self.table.rowCount() and target_c < self.table.columnCount():
                        item = self._ensure_item(target_r, target_c)
                        item.setText(val)
                        if val:
                            self.raw_values[(target_r, target_c)] = val
                        else:
                            self.raw_values.pop((target_r, target_c), None)
        finally:
            self._updating = False
        self.recalculate()
        self._mark_changed()

    def insert_row_above(self) -> None:
        row = self.table.currentRow()
        if row < 0:
            row = 0
        self._shift_data_down(row)
        self.table.insertRow(row)
        self._mark_changed()

    def insert_row_below(self) -> None:
        row = self.table.currentRow()
        if row < 0:
            row = self.table.rowCount() - 1
        self._shift_data_down(row + 1)
        self.table.insertRow(row + 1)
        self._mark_changed()

    def insert_column_left(self) -> None:
        col = self.table.currentColumn()
        if col < 0:
            col = 0
        self._shift_data_right(col)
        self.table.insertColumn(col)
        self._update_column_headers()
        self._mark_changed()

    def insert_column_right(self) -> None:
        col = self.table.currentColumn()
        if col < 0:
            col = self.table.columnCount() - 1
        self._shift_data_right(col + 1)
        self.table.insertColumn(col + 1)
        self._update_column_headers()
        self._mark_changed()

    def delete_row(self) -> None:
        row = self.table.currentRow()
        if row < 0:
            return
        # Remove data for this row, shift everything up
        new_values = {}
        for (r, c), v in self.raw_values.items():
            if r == row:
                continue
            elif r > row:
                new_values[(r - 1, c)] = v
            else:
                new_values[(r, c)] = v
        self.raw_values = new_values
        self._updating = True
        self.table.removeRow(row)
        self._updating = False
        self.recalculate()
        self._mark_changed()

    def delete_column(self) -> None:
        col = self.table.currentColumn()
        if col < 0:
            return
        new_values = {}
        for (r, c), v in self.raw_values.items():
            if c == col:
                continue
            elif c > col:
                new_values[(r, c - 1)] = v
            else:
                new_values[(r, c)] = v
        self.raw_values = new_values
        self._updating = True
        self.table.removeColumn(col)
        self._update_column_headers()
        self._updating = False
        self.recalculate()
        self._mark_changed()

    def _shift_data_down(self, from_row: int) -> None:
        new_values = {}
        for (r, c), v in self.raw_values.items():
            if r >= from_row:
                new_values[(r + 1, c)] = v
            else:
                new_values[(r, c)] = v
        self.raw_values = new_values

    def _shift_data_right(self, from_col: int) -> None:
        new_values = {}
        for (r, c), v in self.raw_values.items():
            if c >= from_col:
                new_values[(r, c + 1)] = v
            else:
                new_values[(r, c)] = v
        self.raw_values = new_values

    def _update_column_headers(self) -> None:
        for c in range(self.table.columnCount()):
            header = self.table.horizontalHeaderItem(c)
            if not header or not header.text() or header.text() == column_name(c - 1) or header.text() == column_name(c + 1):
                self.table.setHorizontalHeaderItem(c, QTableWidgetItem(column_name(c)))

    def _show_tool_panel(self, name: str) -> None:
        """Signal parent MainWindow to show a tool panel."""
        window = self.window()
        if hasattr(window, 'show_tool_panel'):
            window.show_tool_panel(name)

    # ── Search & Filter ──

    def _toggle_search_bar(self) -> None:
        if self.search_bar.isVisible():
            self._hide_search_bar()
        else:
            self.search_bar.show()
            self.search_input.setFocus()
            self.search_input.selectAll()

    def _hide_search_bar(self) -> None:
        self.search_bar.hide()
        self._clear_search_highlights()
        self.search_input.clear()
        self._search_matches.clear()
        self._search_index = -1
        self.search_count_label.setText("")

    def _on_search(self, text: str) -> None:
        self._clear_search_highlights()
        self._search_matches.clear()
        self._search_index = -1
        if not text.strip():
            self.search_count_label.setText("")
            return
        needle = text.strip().lower()
        for row in range(self.table.rowCount()):
            for col in range(self.table.columnCount()):
                val = self._display_value(row, col)
                if val and needle in val.lower():
                    self._search_matches.append((row, col))
                    item = self._ensure_item(row, col)
                    item.setBackground(QColor("#fef3c7"))
        count = len(self._search_matches)
        self.search_count_label.setText(f"{count} found" if count else "No matches")
        if self._search_matches:
            self._search_index = 0
            self._goto_search_match()

    def _search_next(self) -> None:
        if not self._search_matches:
            return
        self._search_index = (self._search_index + 1) % len(self._search_matches)
        self._goto_search_match()

    def _search_prev(self) -> None:
        if not self._search_matches:
            return
        self._search_index = (self._search_index - 1) % len(self._search_matches)
        self._goto_search_match()

    def _goto_search_match(self) -> None:
        if 0 <= self._search_index < len(self._search_matches):
            row, col = self._search_matches[self._search_index]
            self.table.setCurrentCell(row, col)
            self.search_count_label.setText(
                f"{self._search_index + 1}/{len(self._search_matches)}"
            )

    def _clear_search_highlights(self) -> None:
        for row, col in self._search_matches:
            item = self.table.item(row, col)
            if item and item.background().color() == QColor("#fef3c7"):
                item.setBackground(QColor(0, 0, 0, 0))

    def clear_sheet(self) -> None:
        self._updating = True
        try:
            self.table.clearContents()
            self.raw_values.clear()
            for start_row, start_col, _row_span, _col_span in self.merged_ranges:
                self.table.setSpan(start_row, start_col, 1, 1)
            self.merged_ranges.clear()
        finally:
            self._updating = False
        self._mark_changed()

    def undo(self) -> None:
        if self.history_index > 0:
            self.history_index -= 1
            state = self.history[self.history_index]
            self._updating = True
            try:
                self.load_state(state)
            finally:
                self._updating = False
            self._dirty = True
            self.data_changed.emit()

    def redo(self) -> None:
        if self.history_index < len(self.history) - 1:
            self.history_index += 1
            state = self.history[self.history_index]
            self._updating = True
            try:
                self.load_state(state)
            finally:
                self._updating = False
            self._dirty = True
            self.data_changed.emit()

    def set_template_columns(self, template: ExtractionTemplate, include_money_check: bool = False) -> None:
        names = template.field_names()
        if include_money_check and MONEY_CHECK_COLUMN not in names:
            names = names + [MONEY_CHECK_COLUMN]
        while self.table.columnCount() < max(1, len(names)):
            self.add_column()
        for col, name in enumerate(names):
            self.table.setHorizontalHeaderItem(col, QTableWidgetItem(name))

    def append_extraction_row(self, template: ExtractionTemplate, values: dict[str, str]) -> int:
        first_row, _count = self.append_extraction_rows(template, [values])
        return first_row

    def append_extraction_rows(self, template: ExtractionTemplate, rows: list[dict[str, str]]) -> tuple[int, int]:
        template_fields = template.field_names()
        rows = [
            row for row in rows
            if any(_is_meaningful_extracted_value(row.get(field_name, "")) for field_name in template_fields)
        ]
        if not rows:
            return self._next_empty_row(), 0
        checks = [money_check(row) for row in rows]
        include_money_check = any(check for check in checks)
        self.set_template_columns(template, include_money_check)
        first_row = self._next_empty_row()
        current_row = first_row
        fields = template_fields
        if include_money_check:
            fields = fields + [MONEY_CHECK_COLUMN]
        self._updating = True
        try:
            for values, check in zip(rows, checks):
                if current_row >= self.table.rowCount():
                    self.table.insertRow(self.table.rowCount())
                max_lines = 1
                for col, field_name in enumerate(fields):
                    item = self._ensure_item(current_row, col)
                    value = check if field_name == MONEY_CHECK_COLUMN else values.get(field_name, "")
                    item.setText(value)
                    item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
                    if value:
                        self.raw_values[(current_row, col)] = value
                    else:
                        self.raw_values.pop((current_row, col), None)
                    max_lines = max(max_lines, value.count("\n") + 1)
                self.table.setRowHeight(current_row, max(28, min(120, 24 * max_lines)))
                current_row += 1
        finally:
            self._updating = False
        self._mark_changed()
        return first_row, len(rows)

    def _next_empty_row(self) -> int:
        for row in range(self.table.rowCount()):
            has_value = False
            for col in range(self.table.columnCount()):
                if self.raw_values.get((row, col)) or self._display_value(row, col):
                    has_value = True
                    break
            if not has_value:
                return row
        return self.table.rowCount()

    def table_context_csv(self, max_rows: int = 60) -> str:
        from bamini import bamini_to_unicode
        headers = [self.table.horizontalHeaderItem(col).text() if self.table.horizontalHeaderItem(col) else column_name(col) for col in range(self.table.columnCount())]
        active_cols = self._active_columns()
        if not active_cols:
            return ""
        writer_buffer: list[str] = []
        writer_buffer.append(",".join(csv_escape(headers[col]) for col in active_cols))
        rows_added = 0
        for row in range(self.table.rowCount()):
            values = []
            for col in active_cols:
                item = self.table.item(row, col)
                value = item.text() if item else ""
                if item and item.font().family().casefold().startswith("bamini") and not has_tamil_unicode(value):
                    value = bamini_to_unicode(value)
                values.append(value)
            if not any(values):
                continue
            writer_buffer.append(",".join(csv_escape(value) for value in values))
            rows_added += 1
            if rows_added >= max_rows:
                break
        return "\n".join(writer_buffer)

    def _active_columns(self) -> list[int]:
        columns: list[int] = []
        for col in range(self.table.columnCount()):
            header = self.table.horizontalHeaderItem(col)
            header_text = header.text() if header else column_name(col)
            has_data = bool(header_text and header_text != column_name(col))
            if not has_data:
                for row in range(self.table.rowCount()):
                    if self._display_value(row, col):
                        has_data = True
                        break
            if has_data:
                columns.append(col)
        return columns

    def open_native(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Open saved sheet", "", NATIVE_FILE_FILTER)
        if not path:
            return
        self.load_native(Path(path))

    def save_native(self) -> None:
        if self.current_path is None:
            self.save_native_as()
            return
        self.write_native(self.current_path)
        QMessageBox.information(self, "Saved", f"Saved {self.current_path}")

    def save_native_as(self) -> None:
        path, _ = QFileDialog.getSaveFileName(self, "Save sheet inside app", "local_ai_sheet.laisheet", NATIVE_FILE_FILTER)
        if not path:
            return
        target = Path(path)
        if target.suffix.lower() != ".laisheet":
            target = target.with_suffix(".laisheet")
        self.current_path = target
        self.write_native(target)
        QMessageBox.information(self, "Saved", f"Saved {target}")

    def write_native(self, path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(self.to_state(), indent=2, ensure_ascii=False), encoding="utf-8")
        self._dirty = False
        self.status_label.setText(f"Saved {datetime.now().strftime('%H:%M:%S')}")

    def load_native(self, path: Path) -> None:
        payload = json.loads(path.read_text(encoding="utf-8"))
        self.load_state(payload)
        self.history = [self.to_state()]
        self.history_index = 0
        self.current_path = path
        self._dirty = False
        self.status_label.setText(f"Loaded {path.name}")

    def autosave(self) -> None:
        if not self._dirty:
            return
        try:
            AUTOSAVE_FILE.parent.mkdir(parents=True, exist_ok=True)
            AUTOSAVE_FILE.write_text(json.dumps(self.to_state(), indent=2, ensure_ascii=False), encoding="utf-8")
        except OSError as exc:
            self.status_label.setText(f"Autosave failed: {exc}")
            return
        self._dirty = False
        self.status_label.setText(f"Autosaved {datetime.now().strftime('%H:%M:%S')}")

    def load_autosave_if_exists(self) -> None:
        if not AUTOSAVE_FILE.exists():
            return
        try:
            payload = json.loads(AUTOSAVE_FILE.read_text(encoding="utf-8"))
            self.load_state(payload)
        except (OSError, json.JSONDecodeError) as exc:
            self.status_label.setText(f"Autosave not loaded: {exc}")
            return
        self.status_label.setText("Recovered autosave")

    def to_state(self) -> dict[str, object]:
        cells: list[dict[str, object]] = []
        for row in range(self.table.rowCount()):
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                raw = self.raw_values.get((row, col), "")
                if not raw and not item_has_custom_style(item):
                    continue
                cells.append(
                    {
                        "row": row,
                        "col": col,
                        "raw": raw,
                        "style": item_to_style(item),
                    }
                )
        headers = [
            self.table.horizontalHeaderItem(col).text() if self.table.horizontalHeaderItem(col) else column_name(col)
            for col in range(self.table.columnCount())
        ]
        return {
            "version": 1,
            "saved_at": datetime.now().isoformat(timespec="seconds"),
            "row_count": self.table.rowCount(),
            "column_count": self.table.columnCount(),
            "headers": headers,
            "column_widths": [self.table.columnWidth(col) for col in range(self.table.columnCount())],
            "merged_ranges": self.merged_ranges,
            "cells": cells,
        }

    def load_state(self, state: dict[str, object]) -> None:
        current_row, current_col = self.table.currentRow(), self.table.currentColumn()
        for start_row, start_col, _row_span, _col_span in self.merged_ranges:
            self.table.setSpan(start_row, start_col, 1, 1)
        row_count = max(200, int(state.get("row_count", 200)))
        column_count = max(26, int(state.get("column_count", 26)))
        headers = state.get("headers", [])
        column_widths = state.get("column_widths", [])
        self._updating = True
        try:
            self.table.clear()
            self.table.setRowCount(row_count)
            self.table.setColumnCount(column_count)
            self.raw_values = {}
            self.merged_ranges = []
            for col in range(column_count):
                label = headers[col] if isinstance(headers, list) and col < len(headers) else column_name(col)
                self.table.setHorizontalHeaderItem(col, QTableWidgetItem(str(label)))
                if isinstance(column_widths, list) and col < len(column_widths):
                    try:
                        self.table.setColumnWidth(col, int(column_widths[col]))
                    except (TypeError, ValueError):
                        pass
            for cell in state.get("cells", []):
                if not isinstance(cell, dict):
                    continue
                row = int(cell.get("row", 0))
                col = int(cell.get("col", 0))
                if row >= self.table.rowCount() or col >= self.table.columnCount():
                    continue
                raw = str(cell.get("raw", ""))
                item = self._ensure_item(row, col)
                item.setText(raw)
                apply_item_style(item, cell.get("style"))
                if raw:
                    self.raw_values[(row, col)] = raw
                if raw.count("\n"):
                    self.table.setRowHeight(row, max(28, min(120, 24 * (raw.count("\n") + 1))))
                if item.font().pointSize() != DEFAULT_CELL_FONT_SIZE:
                    self.table.resizeRowToContents(row)
            for merged in state.get("merged_ranges", []):
                if not isinstance(merged, list) and not isinstance(merged, tuple):
                    continue
                if len(merged) != 4:
                    continue
                start_row, start_col, row_span, col_span = [int(value) for value in merged]
                self.table.setSpan(start_row, start_col, row_span, col_span)
                self.merged_ranges.append((start_row, start_col, row_span, col_span))
        finally:
            self._updating = False
        self.recalculate()
        if 0 <= current_row < row_count and 0 <= current_col < column_count:
            self.table.setCurrentCell(current_row, current_col)
        self._sync_font_size()

    def export_xlsx(self) -> None:
        path, _ = QFileDialog.getSaveFileName(self, "Export Excel workbook", "extracted_data.xlsx", "Excel Workbook (*.xlsx)")
        if not path:
            return
        self.save_xlsx(Path(path))
        QMessageBox.information(self, "Export complete", f"Saved {path}")

    def save_xlsx(self, path: Path) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Sheet1"
        active_cols = self._active_columns() or list(range(self.table.columnCount()))

        for out_col, col in enumerate(active_cols, start=1):
            header = self.table.horizontalHeaderItem(col)
            c = sheet.cell(row=1, column=out_col, value=header.text() if header else column_name(col))
            c.font = Font(name="Segoe UI", bold=True)
            if has_tamil_unicode(c.value):
                self._apply_xlsx_tamil(c)

        for row in range(self.table.rowCount()):
            if not any(self._display_value(row, col) or self.raw_values.get((row, col)) for col in active_cols):
                continue
            out_row = sheet.max_row + 1
            for out_col, col in enumerate(active_cols, start=1):
                raw = self.raw_values.get((row, col), self._display_value(row, col))
                cell = sheet.cell(row=out_row, column=out_col, value=raw)
                item = self.table.item(row, col)
                if item:
                    font = item.font()
                    foreground = item.foreground()
                    background = item.background()
                    font_color = qcolor_to_rgb(foreground.color()) if foreground.style() != Qt.BrushStyle.NoBrush else None
                    cell.font = Font(
                        name=font.family(),
                        size=font.pointSize() if font.pointSize() > 0 else DEFAULT_CELL_FONT_SIZE,
                        bold=font.bold(),
                        italic=font.italic(),
                        underline="single" if font.underline() else None,
                        color=font_color,
                    )
                    if background.style() != Qt.BrushStyle.NoBrush:
                        cell.fill = PatternFill("solid", fgColor=qcolor_to_rgb(background.color()))
                    cell.alignment = Alignment(
                        horizontal=qt_alignment_to_openpyxl(item.textAlignment()),
                        vertical="top" if "\n" in str(raw) else "center",
                        wrap_text="\n" in str(raw),
                    )
                if has_tamil_unicode(raw) and not raw.startswith("="):
                    self._apply_xlsx_tamil(cell)

        for start_row, start_col, row_span, col_span in self.merged_ranges:
            if start_col in active_cols:
                out_start_col = active_cols.index(start_col) + 1
                out_end_col = out_start_col + col_span - 1
                sheet.merge_cells(
                    start_row=start_row + 2,
                    start_column=out_start_col,
                    end_row=start_row + row_span + 1,
                    end_column=out_end_col,
                )

        for index, col in enumerate(active_cols, start=1):
            sheet.column_dimensions[get_column_letter(index)].width = max(12, self.table.columnWidth(col) / 7)
        workbook.save(path)

    def _apply_xlsx_tamil(self, cell) -> None:
        font = copy(cell.font)
        bamini = self.table.itemDelegate().bamini_family or "Bamini TAMIL"
        english = font.name or "Segoe UI"
        if english.casefold().startswith("bamini"):
            english = "Segoe UI"
        runs = bamini_display_runs(cell.value)
        if not any(is_bamini for _, is_bamini in runs):
            return
        cell.value = CellRichText([
            TextBlock(
                InlineFont(
                    rFont=bamini if is_bamini else english,
                    sz=font.sz, b=font.b, i=font.i, u=font.u, color=font.color,
                ),
                text,
            )
            for text, is_bamini in runs
        ])


class TemplateManagerWidget(QWidget):
    templates_changed = Signal()

    def __init__(self, store: TemplateStore, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.store = store
        self.current_template_id = uuid.uuid4().hex
        self.current_fields: list[TemplateField] = []

        self.template_combo = QComboBox()
        self.template_combo.currentIndexChanged.connect(self._load_selected_template)
        self.template_name = QLineEdit()
        self.field_name = QLineEdit()
        self.field_list = QListWidget()
        self.field_list.currentRowChanged.connect(self._field_selected)
        self.canvas = ImageSelectionCanvas()
        self.canvas.selection_finished.connect(self._selection_finished)
        self.pending_rect: tuple[float, float, float, float] | None = None

        left = QVBoxLayout()
        form = QFormLayout()
        form.addRow("Template", self.template_combo)
        form.addRow("Name", self.template_name)
        form.addRow("Field", self.field_name)
        left.addLayout(form)

        buttons = QGridLayout()
        self._add_button(buttons, "New", self.new_template, 0, 0)
        self._add_button(buttons, "Save", self.save_template, 0, 1)
        self._add_button(buttons, "Delete", self.delete_template, 0, 2)
        self._add_button(buttons, "Add Field", self.add_field, 1, 0)
        self._add_button(buttons, "Remove Field", self.remove_field, 1, 1)
        self._add_button(buttons, "Assign Box", self.assign_box_to_field, 1, 2)
        self._add_button(buttons, "Load Sample", self.load_sample, 2, 0, 1, 3)
        left.addLayout(buttons)
        left.addWidget(self.field_list)

        container = QWidget()
        container.setLayout(left)
        container.setMaximumWidth(360)

        layout = QHBoxLayout(self)
        layout.addWidget(container)
        layout.addWidget(self.canvas, 1)

        self.refresh_templates()
        self.new_template()

    def _add_button(self, layout: QGridLayout, label: str, handler, row: int, col: int, row_span: int = 1, col_span: int = 1) -> None:
        button = QPushButton(label)
        button.clicked.connect(handler)
        tones = {
            "New": "neutral",
            "Save": "success",
            "Delete": "danger",
            "Add Field": "structure",
            "Remove Field": "danger",
            "Assign Box": "merge",
            "Load Sample": "export",
        }
        style_tool_button(button, tones.get(label, "neutral"))
        layout.addWidget(button, row, col, row_span, col_span)

    def refresh_templates(self) -> None:
        selected = self.current_template_id
        self.template_combo.blockSignals(True)
        self.template_combo.clear()
        self.template_combo.addItem("Select saved template", "")
        for template in self.store.all():
            self.template_combo.addItem(template.name, template.template_id)
        index = self.template_combo.findData(selected)
        self.template_combo.setCurrentIndex(index if index >= 0 else 0)
        self.template_combo.blockSignals(False)

    def new_template(self) -> None:
        self.current_template_id = uuid.uuid4().hex
        self.template_combo.setCurrentIndex(0)
        self.template_name.setText("")
        self.current_fields = []
        self.field_list.clear()
        self.canvas.set_rects([])

    def _load_selected_template(self) -> None:
        template_id = self.template_combo.currentData()
        if not template_id:
            return
        template = self.store.get(template_id)
        if not template:
            return
        self.current_template_id = template.template_id
        self.template_name.setText(template.name)
        self.current_fields = [TemplateField(**field.__dict__) for field in template.fields]
        self._render_fields()

    def add_field(self) -> None:
        name = self.field_name.text().strip()
        if not name:
            return
        if any(field.name.lower() == name.lower() for field in self.current_fields):
            QMessageBox.warning(self, "Duplicate field", "That field already exists in this template.")
            return
        self.current_fields.append(TemplateField(name=name))
        self.field_name.clear()
        self._render_fields()

    def remove_field(self) -> None:
        row = self.field_list.currentRow()
        if row < 0:
            return
        self.current_fields.pop(row)
        self._render_fields()

    def load_sample(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Load sample document", "", supported_file_filter())
        if not path:
            return
        try:
            image = render_page(path)
        except Exception as exc:
            QMessageBox.critical(self, "Could not load document", str(exc))
            return
        self.canvas.set_image(image)
        self._render_fields()

    def _selection_finished(self, rect: object) -> None:
        self.pending_rect = rect  # type: ignore[assignment]

    def assign_box_to_field(self) -> None:
        row = self.field_list.currentRow()
        if row < 0:
            QMessageBox.information(self, "Select field", "Select a field first.")
            return
        rect = self.canvas.selected_rect() or self.pending_rect
        if not rect:
            QMessageBox.information(self, "Draw box", "Draw a box on the sample document first.")
            return
        field = self.current_fields[row]
        field.x, field.y, field.width, field.height = rect
        self._render_fields()

    def save_template(self) -> None:
        name = self.template_name.text().strip()
        if not name:
            QMessageBox.warning(self, "Missing name", "Enter a template name.")
            return
        if not self.current_fields:
            QMessageBox.warning(self, "Missing fields", "Add at least one field.")
            return
        template = ExtractionTemplate(name=name, fields=self.current_fields, template_id=self.current_template_id)
        self.store.upsert(template)
        self.refresh_templates()
        self.templates_changed.emit()
        QMessageBox.information(self, "Template saved", f"Saved template '{name}'.")

    def delete_template(self) -> None:
        if not self.current_template_id:
            return
        self.store.delete(self.current_template_id)
        self.new_template()
        self.refresh_templates()
        self.templates_changed.emit()

    def _field_selected(self, row: int) -> None:
        if 0 <= row < len(self.current_fields):
            self.canvas.selected_name = self.current_fields[row].name
            self.canvas.update()

    def _render_fields(self) -> None:
        self.field_list.clear()
        for field in self.current_fields:
            suffix = " [boxed]" if field.has_box else ""
            self.field_list.addItem(QListWidgetItem(f"{field.name}{suffix}"))
        self.canvas.set_rects(
            (field.name, (field.x, field.y, field.width, field.height))
            for field in self.current_fields
            if field.has_box
        )


class ExtractionWidget(QWidget):
    def __init__(self, store: TemplateStore, spreadsheet: SpreadsheetWidget, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.store = store
        self.spreadsheet = spreadsheet
        self.files = QListWidget()
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.template_combo = QComboBox()
        self.ollama_host = QLineEdit("http://localhost:11434")
        self.model_preset = QComboBox()
        for label, value in VISION_MODEL_PRESETS:
            self.model_preset.addItem(label, value)
        self.model_preset.currentIndexChanged.connect(self.apply_model_preset)
        self.vision_model = QLineEdit("qwen2.5vl:7b")
        self.language_hint = QLineEdit("Tamil and English")
        self.use_ai = QCheckBox("Use Ollama vision")
        self.use_ai.setChecked(True)
        self.table_rows = QCheckBox("AI table rows")
        self.table_rows.setChecked(True)
        self.all_pages = QCheckBox("Read every PDF page")
        self.all_pages.setChecked(True)

        controls = QGroupBox("Extraction")
        form = QFormLayout(controls)
        form.addRow("Template", self.template_combo)
        form.addRow("Ollama host", self.ollama_host)
        form.addRow("Model preset", self.model_preset)
        form.addRow("Vision model", self.vision_model)
        form.addRow("Languages", self.language_hint)
        form.addRow("", self.use_ai)
        form.addRow("", self.table_rows)
        form.addRow("", self.all_pages)

        buttons = QHBoxLayout()
        add_button = QPushButton("Add Files")
        add_button.clicked.connect(self.add_files)
        style_tool_button(add_button, "neutral")
        extract_button = QPushButton("Extract")
        extract_button.clicked.connect(self.extract)
        style_tool_button(extract_button, "success")
        clear_button = QPushButton("Clear")
        clear_button.clicked.connect(self.files.clear)
        style_tool_button(clear_button, "danger")
        buttons.addWidget(add_button)
        buttons.addWidget(extract_button)
        buttons.addWidget(clear_button)

        layout = QVBoxLayout(self)
        layout.addWidget(controls)
        layout.addLayout(buttons)
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.progress.setTextVisible(True)
        self.progress.setFormat("%v / %m files")
        layout.addWidget(self.progress)
        layout.addWidget(self.files, 1)
        layout.addWidget(self.log, 1)

        # Enable drag-and-drop
        self.setAcceptDrops(True)
        self.files.setAcceptDrops(True)

        self.refresh_templates()

    def apply_model_preset(self) -> None:
        model = self.model_preset.currentData()
        if model:
            self.vision_model.setText(str(model))

    def refresh_templates(self) -> None:
        current = self.template_combo.currentData()
        self.template_combo.clear()
        for template in self.store.all():
            self.template_combo.addItem(template.name, template.template_id)
        if current:
            index = self.template_combo.findData(current)
            if index >= 0:
                self.template_combo.setCurrentIndex(index)

    def add_files(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(self, "Add PDFs or images", "", supported_file_filter())
        for path in paths:
            self.files.addItem(path)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent) -> None:
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path:
                self.files.addItem(path)

    def extract(self) -> None:
        template_id = self.template_combo.currentData()
        template = self.store.get(template_id) if template_id else None
        if not template:
            QMessageBox.warning(self, "No template", "Choose a saved template first.")
            return
        if self.files.count() == 0:
            QMessageBox.warning(self, "No files", "Add at least one PDF or image.")
            return

        client = OllamaClient(
            host=self.ollama_host.text().strip(),
            vision_model=self.vision_model.text().strip(),
            language_hint=self.language_hint.text().strip() or "Tamil and English",
        )
        use_ai_for_run = self.use_ai.isChecked()
        if use_ai_for_run:
            available, message = client.check_connection()
            if not available:
                use_ai_for_run = False
                self.log.append(
                    f"Ollama is not running at {client.host}. Continuing with OCR only.\n"
                    f"{message}\n"
                    "Start Ollama with `ollama serve`, then run the import again for AI vision extraction.\n"
                )

        if not use_ai_for_run:
            ready, message = tesseract_status()
            if not ready:
                self.log.append(
                    "OCR fallback is not ready either, so extracted cells may be blank.\n"
                    f"{message}\n"
                    "Install Tesseract OCR or start Ollama to extract values from images/PDFs.\n"
                )
            if self.table_rows.isChecked():
                self.log.append("AI table rows needs Ollama vision. Falling back to box/OCR extraction.\n")

        total_files = self.files.count()
        self.progress.setMaximum(total_files)
        self.progress.setValue(0)
        self.progress.setVisible(True)

        for index in range(total_files):
            path = self.files.item(index).text()
            self.progress.setValue(index)
            self.progress.setFormat(f"{index + 1} / {total_files}: {Path(path).name}")
            QApplication.processEvents()
            try:
                rows = self._extract_file_rows(path, template, client, use_ai_for_run)
            except Exception as exc:
                self.log.append(f"Failed: {path}\n{exc}\n")
                continue
            first_row, count = self.spreadsheet.append_extraction_rows(template, rows)
            if count == 0:
                self.log.append(f"No rows extracted: {Path(path).name}\n")
                continue
            if count == 1:
                self.log.append(f"Row {first_row + 1}: {Path(path).name}\n{extraction_summary(template, rows[0])}\n")
            else:
                self.log.append(f"Rows {first_row + 1}-{first_row + count}: {Path(path).name}\nExtracted {count} table rows.\n")
        self.progress.setVisible(False)
        self.log.append("\n✅ Extraction complete.\n")

    def _extract_file_rows(self, path: str, template: ExtractionTemplate, client: OllamaClient, use_ai: bool) -> list[dict[str, str]]:
        if self.table_rows.isChecked() and use_ai:
            rows: list[dict[str, str]] = []
            pages = page_count(path) if self.all_pages.isChecked() else 1
            for page_number in range(pages):
                QApplication.processEvents()
                image = render_page(path, page_number)
                page_rows = client.extract_table_rows(template, image)
                rows.extend(page_rows)
                self.log.append(f"Page {page_number + 1}: found {len(page_rows)} table rows.")
            if rows:
                return rows
            self.log.append("No JSON table rows were returned by the model. Falling back to template boxes.\n")
        return [self._extract_file_box(path, template, client, use_ai)]

    def _extract_file_box(self, path: str, template: ExtractionTemplate, client: OllamaClient, use_ai: bool) -> dict[str, str]:
        page_cache: dict[int, object] = {}
        values: dict[str, str] = {}
        ai_disabled_after_error = False
        for field in template.fields:
            QApplication.processEvents()
            if not field.has_box:
                values[field.name] = ""
                continue
            if field.page not in page_cache:
                page_cache[field.page] = render_page(path, field.page)
            image = page_cache[field.page]
            crop = crop_normalized(image, (field.x, field.y, field.width, field.height))  # type: ignore[arg-type]
            value = ""
            if use_ai and not ai_disabled_after_error:
                try:
                    value = client.extract_crop(field.name, crop)
                except Exception as exc:
                    ai_disabled_after_error = True
                    self.log.append(f"Ollama failed during extraction; using OCR for the rest of this file.\n{exc}\n")
            if not value:
                value = ocr_crop(crop)
            values[field.name] = value.strip()
        return values


class SourcesWidget(QWidget):
    def __init__(self, rag, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.rag = rag
        self.files = QListWidget()
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        
        buttons = QHBoxLayout()
        add_button = QPushButton("Add to Notebook")
        add_button.clicked.connect(self.add_files)
        style_tool_button(add_button, "neutral")
        clear_button = QPushButton("Clear Notebook")
        clear_button.clicked.connect(self.clear_notebook)
        style_tool_button(clear_button, "danger")
        
        buttons.addWidget(add_button)
        buttons.addWidget(clear_button)
        
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Upload PDFs, images, or text files to your local NotebookLM knowledge base."))
        layout.addLayout(buttons)
        layout.addWidget(self.files, 1)
        layout.addWidget(self.log, 1)

    def add_files(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(self, "Add Documents", "", supported_file_filter())
        for path in paths:
            self.log.append(f"Ingesting {Path(path).name}...")
            QApplication.processEvents()
            try:
                chunks = self.rag.ingest_document(Path(path))
                self.files.addItem(f"{Path(path).name} ({chunks} chunks)")
                self.log.append(f"Success: {chunks} chunks embedded.\n")
            except Exception as exc:
                self.log.append(f"Failed to ingest {path}: {exc}\n")

    def clear_notebook(self) -> None:
        self.rag.clear()
        self.files.clear()
        self.log.append("Notebook cleared.\n")


class ChatWorker(QThread):
    """Run chat inference off the main thread."""
    finished = Signal(str)
    error = Signal(str)

    def __init__(self, client: OllamaClient, prompt: str) -> None:
        super().__init__()
        self.client = client
        self.prompt = prompt

    def run(self) -> None:
        try:
            answer = self.client.chat(self.prompt)
            self.finished.emit(answer)
        except Exception as exc:
            self.error.emit(str(exc))


class ChatPanel(QWidget):
    def __init__(self, spreadsheet: SpreadsheetWidget, rag, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.spreadsheet = spreadsheet
        self.rag = rag
        self.history: list[tuple[str, str]] = []
        self._worker: ChatWorker | None = None
        self._pending_citations: list[str] = []

        self.chat_view = QTextBrowser()
        self.chat_view.setOpenExternalLinks(False)
        self.chat_view.setHtml(CHAT_CSS + '<div class="clear"></div>')

        self.message = QLineEdit()
        self.message.setPlaceholderText("Ask a question about your data...")
        self.message.returnPressed.connect(self.send)
        self.host = QLineEdit("http://localhost:11434")
        self.model = QLineEdit("gemma3:12b")

        send_button = QPushButton("⬆ Send")
        send_button.clicked.connect(self.send)
        style_tool_button(send_button, "primary")
        clear_button = QPushButton("🗑")
        clear_button.setFixedWidth(36)
        clear_button.setToolTip("Clear chat")
        clear_button.clicked.connect(self.clear_chat)
        style_tool_button(clear_button, "danger")

        self.typing_label = QLabel("")
        self.typing_label.setStyleSheet("color: #6366f1; font-style: italic; font-size: 12px;")

        form = QFormLayout()
        form.addRow("Host", self.host)
        form.addRow("Chat model", self.model)

        bottom = QHBoxLayout()
        bottom.addWidget(self.message, 1)
        bottom.addWidget(send_button)
        bottom.addWidget(clear_button)

        layout = QVBoxLayout(self)
        title = QLabel("🤖 AI Studio")
        title.setStyleSheet("font-size: 16px; font-weight: 800; color: #a78bfa; padding: 4px 0;")
        layout.addWidget(title)
        layout.addLayout(form)
        layout.addWidget(self.chat_view, 1)
        layout.addWidget(self.typing_label)
        layout.addLayout(bottom)

    def send(self) -> None:
        question = self.message.text().strip()
        if not question:
            return
        self.message.clear()
        self.message.setEnabled(False)
        self.history.append(("User", question))
        self._append_bubble("user", question)
        self.typing_label.setText("● AI is thinking...")
        QApplication.processEvents()

        table_context = self.spreadsheet.table_context_csv(max_rows=60)

        rag_context = ""
        self._pending_citations = []
        try:
            chunks = self.rag.search(question, n_results=5)
            if chunks:
                lines = []
                for idx, c in enumerate(chunks, start=1):
                    lines.append(f"--- Source {idx} ({c.source}, Page {c.page}) ---\n{c.text}\n")
                    self._pending_citations.append(f"[{idx}] {c.source}")
                rag_context = "\n".join(lines)
        except Exception as exc:
            rag_context = f"(RAG Search Failed: {exc})"

        prompt = self.build_prompt(table_context, rag_context, self.history, question)
        client = OllamaClient(host=self.host.text().strip(), chat_model=self.model.text().strip())

        self._worker = ChatWorker(client, prompt)
        self._worker.finished.connect(self._on_chat_response)
        self._worker.error.connect(self._on_chat_error)
        self._worker.start()

    def _on_chat_response(self, answer: str) -> None:
        self.typing_label.setText("")
        self.message.setEnabled(True)
        self.history.append(("Assistant", answer))
        self._append_bubble("ai", answer)
        if self._pending_citations:
            cites = "<br>".join(set(self._pending_citations))
            self.chat_view.append(f'<div class="sources">📎 Sources:<br>{cites}</div>')
        self.message.setFocus()

    def _on_chat_error(self, error: str) -> None:
        self.typing_label.setText("")
        self.message.setEnabled(True)
        self._append_bubble("ai", f"❌ Error: {error}")
        self.message.setFocus()

    def _append_bubble(self, role: str, text: str) -> None:
        escaped = escape_html(text)
        cls = "user-msg" if role == "user" else "ai-msg"
        self.chat_view.append(f'<div class="{cls}">{escaped}</div><div class="clear"></div>')

    def clear_chat(self) -> None:
        self.history.clear()
        self.chat_view.setHtml(CHAT_CSS + '<div class="clear"></div>')

    def build_prompt(self, table_context: str, rag_context: str, history: list[tuple[str, str]], question: str) -> str:
        recent_history = history[-8:]
        lines = [
            "You are a local AI assistant. You have access to a structured spreadsheet and an unstructured knowledge base.",
            "Answer using the provided contexts when relevant. Cite sources when appropriate.",
            "",
            "Spreadsheet context:",
            table_context or "(no spreadsheet data)",
            "",
            "Knowledge Base context:",
            rag_context or "(no documents found)",
            "",
            "Recent conversation:",
        ]
        for role, content in recent_history:
            lines.append(f"{role}: {content}")
        lines.extend(["", f"User: {question}", "Assistant:"])
        return "\n".join(lines)


# ── Petition Extractor Widget (merged from EOWassist) ──

DEFAULT_PETITION_HEADERS = [
    "Sender Name & Address", "Petitioner Name & Address", "Petitioner Mobile No",
    "Deposit Amount", "Repaid Amount", "Default Amount",
    "ID Number", "Plan Name", "161 Statement", "Agent", "Crime No",
]

DEFAULT_PETITION_PROMPT = (
    "You are an expert document analyzer. Your task is to extract structured data from the provided document (which may be a petition, an order, or a legal statement).\n"
    "1. Understand the context and layout of the document.\n"
    "2. Identify the most relevant values for each requested field, even if the format varies.\n"
    "3. Handle Tamil and English text correctly. Preserve Tamil script exactly as Unicode Tamil.\n"
    "4. If a field spans multiple lines or is fragmented, reconstruct it logically.\n"
    "5. Return a JSON array of objects representing the extracted data entries."
)


EMPTY_EXTRACTED_VALUES = {"", "not specified", "not found", "n/a", "na", "none", "null", "-", "--"}


def _is_meaningful_extracted_value(value: object) -> bool:
    text = str(value or "").strip()
    return bool(text) and text.lower() not in EMPTY_EXTRACTED_VALUES


def _clean_extracted_rows(rows: list[dict[str, str]], headers: list[str]) -> list[dict[str, str]]:
    cleaned: list[dict[str, str]] = []
    for row in rows:
        values = {header: str(row.get(header, "") or "").strip() for header in headers}
        if any(_is_meaningful_extracted_value(value) for value in values.values()):
            cleaned.append(values)
    return cleaned


class PetitionThread(QThread):
    row_extracted = Signal(dict)
    all_finished = Signal()
    error = Signal(str)
    progress = Signal(str, int, int)
    file_finished = Signal(str, int)
    file_warning = Signal(str)

    def __init__(
        self,
        file_paths,
        headers,
        instruction,
        model_name,
        host="http://localhost:11434",
        engine="local",
        gemini_api_key="",
        fewshot_template=None,
    ):
        super().__init__()
        self.file_paths = file_paths
        self.headers = headers
        self.instruction = instruction
        self.model_name = model_name
        self.host = host
        self.engine = engine
        self.gemini_api_key = gemini_api_key
        self.fewshot_template = fewshot_template

    def run(self):
        from PIL import Image as PILImage
        import pymupdf as fitz

        if self.engine == "gemini":
            client = GeminiClient(api_key=self.gemini_api_key, model=self.model_name or DEFAULT_GEMINI_MODEL)
            is_vision = True
        else:
            client = OllamaClient(host=self.host, vision_model=self.model_name, chat_model=self.model_name)
            is_vision = client.is_vision_model(self.model_name)

        try:
            total_files = len(self.file_paths)
            for i, file_path in enumerate(self.file_paths, 1):
                base_name = Path(file_path).name
                self.progress.emit(f"[{i}/{total_files}] Processing: {base_name}", i - 1, total_files)
                images: list[PILImage.Image] = []
                text = ""
                try:
                    if file_path.lower().endswith(".pdf"):
                        doc = fitz.open(file_path)
                        try:
                            for page in doc:
                                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                                img = PILImage.frombytes("RGB", [pix.width, pix.height], pix.samples)
                                if is_vision:
                                    images.append(img)
                                else:
                                    text += self._ocr_image(img) + "\n\n"
                        finally:
                            doc.close()
                    else:
                        img = PILImage.open(file_path).convert("RGB")
                        if is_vision:
                            images.append(img)
                        else:
                            text += self._ocr_image(img) + "\n\n"

                    self.progress.emit(f"[{i}/{total_files}] Analyzing: {base_name}", i, total_files)
                    results = client.extract_flexible(
                        headers=self.headers,
                        images=images if images else None,
                        text=text if text else None,
                        instruction=self.instruction,
                        fewshot_template=self.fewshot_template,
                    )
                except Exception as exc:
                    self.file_warning.emit(f"{base_name}: {exc}")
                    self.file_finished.emit(base_name, 0)
                    continue

                results = _clean_extracted_rows(results, self.headers)
                for row in results:
                    row["_file_path"] = file_path
                    self.row_extracted.emit(row)
                self.file_finished.emit(base_name, len(results))

        except Exception as e:
            self.error.emit(f"Extraction failed: {str(e)}")
        finally:
            self.all_finished.emit()

    def _ocr_image(self, image) -> str:
        try:
            import pytesseract as tess  # type: ignore
        except ImportError:
            return ""
        try:
            return tess.image_to_string(image, lang="tam+eng", config="--psm 3")
        except Exception:
            try:
                return tess.image_to_string(image, config="--psm 3")
            except Exception:
                return ""


class ModelComboBox(QComboBox):
    textChanged = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setEditable(True)
        self.lineEdit().textChanged.connect(self.textChanged.emit)
        self.currentIndexChanged.connect(self._on_index_changed)

    def _on_index_changed(self, idx):
        self.textChanged.emit(self.text())

    def text(self) -> str:
        idx = self.currentIndex()
        if idx >= 0:
            data = self.itemData(idx)
            if data:
                return str(data)
        return self.currentText().strip()

    def setText(self, val: str):
        self.blockSignals(True)
        idx = -1
        for i in range(self.count()):
            if self.itemData(i) == val:
                idx = i
                break
        if idx >= 0:
            self.setCurrentIndex(idx)
        else:
            self.setEditText(val)
        self.blockSignals(False)
        self.textChanged.emit(val)


class PetitionWidget(QWidget):
    """Tamil Petition OCR+LLM extraction — merged from EOWassist."""

    def __init__(self, spreadsheet: SpreadsheetWidget, fewshot_store: FewShotStore, parent=None):
        super().__init__(parent)
        self.spreadsheet = spreadsheet
        self.fewshot_store = fewshot_store
        self.model_name = "qwen2.5vl:7b"
        self.engine_name = "local"
        self._local_model_name = self.model_name
        self._gemini_model_name = DEFAULT_GEMINI_MODEL
        self.current_prompt = DEFAULT_PETITION_PROMPT
        self._thread = None
        self._active_headers: list[str] = []
        self._cloud_version_col: int | None = None
        self._queued_files = 0
        self._finished_files = 0
        self._rows_extracted = 0
        self._skipped_files = 0
        self.setAcceptDrops(True)

        # CHANGE 3: Persist Gemini API Key — load saved key
        settings = load_settings()
        saved_key = settings.get('gemini_api_key', os.environ.get('GEMINI_API_KEY', ''))

        # Controls
        self.engine_combo = QComboBox()
        self.engine_combo.setMinimumWidth(160)
        self.engine_combo.addItem("Local Ollama", "local")
        self.engine_combo.addItem("Cloud Gemini", "gemini")
        self.engine_combo.currentIndexChanged.connect(self._apply_engine_combo)

        model_combo = QComboBox()
        model_combo.setMinimumWidth(160)
        for label, val in VISION_MODEL_PRESETS:
            model_combo.addItem(label, val)
        model_combo.currentIndexChanged.connect(lambda: self._apply_model_combo(model_combo))
        self.model_combo = model_combo
        self.model_input = ModelComboBox()
        self.model_input.setText(self.model_name)
        self.model_input.setPlaceholderText("Ollama model tag")
        self.model_input.textChanged.connect(self._model_text_changed)

        self.gemini_key_label = QLabel("Gemini key:")
        self.gemini_key = QLineEdit(saved_key)
        self.gemini_key.setPlaceholderText("GEMINI_API_KEY")
        self.gemini_key.setEchoMode(QLineEdit.EchoMode.Password)
        self.gemini_key_label.hide()
        self.gemini_key.hide()

        # CHANGE 5: FewShot template dropdown
        self.fewshot_combo = QComboBox()
        self.fewshot_combo.setMinimumWidth(160)
        self.refresh_fewshot_templates()

        btn_load = QPushButton("📄 Load Petitions")
        btn_load.clicked.connect(self.load_files)
        style_sidebar_button(btn_load, "primary")
        self._btn_load = btn_load

        btn_prompt = QPushButton("✍ Edit Prompt")
        btn_prompt.clicked.connect(self.edit_prompt)
        style_sidebar_button(btn_prompt, "neutral")

        btn_set_headers = QPushButton("📋 Set Headers")
        btn_set_headers.clicked.connect(self._set_default_headers)
        style_sidebar_button(btn_set_headers, "structure")

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.progress.setTextVisible(True)
        self.progress.setFormat("%v / %m files")

        self.status = QLabel("🛡 KAAVAL — Drop PDF/image files or click Load to begin extraction")
        self.status.setWordWrap(True)
        self.status.setStyleSheet(PETITION_STATUS_STYLE)
        self.stats = QLabel("")
        self.stats.setStyleSheet(PETITION_STATS_STYLE)
        self._update_stats()

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setPlaceholderText("Extraction log will appear here...")
        self.log.setStyleSheet(PETITION_LOG_STYLE)

        fields = QFormLayout()
        fields.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapLongRows)
        fields.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)
        fields.addRow("Engine", self.engine_combo)
        fields.addRow("Preset", model_combo)
        fields.addRow("Model", self.model_input)
        fields.addRow(self.gemini_key_label, self.gemini_key)
        fields.addRow("Training", self.fewshot_combo)

        actions = QHBoxLayout()
        actions.addWidget(btn_prompt)
        actions.addWidget(btn_set_headers)
        self.stats.setWordWrap(True)

        layout = QVBoxLayout(self)
        layout.setSpacing(6)
        layout.addWidget(btn_load)
        layout.addLayout(fields)
        layout.addLayout(actions)
        layout.addWidget(self.status)
        layout.addWidget(self.stats)
        layout.addWidget(self.progress)
        layout.addWidget(self.log, 1)

        self.refresh_ollama_models()

    def refresh_ollama_models(self):
        """Fetch and populate all pulled Ollama models, highlighting those without vision support in red."""
        if self.engine_name != "local":
            return
            
        client = OllamaClient()
        models = client.list_models()
        
        self.model_input.blockSignals(True)
        curr_text = self.model_input.text()
        self.model_input.clear()
        
        if not models:
            # Fallback presets when offline or not loaded yet
            presets = ["qwen2.5-vl", "minicpm-v", "llama3.2-vision", "gemma3:12b"]
            for idx, m in enumerate(presets):
                has_vision = client.is_vision_model(m)
                if has_vision:
                    label = f"👁 {m}"
                    self.model_input.addItem(label, m)
                else:
                    label = f"❌ {m} (No Vision)"
                    self.model_input.addItem(label, m)
                    self.model_input.setItemData(idx, QColor("#ef4444"), Qt.ItemDataRole.ForegroundRole)
        else:
            for idx, m in enumerate(models):
                has_vision = client.is_vision_model(m)
                if has_vision:
                    label = f"👁 {m}"
                    self.model_input.addItem(label, m)
                else:
                    label = f"❌ {m} (No Vision)"
                    self.model_input.addItem(label, m)
                    # Color it red in the dropdown
                    self.model_input.setItemData(idx, QColor("#ef4444"), Qt.ItemDataRole.ForegroundRole)
                    
        # Restore selection
        if curr_text:
            found = False
            for i in range(self.model_input.count()):
                if self.model_input.itemData(i) == curr_text:
                    self.model_input.setCurrentIndex(i)
                    found = True
                    break
            if not found:
                self.model_input.setEditText(curr_text)
        else:
            self.model_input.setCurrentIndex(0)
            
        self.model_input.blockSignals(False)

    def refresh_gemini_models(self):
        """Fetch and populate all supported Gemini models in the dropdown."""
        if self.engine_name != "gemini":
            return
            
        self.model_input.blockSignals(True)
        curr_text = self.model_input.text()
        self.model_input.clear()
        
        gemini_presets = [
            ("gemini-3.5-flash", "Gemini 3.5 Flash (Latest)"),
            ("gemini-3.1-flash-lite", "Gemini 3.1 Flash-Lite (Efficient)"),
            ("gemini-3.1-pro-preview", "Gemini 3.1 Pro (Advanced)"),
            ("gemini-2.5-flash", "Gemini 2.5 Flash"),
            ("gemini-2.5-pro", "Gemini 2.5 Pro"),
            ("gemini-2.5-flash-lite", "Gemini 2.5 Flash-Lite"),
            ("gemini-1.5-flash", "Gemini 1.5 Flash (Cheaper)"),
            ("gemini-1.5-flash-8b", "Gemini 1.5 Flash 8B (Budget)"),
            ("gemini-1.5-pro", "Gemini 1.5 Pro (Standard)"),
        ]
        
        for idx, (m, label) in enumerate(gemini_presets):
            # All these models have native vision support
            display_label = f"👁 {label} [{m}]"
            self.model_input.addItem(display_label, m)
            
        # Restore selection
        if curr_text:
            found = False
            for i in range(self.model_input.count()):
                if self.model_input.itemData(i) == curr_text:
                    self.model_input.setCurrentIndex(i)
                    found = True
                    break
            if not found:
                self.model_input.setEditText(curr_text)
        else:
            # Default to gemini-3.5-flash
            self.model_input.setCurrentIndex(0)
            
        self.model_input.blockSignals(False)

    def refresh_fewshot_templates(self):
        """Refresh the fewshot template dropdown."""
        current = self.fewshot_combo.currentData()
        self.fewshot_combo.blockSignals(True)
        self.fewshot_combo.clear()
        self.fewshot_combo.addItem("No training template", "")
        for t in self.fewshot_store.all():
            self.fewshot_combo.addItem(f"🧠 {t.name}", t.template_id)
        if current:
            idx = self.fewshot_combo.findData(current)
            if idx >= 0:
                self.fewshot_combo.setCurrentIndex(idx)
        self.fewshot_combo.blockSignals(False)

    def _apply_model_combo(self, model_combo: QComboBox) -> None:
        model = str(model_combo.currentData() or "gemma3:12b")
        self.model_input.setText(model)

    def _apply_engine_combo(self) -> None:
        engine = str(self.engine_combo.currentData() or "local")
        self.engine_name = engine
        
        # Block signals on model_combo to prevent trigger loops while repopulating
        self.model_combo.blockSignals(True)
        self.model_combo.clear()
        
        if engine == "gemini":
            self._local_model_name = self.model_input.text().strip() or self._local_model_name
            self.model_combo.setEnabled(True)
            
            gemini_combo_presets = [
                ("Gemini 3.5 Flash (Latest)", "gemini-3.5-flash"),
                ("Gemini 3.1 Flash-Lite", "gemini-3.1-flash-lite"),
                ("Gemini 3.1 Pro", "gemini-3.1-pro-preview"),
                ("Gemini 2.5 Flash", "gemini-2.5-flash"),
                ("Gemini 2.5 Pro", "gemini-2.5-pro"),
                ("Gemini 2.5 Flash-Lite", "gemini-2.5-flash-lite"),
                ("Gemini 1.5 Flash (Cheaper)", "gemini-1.5-flash"),
                ("Gemini 1.5 Pro (Standard)", "gemini-1.5-pro"),
            ]
            for label, val in gemini_combo_presets:
                self.model_combo.addItem(label, val)
                
            self.model_input.setPlaceholderText("Gemini model")
            self.refresh_gemini_models()
            self.model_input.setText(self._gemini_model_name)
            
            # Synchronize preset dropdown
            idx = -1
            for i in range(self.model_combo.count()):
                if self.model_combo.itemData(i) == self._gemini_model_name:
                    idx = i
                    break
            if idx >= 0:
                self.model_combo.setCurrentIndex(idx)
            else:
                self.model_combo.setCurrentIndex(0)
                
            self.gemini_key_label.show()
            self.gemini_key.show()
            self.status.setText(f"Cloud Gemini ready: {self._gemini_model_name}")
        else:
            self._gemini_model_name = self.model_input.text().strip() or self._gemini_model_name
            self.model_combo.setEnabled(True)
            
            # Local presets
            for label, val in VISION_MODEL_PRESETS:
                self.model_combo.addItem(label, val)
                
            self.model_input.setPlaceholderText("Ollama model tag")
            self.refresh_ollama_models()
            self.model_input.setText(self._local_model_name)
            
            # Synchronize preset dropdown
            idx = -1
            for i in range(self.model_combo.count()):
                if self.model_combo.itemData(i) == self._local_model_name:
                    idx = i
                    break
            if idx >= 0:
                self.model_combo.setCurrentIndex(idx)
            else:
                self.model_combo.setCurrentIndex(0)
                
            self.gemini_key_label.hide()
            self.gemini_key.hide()
            self.status.setText("Local Ollama ready")
            
        self.model_combo.blockSignals(False)

    def _model_text_changed(self, text: str) -> None:
        model = text.strip()
        if self.engine_name == "gemini":
            self._gemini_model_name = model or DEFAULT_GEMINI_MODEL
            self.model_name = self._gemini_model_name
        else:
            self._local_model_name = model or "gemma3:12b"
            self.model_name = self._local_model_name

    def _update_stats(self) -> None:
        remaining = max(0, self._queued_files - self._finished_files)
        self.stats.setText(
            f"Queued: {self._queued_files}    Done: {self._finished_files}    "
            f"Rows: {self._rows_extracted}    Needs review: {self._skipped_files}    Remaining: {remaining}"
        )

    def _set_default_headers(self):
        for col, name in enumerate(DEFAULT_PETITION_HEADERS):
            self.spreadsheet.raw_values[(0, col)] = name
        self.spreadsheet.recalculate()
        self.spreadsheet._mark_changed()
        ts = datetime.now().strftime('%H:%M:%S')
        self.log.append(f"[{ts}] ✅ Default petition headers set in Row 1.\n")

    def _get_headers(self) -> list[str]:
        headers = []
        for col in range(200):
            val = self.spreadsheet.raw_values.get((0, col), "").strip()
            if val and val != CLOUD_VERSION_COLUMN:
                headers.append(val)
        return headers

    def _ensure_cloud_version_column(self) -> int:
        for col in range(self.spreadsheet.table.columnCount()):
            if self.spreadsheet.raw_values.get((0, col), "").strip() == CLOUD_VERSION_COLUMN:
                self._cloud_version_col = col
                return col
        col = 0
        while col < self.spreadsheet.table.columnCount() and self.spreadsheet.raw_values.get((0, col), "").strip():
            col += 1
        while col >= self.spreadsheet.table.columnCount():
            self.spreadsheet.add_column()
        self.spreadsheet.raw_values[(0, col)] = CLOUD_VERSION_COLUMN
        self.spreadsheet.recalculate()
        self.spreadsheet._mark_changed()
        self._cloud_version_col = col
        return col

    def load_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Select Petitions", "", supported_file_filter())
        if paths:
            self._start_extraction(paths)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        paths = existing_supported(url.toLocalFile() for url in event.mimeData().urls() if url.toLocalFile())
        if paths:
            self._start_extraction(paths)

    def _start_extraction(self, paths):
        # CHANGE 8: Prevent double-start
        if self._thread is not None and self._thread.isRunning():
            QMessageBox.warning(self, "Already Running", "An extraction is already in progress. Please wait for it to finish.")
            return

        paths = existing_supported(paths)
        if not paths:
            QMessageBox.warning(self, "No supported files", "Drop or select PDF/image files.")
            return
        headers = self._get_headers()
        if not headers:
            self._set_default_headers()
            headers = DEFAULT_PETITION_HEADERS
        self._active_headers = list(headers)
        engine = str(self.engine_combo.currentData() or "local")
        model_name = self.model_input.text().strip() or (DEFAULT_GEMINI_MODEL if engine == "gemini" else self.model_name)
        gemini_key = self.gemini_key.text().strip()

        # CHANGE 3: Save Gemini key if provided
        if gemini_key:
            save_settings({'gemini_api_key': gemini_key})

        if engine == "gemini":
            if not gemini_key and not os.environ.get("GEMINI_API_KEY"):
                QMessageBox.warning(self, "Gemini API key needed", "Enter a Gemini API key or set GEMINI_API_KEY.")
                return
            self._ensure_cloud_version_column()

        # Initialize fallback structures and states
        if not getattr(self, "_in_fallback", False):
            self._fallback_paths = []
            self._paths_map = {Path(p).name: p for p in paths}
            self._file_to_rows = {}
            self._fallback_allocated_rows = {}
            self._in_fallback = False

        # CHANGE 5: Get selected fewshot template
        fewshot_tid = self.fewshot_combo.currentData()
        fewshot_template = self.fewshot_store.get(fewshot_tid) if fewshot_tid else None

        self._btn_load.setEnabled(False)
        self._queued_files = len(paths)
        self._finished_files = 0
        self._rows_extracted = 0
        self._skipped_files = 0
        self._update_stats()
        self.progress.setMaximum(len(paths))
        self.progress.setValue(0)
        self.progress.setVisible(True)
        self.status.setStyleSheet(PETITION_STATUS_STYLE)
        ts = datetime.now().strftime('%H:%M:%S')
        
        prefix = "[Fallback] 🔄 " if getattr(self, "_in_fallback", False) else ""
        self.log.append(f"\n{'─' * 40}\n[{ts}] {prefix}Extracting {len(paths)} file(s)...\n")
        label = f"Cloud Gemini ({model_name})" if engine == "gemini" else f"Local Ollama ({model_name})"
        self.log.append(f"[{ts}] Engine: {label}")
        if fewshot_template:
            self.log.append(f"[{ts}] Training template: {fewshot_template.name} ({len(fewshot_template.examples)} examples)")

        # CHANGE 8: Safe host retrieval with try/except fallback
        try:
            host = self.spreadsheet.window().chat.host.text()
        except Exception:
            host = "http://localhost:11434"

        self._thread = PetitionThread(
            paths,
            headers,
            self.current_prompt,
            model_name,
            host=host,
            engine=engine,
            gemini_api_key=gemini_key,
            fewshot_template=fewshot_template,
        )
        self._thread.progress.connect(self._on_progress)
        self._thread.row_extracted.connect(self._on_row)
        self._thread.file_finished.connect(self._on_file_finished)
        self._thread.file_warning.connect(self._on_warning)
        self._thread.error.connect(self._on_error)
        self._thread.all_finished.connect(self._on_done)
        self._thread.start()

    def _on_progress(self, msg, cur, total):
        self.status.setText(msg)
        self.progress.setValue(cur)

    def _on_row(self, data):
        headers = self._active_headers or self._get_headers()
        if not any(_is_meaningful_extracted_value(data.get(header, "")) for header in headers):
            return

        file_path = data.get("_file_path")
        allocated_row = None
        if getattr(self, "_in_fallback", False) and file_path:
            allocated_rows = getattr(self, "_fallback_allocated_rows", {}).get(file_path, [])
            if allocated_rows:
                allocated_row = allocated_rows.pop(0)

        if allocated_row is not None:
            row = allocated_row
        else:
            row = self.spreadsheet._next_empty_row()

        for col, h in enumerate(headers):
            val = str(data.get(h, ""))
            if _is_meaningful_extracted_value(val):
                self.spreadsheet.raw_values[(row, col)] = val

        if self.engine_name == "gemini" or getattr(self, "_in_fallback", False):
            cloud_col = self._cloud_version_col if self._cloud_version_col is not None else self._ensure_cloud_version_column()
            model_used = "gemini-3.1-flash-lite" if getattr(self, "_in_fallback", False) else self._gemini_model_name
            self.spreadsheet.raw_values[(row, cloud_col)] = model_used

        self.spreadsheet.recalculate()
        self.spreadsheet._mark_changed()
        self._rows_extracted += 1
        self._update_stats()
        ts = datetime.now().strftime('%H:%M:%S')
        
        prefix = "[Fallback] " if getattr(self, "_in_fallback", False) else ""
        self.log.append(f"[{ts}] ✅ {prefix}Row {row + 1} extracted")

        # Fallback tracking
        if file_path:
            if file_path not in self._file_to_rows:
                self._file_to_rows[file_path] = []
            if row not in self._file_to_rows[file_path]:
                self._file_to_rows[file_path].append(row)

            # Check if this row is unfilled or contains "no data found" / "no table found"
            is_unfilled = False
            for h in headers:
                val = str(data.get(h, "")).strip().lower()
                if "no data found" in val or "no table found" in val:
                    is_unfilled = True
                    break

            # Also check if it has very few meaningful fields
            meaningful_fields = 0
            for h in headers:
                val = data.get(h, "")
                if _is_meaningful_extracted_value(val):
                    meaningful_fields += 1
            if meaningful_fields <= 2:
                is_unfilled = True

            if self.engine_name == "local" and not getattr(self, "_in_fallback", False) and is_unfilled:
                if file_path not in self._fallback_paths:
                    self._fallback_paths.append(file_path)

    def _on_file_finished(self, file_name: str, row_count: int):
        self._finished_files += 1
        ts = datetime.now().strftime('%H:%M:%S')
        if row_count == 0:
            self._skipped_files += 1
            prefix = "[Fallback] " if getattr(self, "_in_fallback", False) else ""
            self.log.append(f"[{ts}] ⚠ {prefix}No row found: {file_name}")

            # Queue failed files for cloud fallback
            file_path = self._paths_map.get(file_name)
            if file_path:
                if self.engine_name == "local" and not getattr(self, "_in_fallback", False):
                    if file_path not in self._fallback_paths:
                        self._fallback_paths.append(file_path)

                    # Reserve placeholder row to maintain perfect ordering
                    row = self.spreadsheet._next_empty_row()
                    if file_path not in self._file_to_rows:
                        self._file_to_rows[file_path] = []
                    self._file_to_rows[file_path].append(row)

                    headers = self._active_headers or self._get_headers()
                    for col, h in enumerate(headers):
                        if col == 0:
                            val = f"[No data found in local model for {file_name}]"
                        else:
                            val = "[Pending Cloud Fallback]"
                        self.spreadsheet.raw_values[(row, col)] = val

                    self.spreadsheet.recalculate()
                    self.spreadsheet._mark_changed()
                elif getattr(self, "_in_fallback", False):
                    # Gemini failed too! Let's update the pre-allocated placeholder with a permanent notice
                    allocated_rows = getattr(self, "_fallback_allocated_rows", {}).get(file_path, [])
                    for row in allocated_rows:
                        headers = self._active_headers or self._get_headers()
                        for col, h in enumerate(headers):
                            if col == 0:
                                val = f"[No data found by Cloud Gemini for {file_name}]"
                            else:
                                val = "[Failed in both models]"
                            self.spreadsheet.raw_values[(row, col)] = val
                    self.spreadsheet.recalculate()
                    self.spreadsheet._mark_changed()
        else:
            prefix = "[Fallback] " if getattr(self, "_in_fallback", False) else ""
            self.log.append(f"[{ts}] ✓ {prefix}{file_name}: {row_count} row(s)")
        self._update_stats()

    def _on_warning(self, msg):
        ts = datetime.now().strftime('%H:%M:%S')
        self.log.append(f"[{ts}] ⚠ {msg}")

    def _on_error(self, msg):
        ts = datetime.now().strftime('%H:%M:%S')
        self.log.append(f"[{ts}] ❌ {msg}")

    def _on_done(self):
        ts = datetime.now().strftime('%H:%M:%S')
        
        # Check if cloud fallback is needed
        if self.engine_name == "local" and not getattr(self, "_in_fallback", False) and getattr(self, "_fallback_paths", None):
            gemini_key = self.gemini_key.text().strip() or os.environ.get("GEMINI_API_KEY") or load_settings().get('gemini_api_key', '')
            if not gemini_key:
                self.log.append(f"[{ts}] ⚠ Cloud fallback needed for {len(self._fallback_paths)} file(s), but no Gemini API key was found.")
                self.log.append("Please enter a Gemini key or set GEMINI_API_KEY env variable to enable auto-fallback.\n")
                
                self._btn_load.setEnabled(True)
                self.progress.setVisible(False)
                self.status.setText(f"✅ Extraction complete: {self._rows_extracted} row(s), {self._skipped_files} file(s) need review")
                self.status.setStyleSheet(PETITION_STATUS_DONE_STYLE)
                self.log.append(f"[{ts}] ✅ Done.\n{'─' * 40}\n")
                return

            self._in_fallback = True
            fallback_model = "gemini-3.1-flash-lite"

            self.log.append(f"[{ts}] 🔄 [Fallback] Auto-switching to Cloud Gemini ({fallback_model}) to redo unfilled/failed files...")
            self.log.append(f"[{ts}] [Fallback] Files to redo: {', '.join(Path(p).name for p in self._fallback_paths)}")

            # Copy all pre-allocated and written rows to fallback allocation dictionary
            self._fallback_allocated_rows = {k: list(v) for k, v in self._file_to_rows.items()}

            # Clear old unfilled rows for these files in the spreadsheet
            cleared_count = 0
            for file_path in list(self._fallback_paths):
                rows_to_clear = self._file_to_rows.get(file_path, [])
                for r in rows_to_clear:
                    for col in range(self.spreadsheet.table.columnCount()):
                        self.spreadsheet.raw_values.pop((r, col), None)
                    cleared_count += 1

            if cleared_count > 0:
                self.spreadsheet.recalculate()
                self.spreadsheet._mark_changed()
                self.log.append(f"[{ts}] [Fallback] Cleared {cleared_count} old unfilled row(s) in the spreadsheet.")

            # Run the fallback thread
            paths_to_run = list(self._fallback_paths)
            self._fallback_paths = [] # Clear to prevent recursion

            self._queued_files = len(paths_to_run)
            self._finished_files = 0
            self._rows_extracted = 0
            self._skipped_files = 0
            self._update_stats()
            self.progress.setMaximum(len(paths_to_run))
            self.progress.setValue(0)
            self.progress.setVisible(True)
            self.status.setText("Cloud fallback in progress...")

            try:
                host = self.spreadsheet.window().chat.host.text()
            except Exception:
                host = "http://localhost:11434"

            self._ensure_cloud_version_column()

            fewshot_tid = self.fewshot_combo.currentData()
            fewshot_template = self.fewshot_store.get(fewshot_tid) if fewshot_tid else None

            self._thread = PetitionThread(
                paths_to_run,
                self._active_headers,
                self.current_prompt,
                fallback_model,
                host=host,
                engine="gemini",
                gemini_api_key=gemini_key,
                fewshot_template=fewshot_template,
            )
            self._thread.progress.connect(self._on_progress)
            self._thread.row_extracted.connect(self._on_row)
            self._thread.file_finished.connect(self._on_file_finished)
            self._thread.file_warning.connect(self._on_warning)
            self._thread.error.connect(self._on_error)
            self._thread.all_finished.connect(self._on_done)
            self._thread.start()
            return

        # If we completed a fallback run, reset fallback status
        is_fallback_run = getattr(self, "_in_fallback", False)
        if is_fallback_run:
            self._in_fallback = False
            self.log.append(f"[{ts}] ✅ [Fallback] Cloud fallback run completed successfully.")

        self._btn_load.setEnabled(True)
        self.progress.setVisible(False)
        self.status.setText(f"✅ Extraction complete: {self._rows_extracted} row(s), {self._skipped_files} file(s) need review")
        self.status.setStyleSheet(PETITION_STATUS_DONE_STYLE)
        self.log.append(f"[{ts}] ✅ Done.\n{'─' * 40}\n")

    def edit_prompt(self):
        from PySide6.QtWidgets import QInputDialog
        dlg = QInputDialog(self)
        dlg.setWindowTitle("Edit Petition Prompt")
        dlg.setLabelText("Modify instructions sent to Ollama:")
        dlg.setTextValue(self.current_prompt)
        dlg.setOption(QInputDialog.UsePlainTextEditForTextInput, True)
        dlg.resize(600, 400)
        if dlg.exec() == QInputDialog.Accepted:
            self.current_prompt = dlg.textValue()


class ExpectedOutputThread(QThread):
    success = Signal(str)
    error = Signal(str)
    progress = Signal(str)

    def __init__(self, file_path, headers, engine, model_name, gemini_key, host):
        super().__init__()
        self.file_path = file_path
        self.headers = headers
        self.engine = engine
        self.model_name = model_name
        self.gemini_key = gemini_key
        self.host = host

    def run(self):
        try:
            from PIL import Image as PILImage
            import pymupdf as fitz
            import json
            from ai import GeminiClient, OllamaClient, parse_json_rows

            self.progress.emit("Reading document pages...")
            images = []
            text = ""

            # Check engine
            if self.engine == "gemini":
                client = GeminiClient(api_key=self.gemini_key, model=self.model_name or DEFAULT_GEMINI_MODEL)
                is_vision = True
            else:
                client = OllamaClient(host=self.host, vision_model=self.model_name, chat_model=self.model_name)
                is_vision = client.is_vision_model(self.model_name)

            if self.file_path.lower().endswith(".pdf"):
                doc = fitz.open(self.file_path)
                try:
                    for page in doc:
                        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                        img = PILImage.frombytes("RGB", [pix.width, pix.height], pix.samples)
                        if is_vision:
                            images.append(img)
                        else:
                            try:
                                import pytesseract as tess
                                text += tess.image_to_string(img, lang="tam+eng", config="--psm 3") + "\n\n"
                            except Exception:
                                text += tess.image_to_string(img, config="--psm 3") + "\n\n"
                finally:
                    doc.close()
            else:
                img = PILImage.open(self.file_path).convert("RGB")
                if is_vision:
                    images.append(img)
                else:
                    try:
                        import pytesseract as tess
                        text += tess.image_to_string(img, lang="tam+eng", config="--psm 3") + "\n\n"
                    except Exception:
                        text += tess.image_to_string(img, config="--psm 3") + "\n\n"

            self.progress.emit("AI extracting structured expected values...")
            
            headers_json = json.dumps({h: "" for h in self.headers}, ensure_ascii=False)
            prompt = (
                "Extract structured values from this sample expected output document.\n"
                "Understand the context and identify values for each of the requested keys/columns.\n"
                f"Use exactly these JSON keys: {headers_json}\n"
                "Preserve Tamil script exactly as Unicode Tamil.\n"
                "Return a JSON array containing a single object with the extracted keys and values. "
                "Return ONLY raw valid JSON array, with no markdown code blocks, explanation, or other text."
            )

            system = "You are a precise document extraction model. Output valid JSON array only."
            
            if self.engine == "gemini":
                schema = {
                    "type": "object",
                    "properties": {
                        "rows": {
                            "type": "array",
                            "items": {
                                "type": "object",
                                "properties": {header: {"type": "string"} for header in self.headers},
                                "required": self.headers,
                            },
                        }
                    },
                    "required": ["rows"],
                }
                res = client.generate(prompt, images=images if images else None, system=system, schema=schema)
            else:
                res = client.generate(prompt, model=self.model_name, images=images if images else None, json_mode=True, system=system)

            rows = parse_json_rows(res, self.headers)
            if rows:
                self.success.emit(json.dumps(rows, indent=2, ensure_ascii=False))
            else:
                self.success.emit(json.dumps([{h: "Not Specified" for h in self.headers}], indent=2, ensure_ascii=False))
        except Exception as e:
            self.error.emit(str(e))


# ── CHANGE 4: Training Widget ──

class TrainingWidget(QWidget):
    """AI Training Session — create few-shot examples for guided extraction."""
    templates_changed = Signal()

    def __init__(self, fewshot_store, spreadsheet, parent=None):
        super().__init__(parent)
        self.store = fewshot_store
        self.spreadsheet = spreadsheet
        self.current_template_id = uuid.uuid4().hex
        self.current_examples: list = []

        # Template management
        self.template_combo = QComboBox()
        self.template_combo.setMinimumWidth(160)
        self.template_combo.currentIndexChanged.connect(self._load_selected)
        self.template_name = QLineEdit()
        self.template_name.setPlaceholderText("Training template name...")

        # Example input
        self.example_image_path = QLabel("No image selected")
        self.example_image_path.setStyleSheet("color: #64748b; font-size: 12px; padding: 4px;")
        self.example_output = QTextEdit()
        self.example_output.setPlaceholderText('Expected JSON output for the sample...\ne.g. [{"Sender Name": "ராஜா", "Amount": "50000"}]')
        self.example_output.setMaximumHeight(200)

        # Example list
        self.example_list = QListWidget()
        self.example_list.currentRowChanged.connect(self._example_selected)

        self._current_image_path = ""

        # Layout
        header = QLabel("🧠 AI Training Session")
        header.setStyleSheet("font-size: 16px; font-weight: 800; color: #d97706; padding: 4px 0;")

        desc = QLabel("Train the AI by providing sample documents and their expected outputs.\nSave as a template and select it in the Petition Extractor.")
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #64748b; font-size: 12px; padding: 4px 0 8px 0;")

        # Template section
        tmpl_row = QHBoxLayout()
        tmpl_row.addWidget(QLabel("Template:"))
        tmpl_row.addWidget(self.template_combo, 1)

        name_row = QHBoxLayout()
        name_row.addWidget(QLabel("Name:"))
        name_row.addWidget(self.template_name, 1)

        # Template buttons
        btn_new = QPushButton("✨ New")
        btn_new.clicked.connect(self.new_template)
        style_sidebar_button(btn_new, "neutral")
        btn_save = QPushButton("💾 Save")
        btn_save.clicked.connect(self.save_template)
        style_sidebar_button(btn_save, "success")
        btn_delete = QPushButton("🗑 Delete")
        btn_delete.clicked.connect(self.delete_template)
        style_sidebar_button(btn_delete, "danger")

        tmpl_btns = QHBoxLayout()
        tmpl_btns.setSpacing(6)
        tmpl_btns.addWidget(btn_new)
        tmpl_btns.addWidget(btn_save)
        tmpl_btns.addWidget(btn_delete)

        # Example section
        example_header = QLabel("TRAINING EXAMPLES")
        example_header.setStyleSheet(SECTION_LABEL_STYLE)

        btn_load_img = QPushButton("📄 Load Sample Document")
        btn_load_img.clicked.connect(self._load_sample_image)
        style_sidebar_button(btn_load_img, "primary")

        btn_add_example = QPushButton("➕ Add Example")
        btn_add_example.clicked.connect(self._add_example)
        style_sidebar_button(btn_add_example, "structure")

        btn_remove_example = QPushButton("❌ Remove Selected")
        btn_remove_example.clicked.connect(self._remove_example)
        style_sidebar_button(btn_remove_example, "danger")

        example_btns = QHBoxLayout()
        example_btns.addWidget(btn_add_example)
        example_btns.addWidget(btn_remove_example)

        layout = QVBoxLayout(self)
        layout.setSpacing(6)
        layout.addWidget(header)
        layout.addWidget(desc)
        layout.addLayout(tmpl_row)
        layout.addLayout(name_row)
        layout.addLayout(tmpl_btns)

        # Separator
        sep = QWidget()
        sep.setFixedHeight(1)
        sep.setStyleSheet("background: #e2e8f0;")
        layout.addWidget(sep)

        layout.addWidget(example_header)
        layout.addWidget(btn_load_img)
        layout.addWidget(self.example_image_path)

        output_label_row = QHBoxLayout()
        output_label_row.addWidget(QLabel("Expected Output (JSON):"))
        btn_load_output_pdf = QPushButton("📄 Load Expected Output PDF/Image")
        btn_load_output_pdf.clicked.connect(self._load_expected_output_pdf)
        style_sidebar_button(btn_load_output_pdf, "primary")
        output_label_row.addWidget(btn_load_output_pdf)

        layout.addLayout(output_label_row)
        layout.addWidget(self.example_output)
        layout.addLayout(example_btns)
        layout.addWidget(QLabel("Saved Examples:"))
        layout.addWidget(self.example_list, 1)

        self.refresh_templates()
        self.new_template()

    def refresh_templates(self):
        selected = self.current_template_id
        self.template_combo.blockSignals(True)
        self.template_combo.clear()
        self.template_combo.addItem("Select training template...", "")
        for t in self.store.all():
            self.template_combo.addItem(t.name, t.template_id)
        idx = self.template_combo.findData(selected)
        self.template_combo.setCurrentIndex(idx if idx >= 0 else 0)
        self.template_combo.blockSignals(False)

    def new_template(self):
        self.current_template_id = uuid.uuid4().hex
        self.template_combo.setCurrentIndex(0)
        self.template_name.clear()
        self.current_examples = []
        self.example_list.clear()
        self._current_image_path = ""
        self.example_image_path.setText("No image selected")
        self.example_output.clear()

    def _load_selected(self):
        tid = self.template_combo.currentData()
        if not tid:
            return
        t = self.store.get(tid)
        if not t:
            return
        self.current_template_id = t.template_id
        self.template_name.setText(t.name)
        self.current_examples = [FewShotExample(image_path=e.image_path, output_json=e.output_json) for e in t.examples]
        self._render_examples()

    def _load_sample_image(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Sample Document", "", supported_file_filter())
        if path:
            self._current_image_path = path
            self.example_image_path.setText(f"📎 {Path(path).name}")
            self.example_image_path.setToolTip(path)

    def _add_example(self):
        output = self.example_output.toPlainText().strip()
        if not output:
            QMessageBox.warning(self, "Missing Output", "Enter the expected JSON output for this example.")
            return
        if not self._current_image_path:
            QMessageBox.warning(self, "Missing Input", "Load a sample document image first.")
            return
        self.current_examples.append(FewShotExample(
            image_path=self._current_image_path,
            output_json=output
        ))
        self._current_image_path = ""
        self.example_image_path.setText("No image selected")
        self.example_output.clear()
        self._render_examples()

    def _remove_example(self):
        row = self.example_list.currentRow()
        if 0 <= row < len(self.current_examples):
            self.current_examples.pop(row)
            self._render_examples()

    def _example_selected(self, row):
        if 0 <= row < len(self.current_examples):
            ex = self.current_examples[row]
            self.example_output.setPlainText(ex.output_json)
            self.example_image_path.setText(f"📎 {Path(ex.image_path).name}" if ex.image_path else "No image")

    def _render_examples(self):
        self.example_list.clear()
        for i, ex in enumerate(self.current_examples, 1):
            name = Path(ex.image_path).name if ex.image_path else "No image"
            preview = ex.output_json[:60].replace('\n', ' ') + ('...' if len(ex.output_json) > 60 else '')
            self.example_list.addItem(f"Example {i}: {name} → {preview}")

    def save_template(self):
        name = self.template_name.text().strip()
        if not name:
            QMessageBox.warning(self, "Missing Name", "Enter a template name.")
            return
        if not self.current_examples:
            QMessageBox.warning(self, "No Examples", "Add at least one training example.")
            return
        t = FewShotTemplate(name=name, examples=list(self.current_examples), template_id=self.current_template_id)
        self.store.upsert(t)
        self.refresh_templates()
        self.templates_changed.emit()
        QMessageBox.information(self, "Saved", f"Training template '{name}' saved with {len(self.current_examples)} example(s).")

    def delete_template(self):
        if not self.current_template_id:
            return
        self.store.delete(self.current_template_id)
        self.new_template()
        self.refresh_templates()
        self.templates_changed.emit()

    def _get_headers(self) -> list[str]:
        headers = []
        for col in range(200):
            val = self.spreadsheet.raw_values.get((0, col), "").strip()
            if val and val != CLOUD_VERSION_COLUMN:
                headers.append(val)
        return headers

    def _load_expected_output_pdf(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Expected Output PDF/Image", "", supported_file_filter())
        if not path:
            return

        headers = self._get_headers()
        if not headers:
            headers = DEFAULT_PETITION_HEADERS

        # Safely extract LLM settings from the active petition widget if available
        engine = "local"
        model_name = "minicpm-v"
        gemini_key = ""
        try:
            petition = self.spreadsheet.window().petition
            engine = str(petition.engine_combo.currentData() or "local")
            model_name = petition.model_input.text().strip() or (DEFAULT_GEMINI_MODEL if engine == "gemini" else petition.model_name)
            gemini_key = petition.gemini_key.text().strip()
        except Exception:
            pass

        if engine == "gemini":
            if not gemini_key and not os.environ.get("GEMINI_API_KEY"):
                QMessageBox.warning(self, "Gemini API key needed", "Please configure the Gemini API key in the Petition Extractor tab first.")
                return

        try:
            host = self.spreadsheet.window().chat.host.text()
        except Exception:
            host = "http://localhost:11434"

        # Show a progress dialog
        self.progress_dialog = QProgressDialog("Initializing expected output loader...", "Cancel", 0, 0, self)
        self.progress_dialog.setWindowTitle("Extracting Expected Output")
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setAutoClose(True)
        
        # Style progress dialog to match premium look
        self.progress_dialog.setStyleSheet(
            "QProgressDialog { background-color: #ffffff; color: #1e293b; font-family: 'Segoe UI'; font-size: 13px; }"
            "QLabel { color: #1e293b; font-weight: 500; }"
            "QPushButton { background: #f1f5f9; color: #475569; border: 1px solid #cbd5e1; border-radius: 6px; padding: 4px 8px; }"
            "QPushButton:hover { background: #e2e8f0; }"
        )
        self.progress_dialog.show()

        # Launch the extraction thread
        self._output_thread = ExpectedOutputThread(
            path,
            headers,
            engine,
            model_name,
            gemini_key,
            host
        )
        self._output_thread.progress.connect(self._on_output_progress)
        self._output_thread.success.connect(self._on_output_success)
        self._output_thread.error.connect(self._on_output_error)
        
        self.progress_dialog.canceled.connect(self._output_thread.terminate)
        self._output_thread.start()

    def _on_output_progress(self, msg: str):
        if hasattr(self, 'progress_dialog') and self.progress_dialog.isVisible():
            self.progress_dialog.setLabelText(msg)

    def _on_output_success(self, json_data: str):
        if hasattr(self, 'progress_dialog'):
            self.progress_dialog.close()
        self.example_output.setPlainText(json_data)
        try:
            self.spreadsheet.window().statusBar().showMessage("✅ Expected output loaded successfully from document!", 5000)
        except Exception:
            pass

    def _on_output_error(self, err_msg: str):
        if hasattr(self, 'progress_dialog'):
            self.progress_dialog.close()
        QMessageBox.critical(self, "Extraction Error", f"Failed to extract expected output:\n{err_msg}")


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("KAAVAL — Knowledge & Advanced Analysis for Vigilance & Asset Loss")
        screen = self.screen().availableGeometry()
        self.resize(min(1500, screen.width() - 32), min(920, screen.height() - 64))

        self.store = TemplateStore()
        self.fewshot_store = FewShotStore()

        from rag import RAGDatabase
        ollama_client = OllamaClient()
        db_path = APP_DIR / "vector_db"
        self.rag = RAGDatabase(db_path, ollama_client)

        self.spreadsheet = SpreadsheetWidget()
        self.templates = TemplateManagerWidget(self.store)
        self.extraction = ExtractionWidget(self.store, self.spreadsheet)
        self.petition = PetitionWidget(self.spreadsheet, self.fewshot_store)
        self.sources = SourcesWidget(self.rag)
        self.chat = ChatPanel(self.spreadsheet, self.rag)
        self.training = TrainingWidget(self.fewshot_store, self.spreadsheet)

        self.templates.templates_changed.connect(self.extraction.refresh_templates)
        # When training templates change, refresh petition's fewshot dropdown
        self.training.templates_changed.connect(self.petition.refresh_fewshot_templates)

        # CHANGE 2: Right side — tool panels in a stacked widget, petition first
        self.tool_stack = QStackedWidget()
        self.tool_panels = {}
        for name, widget, label in [
            ("petition", self.petition, "🛡 Petition Extractor"),
            ("chat", self.chat, "🤖 AI Chat"),
            ("training", self.training, "🧠 AI Training"),
            ("vision", self.extraction, "📥 Vision Import"),
            ("template", self.templates, "📋 Templates"),
            ("notebook", self.sources, "📚 Notebook"),
        ]:
            pane = QScrollArea()
            pane.setWidgetResizable(True)
            pane.setFrameShape(QScrollArea.Shape.NoFrame)
            pane.setWidget(widget)
            self.tool_stack.addWidget(pane)
            self.tool_panels[name] = (pane, label)

        # Side panel header with back button and title — default to Petition
        self.panel_title = QLabel("🛡 Petition Extractor")
        self.panel_title.setStyleSheet(SIDEBAR_HEADER_STYLE)
        self.panel_back_btn = QPushButton("← Petition")
        self.panel_back_btn.setFixedWidth(90)
        self.panel_back_btn.clicked.connect(lambda: self.show_tool_panel("petition"))
        self.panel_back_btn.setStyleSheet(
            "QPushButton { background: #f1f5f9; color: #1e40af; border: 1px solid #e2e8f0;"
            "border-radius: 6px; padding: 4px 8px; font-size: 11px; font-weight: 600; }"
            "QPushButton:hover { background: #dbeafe; }"
        )
        self.panel_back_btn.hide()

        panel_header = QHBoxLayout()
        panel_header.setContentsMargins(0, 0, 0, 0)
        panel_header.addWidget(self.panel_back_btn)
        panel_header.addWidget(self.panel_title, 1)

        # CHANGE 7: wider right panel
        right_panel = self.right_panel = QWidget()
        right_panel.setMinimumWidth(300)
        self.spreadsheet.setMinimumWidth(420)
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(4)
        right_layout.addLayout(panel_header)
        right_layout.addWidget(self.tool_stack, 1)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.addWidget(self.spreadsheet)
        splitter.addWidget(right_panel)
        splitter.setSizes([1100, 400])
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 0)
        splitter.setHandleWidth(6)
        splitter.setChildrenCollapsible(False)

        self.setCentralWidget(splitter)
        self.spreadsheet.load_autosave_if_exists()

        # CHANGE 2: Set petition as default panel
        self.show_tool_panel("petition")

        # Status Bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("KAAVAL v1.0 — Tamil Nadu Police EOW")
        self.ollama_status = QLabel("● Checking...")
        self.ollama_status.setStyleSheet("color: #64748b; font-size: 12px; padding: 0 8px;")
        self.row_count_label = QLabel("")
        self.row_count_label.setStyleSheet("color: #475569; font-size: 12px;")
        self.status_bar.addPermanentWidget(self.row_count_label)
        self.status_bar.addPermanentWidget(self.ollama_status)
        self._ollama_connected = False
        self._check_ollama_connection()

        self.autosave_timer = QTimer(self)
        self.autosave_timer.setInterval(10000)
        self.autosave_timer.timeout.connect(self.spreadsheet.autosave)
        self.autosave_timer.start()

        self.connection_timer = QTimer(self)
        self.connection_timer.setInterval(30000)
        self.connection_timer.timeout.connect(self._check_ollama_connection)
        self.connection_timer.start()

        self.spreadsheet.data_changed.connect(self._update_row_count)
        self._update_row_count()

    def show_tool_panel(self, name: str) -> None:
        if name in self.tool_panels:
            self.right_panel.show()
            widget, label = self.tool_panels[name]
            self.tool_stack.setCurrentWidget(widget)
            self.panel_title.setText(label)
            # CHANGE 7: back button visibility — hide when on petition (default)
            self.panel_back_btn.setVisible(name != "petition")

    def toggle_tool_panel(self) -> None:
        self.right_panel.setVisible(self.right_panel.isHidden())

    def _check_ollama_connection(self) -> None:
        client = OllamaClient()
        ok, _ = client.check_connection()
        was_connected = getattr(self, "_ollama_connected", False)
        self._ollama_connected = ok
        if ok:
            self.ollama_status.setText("● Ollama Connected")
            self.ollama_status.setStyleSheet("color: #059669; font-size: 12px; font-weight: 600; padding: 0 8px;")
            if not was_connected:
                self.petition.refresh_ollama_models()
        else:
            self.ollama_status.setText("○ Ollama Offline")
            self.ollama_status.setStyleSheet("color: #dc2626; font-size: 12px; font-weight: 600; padding: 0 8px;")

    def _update_row_count(self) -> None:
        count = sum(1 for r in range(self.spreadsheet.table.rowCount())
                    if any(self.spreadsheet.raw_values.get((r, c)) for c in range(self.spreadsheet.table.columnCount())))
        self.row_count_label.setText(f"{count} rows")

    def closeEvent(self, event) -> None:
        # CHANGE 8: Stop any running PetitionThread
        if self.petition._thread is not None and self.petition._thread.isRunning():
            self.petition._thread.quit()
            self.petition._thread.wait(3000)
        self.spreadsheet.autosave()
        super().closeEvent(event)


# style_tool_button is now imported from theme.py


def item_has_custom_style(item: QTableWidgetItem | None) -> bool:
    if item is None:
        return False
    font = item.font()
    return (
        font.bold()
        or font.italic()
        or font.underline()
        or (font.pointSize() > 0 and font.pointSize() != DEFAULT_CELL_FONT_SIZE)
        or font.family() != "Segoe UI"
        or item.foreground().style() != Qt.BrushStyle.NoBrush
        or item.background().style() != Qt.BrushStyle.NoBrush
        or item.textAlignment() != 0
    )


def item_to_style(item: QTableWidgetItem | None) -> dict[str, object]:
    if item is None:
        return {}
    font = item.font()
    foreground = item.foreground()
    background = item.background()
    return {
        "bold": font.bold(),
        "italic": font.italic(),
        "underline": font.underline(),
        "family": font.family(),
        "point_size": font.pointSize(),
        "foreground": qcolor_to_rgb(foreground.color()) if foreground.style() != Qt.BrushStyle.NoBrush else None,
        "background": qcolor_to_rgb(background.color()) if background.style() != Qt.BrushStyle.NoBrush else None,
        "alignment": int(item.textAlignment()),
    }


def apply_item_style(item: QTableWidgetItem, style: object) -> None:
    if not isinstance(style, dict):
        return
    font = item.font()
    if style.get("family"):
        font.setFamily(str(style.get("family")))
    size = style.get("point_size")
    if isinstance(size, int) and 6 <= size <= 144:
        font.setPointSize(size)
    font.setBold(bool(style.get("bold")))
    font.setItalic(bool(style.get("italic")))
    font.setUnderline(bool(style.get("underline")))
    item.setFont(font)
    foreground = style.get("foreground")
    background = style.get("background")
    if isinstance(foreground, str) and foreground:
        item.setForeground(QColor(f"#{foreground}"))
    if isinstance(background, str) and background:
        item.setBackground(QColor(f"#{background}"))
    try:
        alignment = int(style.get("alignment", 0))
    except (TypeError, ValueError):
        alignment = 0
    if alignment:
        try:
            item.setTextAlignment(Qt.AlignmentFlag(alignment))
        except ValueError:
            pass


TOTAL_KEYWORDS = ("total", "invested", "investment", "borrowed", "principal", "amount paid", "deposited", "செலுத்த")
RETURNED_KEYWORDS = ("returned", "return", "received", "receipt", "paid back", "பெற்ற")
BALANCE_KEYWORDS = ("balance", "remaining", "reminding", "pending", "receivable", "வேண்டிய", "மீத")


def money_check(values: dict[str, str]) -> str:
    total_key = find_money_key(values, TOTAL_KEYWORDS, RETURNED_KEYWORDS + BALANCE_KEYWORDS)
    returned_key = find_money_key(values, RETURNED_KEYWORDS, BALANCE_KEYWORDS)
    balance_key = find_money_key(values, BALANCE_KEYWORDS, RETURNED_KEYWORDS)
    if not total_key or not (returned_key or balance_key):
        return ""
    total = parse_money(values.get(total_key, ""))
    returned = parse_money(values.get(returned_key, "")) if returned_key else 0.0
    balance = parse_money(values.get(balance_key, "")) if balance_key else 0.0
    if total is None:
        return ""
    returned = returned or 0.0
    balance = balance or 0.0
    difference = total - returned - balance
    if abs(difference) <= 1:
        return "OK"
    return f"Mismatch: diff {format_money(difference)}"


def find_money_key(values: dict[str, str], keywords: tuple[str, ...], exclude: tuple[str, ...] = ()) -> str | None:
    for key in values:
        lowered = key.lower()
        if any(word in lowered for word in exclude):
            continue
        if any(word in lowered for word in keywords):
            return key
    return None


def parse_money(value: str) -> float | None:
    text = str(value).replace(",", "").replace("₹", "").strip()
    if not text or text in {"-", "--"}:
        return None
    total = 0.0
    found = False
    for line in text.splitlines() or [text]:
        numbers = re.findall(r"-?\d+(?:\.\d+)?", line)
        if not numbers:
            continue
        found = True
        if "=" in line:
            total += float(numbers[-1])
        elif len(numbers) == 1:
            total += float(numbers[0])
        else:
            total += float(numbers[-1])
    return total if found else None


def format_money(value: float) -> str:
    if abs(value - round(value)) < 0.01:
        return str(int(round(value)))
    return f"{value:.2f}"


def qcolor_to_rgb(color: QColor) -> str:
    if not color.isValid():
        return "000000"
    return f"{color.red():02X}{color.green():02X}{color.blue():02X}"


def qt_alignment_to_openpyxl(alignment: int) -> str:
    value = int(alignment)
    if value & int(Qt.AlignmentFlag.AlignRight):
        return "right"
    if value & int(Qt.AlignmentFlag.AlignHCenter) or value & int(Qt.AlignmentFlag.AlignCenter):
        return "center"
    return "left"


def csv_escape(value: str) -> str:
    if any(char in value for char in [",", '"', "\n"]):
        return '"' + value.replace('"', '""') + '"'
    return value


def escape_html(value: str) -> str:
    return value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br>")


def main() -> None:
    app = QApplication(sys.argv)
    app.setStyleSheet(LIGHT_STYLESHEET)
    # Set app icon
    from PySide6.QtGui import QIcon
    icon_path = Path(__file__).parent / "kaaval.ico"
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
