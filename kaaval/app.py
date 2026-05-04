from __future__ import annotations

import csv
import json
import re
import sys
import uuid
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from PySide6.QtCore import QTimer, Qt, Signal, QThread, QMimeData
from PySide6.QtGui import QColor, QFont, QShortcut, QKeySequence, QDragEnterEvent, QDropEvent, QPen, QPainter
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QColorDialog,
    QComboBox,
    QFileDialog,
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
    QPushButton,
    QSplitter,
    QStackedWidget,
    QStatusBar,
    QStyle,
    QStyledItemDelegate,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextBrowser,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from ai import OllamaClient, build_chat_prompt, extraction_summary, ocr_crop, tesseract_status
from document import crop_normalized, page_count, render_page, supported_file_filter
from formula import FormulaEngine, column_name
from templates import APP_DIR, ExtractionTemplate, TemplateField, TemplateStore
from widgets import ImageSelectionCanvas
from theme import LIGHT_STYLESHEET, style_tool_button, CHAT_CSS

try:
    import pytesseract  # type: ignore
    import os
    tess_path = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    if os.path.exists(tess_path):
        pytesseract.pytesseract.tesseract_cmd = tess_path
except ImportError:
    pass

NATIVE_FILE_FILTER = "Local AI Sheet (*.laisheet)"
AUTOSAVE_FILE = APP_DIR / "autosave.laisheet"
MONEY_CHECK_COLUMN = "Money Check"

VISION_MODEL_PRESETS = [
    ("Gemma 3 12B - Tamil vision", "gemma3:12b"),
    ("Gemma 4 E2B - lighter", "gemma4:e2b"),
    ("Gemma 4 26B - stronger", "gemma4:26b"),
    ("Gemma 4 31B - strongest", "gemma4:31b"),
    ("Qwen 2.5 VL 7B", "qwen2.5vl:7b"),
]


class SpreadsheetDelegate(QStyledItemDelegate):
    def paint(self, painter: QPainter, option: 'QStyleOptionViewItem', index) -> None:
        painter.save()
        
        # 1. Let Qt draw the standard background, fill color, and text
        from PySide6.QtWidgets import QStyleOptionViewItem
        opt = QStyleOptionViewItem(option)
        is_selected = opt.state & QStyle.StateFlag.State_Selected
        
        # Prevent default blue selection styling by masking out the selected state
        if is_selected:
            opt.state &= ~QStyle.StateFlag.State_Selected

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
        formula_bar.addWidget(self.formula_input)
        formula_bar.addWidget(self.status_label)

        layout = QVBoxLayout(self)
        layout.setSpacing(4)
        layout.addWidget(toolbar)
        layout.addLayout(formula_bar)
        layout.addWidget(self.search_bar)
        layout.addWidget(self.table)

        self.engine = FormulaEngine(self._raw_value)
        
        self.history = [self.to_state()]
        self.history_index = 0

    def _build_toolbar(self) -> QWidget:
        toolbar = QWidget()
        toolbar.setStyleSheet("QWidget { background: qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #ffffff, stop:1 #f8fafc); border: 1px solid #e2e8f0; border-radius: 10px; padding: 3px; }")
        
        main_layout = QVBoxLayout(toolbar)
        main_layout.setContentsMargins(4, 4, 4, 4)
        main_layout.setSpacing(4)
        
        row1_layout = QHBoxLayout()
        row1_layout.setContentsMargins(0, 0, 0, 0)
        row1_layout.setSpacing(3)
        
        row2_layout = QHBoxLayout()
        row2_layout.setContentsMargins(0, 0, 0, 0)
        row2_layout.setSpacing(3)

        actions_row1 = [
            ("📂 Open", self.open_native, "neutral", "Ctrl+O"),
            ("💾 Save", self.save_native, "success", "Ctrl+S"),
            ("📄 Save As", self.save_native_as, "success", "Ctrl+Shift+S"),
            ("📊 Export", self.export_xlsx, "export", "Ctrl+E"),
            (None, None, None, None),
            ("↩ Undo", self.undo, "neutral", "Ctrl+Z"),
            ("↪ Redo", self.redo, "neutral", "Ctrl+Y"),
            ("🗑 Clear Sheet", self.clear_sheet, "danger", None),
        ]
        
        actions_row2 = [
            ("𝐁", self.apply_bold, "format", "Ctrl+B"),
            ("𝐼", self.apply_italic, "format", "Ctrl+I"),
            ("U̲", self.apply_underline, "format", "Ctrl+U"),
            ("🎨 Txt", self.apply_text_color, "format", None),
            ("🪣 Fill", self.apply_fill_color, "format", None),
            ("🔲 Borders", self.apply_borders, "format", None),
            (None, None, None, None),
            ("◀", lambda: self.apply_alignment(Qt.AlignmentFlag.AlignLeft), "align", None),
            ("◆", lambda: self.apply_alignment(Qt.AlignmentFlag.AlignCenter), "align", None),
            ("▶", lambda: self.apply_alignment(Qt.AlignmentFlag.AlignRight), "align", None),
            ("↩ Wrap", self.toggle_word_wrap, "align", None),
            (None, None, None, None),
            ("⊞ Merge", self.merge_selection, "merge", None),
            ("⊟ Unmerge", self.unmerge_selection, "danger", None),
            (None, None, None, None),
            ("↕ Fit Rows", self.auto_fit_rows, "structure", None),
            ("↔ Fit Cols", self.auto_fit_columns, "structure", None),
            ("⬆➕ Row", self.insert_row_above, "structure", None),
            ("⬅➕ Col", self.insert_column_left, "structure", None),
            ("❌ Row", self.delete_row, "danger", None),
            ("❌ Col", self.delete_column, "danger", None),
        ]
        
        for layout, actions in [(row1_layout, actions_row1), (row2_layout, actions_row2)]:
            for label, handler, tone, shortcut in actions:
                if label is None:
                    sep = QWidget()
                    sep.setFixedWidth(1)
                    sep.setMinimumHeight(24)
                    sep.setStyleSheet("background: #cbd5e1;")
                    layout.addWidget(sep)
                    continue
                button = QPushButton(label)
                button.clicked.connect(handler)
                if shortcut:
                    button.setToolTip(f"{label}  ({shortcut})")
                style_tool_button(button, tone)
                layout.addWidget(button)

        row1_layout.addStretch(1)
        row2_layout.addStretch(1)

        # Tools dropdown button on Row 1
        tools_btn = QPushButton("⚙ Tools ▼")
        style_tool_button(tools_btn, "primary")
        tools_menu = QMenu(self)
        tools_menu.addAction("🛡 Petition Extractor", lambda: self._show_tool_panel("petition"))
        tools_menu.addAction("📥 Vision Import", lambda: self._show_tool_panel("vision"))
        tools_menu.addAction("📋 Template Editor", lambda: self._show_tool_panel("template"))
        tools_menu.addAction("📚 Notebook / RAG", lambda: self._show_tool_panel("notebook"))
        tools_btn.setMenu(tools_menu)
        row1_layout.addWidget(tools_btn)
        
        main_layout.addLayout(row1_layout)
        main_layout.addLayout(row2_layout)

        return toolbar

    def _raw_value(self, row: int, col: int) -> str:
        return self.raw_values.get((row, col), "")

    def _ensure_item(self, row: int, col: int) -> QTableWidgetItem:
        item = self.table.item(row, col)
        if item is None:
            item = QTableWidgetItem("")
            self.table.setItem(row, col, item)
        return item

    def _cell_changed(self, row: int, col: int) -> None:
        if self._updating:
            return
        item = self.table.item(row, col)
        text = item.text() if item else ""
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
        for item in self.selected_items_or_cells():
            font = item.font()
            font.setBold(not font.bold())
            item.setFont(font)
        self._mark_changed()

    def apply_italic(self) -> None:
        for item in self.selected_items_or_cells():
            font = item.font()
            font.setItalic(not font.italic())
            item.setFont(font)
        self._mark_changed()

    def apply_underline(self) -> None:
        for item in self.selected_items_or_cells():
            font = item.font()
            font.setUnderline(not font.underline())
            item.setFont(font)
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
        if not rows:
            return self._next_empty_row(), 0
        checks = [money_check(row) for row in rows]
        include_money_check = any(check for check in checks)
        self.set_template_columns(template, include_money_check)
        first_row = self._next_empty_row()
        current_row = first_row
        fields = template.field_names()
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
        output = []
        headers = [self.table.horizontalHeaderItem(col).text() if self.table.horizontalHeaderItem(col) else column_name(col) for col in range(self.table.columnCount())]
        active_cols = self._active_columns()
        if not active_cols:
            return ""
        writer_buffer: list[str] = []
        writer_buffer.append(",".join(csv_escape(headers[col]) for col in active_cols))
        rows_added = 0
        for row in range(self.table.rowCount()):
            values = [self._display_value(row, col) for col in active_cols]
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
            sheet.cell(row=1, column=out_col, value=header.text() if header else column_name(col))

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
        self.vision_model = QLineEdit("gemma3:12b")
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
    "You are an expert data extractor. I will provide you with a transcription of Tamil legal/financial petitions. "
    "This text may represent one single petition spanning multiple pages, or multiple different petitions.\n"
    "The text is from an OCR engine, expect typos and broken spacing.\n"
    "Pay extreme attention to Crime Numbers (e.g., Cr.No, FIR No, Crime No) and Scheme Names/Company Names. "
    "Use logical deduction to reconstruct fragmented Tamil or English names.\n"
    "IMPORTANT: For fields like 'ID Number', 'Plan Name', or 'Scheme Name', if they appear in English in the "
    "original text, preserve the exact English text. Do NOT translate them into Tamil."
)


class PetitionThread(QThread):
    row_extracted = Signal(dict)
    all_finished = Signal()
    error = Signal(str)
    progress = Signal(str, int, int)

    def __init__(self, file_paths, headers, instruction, model_name):
        super().__init__()
        self.file_paths = file_paths
        self.headers = headers
        self.instruction = instruction
        self.model_name = model_name

    def run(self):
        import json as _json
        try:
            import pytesseract
        except ImportError:
            self.error.emit("pytesseract is not installed.")
            self.all_finished.emit()
            return
        try:
            self.progress.emit("OCR processing...", 0, len(self.file_paths))
            combined_text = ""
            total = len(self.file_paths)
            for i, file_path in enumerate(self.file_paths, 1):
                base_name = Path(file_path).name
                self.progress.emit(f"[{i}/{total}] OCR: {base_name}", i, total)
                if file_path.lower().endswith('.pdf'):
                    import fitz
                    from PIL import Image as PILImage
                    doc = fitz.open(file_path)
                    for page in doc:
                        # Render page to image with 2x zoom (DPI=144 is 2x zoom if default is 72)
                        # fitz.Matrix(2, 2) is clearer for 2x zoom
                        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                        img = PILImage.frombytes("RGB", [pix.width, pix.height], pix.samples)
                        combined_text += pytesseract.image_to_string(img, lang='tam+eng') + "\n\n"
                    doc.close()
                else:
                    combined_text += pytesseract.image_to_string(file_path, lang='tam+eng') + "\n\n"
            if not combined_text.strip():
                self.error.emit("No text extracted from files.")
                self.all_finished.emit()
                return
            self.progress.emit("Analyzing with LLM...", total, total)
            headers_json = _json.dumps({h: "" for h in self.headers}, indent=2, ensure_ascii=False)
            prompt = f"""{self.instruction}\n\nExtract and output ONLY a raw JSON array of objects with these keys for EACH petition:\n[\n{headers_json}\n]\n\nText:\n{combined_text}"""
            import urllib.request
            payload = _json.dumps({"model": self.model_name, "messages": [{"role": "user", "content": prompt}], "stream": False, "options": {"temperature": 0}}).encode()
            req = urllib.request.Request("http://localhost:11434/api/chat", data=payload, headers={"Content-Type": "application/json"})
            with urllib.request.urlopen(req, timeout=300) as resp:
                data = _json.loads(resp.read().decode())
            output_text = data.get("message", {}).get("content", "")
            if "```json" in output_text:
                output_text = output_text.split("```json")[1].split("```")[0].strip()
            elif "```" in output_text:
                output_text = output_text.split("```")[1].split("```")[0].strip()
            parsed = _json.loads(output_text)
            if isinstance(parsed, list):
                for petition in parsed:
                    self.row_extracted.emit(petition)
            elif isinstance(parsed, dict):
                self.row_extracted.emit(parsed)
        except Exception as e:
            self.error.emit(str(e))
        self.all_finished.emit()


class PetitionWidget(QWidget):
    """Tamil Petition OCR+LLM extraction — merged from EOWassist."""

    def __init__(self, spreadsheet: SpreadsheetWidget, parent=None):
        super().__init__(parent)
        self.spreadsheet = spreadsheet
        self.model_name = "gemma3:12b"
        self.current_prompt = DEFAULT_PETITION_PROMPT
        self._thread = None
        self.setAcceptDrops(True)

        # Controls
        model_combo = QComboBox()
        for label, val in VISION_MODEL_PRESETS:
            model_combo.addItem(label, val)
        model_combo.currentIndexChanged.connect(lambda: setattr(self, 'model_name', model_combo.currentData() or 'gemma3:12b'))

        btn_load = QPushButton("📄 Load Petitions")
        btn_load.clicked.connect(self.load_files)
        style_tool_button(btn_load, "primary")
        self._btn_load = btn_load

        btn_prompt = QPushButton("✍ Edit Prompt")
        btn_prompt.clicked.connect(self.edit_prompt)
        style_tool_button(btn_prompt, "neutral")

        btn_set_headers = QPushButton("📋 Set Default Headers")
        btn_set_headers.clicked.connect(self._set_default_headers)
        style_tool_button(btn_set_headers, "structure")

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.progress.setTextVisible(True)

        self.status = QLabel("✦ Drop PDF/image files or click Load")
        self.status.setStyleSheet("color: #1e40af; font-weight: 600; font-size: 13px; padding: 4px;")

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setPlaceholderText("Extraction log...")

        controls = QHBoxLayout()
        controls.addWidget(btn_load)
        controls.addWidget(btn_set_headers)
        controls.addWidget(btn_prompt)
        controls.addWidget(QLabel("Model:"))
        controls.addWidget(model_combo, 1)

        layout = QVBoxLayout(self)
        layout.addLayout(controls)
        layout.addWidget(self.status)
        layout.addWidget(self.progress)
        layout.addWidget(self.log, 1)

    def _set_default_headers(self):
        for col, name in enumerate(DEFAULT_PETITION_HEADERS):
            self.spreadsheet.raw_values[(0, col)] = name
        self.spreadsheet.recalculate()
        self.spreadsheet._mark_changed()
        self.log.append("✅ Default petition headers set in Row 1.\n")

    def _get_headers(self) -> list[str]:
        headers = []
        for col in range(200):
            val = self.spreadsheet.raw_values.get((0, col), "").strip()
            if val:
                headers.append(val)
        return headers

    def load_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Select Petitions", "", "Documents (*.pdf *.png *.jpg *.jpeg)")
        if paths:
            self._start_extraction(paths)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        paths = [url.toLocalFile() for url in event.mimeData().urls() if url.toLocalFile()]
        if paths:
            self._start_extraction(paths)

    def _start_extraction(self, paths):
        headers = self._get_headers()
        if not headers:
            self._set_default_headers()
            headers = DEFAULT_PETITION_HEADERS
        self._btn_load.setEnabled(False)
        self.progress.setMaximum(len(paths))
        self.progress.setValue(0)
        self.progress.setVisible(True)
        self.log.append(f"\n{'─' * 40}\n🔄 Extracting {len(paths)} file(s)...\n")
        self._thread = PetitionThread(paths, headers, self.current_prompt, self.model_name)
        self._thread.progress.connect(self._on_progress)
        self._thread.row_extracted.connect(self._on_row)
        self._thread.error.connect(self._on_error)
        self._thread.all_finished.connect(self._on_done)
        self._thread.start()

    def _on_progress(self, msg, cur, total):
        self.status.setText(msg)
        self.progress.setValue(cur)

    def _on_row(self, data):
        headers = self._get_headers()
        row = self.spreadsheet._next_empty_row()
        for col, h in enumerate(headers):
            val = str(data.get(h, ""))
            if val and val != "Not Specified":
                self.spreadsheet.raw_values[(row, col)] = val
        self.spreadsheet.recalculate()
        self.spreadsheet._mark_changed()
        self.log.append(f"✅ Row {row + 1} extracted")

    def _on_error(self, msg):
        self.log.append(f"❌ {msg}")

    def _on_done(self):
        self._btn_load.setEnabled(True)
        self.progress.setVisible(False)
        self.status.setText("✅ Extraction complete")
        self.status.setStyleSheet("color: #059669; font-weight: 600; font-size: 13px; padding: 4px;")
        self.log.append(f"✅ Done.\n{'─' * 40}\n")

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


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("KAAVAL — Knowledge & Advanced Analysis for Vigilance & Asset Loss")
        self.resize(1500, 920)

        self.store = TemplateStore()

        from rag import RAGDatabase
        ollama_client = OllamaClient()
        db_path = APP_DIR / "vector_db"
        self.rag = RAGDatabase(db_path, ollama_client)

        self.spreadsheet = SpreadsheetWidget()
        self.templates = TemplateManagerWidget(self.store)
        self.extraction = ExtractionWidget(self.store, self.spreadsheet)
        self.petition = PetitionWidget(self.spreadsheet)
        self.sources = SourcesWidget(self.rag)
        self.chat = ChatPanel(self.spreadsheet, self.rag)

        self.templates.templates_changed.connect(self.extraction.refresh_templates)

        # Right side: tool panels in a stacked widget
        self.tool_stack = QStackedWidget()
        self.tool_panels = {}
        for name, widget, label in [
            ("chat", self.chat, "🤖 AI Chat"),
            ("petition", self.petition, "🛡 Petitions"),
            ("vision", self.extraction, "📥 Vision Import"),
            ("template", self.templates, "📋 Templates"),
            ("notebook", self.sources, "📚 Notebook"),
        ]:
            self.tool_stack.addWidget(widget)
            self.tool_panels[name] = (widget, label)

        # Side panel header with back button and title
        self.panel_title = QLabel("🤖 AI Chat")
        self.panel_title.setStyleSheet(
            "font-size: 15px; font-weight: 700; color: #1e40af; padding: 6px 10px;"
            "background: qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #ffffff, stop:1 #f8fafc);"
            "border: 1px solid #e2e8f0; border-radius: 8px;"
        )
        self.panel_back_btn = QPushButton("← Chat")
        self.panel_back_btn.setFixedWidth(80)
        self.panel_back_btn.clicked.connect(lambda: self.show_tool_panel("chat"))
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

        right_panel = QWidget()
        right_panel.setMinimumWidth(150)
        self.spreadsheet.setMinimumWidth(200)
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(4)
        right_layout.addLayout(panel_header)
        right_layout.addWidget(self.tool_stack, 1)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.addWidget(self.spreadsheet)
        splitter.addWidget(right_panel)
        splitter.setSizes([1100, 400])
        splitter.setHandleWidth(6)
        splitter.setChildrenCollapsible(False)

        self.setCentralWidget(splitter)
        self.spreadsheet.load_autosave_if_exists()

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
            widget, label = self.tool_panels[name]
            self.tool_stack.setCurrentWidget(widget)
            self.panel_title.setText(label)
            self.panel_back_btn.setVisible(name != "chat")

    def _check_ollama_connection(self) -> None:
        client = OllamaClient()
        ok, _ = client.check_connection()
        if ok:
            self.ollama_status.setText("● Ollama Connected")
            self.ollama_status.setStyleSheet("color: #059669; font-size: 12px; font-weight: 600; padding: 0 8px;")
        else:
            self.ollama_status.setText("○ Ollama Offline")
            self.ollama_status.setStyleSheet("color: #dc2626; font-size: 12px; font-weight: 600; padding: 0 8px;")

    def _update_row_count(self) -> None:
        count = sum(1 for r in range(self.spreadsheet.table.rowCount())
                    if any(self.spreadsheet.raw_values.get((r, c)) for c in range(self.spreadsheet.table.columnCount())))
        self.row_count_label.setText(f"{count} rows")

    def closeEvent(self, event) -> None:
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
        "foreground": qcolor_to_rgb(foreground.color()) if foreground.style() != Qt.BrushStyle.NoBrush else None,
        "background": qcolor_to_rgb(background.color()) if background.style() != Qt.BrushStyle.NoBrush else None,
        "alignment": int(item.textAlignment()),
    }


def apply_item_style(item: QTableWidgetItem, style: object) -> None:
    if not isinstance(style, dict):
        return
    font = item.font()
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
