import sys
import os
import json
import io
from PIL import Image
import pytesseract
import ollama

from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QPushButton, QFileDialog, QLabel,
                             QMessageBox, QComboBox, QProgressBar, QTextEdit,
                             QStatusBar, QSplitter, QLineEdit, QFormLayout,
                             QGroupBox)
from PySide6.QtCore import Qt, QThread, Signal as pyqtSignal, QTimer, QMimeData
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QFont

# Append parent dir so we can import local_ai_sheet
if getattr(sys, 'frozen', False):
    sys.path.append(sys._MEIPASS)
else:
    sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from local_ai_sheet.app import SpreadsheetWidget

# Configure this to your Tesseract installation path if on Windows
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# ── Dark Theme ──
DARK_STYLESHEET = """
QWidget {
    background-color: #0f1117;
    color: #e2e8f0;
    font-family: "Segoe UI", system-ui, sans-serif;
    font-size: 13px;
}
QMainWindow { background-color: #0f1117; }
QPushButton {
    background: #1e293b; color: #e2e8f0; border: 1px solid #334155;
    border-radius: 6px; padding: 8px 16px; font-weight: 600;
}
QPushButton:hover { border: 1px solid #6366f1; background: #334155; }
QPushButton:pressed { padding-top: 9px; padding-bottom: 7px; }
QPushButton:disabled { background: #0f1117; color: #475569; border: 1px solid #1e293b; }
QComboBox {
    background: #1a1d27; border: 1px solid #2d3348; border-radius: 6px;
    padding: 6px 10px; color: #e2e8f0; min-height: 22px;
}
QComboBox:hover { border: 1px solid #6366f1; }
QComboBox QAbstractItemView { background: #1a1d27; border: 1px solid #2d3348; selection-background-color: #6366f1; }
QLineEdit {
    background: #1a1d27; border: 1px solid #2d3348; border-radius: 6px;
    padding: 6px 10px; color: #e2e8f0;
}
QLineEdit:focus { border: 1px solid #6366f1; }
QTextEdit {
    background: #12141d; border: 1px solid #1e293b; border-radius: 6px;
    color: #e2e8f0; padding: 8px;
}
QProgressBar {
    background: #1a1d27; border: 1px solid #2d3348; border-radius: 6px;
    text-align: center; color: #e2e8f0; font-weight: 600; min-height: 22px;
}
QProgressBar::chunk {
    background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #6366f1, stop:1 #a78bfa);
    border-radius: 5px;
}
QGroupBox {
    border: 1px solid #1e293b; border-radius: 8px; margin-top: 12px;
    padding-top: 16px; background: #12141d; font-weight: 700; color: #a78bfa;
}
QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top left; padding: 2px 12px; }
QLabel { color: #cbd5e1; background: transparent; }
QStatusBar { background: #0d0f15; border-top: 1px solid #1e293b; color: #64748b; font-size: 12px; }
QSplitter::handle { background: #1e293b; width: 3px; }
QSplitter::handle:hover { background: #6366f1; }
QScrollBar:vertical { background: #0f1117; width: 10px; }
QScrollBar::handle:vertical { background: #334155; border-radius: 5px; min-height: 30px; }
QScrollBar::handle:vertical:hover { background: #6366f1; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar:horizontal { background: #0f1117; height: 10px; }
QScrollBar::handle:horizontal { background: #334155; border-radius: 5px; }
QScrollBar::handle:horizontal:hover { background: #6366f1; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0; }
QToolTip { background: #1e293b; color: #e2e8f0; border: 1px solid #334155; border-radius: 4px; padding: 4px 8px; }
"""

MODEL_PRESETS = [
    ("Gemma 3 12B", "gemma3:12b"),
    ("Gemma 4 E2B", "gemma4:e2b"),
    ("Gemma 4 26B", "gemma4:26b"),
    ("Qwen 2.5 VL 7B", "qwen2.5vl:7b"),
]

DEFAULT_INSTRUCTION = """You are an expert data extractor. I will provide you with a transcription of Tamil legal/financial petitions. This text may represent one single petition spanning multiple pages, or multiple different petitions.
The text is from an OCR engine, expect typos and broken spacing.
Pay extreme attention to Crime Numbers (e.g., Cr.No, FIR No, Crime No) and Scheme Names/Company Names. Use logical deduction to reconstruct fragmented Tamil or English names.
IMPORTANT: For fields like 'ID Number', 'Plan Name', or 'Scheme Name', if they appear in English in the original text, preserve the exact English text. Do NOT translate them into Tamil."""

class ExtractionThread(QThread):
    finished = pyqtSignal(dict)
    all_finished = pyqtSignal()
    error = pyqtSignal(str)
    progress = pyqtSignal(str, int, int)  # message, current, total

    def __init__(self, file_paths, headers, instruction, model_name):
        super().__init__()
        self.file_paths = file_paths
        self.headers = headers
        self.instruction = instruction
        self.model_name = model_name

    def run(self):
        try:
            self.progress.emit("Reading text from selected files via OCR...", 0, len(self.file_paths))
            combined_text = ""
            total = len(self.file_paths)

            for i, file_path in enumerate(self.file_paths, 1):
                base_name = os.path.basename(file_path)
                self.progress.emit(f"[{i}/{total}] OCR processing {base_name}...", i, total)

                if file_path.lower().endswith('.pdf'):
                    import fitz
                    doc = fitz.open(file_path)
                    for page_num in range(len(doc)):
                        page = doc.load_page(page_num)
                        # Render page to image with 2x zoom for better OCR accuracy
                        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                        page_img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                        combined_text += pytesseract.image_to_string(page_img, lang='tam+eng') + "\n\n"
                    doc.close()
                else:
                    combined_text += pytesseract.image_to_string(file_path, lang='tam+eng') + "\n\n"

            if not combined_text.strip():
                raise ValueError("No text could be extracted from the selected files.")

            self.progress.emit("Analyzing combined text with local LLM...", total, total)

            headers_json = json.dumps({h: "" for h in self.headers}, indent=2, ensure_ascii=False)
            prompt = f"""{self.instruction}

Extract the relevant details and output ONLY a raw JSON array of objects with the following exact keys for EACH petition found. If a detail is missing, put "Not Specified".

[
{headers_json}
]

Tamil Text Transcription:
{combined_text}
"""

            response = ollama.chat(model=self.model_name, messages=[
                {'role': 'user', 'content': prompt}
            ])

            output_text = response['message']['content']

            if "```json" in output_text:
                output_text = output_text.split("```json")[1].split("```")[0].strip()
            elif "```" in output_text:
                output_text = output_text.split("```")[1].split("```")[0].strip()

            data = json.loads(output_text)

            if isinstance(data, list):
                for petition in data:
                    self.finished.emit(petition)
            elif isinstance(data, dict):
                self.finished.emit(data)

        except Exception as e:
            self.error.emit(f"Error processing files: {str(e)}")

        self.all_finished.emit()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("✦ Tamil Petition Data Extractor")
        self.setGeometry(100, 100, 1400, 900)

        self.model_name = "gemma3:12b"
        self.current_prompt = DEFAULT_INSTRUCTION
        self.init_ui()
        self.populate_default_headers()

        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.ollama_status = QLabel("● Checking...")
        self.ollama_status.setStyleSheet("color: #64748b; font-size: 12px; padding: 0 8px;")
        self.row_count_label = QLabel("0 rows")
        self.row_count_label.setStyleSheet("color: #475569; font-size: 12px;")
        self.status_bar.addPermanentWidget(self.row_count_label)
        self.status_bar.addPermanentWidget(self.ollama_status)
        self._check_ollama()

        # Timers
        self.connection_timer = QTimer(self)
        self.connection_timer.setInterval(30000)
        self.connection_timer.timeout.connect(self._check_ollama)
        self.connection_timer.start()

    def _check_ollama(self):
        try:
            import urllib.request
            req = urllib.request.Request("http://localhost:11434/api/tags", method="GET")
            with urllib.request.urlopen(req, timeout=3):
                self.ollama_status.setText("● Ollama Connected")
                self.ollama_status.setStyleSheet("color: #22c55e; font-size: 12px; padding: 0 8px;")
        except Exception:
            self.ollama_status.setText("○ Ollama Offline")
            self.ollama_status.setStyleSheet("color: #ef4444; font-size: 12px; padding: 0 8px;")

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(8)

        # Top Controls Group
        controls = QGroupBox("⚙ Extraction Controls")
        control_layout = QHBoxLayout(controls)

        # Left: buttons
        btn_col = QVBoxLayout()
        self.btn_load = QPushButton("📄 Load PDF / Image")
        self.btn_load.clicked.connect(self.load_document)
        self.btn_load.setStyleSheet(
            "QPushButton { background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #6366f1, stop:1 #818cf8);"
            "color: white; border: none; border-radius: 6px; padding: 10px 20px; font-weight: 700; font-size: 14px; }"
            "QPushButton:hover { border: 1px solid #a78bfa; }"
            "QPushButton:disabled { background: #1e293b; color: #475569; }"
        )
        self.btn_export = QPushButton("📊 Export to Excel")
        self.btn_export.setStyleSheet(
            "QPushButton { background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #7c3aed, stop:1 #a855f7);"
            "color: white; border: none; border-radius: 6px; padding: 8px 16px; font-weight: 600; }"
            "QPushButton:hover { border: 1px solid #c084fc; }"
        )
        self.btn_prompt = QPushButton("✍ Edit Prompt")
        btn_col.addWidget(self.btn_load)
        btn_col.addWidget(self.btn_export)
        btn_col.addWidget(self.btn_prompt)
        self.btn_prompt.clicked.connect(self.edit_prompt)

        # Right: model settings
        settings_col = QFormLayout()
        self.model_combo = QComboBox()
        for label, value in MODEL_PRESETS:
            self.model_combo.addItem(label, value)
        self.model_combo.currentIndexChanged.connect(self._on_model_changed)
        settings_col.addRow("Model", self.model_combo)

        self.status_label = QLabel("✦ Ready — Drop files or click Load")
        self.status_label.setStyleSheet("color: #a78bfa; font-weight: 600; font-size: 13px;")

        control_layout.addLayout(btn_col)
        control_layout.addLayout(settings_col)
        control_layout.addStretch()
        control_layout.addWidget(self.status_label)

        main_layout.addWidget(controls)

        # Progress bar
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.progress.setTextVisible(True)
        main_layout.addWidget(self.progress)

        # Splitter: Spreadsheet + Log
        splitter = QSplitter(Qt.Orientation.Vertical)

        self.sheet = SpreadsheetWidget()
        splitter.addWidget(self.sheet)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumHeight(200)
        self.log.setPlaceholderText("Extraction log will appear here...")
        splitter.addWidget(self.log)
        splitter.setSizes([650, 200])

        main_layout.addWidget(splitter, 1)

        # Connect export button
        self.btn_export.clicked.connect(self.sheet.export_xlsx)

        # Enable drag-and-drop
        self.setAcceptDrops(True)

    def _on_model_changed(self):
        model = self.model_combo.currentData()
        if model:
            self.model_name = model
            self.status_label.setText(f"Model: {self.model_combo.currentText()}")

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        paths = []
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path and path.lower().endswith(('.pdf', '.png', '.jpg', '.jpeg')):
                paths.append(path)
        if paths:
            self._start_extraction(paths)

    def populate_default_headers(self):
        default_headers = [
            "Sender Name & Address",
            "Petitioner Name & Address",
            "Petitioner Mobile No",
            "Deposit Amount",
            "Repaid Amount",
            "Default Amount",
            "ID Number",
            "Plan Name",
            "161 Statement",
            "Agent",
            "Crime No"
        ]
        for col, header in enumerate(default_headers):
            self.sheet.raw_values[(0, col)] = header
        self.sheet.recalculate()

    def get_headers(self):
        headers = []
        for col in range(200):
            val = self.sheet.raw_values.get((0, col), "").strip()
            if val:
                headers.append(val)
        return headers

    def load_document(self):
        headers = self.get_headers()
        if not headers:
            QMessageBox.warning(self, "No Headers", "Please define at least one header in Row 1.")
            return

        file_paths, _ = QFileDialog.getOpenFileNames(
            self, "Select Documents", "", "Documents/Images (*.pdf *.png *.jpg *.jpeg)"
        )
        if file_paths:
            self._start_extraction(file_paths)

    def _start_extraction(self, file_paths):
        headers = self.get_headers()
        if not headers:
            QMessageBox.warning(self, "No Headers", "Please define at least one header in Row 1.")
            return

        self.btn_load.setEnabled(False)
        self.progress.setMaximum(len(file_paths))
        self.progress.setValue(0)
        self.progress.setVisible(True)
        self.log.append(f"\n{'─' * 40}\n🔄 Starting extraction of {len(file_paths)} file(s)...\n")

        self.thread = ExtractionThread(file_paths, headers, self.current_prompt, self.model_name)
        self.thread.progress.connect(self.update_status)
        self.thread.finished.connect(self.handle_extraction_success)
        self.thread.all_finished.connect(self.handle_all_finished)
        self.thread.error.connect(self.handle_extraction_error)
        self.thread.start()

    def update_status(self, msg, current, total):
        self.status_label.setText(msg)
        self.progress.setValue(current)
        self.progress.setFormat(f"{current}/{total}: {msg}")

    def handle_extraction_success(self, data):
        headers = self.get_headers()

        next_row = 1
        for r in range(1, 1000):
            is_empty = True
            for c in range(len(headers)):
                if self.sheet.raw_values.get((r, c)):
                    is_empty = False
                    break
            if is_empty:
                next_row = r
                break

        for col, header in enumerate(headers):
            val = str(data.get(header, ""))
            if val and val != "Not Specified":
                self.sheet.raw_values[(next_row, col)] = val
        self.sheet.recalculate()
        self.sheet.table.resizeRowsToContents()
        self.sheet._mark_changed()

        self.log.append(f"✅ Extracted row {next_row}")
        self._update_row_count()

    def handle_all_finished(self):
        self.status_label.setText("✅ All Extractions Complete!")
        self.status_label.setStyleSheet("color: #22c55e; font-weight: 600; font-size: 13px;")
        self.btn_load.setEnabled(True)
        self.progress.setVisible(False)
        self.log.append(f"\n✅ Extraction complete.\n{'─' * 40}\n")

    def handle_extraction_error(self, error_msg):
        self.btn_load.setEnabled(True)
        self.progress.setVisible(False)
        self.log.append(f"❌ Error: {error_msg}\n")
        QMessageBox.critical(self, "Extraction Error", f"An error occurred:\n{error_msg}")

    def _update_row_count(self):
        count = sum(1 for r in range(self.sheet.table.rowCount())
                    if any(self.sheet.raw_values.get((r, c)) for c in range(self.sheet.table.columnCount())))
        self.row_count_label.setText(f"{count} rows")

    def edit_prompt(self):
        from PySide6.QtWidgets import QInputDialog
        dialog = QInputDialog(self)
        dialog.setWindowTitle("Edit LLM Prompt")
        dialog.setLabelText("Modify the instructions sent to Ollama:")
        dialog.setTextValue(self.current_prompt)
        dialog.setOption(QInputDialog.UsePlainTextEditForTextInput, True)
        dialog.resize(600, 400)

        if dialog.exec() == QInputDialog.Accepted:
            self.current_prompt = dialog.textValue()
            QMessageBox.information(self, "Success", "Prompt updated successfully.")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyleSheet(DARK_STYLESHEET)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())