# ✦ KAAVAL & EOWassist
### Tamil Nadu Police Economic Offences Wing (EOW)
**Knowledge & Advanced Analysis for Vigilance & Asset Loss**

A high-end, fully local Python desktop application suite for automated document extraction, bilingual (Tamil/English) OCR, and intelligent spreadsheet management.

---

## 🚀 Overview

This repository contains two primary modules designed for the Tamil Nadu Police EOW:

1. **KAAVAL**: A professional spreadsheet-based interface for managing petition data, featuring custom cell styling, borders, and integrated AI chat.
2. **EOWassist (Petition Extractor)**: A dedicated tool for batch-processing scanned petitions (PDF/Image) using high-quality OCR and Local LLMs (via Ollama) to extract structured data directly into a spreadsheet.

### Key Features
- **Native PDF Rendering**: Uses `PyMuPDF` (fitz) for high-performance, high-quality page rendering without external dependencies like Poppler.
- **Bilingual OCR**: Advanced `pytesseract` integration for seamless Tamil and English text extraction.
- **Intelligent Extraction**: Uses Local Vision Models (Gemma/Qwen via Ollama) to parse complex tabular data from images.
- **Professional Spreadsheet**: A custom-built PyQt6 spreadsheet engine with support for formulas, merged cells, custom borders, and Excel (.xlsx) export.
- **Local AI Chat**: Chat with your spreadsheet data using RAG (Retrieval Augmented Generation) to find information across thousands of rows.

---

## 🛠 Setup & Installation

### 1. Prerequisites
- **Python 3.12+**
- **Tesseract OCR**: [Download and Install](https://github.com/UB-Mannheim/tesseract/wiki). 
  - *Note: Ensure you include Tamil language data during installation.*
- **Ollama**: [Download and Install](https://ollama.com).
  - *Required Models*: 
    ```bash
    ollama pull gemma3:12b
    ollama pull qwen2.5-vl
    ```

### 2. Environment Setup
```powershell
# Clone the repository
git clone <your-repo-url>
cd "New project"

# Create and activate virtual environment
python -m venv .venv
.\.venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

---

## 🚦 How to Run

### **KAAVAL (Main App)**
The central hub for data management and AI chat.
```powershell
python kaaval_app.py
```

### **EOWassist (Petition Extractor)**
The standalone tool for rapid data entry from scanned files.
```powershell
python EOWassist\Dataext.py
```

---

## 🏗 Building Executables
To create standalone Windows `.exe` files that don't require Python to be installed:

1. Ensure the virtual environment is active.
2. Run the build script:
   ```powershell
   .\build_kaaval.bat
   ```
This will generate a portable `KAAVAL.exe` in the `dist/` folder and create a Desktop shortcut.

---

## 📖 Walkthrough & Usage

### 1. Extracting Data from Petitions
1. Open **EOWassist**.
2. Click **📄 Load PDF / Image** and select your scanned petitions.
3. The app will render the PDF at 2x zoom for maximum OCR accuracy.
4. The local AI will analyze the text and populate the spreadsheet rows automatically.
5. Click **📊 Export to Excel** to save your work.

### 2. Managing Data in KAAVAL
1. Open **KAAVAL**.
2. Load your `.laisheet` file or start fresh.
3. Use the **Toolbar** to apply Bold, Italic, custom Colors, and **Custom Borders**.
4. Right-click any cell for advanced operations (Merge, Unmerge, Auto-fit).
5. Use the **🤖 AI Studio** on the right to ask questions like: *"How many petitions are from the scheme 'Fixed Deposit'?"*

---

## 🔒 Security & Privacy
This application is designed to be **100% Offline**.
- **No data** is ever sent to external servers.
- **OCR and LLM** processing happens entirely on your local GPU/CPU.
- **Persistence** is handled via local `.laisheet` and `.xlsx` files.

---

## 📜 License
Developed for the **Tamil Nadu Police Economic Offences Wing**.
