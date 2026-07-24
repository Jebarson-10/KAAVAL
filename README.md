# 🛡️ KAAVAL & EOWassist
### Tamil Nadu Police Economic Offences Wing (EOW)
**Knowledge & Advanced Analysis for Vigilance & Asset Loss**

A high-end, fully local Python desktop application suite for automated document extraction, bilingual (Tamil/English) OCR, and intelligent spreadsheet management.

---

## 📖 Overview

This repository contains two primary modules designed for the Tamil Nadu Police EOW:

1. **KAAVAL**: A professional spreadsheet-based interface for managing petition data, featuring custom cell styling, borders, and integrated AI chat.
2. **EOWassist (Petition Extractor)**: A dedicated tool for batch-processing scanned petitions (PDF/Image) using high-quality OCR and Cloud/Local LLMs (via Gemini/Ollama) to extract structured data directly into a spreadsheet.

### Key Features
- **Native PDF Rendering**: Uses `PyMuPDF` (fitz) for high-performance, high-quality page rendering without external dependencies like Poppler.
- **Bilingual Extraction & OCR**: Seamless Tamil and English text extraction with auto-translation and formatting capabilities built-in.
- **Automatic Bamini Tamil Encoding**: Intelligently auto-detects extracted Tamil text and natively converts and renders it into the legacy `Bamini` font layout perfectly, while leaving English letters and numbers unaffected. This ensures maximum compatibility with older regional government systems.
- **Intelligent Extraction Engines**: 
  - *Cloud Mode (Gemini)*: Fast and highly accurate data extraction using Google's Gemini 1.5 Pro / Flash models via API.
  - *Local Mode (Ollama)*: Fully offline, private data extraction using Local Vision Models (Gemma/Qwen).
- **Professional Spreadsheet**: A custom-built PyQt6 spreadsheet engine with support for formulas, merged cells, custom borders, advanced styling, and Excel (.xlsx) import/export.
- **Local AI Chat / RAG**: Chat securely with your spreadsheet data using Retrieval Augmented Generation (ChromaDB + Local embeddings) to find deep insights across thousands of rows.

---

## 🛠️ Setup & Installation

### 1. Prerequisites
- **Python 3.12+**
- *(Optional)* **Ollama**: [Download and Install](https://ollama.com) (if you intend to use local extraction models).
  - *Recommended Models*: 
    ```bash
    ollama pull gemma3:12b
    ollama pull qwen2.5-vl
    ```

### 2. Environment Setup
```powershell
# Clone the repository
git clone https://github.com/Jebarson-10/KAAVAL.git
cd KAAVAL

# Create and activate virtual environment
python -m venv .venv
.\.venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

---

## 🚀 How to Run (Development)

### **KAAVAL (Main App)**
The central hub for data management, PDF extraction, and AI chat.
```powershell
python kaaval_app.py
```

---

## 📦 Building Standalone Executables (Production)
To create standalone Windows `.exe` files that don't require Python to be installed on target machines:

1. Ensure the virtual environment is active.
2. Run the build script:
   ```powershell
   .\build_kaaval.bat
   ```
This will run `PyInstaller`, cleanly bundle all required AI backends, Chromadb database plugins, and PDF engines, and generate a portable `KAAVAL.exe` in the `dist/` folder alongside a Desktop shortcut.

---

## 🎯 Walkthrough & Usage

### 1. Extracting Data from Petitions
1. Open the **KAAVAL** app.
2. Provide your **Gemini API Key** in the settings (or select Local Ollama).
3. Drag and drop your scanned petitions (PDFs/Images) into the extraction queue.
4. The app will render the PDF natively in the background and pass it to the selected AI engine.
5. The AI will parse the complex tabular data and translate fields as necessary, automatically converting Tamil outputs into perfectly legible Bamini encoding.
6. The structured data is populated dynamically row-by-row into the spreadsheet.

### 2. Managing Data in KAAVAL
1. Use the **Toolbar** to apply Bold, Italic, custom Colors, and **Custom Borders**.
2. Right-click any cell for advanced operations (Merge, Unmerge, Auto-fit).
3. Easily toggle cells back and forth between Unicode Tamil and Bamini Tamil if manual corrections are needed.
4. Save locally as `.laisheet` or click **Export to Excel** to distribute `.xlsx` files.

### 3. AI Studio Chat
1. Use the **🤖 AI Studio** panel on the right sidebar.
2. Ask questions like: *"How many petitions are from the scheme 'Fixed Deposit' in Thanjavur?"*
3. The RAG engine will instantly search your spreadsheet rows and provide a synthesized answer.

---

## 🛡️ Security & Privacy
This application features a hybrid privacy approach:
- **Offline Core**: The spreadsheet interface, RAG database (ChromaDB), vector embeddings, and manual data operations are strictly local.
- **Choice of Extractors**: You can choose to use local Ollama models (100% offline, highly secure) or Google Gemini API (cloud, fast) depending on your data classification needs. 

---

## 📄 License
Developed and open-sourced for the **Tamil Nadu Police Economic Offences Wing**.
