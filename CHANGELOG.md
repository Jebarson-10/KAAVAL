# Changelog

All notable changes to KAAVAL are documented here.
Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

---

## [Unreleased]

### Added
- AI Training Session with few-shot template support
- Training template dropdown in Petition Extractor
- Gemini API key persistence across sessions
- Production-grade UI polish and sidebar layout fixes
- Thread-safe extraction with proper cleanup on close

---

## [v1.0.1]

### Fixed
- Reorganized the spreadsheet ribbon and petition tools to fit smaller windows.
- Added a font-size control and preserved formulas during cell formatting.
- Rendered Tamil in Bamini and retained source English in mixed-language cells.

---

## [v1.0.0] — Initial Release

### Added
- Petition document extraction via local Ollama models
- Gemini online extraction support
- Spreadsheet view with formula engine
- AI Chat panel (RAG-based Notebook)
- Export to Excel / CSV / HTML
- Tamil OCR support (Tesseract)
