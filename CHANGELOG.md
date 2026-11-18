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

## [v1.0.3]

### Fixed
- Display original Unicode text in the spreadsheet by default.
- Added an optional Bamini preview for Tamil text without changing cell values.

## [v1.0.2]

### Fixed
- Kept ordinary English readable when opening sheets with an old Bamini cell font.
- Widened the ribbon font-size field and added the build version to the window title.
- Made Gemini 3.8 Flash the default and added Gemini 4 Hydrogen as a preset.

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
