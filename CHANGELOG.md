# Changelog

All notable changes to the Priceline ORION Enterprise Platform.

Format: [Semantic Versioning](https://semver.org/) — MAJOR.MINOR.PATCH

---

## [2.7.0] — 2025-02-10

### Added
- **FA Coaching Module** (`FACoachingCore.gs`, `FACoaching.html`)
  - Date-based coaching case lookup from Google Sheet
  - Autocomplete dropdown for coaching date selection
  - Case list view (Analyst, ITN, Notes)
  - Case detail view with full QA notes
  - SME notes input with save-to-sheet functionality
  - Auto-creates `SME Notes` column header if absent
- FA Coaching route (`?page=fa`) in `doGet()` router
- Unlocked FA module card on Home.html (was "Coming Soon")
- `Config.MODULES.FAQ` now enabled

### Changed
- `Config.MODULES.FAQ.name` renamed from "Fraud Analysis Queue" to "FA Coaching"
- Home page grid now shows 3 active modules (CB, FA, RT)

---

## [2.6.0] — 2025-01-10

### Added
- Route fixes and stability improvements
- Restored HTML pages deleted in v2.4.0

### Fixed
- All pages use consistent ORION dark theme
- Proper file structure for GAS compatibility

---

## [2.5.0] — 2025-01-10

### Fixed
- Restored HTML pages that were accidentally removed in v2.4.0
- Fixed routing errors for all pages
- Total: 19 code files (6 .gs + 13 .html)

---

## [2.3.0] — 2025-01-10

### Changed
- Merged `RepManager.gs` + `RebuttalGenerator.gs` into `RCACore.gs`
- Moved docs back to root (`README.md`, `ORION_PROMPT.md`)
- Total: 22 code files (9 .gs + 13 .html)

---

## [2.2.0] — 2025-01-10

### Changed
- Created `docs/` folder for documentation
- Merged 8 HTML partials → 2 (`RCAPartials.html`, `RTPartials.html`)
- Total: 24 code files (13 HTML + 11 .gs)
- Updated `MainRCAPage.html` includes

---

## [2.1.0] — 2025-01-10

### Added
- **Universal Theme System** (`Theme.html`)
  - One-line import for complete design system
  - Dark/light mode toggle with localStorage persistence
  - Unified color system, typography, spacing, components
  - Dotted grid background pattern
- `ThemeState.html` for immediate theme application (prevents flash)
- `ORION_PROMPT.md` context file for AI assistants

---

## [2.0.0] — 2025-01-10

### Added
- Enterprise-level consolidation
- Unified `InfoPage.html` for About/Privacy/Data Retention
- Naming schema and semantic versioning
- Case-insensitive `isAdmin()` comparison

### Changed
- Consolidated 45 files → 29 files
- Merged 32 .gs files → 11 .gs files
- Hidden Word Analyser for non-admins (removed grayed-out state)

---

## [1.0.0] — 2025-01-09

### Added
- **Initial Release**
- 6-module landing page (CB, CBQA, FAQ, FAQA, RT, ON)
- Chargeback sub-modules (Rebuttal Engine, Rep Manager, Word Analyser)
- AI-powered text classification (GPT-4o-mini)
- Reporting governance with RACI workflow
- ORION dark theme UI
- Role-based access control
- Google Sheets data layer
- Google Drive file management
