# Priceline ORION Enterprise Platform

> **Unified Operations Platform** — Chargeback Management, Fraud Analysis Coaching, Reporting Governance & Text Analytics  
> Built entirely on Google Apps Script (GAS) + Google Sheets as the data layer

---

## Table of Contents

- [Overview](#overview)
- [Architecture](#architecture)
- [Modules](#modules)
  - [CB — Chargeback](#cb--chargeback)
  - [FA — FA Coaching](#fa--fa-coaching)
  - [RT — Reporting](#rt--reporting)
  - [Admin Panel](#admin-panel)
- [File Structure](#file-structure)
- [Configuration](#configuration)
- [Deployment](#deployment)
- [Authentication & Authorization](#authentication--authorization)
- [API Reference](#api-reference)
- [Design System](#design-system)
- [Changelog](#changelog)

---

## Overview

ORION is an enterprise-grade internal operations platform built for Priceline's fraud and chargeback teams. It runs as a Google Apps Script Web App, serving an SPA-like experience from Google's infrastructure with zero external dependencies.

**Key Capabilities:**
- RCA case management with rebuttal PDF generation
- Representation Manager for case submission workflow
- AI-powered text analytics (GPT-4o-mini integration)
- Reporting governance with RACI-based approval workflows
- FA Coaching module for SME case review
- Email confirmation upload pipeline
- Role-based access control (RBAC)
- Dark/light theme with unified design system

---

## Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                    Google Apps Script                        │
│                                                             │
│  Code.gs (consolidated)                                     │
│  ├── Config Core (CONFIG, Config, VERSION)                  │
│  ├── Auth Core (ORION_STANDARDS, Authorization, Auth)       │
│  ├── Router (doGet → page routing)                          │
│  ├── Utils (Utils, SheetsService)                           │
│  ├── RCA Core (case mgmt, recovery, upload, rebuttal gen)   │
│  ├── CBQ Core (text analytics, AI pipeline)                 │
│  ├── AI Core (GPT-4o-mini, tokenizer, batch processing)     │
│  ├── Reporting Core (RACI workflow, DAOs, importer)         │
│  └── FA Coaching Core (coaching data, SME notes)            │
│                                                             │
│  HTML Pages (GAS include() system)                          │
│  ├── Theme.html (universal design system)                   │
│  ├── ThemeState.html (dark/light toggle persistence)        │
│  ├── Home.html (6-module landing grid)                      │
│  ├── CBModule.html (Chargeback sub-module selector)         │
│  ├── MainRCAPage.html + RCAPartials.html                    │
│  ├── RepManagerUI.html                                      │
│  ├── CBQMain.html (Word Analyser)                           │
│  ├── FACoaching.html (FA Coaching UI)                       │
│  ├── Dashboard.html + ReportsQueue.html + ReportDetail.html │
│  ├── AdminPanel.html                                        │
│  ├── EmailConfirm.html                                      │
│  └── InfoPage.html (About/Privacy/Data Retention)           │
│                                                             │
└───────────────┬─────────────────────────────────────────────┘
                │
                ▼
┌─────────────────────────────────────────────────────────────┐
│                   Google Sheets (Data Layer)                 │
│                                                             │
│  Live Datasheet ─── RCA cases, Rep Manager, Rebuttals       │
│  CB Controller ──── Chargeback configuration                 │
│  ORION Spreadsheet ─ RT_Reports, RT_AuditLog, RT_Users      │
│  FA Coaching Sheet ─ Coaching dates, analyst cases, notes    │
│  Recovery / VB Exception / Canned Responses sheets          │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

### Tech Stack

| Layer | Technology |
|-------|-----------|
| Runtime | Google Apps Script (V8) |
| Data Store | Google Sheets |
| File Storage | Google Drive |
| AI | OpenAI GPT-4o-mini (via UrlFetchApp) |
| PDF Gen | Google Docs Templates + DriveApp |
| Auth | Google Workspace SSO (Session API) |
| Frontend | Vanilla HTML/CSS/JS (no frameworks) |
| Theme | Custom CSS variables + dark/light mode |

---

## Modules

### CB — Chargeback

The Chargeback module is the primary operational tool, containing four sub-modules:

#### Rebuttal Engine (`?page=rca`)
- ITN-based case lookup from live datasheet
- Multi-tab case research interface (screenshots, accertify data, checkout pages)
- Automated rebuttal PDF generation from Google Docs templates
- Case status workflow (Pending → In Progress → Completed)
- Screenshot upload pipeline to Google Drive folders
- Keyboard shortcuts for power users

#### Representation Manager (`?page=rep`)
- Pending case queue filtered by portal and currency
- One-click actions: Challenge, Accept, Reversal, Autoclosed, Recreate
- Real-time count dashboard
- ITN file search across Drive

#### Email Confirmations (`?page=email-confirm`)
- Bulk PDF upload for email confirmation documents
- Auto-matching filenames to ITN numbers
- Status tracking (Pending → Completed)

#### Word Analyser (`?page=cbq`) — Admin Only
- AI-powered text extraction and classification
- Word frequency analysis with filtering
- Word cloud and bubble chart visualizations
- Processing queue with batch operations
- PDF and folder upload support

---

### FA — FA Coaching

**Route:** `?page=fa`  
**Sheet ID:** `1qvDu_adBcaP__9sCxZ63He349g7ERrReKfmKlOdaNsM`

SME coaching interface for reviewing fraud analysis cases:

1. **Date Selection** — Type or pick a coaching date from autocomplete dropdown
2. **Case List** — Table showing Analyst Name, ITN, and QA Notes
3. **Case Detail** — Full case view with analyst info, resolution, dates, and QA notes
4. **SME Notes** — Textarea for SME to enter coaching feedback, saved to column G

**Schema (Google Sheet):**

| Column | Field |
|--------|-------|
| A | Coaching Date |
| B | Analyst Name |
| C | ITN |
| D | Date |
| E | Resolution |
| F | QA Notes |
| G | SME Notes (auto-created) |

---

### RT — Reporting

**Routes:** `?page=queue`, `?page=dashboard`, `?page=report&id=xxx`

Reporting governance module with RACI-based workflows:

- Report inventory with classification (Master/Derived)
- Frequency management (Daily, Weekly, Monthly, etc.)
- RACI role assignment (Responsible, Accountable, Consulted, Informed)
- Approval workflow (Pending → Accepted/Rejected)
- Audit log for all actions
- User/role management
- Filter and search across all reports
- Import capabilities for bulk report onboarding

---

### Admin Panel

**Route:** `?page=admin`

- Sheet ID configuration (hot-swap data sources)
- Folder ID management
- User administration
- System health monitoring

---

## File Structure

```
Priceline-Orion/
├── Code.gs                 # ALL server-side code (7 modules merged, ~294KB)
├── Index.html              # ALL client-side UI (17 pages merged, ~511KB)
├── appsscript.json         # GAS manifest
├── README.md
└── CHANGELOG.md
```

> **Two-file architecture**: All `.gs` modules are merged into `Code.gs` and all `.html` pages are
> merged into `Index.html` using GAS server-side conditional rendering (`<? if (page === '...') { ?>`).
> The `doGet()` router passes `page` as a template variable to `Index.html`, which conditionally
> renders only the requested page's styles and markup.

### Code.gs Internal Structure

The consolidated `Code.gs` contains these modules in dependency order:

| Section | Lines (approx) | Description |
|---------|----------------|-------------|
| Config Core | ~300 | `CONFIG`, `Config`, sheet/folder IDs, feature flags, case types |
| Version & Naming | ~170 | `VERSION`, changelog, naming conventions |
| Auth Core | ~250 | `ORION_STANDARDS`, `Authorization` RBAC, `Auth` identity |
| Router | ~250 | `doGet()`, `include()`, `isAdmin_()`, menu functions |
| API Endpoints | ~200 | Top-level functions for `google.script.run` |
| Utils | ~280 | `Utils` helpers, `SheetsService` CRUD abstraction |
| RCA Core | ~1500 | Case management, recovery, screenshot upload, admin |
| CBQ Core | ~1100 | CBQ configuration, controller, DAOs |
| AI Core | ~900 | AI processor, text extractor, tokenizer, metrics |
| Reporting Core | ~1800 | Report workflow, importer, DAOs, audit |
| FA Coaching Core | ~170 | Coaching dates, cases by date, SME notes |

---

## Configuration

### Google Sheet IDs

All sheet references are centralized in `CONFIG.DEFAULT_SHEET_IDS`:

| Key | Purpose | Default ID |
|-----|---------|------------|
| `LIVE_DATASHEET` | RCA/Rep Manager/Rebuttal data | `1ApuyYWWd4GCVlwkzRsdvPsrSFw_ppFlYYh6QQFymV28` |
| `CB_CONTROLLER` | Chargeback configuration | `1AKQpG3Glr-mrkmNcuUCSNo8fVqN7HrqkTocypS7II7s` |
| `RECOVERY` | Recovery tracking | `1lwO9EXKgzZmb-R162Vl_hd1Y06jdiAtaAh9t_XBBENE` |
| `VB_EXCEPTION` | VB exception list | `1tRZIM8nqh0q41S4PoV1Fp5BJ0RGGjVsNdFi-p_cSsdA` |
| `EMAIL_CONFIRMATION` | Email confirmation tracker | `1Q430kJMCMBbaiif1hFesbPUELtNyH1PrL9h3Uq8ApGw` |
| `CANNED_RESPONSES` | Canned response templates | `1p8kXzVAIB0AvZ1V6EPTLW9MIwjxEwNPL2DADz9X3BjE` |
| `REBUTTAL_TEMPLATE` | Rebuttal PDF template | `1v5ZPUvBozEGuaEiv1jndad9dvGm3XDA6CNmMFyAwLqU` |

FA Coaching uses its own config:

| Key | Purpose | ID |
|-----|---------|-----|
| `FA_COACHING_CONFIG.SHEET_ID` | SME coaching data | `1qvDu_adBcaP__9sCxZ63He349g7ERrReKfmKlOdaNsM` |

ORION Reporting uses:

| Key | Purpose | ID |
|-----|---------|-----|
| `Config.SPREADSHEET_ID` | RT Reports, Users, Audit Log | `1fbSOgULlPsxE6Mc39gJrQDvn5wf4PfgiCpKBKpz-7eQ` |

### Google Drive Folders

| Key | Purpose |
|-----|---------|
| `IMAGES` | RCA screenshot uploads |
| `CONFIRM_PAGE` | Confirmation page screenshots |
| `CHECKOUT_PAGE` | Checkout page screenshots |
| `WEX_REFUND` | WEX refund documents |
| `ACCERTIFY` | Accertify screenshots |
| `REBUTTAL_TEMP` | Temporary rebuttal files |
| `REBUTTAL_PDF` | Generated rebuttal PDFs |

### Script Properties

Sensitive values are stored via `PropertiesService.getScriptProperties()`:

| Property | Purpose |
|----------|---------|
| `OPENAI_API_KEY` | GPT-4o-mini API key for AI features |
| `SHEET_*` | Override default sheet IDs per-environment |
| `FOLDER_*` | Override default folder IDs per-environment |

---

## Deployment

### Prerequisites
- Google Workspace account with Apps Script access
- Access to all referenced Google Sheets and Drive folders

### Steps

1. **Create a new Google Apps Script project** at [script.google.com](https://script.google.com)
2. **Copy `Code.gs`** as the main script file (replace default Code.gs)
3. **Create `Index.html`** as an HTML file in the project (File → New → HTML file → name it "Index")
4. **Copy `appsscript.json`** to the manifest (View → Show manifest file)
5. **Set Script Properties** (Project Settings → Script Properties):
   - `OPENAI_API_KEY` — your OpenAI key (required for AI features)
6. **Deploy as Web App**:
   - Execute as: **Me** (the deploying user)
   - Who has access: **Anyone within your organization**
7. **Authorize** the required scopes on first run

### Using clasp (recommended for development)

```bash
# Install clasp
npm install -g @google/clasp

# Login
clasp login

# Clone existing project or create new
clasp create --type webapp

# Push files
clasp push

# Deploy
clasp deploy --description "v2.7.0"
```

---

## Authentication & Authorization

### Identity
- Uses `Session.getActiveUser().getEmail()` for identity
- Falls back to `Session.getEffectiveUser().getEmail()` if active user unavailable

### Admin System
Admin access is controlled by the `ORION_STANDARDS.ADMINS` whitelist:
```javascript
ADMINS: [
  'naitik.joshi@priceline.com',
  'naitikjoshipro@gmail.com'
]
```

### RBAC Levels
| Level | Value | Capabilities |
|-------|-------|-------------|
| Viewer | 1 | View reports |
| Analyst | 2 | View + edit reports |
| Manager | 3 | View + edit + approve/reject |
| Leader | 4 | + workflow editing |
| Admin | 5 | Full access including config |

---

## API Reference

All functions are exposed as top-level GAS functions callable via `google.script.run`:

### Router
| Function | Description |
|----------|-------------|
| `doGet(e)` | Main web app entry point, routes `?page=` param |
| `include(filename)` | HTML template inclusion |

### RCA / Rebuttal
| Function | Description |
|----------|-------------|
| `searchCases(itn)` | Search cases by ITN |
| `getCaseData(itn)` | Get full case details |
| `generateRebuttals()` | Batch generate rebuttal PDFs |
| `uploadScreenshot(data)` | Upload screenshot to Drive |

### Representation Manager
| Function | Description |
|----------|-------------|
| `rmGetPendingUploads(portal, currency)` | Get pending rep cases |
| `rmChallenge(row)` | Mark case as challenged |
| `rmAccept(row)` | Mark case as accepted |
| `rmReversal(row)` | Mark case as reversal |
| `rmAutoclosed(row)` | Mark case as autoclosed |
| `rmRecreate(row, res)` | Recreate a case |
| `rmGetCounts()` | Get queue counts |
| `rmSearchFiles(itn)` | Search Drive files by ITN |

### Email Confirmations
| Function | Description |
|----------|-------------|
| `ecGetPendingItems()` | Get pending email confirmations |
| `ecUploadFiles(formData)` | Upload confirmation PDFs |

### Reporting (RT)
| Function | Description |
|----------|-------------|
| `getReports(filters)` | Get filtered reports |
| `getReportById(id)` | Get single report |
| `createReport(data)` | Create new report |
| `updateReport(id, data)` | Update report fields |
| `changeReportState(id, state, comment)` | Workflow state change |
| `deleteReport(id)` | Delete report (admin only) |
| `getReportCounts()` | Get counts by state |
| `getDashboardStats()` | Dashboard aggregates |
| `getAuditLog(filters)` | Audit log entries |
| `getUsersForPicker()` | User list for dropdowns |

### CBQ (Text Analytics)
| Function | Description |
|----------|-------------|
| `cbqUploadText(data)` | Upload text for analysis |
| `cbqUploadTextWithAI(data)` | Upload with AI processing |
| `cbqUploadPDFWithAI(data)` | Upload PDF with AI  |
| `cbqGetDocuments(filters)` | Get analyzed documents |
| `cbqGetStatsSummary()` | Word frequency stats |
| `cbqGetWordCloudData(opts)` | Word cloud data |
| `cbqGetBubbleChartData(opts)` | Bubble chart data |
| `cbqProcessQueue(limit)` | Process AI queue |
| `cbqTestAIConnection()` | Test OpenAI connectivity |

### FA Coaching
| Function | Description |
|----------|-------------|
| `faGetCoachingDates()` | Get unique coaching dates |
| `faGetCasesByDate(dateStr)` | Get cases for a date |
| `faSaveSmeNotes(rowIndex, notes)` | Save SME coaching notes |

### System
| Function | Description |
|----------|-------------|
| `getUserInfo()` | Current user email, name, admin status |
| `getClientConfig()` | Feature flags, UI config |
| `getSheetConfig()` | Current sheet/folder ID configuration (admin) |
| `updateSheetId(key, id)` | Update a sheet ID (admin) |
| `getVersion()` | Current platform version |

---

## Design System

The ORION design system is defined in `Theme.html` — a single include that provides:

### Color System
- **Dark mode** (default): Pure black gradations (`#000000` → `#1a1a1a`)
- **Light mode**: Warm off-white (`#e8e6e1` → `#ffffff`)
- **Monochrome accents**: No colored accents in base theme
- **State colors**: Success (green), Warning (amber), Error (red), Info (blue)

### Typography
- **UI Font**: Space Grotesk (weights: 300–700)
- **Mono Font**: Space Mono
- **Scale**: 10px → 48px (7 steps)

### Components (CSS classes)
| Component | Classes |
|-----------|---------|
| Buttons | `.btn`, `.btn-primary`, `.btn-success`, `.btn-sm`, `.btn-lg` |
| Cards | `.card`, `.card-header`, `.card-title` |
| Forms | `.form-group`, `.form-input`, `.form-select`, `.form-textarea` |
| Badges | `.badge`, `.badge-success`, `.badge-warning`, `.badge-error` |
| Modal | `.modal-overlay`, `.modal`, `.modal-header`, `.modal-body` |
| Toast | `.toast`, `.toast-success`, `.toast-error` |
| Loading | `.loading`, `.spinner` |
| Empty State | `.empty-state`, `.empty-state-icon` |
| Navigation | `.nav-item`, `.nav-item.active` |
| Layout | `.app-container`, `.sidebar`, `.main-content`, `.header` |

### Background
- Dotted grid pattern: `radial-gradient(circle, var(--border) 1px, transparent 1px)` at 24px spacing

---

## Changelog

See [CHANGELOG.md](./CHANGELOG.md) for full version history.

### Latest: v2.7.0 (2025-02-10)
- ✅ Added FA Coaching module for SME case review
- ✅ Unlocked FA module on home page
- ✅ Consolidated all code into single-repo deployment

---

## License

Internal use only — Priceline Group.

## Author

**Naitik Joshi** — naitik.joshi@priceline.com
