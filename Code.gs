/**
 * ═══════════════════════════════════════════════════════════════════════════
 * ORION ENTERPRISE PLATFORM - CONFIG CORE
 * ═══════════════════════════════════════════════════════════════════════════
 * 
 * Consolidated configuration for all modules.
 * Merged from: config.gs (RCA), OrionConfig.gs (ORION/RT)
 * 
 * LAST UPDATED: 2024-01-10
 */

// ═══════════════════════════════════════════════════════════════════════════
// DEV_MODE — Set to true to bypass auth & use mock data for local UI testing
// No Google Sheets, Drive, or enterprise permissions required.
// Set to false before deploying to production.
// ═══════════════════════════════════════════════════════════════════════════
const DEV_MODE = true;

// ═══════════════════════════════════════════════════════════════════════════
// CONFIG - RCA Manager Configuration
// ═══════════════════════════════════════════════════════════════════════════

const CONFIG = {
  // Feature flags for RCA Manager
  FEATURES: {
    ENABLE_CASE_DISTRIBUTION: true,
    ENABLE_ADMIN_PANEL: true,
    ENABLE_TIME_TRACKING: true,
    ENABLE_KEYBOARD_SHORTCUTS: true,
    ENABLE_SEARCH_HISTORY: false,
    ENABLE_QUEUE_DASHBOARD: true,
    ENABLE_RCA_TOOLTIPS: true,
    ENABLE_CONDENSED_LAYOUT: false,
    // UI Control Flags (v3.5)
    ENABLE_ACCENT_COLORS: true  // Global accent color control
  },

  // Admin emails (use ORION_STANDARDS.ADMINS from AuthCore.gs)
  get ADMIN_EMAILS() { 
    return typeof ORION_STANDARDS !== 'undefined' ? ORION_STANDARDS.ADMINS : [
      'naitik.joshi@priceline.com',
      'naitikjoshipro@gmail.com'
    ];
  },

  // Default sheet IDs (can be overridden via Script Properties)
  // ALL sheet IDs should be here - single source of truth
  DEFAULT_SHEET_IDS: {
    // ═══════════════════════════════════════════════════════════════════════
    // LIVE DATABASE - Single source of truth
    // RCA_DATA, REP_MANAGER, and REBUTTAL_DATASHEET all use the same sheet
    // To switch environments, only update LIVE_DATASHEET
    // ═══════════════════════════════════════════════════════════════════════
    LIVE_DATASHEET: "1ApuyYWWd4GCVlwkzRsdvPsrSFw_ppFlYYh6QQFymV28",
    
    // Aliases (all point to LIVE_DATASHEET for backward compatibility)
    get RCA_DATA() { return this.LIVE_DATASHEET; },
    get REP_MANAGER() { return this.LIVE_DATASHEET; },
    get REBUTTAL_DATASHEET() { return this.LIVE_DATASHEET; },
    
    // Other module-specific sheets
    CB_CONTROLLER: "1AKQpG3Glr-mrkmNcuUCSNo8fVqN7HrqkTocypS7II7s",
    RECOVERY: "1lwO9EXKgzZmb-R162Vl_hd1Y06jdiAtaAh9t_XBBENE",
    VB_EXCEPTION: "1tRZIM8nqh0q41S4PoV1Fp5BJ0RGGjVsNdFi-p_cSsdA",
    EMAIL_CONFIRMATION: "1Q430kJMCMBbaiif1hFesbPUELtNyH1PrL9h3Uq8ApGw",
    CANNED_RESPONSES: "1p8kXzVAIB0AvZ1V6EPTLW9MIwjxEwNPL2DADz9X3BjE",
    
    // Rebuttal Generator template (separate from datasheet)
    REBUTTAL_TEMPLATE: "1v5ZPUvBozEGuaEiv1jndad9dvGm3XDA6CNmMFyAwLqU"
  },

  // Default folder IDs (can be overridden via Script Properties)
  // ALL folder IDs should be here - single source of truth
  DEFAULT_FOLDER_IDS: {
    // RCA Screenshots
    IMAGES: "1Z1Z4cRYynbjbZKDzd0dm3jumbdYWZTGT",
    CONFIRM_PAGE: "18GwZnRI2z2kw9nYQVoml20x4VybE1EBd",
    CHECKOUT_PAGE: "1Xkd75sjNAWcFclQ7GdK2wiU0u3j9Ns81",
    WEX_REFUND: "1QVXlVO5lRRXxdJGbKRK34oICv-ctR7zV",
    ACCERTIFY: "1m9NVneBrUEVhUHlXK2At0gQqLhM0PZTr",
    
    // Rebuttal Generator
    REBUTTAL_TEMP: "1-8h4Cugezu0rGoe8EgiCy50pY87tifgK",
    REBUTTAL_PDF: "1Za-O7QgYY4cQqxpulhcgWQSCySG70xAV"
  },

  // Case type categories
  CASE_TYPES: {
    FRAUD: [
      "Risk - Fraudulent Transaction",
      "Risk - Fraudulent Transaction (w/ Contact)",
      "Other - Chip liability shift (fraud)"
    ],
    SERVICE: [
      "Service - Cancel Not Performed",
      "Customer Service Error",
      "Sales - Data Entry Issue (Online Reservation Only)",
      "Sales - Incorrect City/State/Property Booked"
    ],
    CUSTOMER: [
      "Customer - Cancel policy dispute/Non Refundable",
      "Customer - Cancel policy dispute",
      "Customer - Cancel policy dispute (No Contact / Confirmed)",
      "Customer - Cancel Penalty Dispute/Early Checkout",
      "Customer - Cancel policy dispute/No Show",
      "Customer - Taxes and Fees",
      "Customer - Currency",
      "Customer - Charge not recognized",
      "Customer - Booking Error",
      "Customer - Unable to check in/Customer Issue (Age, CC)",
      "Customer - No Receipt"
    ],
    HOTEL: [
      "Hotel - Unsatisfactory",
      "Hotel - Billed Direct",
      "Hotel - Hotel Billed Direct (Wex not charged/Direct Bill)",
      "Hotel - Hotel revealed rate",
      "Hotel - Room Type Not Available",
      "Hotel - Closed/Renovation"
    ]
  },

  SESSION: { STALE_MINUTES: 30, HEARTBEAT_INTERVAL_MS: 60000 },
  UI: { DEFAULT_THEME: 'dark', SEARCH_HISTORY_LIMIT: 10, TOAST_DURATION_MS: 3000 }
};

// ═══════════════════════════════════════════════════════════════════════════
// Config - ORION/Reporting Configuration
// ═══════════════════════════════════════════════════════════════════════════

const Config = {
  // ORION Spreadsheet
  SPREADSHEET_ID: '1fbSOgULlPsxE6Mc39gJrQDvn5wf4PfgiCpKBKpz-7eQ',
  
  SHEETS: {
    REPORTS: 'RT_Reports',
    AUDIT_LOG: 'RT_AuditLog',
    CONFIG: 'RT_Config',
    USERS: 'RT_Users',
    USER_PROCESS_ROLES: 'RT_UserProcessRoles'
  },

  AI_ENABLED: false,
  
  FEATURE_FLAGS: {
    ALLOW_INITIALIZE_SHEETS: true,
    ALLOW_IMPORT_REPORTS: true,
    ALLOW_IMPORT_USERS: true,
    ALLOW_BULK_ACTIONS: false
  },

  VALID_EMAIL_DOMAIN: 'priceline.com',

  MODULES: {
    RT: { id: 'RT', name: 'Reporting', icon: '📈', enabled: true },
    ON: { id: 'ON', name: 'Operations', icon: '📁', enabled: false },
    CB: { id: 'CB', name: 'Chargeback', icon: '📋', enabled: true },
    CBQA: { id: 'CBQA', name: 'Chargeback QA', icon: '✅', enabled: false },
    FAQ: { id: 'FAQ', name: 'FA Coaching', icon: '🔍', enabled: true },
    FAQA: { id: 'FAQA', name: 'Fraud Analysis QA', icon: '📊', enabled: false }
  },

  PROCESS_ROLES: { VIEWER: 'Viewer', ANALYST: 'Analyst', MANAGER: 'Manager', ADMIN: 'Admin' },
  CLASSIFICATIONS: ['Master', 'Derived'],
  
  FREQUENCIES: [
    'Daily', 'Weekly', 'Monday', 'Friday', 'Twice a Week', 'Twice a Month',
    'Biweekly', 'Monthly', 'Hourly / End of Day Summary', 'Daily/As Triggered',
    'Daily (Raw Extract)', 'As Triggered', 'Monthly/As Required', 'Daily / As Required'
  ],

  STATES: { PENDING: 'Pending', ACCEPTED: 'Accepted', REJECTED: 'Rejected' },
  TEAMS: ['Reporting Team', 'Chargeback Team', 'Fraud Prevention Team', 'SME', 'SME (Ops)'],
  RACI_ROLES: { R: 'Responsible', A: 'Accountable', C: 'Consulted', I: 'Informed' },

  REPORTS_SCHEMA: [
    'id', 'name', 'classification', 'description', 'frequency', 'team',
    'responsible', 'accountable', 'consulted', 'informed', 'qaResponsible',
    'finalAccountability', 'owner', 'trainedResources', 'requestedBy',
    'dataSource', 'upstreamDeps', 'downstreamDeps', 'targetRecipients',
    'storageLocation', 'storageLinkType', 'state', 'pendingApprovals',
    'approvalHistory', 'comments', 'isActive', 'createdBy', 'createdAt', 'updatedBy', 'updatedAt'
  ],

  USERS_SCHEMA: ['id', 'email', 'firstName', 'lastName', 'displayName', 'active', 'createdAt', 'updatedAt'],
  USER_PROCESS_ROLES_SCHEMA: ['id', 'email', 'processId', 'role', 'createdAt', 'updatedAt'],
  AUDIT_SCHEMA: ['id', 'timestamp', 'user', 'action', 'module', 'entityType', 'entityId', 'details', 'previousValue', 'newValue'],

  AUDIT_ACTIONS: {
    CREATE: 'CREATE', UPDATE: 'UPDATE', DELETE: 'DELETE', STATE_CHANGE: 'STATE_CHANGE',
    APPROVE: 'APPROVE', REJECT: 'REJECT', COMMENT: 'COMMENT', DEACTIVATE: 'DEACTIVATE',
    ACTIVATE: 'ACTIVATE', VIEW: 'VIEW', LOGIN: 'LOGIN', USER_CREATE: 'USER_CREATE',
    USER_UPDATE: 'USER_UPDATE', ROLE_ASSIGN: 'ROLE_ASSIGN', EXPORT: 'EXPORT'
  }
};

// ═══════════════════════════════════════════════════════════════════════════
// SCRIPT PROPERTIES MANAGEMENT
// ═══════════════════════════════════════════════════════════════════════════

function getSheetId(key) {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('SHEET_' + key) || CONFIG.DEFAULT_SHEET_IDS[key];
}

function getFolderId(key) {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('FOLDER_' + key) || CONFIG.DEFAULT_FOLDER_IDS[key];
}

function getSheetConfig() {
  if (!isAdmin()) return { error: "Unauthorized" };
  
  const props = PropertiesService.getScriptProperties();
  const config = { sheets: {}, folders: {} };
  
  for (const key of Object.keys(CONFIG.DEFAULT_SHEET_IDS)) {
    config.sheets[key] = {
      current: props.getProperty('SHEET_' + key) || CONFIG.DEFAULT_SHEET_IDS[key],
      default: CONFIG.DEFAULT_SHEET_IDS[key],
      isCustom: !!props.getProperty('SHEET_' + key)
    };
  }
  
  for (const key of Object.keys(CONFIG.DEFAULT_FOLDER_IDS)) {
    config.folders[key] = {
      current: props.getProperty('FOLDER_' + key) || CONFIG.DEFAULT_FOLDER_IDS[key],
      default: CONFIG.DEFAULT_FOLDER_IDS[key],
      isCustom: !!props.getProperty('FOLDER_' + key)
    };
  }
  
  return config;
}

function updateSheetId(key, newId) {
  if (!isAdmin()) return { success: false, error: "Unauthorized" };
  if (!CONFIG.DEFAULT_SHEET_IDS.hasOwnProperty(key)) return { success: false, error: "Invalid sheet key" };
  
  try {
    SpreadsheetApp.openById(newId);
    PropertiesService.getScriptProperties().setProperty('SHEET_' + key, newId);
    return { success: true, message: `${key} updated to ${newId}` };
  } catch (e) {
    return { success: false, error: "Cannot access sheet: " + e.message };
  }
}

function updateFolderId(key, newId) {
  if (!isAdmin()) return { success: false, error: "Unauthorized" };
  if (!CONFIG.DEFAULT_FOLDER_IDS.hasOwnProperty(key)) return { success: false, error: "Invalid folder key" };
  
  try {
    DriveApp.getFolderById(newId);
    PropertiesService.getScriptProperties().setProperty('FOLDER_' + key, newId);
    return { success: true, message: `${key} updated to ${newId}` };
  } catch (e) {
    return { success: false, error: "Cannot access folder: " + e.message };
  }
}

function resetSheetId(key) {
  if (!isAdmin()) return { success: false, error: "Unauthorized" };
  PropertiesService.getScriptProperties().deleteProperty('SHEET_' + key);
  return { success: true, message: `${key} reset to default` };
}

function resetFolderId(key) {
  if (!isAdmin()) return { success: false, error: "Unauthorized" };
  PropertiesService.getScriptProperties().deleteProperty('FOLDER_' + key);
  return { success: true, message: `${key} reset to default` };
}

function isAdmin() {
  if (!CONFIG.FEATURES.ENABLE_ADMIN_PANEL) return false;
  const email = Session.getActiveUser().getEmail();
  return CONFIG.ADMIN_EMAILS.includes(email);
}

function getClientConfig() {
  return {
    features: CONFIG.FEATURES,
    isAdmin: isAdmin(),
    ui: CONFIG.UI,
    userEmail: Session.getActiveUser().getEmail()
  };
}

function getCaseTypeCategory(rcaValue) {
  for (const [category, rcaList] of Object.entries(CONFIG.CASE_TYPES)) {
    if (rcaList.includes(rcaValue)) return category;
  }
  return 'OTHER';
}

// Freeze configs
Object.freeze(CONFIG);
Object.freeze(CONFIG.FEATURES);
Object.freeze(CONFIG.DEFAULT_SHEET_IDS);
Object.freeze(CONFIG.DEFAULT_FOLDER_IDS);
Object.freeze(Config);
Object.freeze(Config.SHEETS);
Object.freeze(Config.MODULES);
Object.freeze(Config.STATES);
Object.freeze(Config.RACI_ROLES);
Object.freeze(Config.AUDIT_ACTIONS);
/**
 * ═══════════════════════════════════════════════════════════════════════════
 * ORION ENTERPRISE PLATFORM - VERSION & NAMING SCHEMA
 * ═══════════════════════════════════════════════════════════════════════════
 * 
 * Semantic Versioning: MAJOR.MINOR.PATCH
 * - MAJOR: Breaking changes (incompatible API changes)
 * - MINOR: New features (backwards compatible)
 * - PATCH: Bug fixes (backwards compatible)
 * 
 * LAST UPDATED: 2024-01-10
 */

const VERSION = {
  // Current version
  MAJOR: 2,
  MINOR: 6,
  PATCH: 0,
  
  // Build metadata
  BUILD_DATE: '2025-01-10',
  BUILD_NUMBER: 6,
  
  // Get version string
  get string() {
    return `${this.MAJOR}.${this.MINOR}.${this.PATCH}`;
  },
  
  get full() {
    return `v${this.string} (${this.BUILD_DATE})`;
  }
};

/**
 * ═══════════════════════════════════════════════════════════════════════════
 * CHANGELOG
 * ═══════════════════════════════════════════════════════════════════════════
 * 
 * v2.5.0 (2024-01-10) - Route Fixes & Stability
 * ------------------------------------------------
 * - Restored HTML pages deleted in v2.4.0 (caused route errors)
 * - Total: 19 code files (6 .gs + 13 .html)
 * - All pages use consistent ORION dark theme
 * - Proper file structure for GAS compatibility
 * 
 * v2.3.0 (2024-01-10) - Final Consolidation
 * ------------------------------------------------
 * - Merged RepManager.gs + RebuttalGenerator.gs into RCACore.gs
 * - Moved docs back to root (README.md, ORION_PROMPT.md)
 * - Total: 22 code files (9 .gs + 13 .html)
 * 
 * v2.2.0 (2024-01-10) - GitHub Structure & Partial Consolidation
 * ------------------------------------------------
 * - Created docs/ folder (README.md, ORION_PROMPT.md)
 * - Merged 8 HTML partials → 2 (RCAPartials, RTPartials)
 * - Total: 24 code files (13 HTML + 11 .gs)
 * - Updated MainRCAPage.html includes
 * 
 * v2.1.0 (2024-01-10) - Universal Styling System
 * ------------------------------------------------
 * - Created Theme.html: One-line import for complete design system
 * - Created ORION_PROMPT.md: Context reset prompt for AI assistants
 * - Unified color system, typography, spacing, components
 * - Dark/light mode toggle support
 * 
 * v2.0.0 (2024-01-10) - Enterprise Consolidation
 * ------------------------------------------------
 * - Consolidated 45 files → 29 files
 * - Merged 32 .gs files → 11 .gs files
 * - Created unified InfoPage.html for About/Privacy/DataRetention
 * - Fixed isAdmin() case-insensitive comparison
 * - Hidden Word Analyser for non-admins (not grayed out)
 * - Implemented naming schema and versioning
 * 
 * v1.0.0 (2024-01-09) - Initial Release
 * ------------------------------------------------
 * - 6-module landing page (CB, CBQA, FAQ, FAQA, RT, ON)
 * - Chargeback sub-modules (Rebuttal Engine, Rep Manager, Word Analyser)
 * - AI-powered text classification (GPT-4o-mini)
 * - Reporting governance with RACI workflow
 * - ORION dark theme UI
 */

/**
 * ═══════════════════════════════════════════════════════════════════════════
 * NAMING SCHEMA
 * ═══════════════════════════════════════════════════════════════════════════
 * 
 * FILE NAMING CONVENTIONS:
 * 
 * JavaScript (.gs) Files:
 * ─────────────────────────
 * Format: [Module]Core.gs or [Purpose].gs
 * 
 * Current Files (11 total):
 * 1. Code.gs           - Main entry point, router, doGet
 * 2. AuthCore.gs       - Authentication, Authorization, ADMINS
 * 3. ConfigCore.gs     - CONFIG (RCA) + Config (ORION)
 * 4. Utils.gs          - Shared utilities + SheetsService
 * 5. RCACore.gs        - RCA: Case mgmt, recovery, upload, admin
 * 6. CBQCore.gs        - CBQ: Config, Controller, all DAOs
 * 7. AICore.gs         - AI: Processor, TextExtractor, Tokenizer, Metrics, Batch
 * 8. ReportingCore.gs  - RT: Workflow, Importer, all DAOs
 * 9. RepManager.gs     - Representation Manager logic
 * 10. RebuttalGenerator.gs - PDF rebuttal generation
 * 11. VERSION.gs       - Versioning, changelog, naming schema
 * 
 * HTML Files:
 * ─────────────────────────
 * Format: [Purpose]Page.html or [Module]UI.html
 * 
 * Pages (Full page templates):
 * - Home.html          - 6-module landing page
 * - CBModule.html      - CB sub-module selector
 * - MainRCAPage.html   - Rebuttal Engine UI
 * - RepManagerUI.html  - Representation Manager UI
 * - CBQMain.html       - Word Analyser UI
 * - Dashboard.html     - RT Dashboard
 * - ReportsQueue.html  - RT Reports queue
 * - ReportDetail.html  - RT Single report view
 * - AdminPanel.html    - Admin management
 * - InfoPage.html      - About/Privacy/DataRetention (unified)
 * 
 * Partials (Included in pages):
 * - Theme.html         - UNIVERSAL STYLING (use this!)
 * - JavaScript.html    - RCA client-side JS
 * - Stylesheet.html    - RCA styles (legacy)
 * - Styles.html        - RT styles (legacy)
 * - RCAResponse.html   - RCA response templates
 * - AppJS.html         - RT client-side JS
 * - ApiJS.html         - RT API handlers
 * - QueueJS.html       - RT queue JS
 * 
 * FUNCTION NAMING:
 * ─────────────────────────
 * - Public API: camelCase (getReports, createUser)
 * - Private/Internal: camelCase_ (isAdmin_, validateData_)
 * - Constants: SCREAMING_SNAKE (CONFIG, ADMINS)
 * - Objects/Modules: PascalCase (Authorization, Auth)
 * 
 * OBJECT NAMING:
 * ─────────────────────────
 * - Config objects: PascalCase or all-caps (Config, CONFIG)
 * - Module objects: PascalCase (Authorization, Auth, Utils)
 * - DAO objects: [Entity]DAO pattern (ReportsDAO, UsersDAO)
 */

/**
 * Get current version info
 * @returns {Object} Version information
 */
function getVersion() {
  return {
    version: VERSION.string,
    full: VERSION.full,
    buildDate: VERSION.BUILD_DATE
  };
}

/**
 * Log version on script load
 */
console.log(`ORION Platform ${VERSION.full}`);

// Freeze version
Object.freeze(VERSION);
/**
 * ═══════════════════════════════════════════════════════════════════════════
 * ORION ENTERPRISE PLATFORM - AUTH CORE
 * ═══════════════════════════════════════════════════════════════════════════
 * 
 * Consolidated authentication, authorization, and security module.
 * Merged from: Auth.gs, Authorization.gs, GlobalStandards.gs
 * 
 * LAST UPDATED: 2024-01-10
 */

// ═══════════════════════════════════════════════════════════════════════════
// GLOBAL STANDARDS & CONVENTIONS
// ═══════════════════════════════════════════════════════════════════════════

const ORION_STANDARDS = {
  
  // Naming conventions documentation
  NAMING: {
    // File naming: PascalCase for .gs and .html files
    // Function naming: camelCase, private: underscore suffix
    // Constants: SCREAMING_SNAKE_CASE
    // Module prefixes: CBQ*, Rep*, RT*
  },
  
  // Security rules documentation
  SECURITY: {
    // API Keys: NEVER hardcode. Use Script Properties.
    // Admin checks: Always verify isAdmin before sensitive operations
    // Input sanitization: Escape HTML for display
  },
  
  // Admin email whitelist
  ADMINS: [
    'naitik.joshi@priceline.com',
    'naitikjoshipro@gmail.com'
  ],
  
  // Script property keys
  PROPERTY_KEYS: {
    OPENAI_API_KEY: 'OPENAI_API_KEY',
    RCA_DATA_SHEET_ID: 'RCA_DATA_SHEET_ID',
    CB_CONTROLLER_SHEET_ID: 'CB_CONTROLLER_SHEET_ID',
    CBQ_SPREADSHEET_ID: 'CBQ_SPREADSHEET_ID',
    ORION_SPREADSHEET_ID: 'ORION_SPREADSHEET_ID',
    IMAGES_FOLDER_ID: 'IMAGES_FOLDER_ID',
    REBUTTALS_FOLDER_ID: 'REBUTTALS_FOLDER_ID'
  }
};

// ═══════════════════════════════════════════════════════════════════════════
// AUTHORIZATION - RBAC Implementation
// ═══════════════════════════════════════════════════════════════════════════

const Authorization = {
  
  // Use centralized admin list
  get ADMINS() { return ORION_STANDARDS.ADMINS; },
  
  // Role hierarchy (higher number = more access)
  ROLE_LEVELS: {
    VIEWER: 1,
    ANALYST: 2,
    MANAGER: 3,
    LEADER: 4,
    ADMIN: 5
  },
  
  /**
   * Check if user is an admin
   */
  isAdmin(email) {
    if (!email) return false;
    const lowerEmail = email.toLowerCase();
    return ORION_STANDARDS.ADMINS.some(admin => admin.toLowerCase() === lowerEmail);
  },
  
  /**
   * Get user's roles based on email domain
   */
  getUserRoles(email) {
    if (!email) return [];
    
    const roles = [];
    
    if (Authorization.isAdmin(email)) {
      roles.push('Admin');
    }
    
    const domain = email.split('@')[1];
    if (domain === 'priceline.com') {
      roles.push('Internal');
    }
    
    if (roles.length === 0) {
      roles.push('Viewer');
    }
    
    return roles;
  },
  
  /**
   * Get user's permission level
   */
  getUserLevel(email) {
    if (Authorization.isAdmin(email)) {
      return Authorization.ROLE_LEVELS.ADMIN;
    }
    return Authorization.ROLE_LEVELS.ANALYST;
  },
  
  /**
   * Get all permissions for a user
   */
  getUserPermissions(email) {
    const isAdmin = Authorization.isAdmin(email);
    const level = Authorization.getUserLevel(email);
    
    return {
      canViewReports: level >= Authorization.ROLE_LEVELS.VIEWER,
      canEditReports: level >= Authorization.ROLE_LEVELS.ANALYST,
      canApproveReports: level >= Authorization.ROLE_LEVELS.MANAGER,
      canDeleteReports: isAdmin,
      canEditWorkflows: isAdmin || level >= Authorization.ROLE_LEVELS.LEADER,
      canManageUsers: isAdmin,
      canViewAuditLog: isAdmin,
      canManageConfig: isAdmin,
      canAccessAdmin: isAdmin
    };
  },
  
  /**
   * Check if user can perform action on a report
   */
  canPerformAction(email, action, report) {
    const permissions = Authorization.getUserPermissions(email);
    
    switch (action) {
      case 'view': return permissions.canViewReports;
      case 'edit': 
        return permissions.canEditReports && 
               (Authorization.isAdmin(email) || Authorization.isResponsibleFor(email, report));
      case 'approve':
      case 'reject':
        return permissions.canApproveReports && Authorization.isInApprovalChain(email, report);
      case 'delete': return permissions.canDeleteReports;
      default: return false;
    }
  },
  
  /**
   * Check if user is responsible for a report
   */
  isResponsibleFor(email, report) {
    if (!report || !email) return false;
    if (report.createdBy === email) return true;
    
    const responsible = report.responsible || '';
    if (responsible.includes('@')) {
      return responsible.toLowerCase() === email.toLowerCase();
    }
    return false;
  },
  
  /**
   * Check if user is in the approval chain
   */
  isInApprovalChain(email, report) {
    if (!report || !email) return false;
    if (Authorization.isAdmin(email)) return true;
    
    const userName = Utils.extractNameFromEmail(email).toLowerCase();
    const checkFields = ['accountable', 'consulted', 'informed', 'qaResponsible'];
    
    for (const field of checkFields) {
      const value = (report[field] || '').toLowerCase();
      if (value.includes(email.toLowerCase()) || value.includes(userName)) {
        return true;
      }
    }
    return false;
  },
  
  /**
   * Get reports requiring user's approval
   */
  getReportsRequiringApproval(email, reports) {
    return reports.filter(report => {
      if (report.state !== Config.STATES.PENDING) return false;
      if (!Authorization.isInApprovalChain(email, report)) return false;
      
      const approvalHistory = Utils.safeParseJSON(report.approvalHistory, []);
      return !approvalHistory.some(a => a.user === email && a.action === 'approve');
    });
  },
  
  /**
   * Get reports user is responsible for
   */
  getMyReports(email, reports) {
    return reports.filter(report => Authorization.isResponsibleFor(email, report));
  },
  
  /**
   * Require specific permission, throw if not allowed
   */
  require(email, permission) {
    const permissions = Authorization.getUserPermissions(email);
    if (!permissions[permission]) {
      throw new Error(`Access denied: ${permission} required`);
    }
  }
};

// ═══════════════════════════════════════════════════════════════════════════
// AUTHENTICATION - Google Workspace Identity
// ═══════════════════════════════════════════════════════════════════════════

const Auth = {
  
  /**
   * Get current user's email
   */
  getCurrentUser() {
    try {
      const user = Session.getActiveUser().getEmail();
      if (!user) {
        return Session.getEffectiveUser().getEmail();
      }
      return user;
    } catch (e) {
      console.error('Failed to get current user:', e);
      return null;
    }
  },
  
  /**
   * Get user's display name from email
   */
  getUserDisplayName(email) {
    const userEmail = email || Auth.getCurrentUser();
    return Utils.extractNameFromEmail(userEmail);
  },
  
  /**
   * Get user's first name
   */
  getFirstName(email) {
    const userEmail = email || Auth.getCurrentUser();
    return Utils.extractFirstName(userEmail);
  },
  
  /**
   * Check if user is authenticated
   */
  isAuthenticated() {
    const user = Auth.getCurrentUser();
    return !!user && user.length > 0;
  },
  
  /**
   * Get complete user context
   */
  getUserContext() {
    const email = Auth.getCurrentUser();
    if (!email) return null;
    
    return {
      email: email,
      displayName: Auth.getUserDisplayName(email),
      firstName: Auth.getFirstName(email),
      isAdmin: Authorization.isAdmin(email),
      roles: Authorization.getUserRoles(email),
      permissions: Authorization.getUserPermissions(email)
    };
  }
};

// ═══════════════════════════════════════════════════════════════════════════
// GLOBAL UTILITY FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════

function getCurrentUser_() {
  if (DEV_MODE) return 'dev@priceline.com';
  try {
    return Session.getActiveUser().getEmail() || '';
  } catch (e) {
    return '';
  }
}

function isCurrentUserAdmin_() {
  const email = getCurrentUser_();
  return ORION_STANDARDS.ADMINS.includes(email);
}

function getScriptProperty_(key, fallback) {
  try {
    const value = PropertiesService.getScriptProperties().getProperty(key);
    return value || fallback || '';
  } catch (e) {
    return fallback || '';
  }
}

function escapeHtml_(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function isValidEmail_(email) {
  if (!email) return false;
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function generateId_() {
  return Utilities.getUuid();
}

function getTimestamp_() {
  return new Date().toISOString();
}

// ═══════════════════════════════════════════════════════════════════════════
// FREEZE OBJECTS
// ═══════════════════════════════════════════════════════════════════════════

Object.freeze(ORION_STANDARDS);
Object.freeze(ORION_STANDARDS.ADMINS);
Object.freeze(ORION_STANDARDS.PROPERTY_KEYS);
Object.freeze(Authorization);
Object.freeze(Authorization.ROLE_LEVELS);
Object.freeze(Auth);
/**
 * ═══════════════════════════════════════════════════════════════════════════
 * ORION ENTERPRISE PLATFORM - UNIFIED CODE.gs
 * ═══════════════════════════════════════════════════════════════════════════
 * 
 * Main entry point for the unified enterprise platform containing:
 * - Rebuttal Engine (Case management & rebuttal workflow)
 * - Representation Manager (Case representation)
 * - ORION RT (Reporting governance)
 * - ORION CBQ (Chargeback analytics with AI)
 * - Rebuttal Generator (PDF generation)
 * 
 * Single doGet() router serves all modules.
 */

// ═══════════════════════════════════════════════════════════════════════════
// WEB APP ENTRY POINT
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Serves the web application
 * @param {Object} e - Event object with query parameters
 * @returns {HtmlOutput} The HTML page
 */
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'home';
  
  // Get current user (DEV_MODE: skip Session auth)
  const user = DEV_MODE ? 'dev@priceline.com' : Session.getActiveUser().getEmail();
  if (!user) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;padding:40px;text-align:center;background:#000;color:#fff;min-height:100vh;">' +
      '<h1>Access Denied</h1><p>Please sign in with your Google Workspace account.</p></div>'
    );
  }
  
  let title = 'ORION Enterprise';
  
  // Page-specific variables
  let pageType = '';
  let pageTitle = '';
  let reportId = null;
  
  try {
    // All pages served from single Index.html — just set title & page-specific vars
    switch (page) {
      case 'home':
      case 'orion':
        title = 'ORION Enterprise Platform';
        break;
      case 'cb':
        title = 'Chargeback Module';
        break;
      case 'rca':
        title = 'Rebuttal Engine';
        break;
      case 'rep':
      case 'rep-new':
        title = 'Representation Manager';
        break;
      case 'queue':
        title = 'ORION - Reports Queue';
        break;
      case 'report':
        reportId = (e && e.parameter && e.parameter.id) || null;
        title = 'ORION - Report Details';
        break;
      case 'dashboard':
        title = 'ORION - Dashboard';
        break;
      case 'admin':
        title = 'ORION - Admin Panel';
        break;
      case 'email-confirm':
        title = 'Email Confirmations';
        break;
      case 'cbq':
        title = 'ORION - Text Analytics';
        break;
      case 'fa':
        title = 'FA Coaching';
        break;
      case 'about':
        pageType = 'about';
        pageTitle = 'About';
        title = 'ORION - About';
        break;
      case 'privacy':
        pageType = 'privacy';
        pageTitle = 'Privacy Policy';
        title = 'ORION - Privacy Policy';
        break;
      case 'data-retention':
        pageType = 'data-retention';
        pageTitle = 'Data Retention';
        title = 'ORION - Data Retention';
        break;
      default:
        // Default to RCA (home page equivalent for the rebuttal engine)
        title = 'Rebuttal Engine';
    }
    
    // Normalize: 'orion' → 'home', 'rep-new' → 'rep'
    const effectivePage = (page === 'orion') ? 'home' 
                        : (page === 'rep-new') ? 'rep'
                        : page;
    
    // Single template for all pages
    const template = HtmlService.createTemplateFromFile('Index');
    
    // Set template variables used by Index.html conditional blocks
    template.page = effectivePage;
    template.user = user;
    template.isAdmin = isAdmin_(user);
    template.featureFlags = Config.FEATURE_FLAGS || {
      ALLOW_INITIALIZE_SHEETS: true,
      ALLOW_IMPORT_REPORTS: true,
      ALLOW_IMPORT_USERS: true,
      ALLOW_BULK_ACTIONS: false
    };
    template.reportId = reportId;
    template.pageType = pageType;
    template.pageTitle = pageTitle;
    
    return template.evaluate()
      .setTitle(title)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setFaviconUrl('https://drive.google.com/uc?id=1XOOwMdi1d6IVYlDV0A9grhfXbWGyRd7X&export=download&format=png');
      
  } catch (error) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:monospace;padding:40px;background:#1a0000;color:#ff6b6b;min-height:100vh;">' +
      '<h1>Error Loading Page</h1>' +
      '<p>Page: ' + page + '</p>' +
      '<p>Error: ' + error.message + '</p>' +
      '<pre>' + error.stack + '</pre></div>'
    );
  }
}

/**
 * Check if user is admin
 * Priority: Authorization module → ORION_STANDARDS → CONFIG
 */
function isAdmin_(user) {
  if (DEV_MODE) return true;
  try {
    // Try Authorization module first
    if (typeof Authorization !== 'undefined' && Authorization.isAdmin) {
      return Authorization.isAdmin(user);
    }
  } catch (e) { /* fallback */ }
  
  // Use centralized admin list from GlobalStandards
  if (typeof ORION_STANDARDS !== 'undefined' && ORION_STANDARDS.ADMINS) {
    return ORION_STANDARDS.ADMINS.includes(user);
  }
  
  // Final fallback to CONFIG
  const admins = (typeof CONFIG !== 'undefined' && CONFIG.ADMIN_EMAILS) || [];
  return admins.includes(user);
}

// ═══════════════════════════════════════════════════════════════════════════
// SPREADSHEET MENU
// ═══════════════════════════════════════════════════════════════════════════

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('ORION Platform')
    .addItem('🔗 Open Rebuttal Engine', 'openRCAManager')
    .addItem('📊 Open ORION Dashboard', 'openORION')
    .addItem('📁 Open Rep Manager', 'openRepManager')
    .addSeparator()
    .addItem('⚙️ Initialize Sheets', 'initializeSheets')
    .addToUi();
}

function openRCAManager() {
  openPage_('rca');
}

function openORION() {
  openPage_('orion');
}

function openRepManager() {
  openPage_('rep');
}

function openPage_(page) {
  const url = ScriptApp.getService().getUrl() + '?page=' + page;
  const html = HtmlService.createHtmlOutput(
    `<script>window.open('${url}', '_blank');google.script.host.close();</script>`
  );
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening...');
}

function initializeSheets() {
  if (DEV_MODE) return { success: true, message: 'DEV_MODE: Sheets init skipped' };
  try {
    SheetsService.initializeAllSheets();
    SpreadsheetApp.getUi().alert('Sheets initialized successfully!');
  } catch (e) {
    console.log('Sheets initialized');
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// DEV_MODE MOCK DATA OVERRIDES
// When DEV_MODE = true, all google.script.run endpoints return mock data
// so the UI can be tested without Google Sheets / Drive access.
// ═══════════════════════════════════════════════════════════════════════════

if (typeof DEV_MODE !== 'undefined' && DEV_MODE) {
  
  // ── Mock data templates ──────────────────────────────────────────────
  const MOCK_USER = 'dev@priceline.com';
  const MOCK_DATE = new Date().toISOString();
  
  const MOCK_REPORT = {
    id: 'RPT-001', title: 'Monthly Chargeback Analysis', state: 'Pending',
    type: 'Analysis', priority: 'High', description: 'Monthly review of chargeback patterns.',
    created: MOCK_DATE, modified: MOCK_DATE, createdBy: MOCK_USER,
    responsible: MOCK_USER, accountable: 'manager@priceline.com',
    consulted: '', informed: '', approvalHistory: '[]', comments: '[]',
    process: 'Standard', frequency: 'Monthly'
  };
  
  const MOCK_REPORT_2 = {
    id: 'RPT-002', title: 'Fraud Detection Rate Review', state: 'Accepted',
    type: 'Review', priority: 'Medium', description: 'Quarterly fraud detection metrics.',
    created: MOCK_DATE, modified: MOCK_DATE, createdBy: 'analyst@priceline.com',
    responsible: 'analyst@priceline.com', accountable: MOCK_USER,
    consulted: '', informed: '', approvalHistory: '[]', comments: '[]',
    process: 'Standard', frequency: 'Quarterly'
  };
  
  const MOCK_CASE = {
    portal: 'EXP', itn: '12345678901234', rcaValue: 'Customer - Cancel policy dispute',
    assignedUser: MOCK_USER, status: 'Pending', priority: 'Normal',
    bookingDate: '2024-01-15', checkInDate: '2024-02-01', checkOutDate: '2024-02-03',
    hotelName: 'Grand Hyatt Demo Hotel', guestName: 'John Demo',
    chargebackAmount: 250.00, chargebackCurrency: 'USD', reasonCode: '13.1',
    cardType: 'Visa', caseId: 'CB-2024-001', representmentDeadline: '2024-03-01',
    notes: 'Demo case for UI testing', response: '', row: 2,
    rcaCategory: 'Customer', weight: 3
  };
  
  const MOCK_CASE_2 = {
    portal: 'BKG', itn: '98765432109876', rcaValue: 'Hotel - Billed Direct',
    assignedUser: 'analyst@priceline.com', status: 'Assigned', priority: 'High',
    bookingDate: '2024-01-20', checkInDate: '2024-03-01', checkOutDate: '2024-03-05',
    hotelName: 'Marriott Demo Resort', guestName: 'Jane Sample',
    chargebackAmount: 475.50, chargebackCurrency: 'USD', reasonCode: '13.3',
    cardType: 'Mastercard', caseId: 'CB-2024-002', representmentDeadline: '2024-04-01',
    notes: 'Another demo case', response: '', row: 3,
    rcaCategory: 'Hotel', weight: 5
  };

  const MOCK_COACHING_CASE = {
    row: 2, date: '2024-01-15', analyst: 'analyst@priceline.com',
    itn: '12345678901234', portal: 'EXP', hotelName: 'Grand Hyatt Demo',
    reasonCode: '13.1', outcome: 'Won', amount: 250.00,
    qaScore: 95, smeNotes: '', reviewedBy: ''
  };
  
  // ── ORION RT (Reports/Queue/Dashboard) ───────────────────────────────
  getUserInfo = function() {
    return { email: MOCK_USER, displayName: 'dev', isAdmin: true };
  };
  getReports = function(_filters) {
    return [MOCK_REPORT, MOCK_REPORT_2];
  };
  getReportById = function(id) {
    return id === 'RPT-002' ? MOCK_REPORT_2 : MOCK_REPORT;
  };
  updateReport = function(_id, _data) {
    return { success: true, message: 'DEV_MODE: Report updated' };
  };
  changeReportState = function(_id, _state, _comment) {
    return { success: true, message: 'DEV_MODE: State changed' };
  };
  createReport = function(_data) {
    return { success: true, id: 'RPT-NEW-' + Date.now() };
  };
  deleteReport = function(_id) {
    return { success: true, message: 'DEV_MODE: Report deleted' };
  };
  getReportCounts = function() {
    return { total: 2, pending: 1, accepted: 1, rejected: 0, draft: 0 };
  };
  getDashboardStats = function() {
    return { totalReports: 2, byState: { pending: 1, accepted: 1, rejected: 0 } };
  };
  getFilterOptions = function() {
    return {
      types: ['Analysis', 'Review', 'Audit'],
      states: ['Pending', 'Accepted', 'Rejected', 'Draft'],
      priorities: ['High', 'Medium', 'Low'],
      owners: [MOCK_USER, 'analyst@priceline.com', 'manager@priceline.com'],
      processes: ['Standard', 'Expedited'],
      frequencies: ['Daily', 'Weekly', 'Monthly', 'Quarterly']
    };
  };
  getUsersForPicker = function() {
    return [
      { email: MOCK_USER, name: 'Dev User', role: 'Admin' },
      { email: 'analyst@priceline.com', name: 'Demo Analyst', role: 'Analyst' },
      { email: 'manager@priceline.com', name: 'Demo Manager', role: 'Manager' }
    ];
  };
  getAuditLog = function(_filters) {
    return [
      { timestamp: MOCK_DATE, user: MOCK_USER, action: 'Created', entity: 'RPT-001', details: 'Created report' },
      { timestamp: MOCK_DATE, user: 'analyst@priceline.com', action: 'Updated', entity: 'RPT-002', details: 'Changed state' }
    ];
  };
  importSampleUsers = function() { return { success: true, count: 3 }; };
  importAllReports = function() { return { success: true, count: 2 }; };
  importReportsFromData = function() { return { success: true, count: 0 }; };
  getUsers = function() {
    return [
      { email: MOCK_USER, name: 'Dev User', role: 'Admin', active: true },
      { email: 'analyst@priceline.com', name: 'Demo Analyst', role: 'Analyst', active: true }
    ];
  };
  getAllUsers = function() { return getUsers(); };
  exportReportsCSV = function() { return { success: true, url: '#' }; };
  
  // ── RCA Manager (Rebuttal Engine) ────────────────────────────────────
  searchCase = function(_itn) { return MOCK_CASE; };
  getCase = function(_portal) { return [MOCK_CASE, MOCK_CASE_2]; };
  getPendingCases = function() { return [MOCK_CASE, MOCK_CASE_2]; };
  getReservationData = function(_itn) {
    return {
      found: true, itn: '12345678901234', guestName: 'John Demo',
      hotelName: 'Grand Hyatt Demo Hotel', checkIn: '2024-02-01', checkOut: '2024-02-03',
      amount: 250.00, currency: 'USD', status: 'Confirmed'
    };
  };
  checkClaim = function(_row) { return { claimed: false, claimedBy: null }; };
  resolveRCA = function() { return { success: true }; };
  updateRefundAmount = function() { return { success: true }; };
  updateRoomType = function() { return { success: true }; };
  markInactive = function() { return { success: true }; };
  recordHeartbeat = function() { return { success: true }; };
  uploadImageToDrive = function() { return { success: true, url: 'https://via.placeholder.com/200' }; };
  uploadAllFiles = function() { return { success: true }; };
  sendForRecovery = function() { return { success: true }; };
  sendForVbException = function() { return { success: true }; };
  generateDemoData = function() { return { success: true, count: 5 }; };
  getActiveAnalysts = function() {
    return [
      { email: MOCK_USER, name: 'Dev User', casesAssigned: 3, casesResolved: 10 },
      { email: 'analyst@priceline.com', name: 'Demo Analyst', casesAssigned: 5, casesResolved: 8 }
    ];
  };
  getAnalystStats = function() {
    return [
      { analyst: MOCK_USER, assigned: 3, resolved: 10, avgTime: '2.5h' },
      { analyst: 'analyst@priceline.com', assigned: 5, resolved: 8, avgTime: '3.1h' }
    ];
  };
  getCaseDistributionSummary = function() {
    return { total: 8, byPortal: { EXP: 5, BKG: 3 }, byStatus: { Pending: 2, Assigned: 4, Resolved: 2 } };
  };
  getCannedResponses = function() {
    return [
      { id: 1, name: 'No Show Default', category: 'Customer', text: 'Guest did not show up for reservation.' },
      { id: 2, name: 'Direct Bill', category: 'Hotel', text: 'Hotel billed the guest directly.' }
    ];
  };
  createCannedRes = function() { return { success: true, id: Date.now() }; };
  saveCannedResponse = function() { return { success: true }; };
  initializeAutoCreation = function() { return { success: true }; };
  
  // ── Config / Admin ───────────────────────────────────────────────────
  getSheetConfig = function() {
    return {
      RCA_DATA: CONFIG.DEFAULT_SHEET_IDS.LIVE_DATASHEET,
      CB_CONTROLLER: CONFIG.DEFAULT_SHEET_IDS.CB_CONTROLLER,
      IMAGES: CONFIG.DEFAULT_FOLDER_IDS.IMAGES
    };
  };
  getClientConfig = function() {
    return {
      features: CONFIG.FEATURES,
      caseTypes: CONFIG.CASE_TYPES,
      sessionTimeout: CONFIG.SESSION.STALE_MINUTES
    };
  };
  updateSheetId = function() { return { success: true }; };
  resetSheetId = function() { return { success: true }; };
  updateFolderId = function() { return { success: true }; };
  resetFolderId = function() { return { success: true }; };
  createUser = function() { return { success: true }; };
  updateUser = function() { return { success: true }; };
  deactivateReport = function() { return { success: true }; };
  addComment = function() { return { success: true }; };
  agentCount = function() { return 2; };
  
  // ── Representation Manager ───────────────────────────────────────────
  rmGetPendingUploads = function(_portal, _currency) {
    return [
      { row: 2, itn: '11111111111111', portal: 'EXP', amount: 150, currency: 'USD', status: 'Pending', hotel: 'Demo Hotel A' },
      { row: 3, itn: '22222222222222', portal: 'BKG', amount: 300, currency: 'USD', status: 'Pending', hotel: 'Demo Hotel B' }
    ];
  };
  rmChallenge = function() { return { success: true }; };
  rmAccept = function() { return { success: true }; };
  rmReversal = function() { return { success: true }; };
  rmAutoclosed = function() { return { success: true }; };
  rmRecreate = function() { return { success: true }; };
  rmGetCounts = function() { return { pending: 2, challenged: 1, accepted: 3, total: 6 }; };
  rmSearchFiles = function() { return []; };
  
  // ── Email Confirmations ──────────────────────────────────────────────
  ecGetPendingItems = function() {
    return [
      { analyst: MOCK_USER, itn: '33333333', addedDt: '2024-01-10', dueDt: '2024-02-10', status: 'Pending' },
      { analyst: MOCK_USER, itn: '44444444', addedDt: '2024-01-12', dueDt: '2024-02-12', status: 'Pending' }
    ];
  };
  ecUploadFiles = function() { return { message: 'DEV_MODE: Files uploaded!', updated: [] }; };
  
  // ── CBQ (Text Analytics / Word Analyser) ─────────────────────────────
  cbqGetDocuments = function() { return []; };
  cbqGetQueueStatus = function() { return { queued: 0, processing: 0, completed: 5, failed: 0 }; };
  cbqGetStatsSummary = function() {
    return { totalDocs: 5, totalWords: 1200, uniqueWords: 340, avgWordsPerDoc: 240 };
  };
  cbqGetWordCloudData = function() {
    return [
      { word: 'chargeback', count: 45 }, { word: 'refund', count: 38 },
      { word: 'hotel', count: 32 }, { word: 'guest', count: 28 },
      { word: 'reservation', count: 25 }, { word: 'cancelled', count: 20 }
    ];
  };
  cbqGetBubbleChartData = function() {
    return [
      { word: 'fraud', x: 10, y: 40, r: 30 }, { word: 'dispute', x: 30, y: 60, r: 25 },
      { word: 'refund', x: 50, y: 30, r: 35 }
    ];
  };
  cbqGetConfig = function() { return { aiEnabled: false, maxDocs: 100 }; };
  cbqGetFilteredWords = function() { return ['the', 'a', 'an', 'is', 'was', 'are']; };
  cbqUploadText = function() { return { success: true, docId: 'DOC-' + Date.now() }; };
  cbqUploadTextWithAI = function() { return { success: true, docId: 'DOC-' + Date.now() }; };
  cbqUploadPDFWithAI = function() { return { success: true, docId: 'DOC-' + Date.now() }; };
  cbqUploadFromFolderWithAI = function() { return { success: true, count: 0 }; };
  cbqProcessQueue = function() { return { processed: 0 }; };
  cbqRecalculateStats = function() { return { success: true }; };
  cbqExportStats = function() { return { success: true, url: '#' }; };
  cbqAddFilteredWord = function() { return { success: true }; };
  cbqRemoveFilteredWord = function() { return { success: true }; };
  cbqInitializeFilteredWords = function() { return { success: true }; };
  cbqTestAIConnection = function() { return { success: true, model: 'gpt-4o-mini (mock)' }; };
  
  // ── FA Coaching ──────────────────────────────────────────────────────
  getCoachingDates = function() {
    return ['2024-01-15', '2024-01-22', '2024-01-29', '2024-02-05'];
  };
  getCoachingCases = function(_date) {
    return [MOCK_COACHING_CASE, {
      row: 3, date: '2024-01-15', analyst: 'analyst@priceline.com',
      itn: '98765432109876', portal: 'BKG', hotelName: 'Marriott Demo',
      reasonCode: '13.3', outcome: 'Lost', amount: 475.50,
      qaScore: 72, smeNotes: 'Needs follow-up', reviewedBy: MOCK_USER
    }];
  };
  saveCoachingNotes = function() { return { success: true }; };
  faSaveSmeNotes = function() { return { success: true }; };
  getUserEmail = function() { return MOCK_USER; };
  
  // ── Rebuttal Generator ───────────────────────────────────────────────
  generateRebuttals = function() { return { success: true, message: 'DEV_MODE: No rebuttals generated' }; };
  
  // ── Triggered functions (no-ops in DEV_MODE) ─────────────────────────
  sendForEmailConfirmation = function() { return { success: true }; };
}

// ═══════════════════════════════════════════════════════════════════════════
// RCA MANAGER API (google.script.run endpoints)
// ═══════════════════════════════════════════════════════════════════════════

// These are called from JavaScript.html and MainRCAPage.html
// Kept in separate files: config.gs, casemanagement.gs, admin.gs, etc.

// ═══════════════════════════════════════════════════════════════════════════
// REPRESENTATION MANAGER API
// ═══════════════════════════════════════════════════════════════════════════

function rmGetPendingUploads(portal, currency) {
  return repGetPendingUploads(portal, currency);
}

function rmChallenge(row) {
  return repChallenge(row);
}

function rmAccept(row) {
  return repAccept(row);
}

function rmReversal(row) {
  return repReversal(row);
}

function rmAutoclosed(row) {
  return repAutoclosed(row);
}

function rmRecreate(row, res) {
  return repRecreate(row, res);
}

function rmGetCounts() {
  return repGetCounts();
}

function rmSearchFiles(itn) {
  return repSearchFiles(itn);
}

// ═══════════════════════════════════════════════════════════════════════════
// EMAIL CONFIRMATIONS API
// ═══════════════════════════════════════════════════════════════════════════

const EC_CONFIG = {
  SHEET_ID: "1FqABBeWj9L8F4z_RE5DGAdOFoUMp8F6Aq1i_uj3Yy7w",
  SHEET_NAME: "data",
  FOLDER_ID: "1qybR0A3fOqE-gGp9dsKCiLFEacLPq718"
};

function ecGetPendingItems() {
  try {
    const sheet = SpreadsheetApp.openById(EC_CONFIG.SHEET_ID).getSheetByName(EC_CONFIG.SHEET_NAME);
    const agent = Session.getActiveUser().getEmail();
    const data = sheet.getDataRange().getValues();
    const items = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[4] === "Pending" && row[0] === agent) {
        items.push({
          analyst: row[0],
          itn: row[1],
          addedDt: row[2] ? Utilities.formatDate(new Date(row[2]), "CST", "yyyy-MM-dd") : "",
          dueDt: row[3] ? Utilities.formatDate(new Date(row[3]), "CST", "yyyy-MM-dd") : "",
          status: row[4]
        });
      }
    }
    return items;
  } catch (e) {
    console.error('ecGetPendingItems error:', e);
    return [];
  }
}

function ecUploadFiles(formData) {
  try {
    const folder = DriveApp.getFolderById(EC_CONFIG.FOLDER_ID);
    const sheet = SpreadsheetApp.openById(EC_CONFIG.SHEET_ID).getSheetByName(EC_CONFIG.SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const date = new Date();
    const updated = [];

    formData.files.forEach(fileData => {
      const blob = Utilities.newBlob(
        Utilities.base64Decode(fileData.content), 
        fileData.mimeType, 
        fileData.name
      );
      folder.createFile(blob).setName(fileData.name.split('.')[0].trim() + " - Email Confirmation.pdf");

      // Match filename (without extension) to ITN
      const fileITN = fileData.name.split('.')[0].trim();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][1]) === String(fileITN)) {
          sheet.getRange(i + 1, 5).setValue("Completed");
          sheet.getRange(i + 1, 6).setValue(Utilities.formatDate(date, "CST", "yyyy-MM-dd"));
          updated.push(fileITN);
          break;
        }
      }
    });

    return { message: "Files uploaded!", updated };
  } catch (e) {
    console.error('ecUploadFiles error:', e);
    throw new Error('Upload failed: ' + e.message);
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// REBUTTAL GENERATOR API
// ═══════════════════════════════════════════════════════════════════════════

function generateRebuttals() {
  return createRebuttals();
}

// ═══════════════════════════════════════════════════════════════════════════
// ORION RT API (Reports)
// ═══════════════════════════════════════════════════════════════════════════

function getUserInfo() {
  const user = Session.getActiveUser().getEmail();
  return {
    email: user,
    displayName: user.split('@')[0],
    isAdmin: isAdmin_(user)
  };
}

function getReports(filters) {
  return ReportsDAO.getReports(filters);
}

function getReportById(reportId) {
  return ReportsDAO.getReportById(reportId);
}

function updateReport(reportId, data) {
  const user = Session.getActiveUser().getEmail();
  return ReportsDAO.updateReport(reportId, data, user);
}

function changeReportState(reportId, newState, comment) {
  const user = Session.getActiveUser().getEmail();
  return ReportWorkflow.changeState(reportId, newState, user, comment);
}

function getAuditLog(filters) {
  return AuditDAO.getEntries(filters);
}

function getFilterOptions() {
  return ReportsDAO.getFilterOptions();
}

function createReport(data) {
  const user = Session.getActiveUser().getEmail();
  return ReportsDAO.createReport(data, user, isAdmin_(user));
}

function deleteReport(reportId) {
  const user = Session.getActiveUser().getEmail();
  return ReportsDAO.deleteReport(reportId, user);
}

function getReportCounts() {
  return ReportsDAO.getReportCounts();
}

function getUsersForPicker() {
  return UsersDAO.getUsersForPicker();
}

function getDashboardStats() {
  const reports = ReportsDAO.getReports({});
  const byState = {
    pending: reports.filter(r => r.state === 'Pending').length,
    accepted: reports.filter(r => r.state === 'Accepted').length,
    rejected: reports.filter(r => r.state === 'Rejected').length
  };
  return { totalReports: reports.length, byState };
}

// ═══════════════════════════════════════════════════════════════════════════
// REPORTING MANAGER - IMPORT FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════

function importSampleUsers() {
  return DataImporter.importSampleUsers();
}

function importAllReports() {
  return DataImporter.importAllReports();
}

function getUsers() {
  return UsersDAO.getAll();
}

// ═══════════════════════════════════════════════════════════════════════════
// ORION CBQ API (Chargeback Analytics)
// ═══════════════════════════════════════════════════════════════════════════

function cbqUploadText(data) { return CBQController.uploadText(data); }
function cbqUploadTextWithAI(data) { return CBQController.uploadTextWithAI(data); }
function cbqUploadPDFWithAI(data) { return CBQController.uploadPDFWithAI(data); }
function cbqUploadFromFolderWithAI(data) { return CBQController.uploadFromFolderWithAI(data); }
function cbqGetDocuments(filters) { return CBQController.getDocuments(filters); }
function cbqGetQueueStatus() { return CBQController.getQueueStatus(); }
function cbqProcessQueue(limit) { return CBQController.processQueue(limit); }
function cbqRecalculateStats(rc) { return CBQController.recalculateStats(rc); }
function cbqGetStatsSummary() { return CBQController.getStatsSummary(); }
function cbqGetWordCloudData(opts) { return CBQController.getWordCloudData(opts); }
function cbqGetBubbleChartData(opts) { return CBQController.getBubbleChartData(opts); }
function cbqGetConfig() { return CBQController.getConfig(); }
function cbqExportStats(rc) { return CBQController.exportStats(rc); }
function cbqGetFilteredWords() { return CBQController.getFilteredWords(); }
function cbqAddFilteredWord(word, cat) { return CBQController.addFilteredWord(word, cat); }
function cbqRemoveFilteredWord(word) { return CBQController.removeFilteredWord(word); }
function cbqInitializeFilteredWords() { return CBQController.initializeFilteredWords(); }
function cbqTestAIConnection() { return CBQController.testAIConnection(); }
/**
 * ØN Master System - Utility Functions
 * Shared utilities for all modules
 */

const Utils = {
  /**
   * Generate a UUID v4
   * @returns {string} UUID
   */
  generateUUID() {
    return Utilities.getUuid();
  },
  
  /**
   * Get current timestamp in ISO format
   * @returns {string} ISO timestamp
   */
  now() {
    return new Date().toISOString();
  },
  
  /**
   * Format date for display
   * @param {string|Date} date - Date to format
   * @returns {string} Formatted date
   */
  formatDate(date) {
    if (!date) return '';
    const d = new Date(date);
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  },
  
  /**
   * Format date for display (short)
   * @param {string|Date} date - Date to format
   * @returns {string} Formatted date
   */
  formatDateShort(date) {
    if (!date) return '';
    const d = new Date(date);
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMM dd, yyyy');
  },
  
  /**
   * Extract display name from email
   * e.g., "naitik.joshi@priceline.com" → "Naitik Joshi"
   * @param {string} email - Email address
   * @returns {string} Display name
   */
  extractNameFromEmail(email) {
    if (!email) return 'Unknown';
    
    const localPart = email.split('@')[0];
    const parts = localPart.split(/[._-]/);
    
    return parts
      .map(part => part.charAt(0).toUpperCase() + part.slice(1).toLowerCase())
      .join(' ');
  },
  
  /**
   * Extract first name from email
   * @param {string} email - Email address
   * @returns {string} First name
   */
  extractFirstName(email) {
    if (!email) return 'Unknown';
    const localPart = email.split('@')[0];
    const firstName = localPart.split(/[._-]/)[0];
    return firstName.charAt(0).toUpperCase() + firstName.slice(1).toLowerCase();
  },
  
  /**
   * Safely parse JSON
   * @param {string} json - JSON string
   * @param {*} defaultValue - Default value if parsing fails
   * @returns {*} Parsed value or default
   */
  safeParseJSON(json, defaultValue = null) {
    try {
      return JSON.parse(json);
    } catch (e) {
      return defaultValue;
    }
  },
  
  /**
   * Safely stringify to JSON
   * @param {*} obj - Object to stringify
   * @returns {string} JSON string
   */
  safeStringify(obj) {
    try {
      return JSON.stringify(obj);
    } catch (e) {
      return '{}';
    }
  },
  
  /**
   * Check if a string is a valid URL
   * @param {string} str - String to check
   * @returns {boolean} True if valid URL
   */
  isValidUrl(str) {
    if (!str) return false;
    try {
      new URL(str);
      return true;
    } catch (e) {
      return false;
    }
  },
  
  /**
   * Detect storage link type
   * @param {string} link - Storage link
   * @returns {string} Link type: 'gdrive', 'onedrive', 'path', or 'unknown'
   */
  detectStorageLinkType(link) {
    if (!link) return 'unknown';
    
    const lowerLink = link.toLowerCase();
    
    if (lowerLink.includes('drive.google.com') || lowerLink.includes('docs.google.com')) {
      return 'gdrive';
    }
    if (lowerLink.includes('onedrive') || lowerLink.includes('sharepoint')) {
      return 'onedrive';
    }
    if (lowerLink.startsWith('http://') || lowerLink.startsWith('https://')) {
      return 'url';
    }
    if (lowerLink.includes('>') || lowerLink.includes('/') || lowerLink.includes('\\')) {
      return 'path';
    }
    
    return 'unknown';
  },
  
  /**
   * Truncate string with ellipsis
   * @param {string} str - String to truncate
   * @param {number} maxLength - Maximum length
   * @returns {string} Truncated string
   */
  truncate(str, maxLength = 50) {
    if (!str || str.length <= maxLength) return str;
    return str.substring(0, maxLength - 3) + '...';
  },
  
  /**
   * Deep clone an object
   * @param {*} obj - Object to clone
   * @returns {*} Cloned object
   */
  deepClone(obj) {
    return JSON.parse(JSON.stringify(obj));
  },
  
  /**
   * Check if object is empty
   * @param {Object} obj - Object to check
   * @returns {boolean} True if empty
   */
  isEmpty(obj) {
    if (!obj) return true;
    if (Array.isArray(obj)) return obj.length === 0;
    if (typeof obj === 'object') return Object.keys(obj).length === 0;
    return false;
  },
  
  /**
   * Sanitize HTML to prevent XSS
   * @param {string} str - String to sanitize
   * @returns {string} Sanitized string
   */
  sanitizeHtml(str) {
    if (!str) return '';
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  },
  
  /**
   * Create a standardized error response
   * @param {string} message - Error message
   * @param {string} code - Error code
   * @returns {Object} Error object
   */
  createError(message, code = 'ERROR') {
    return {
      success: false,
      error: {
        code: code,
        message: message
      }
    };
  },
  
  /**
   * Create a standardized success response
   * @param {*} data - Response data
   * @param {string} message - Success message
   * @returns {Object} Success object
   */
  createSuccess(data, message = 'Success') {
    return {
      success: true,
      message: message,
      data: data
    };
  },
  
  /**
   * Wrap a function with error handling
   * @param {Function} fn - Function to wrap
   * @returns {Function} Wrapped function
   */
  withErrorHandling(fn) {
    return function(...args) {
      try {
        const result = fn.apply(this, args);
        return Utils.createSuccess(result);
      } catch (error) {
        console.error('Error in function:', error);
        return Utils.createError(error.message);
      }
    };
  },
  
  /**
   * Log with timestamp and context
   * @param {string} level - Log level
   * @param {string} context - Context/module name
   * @param {string} message - Log message
   * @param {*} data - Additional data
   */
  log(level, context, message, data = null) {
    const timestamp = Utils.formatDate(new Date());
    const logMessage = `[${timestamp}] [${level.toUpperCase()}] [${context}] ${message}`;
    
    if (data) {
      console.log(logMessage, data);
    } else {
      console.log(logMessage);
    }
  },
  
  /**
   * Filter object to only include specified keys
   * @param {Object} obj - Object to filter
   * @param {Array} keys - Keys to include
   * @returns {Object} Filtered object
   */
  pick(obj, keys) {
    const result = {};
    keys.forEach(key => {
      if (obj.hasOwnProperty(key)) {
        result[key] = obj[key];
      }
    });
    return result;
  },
  
  /**
   * Remove specified keys from object
   * @param {Object} obj - Object to filter
   * @param {Array} keys - Keys to remove
   * @returns {Object} Filtered object
   */
  omit(obj, keys) {
    const result = { ...obj };
    keys.forEach(key => {
      delete result[key];
    });
    return result;
  }
};

// Freeze utils to prevent modifications
Object.freeze(Utils);
/**
 * ØN Master System - Sheets Service
 * Google Sheets abstraction layer for data operations
 */

const SheetsService = {
  /**
   * Get the main spreadsheet
   * @returns {Spreadsheet} Spreadsheet object
   */
  getSpreadsheet() {
    return SpreadsheetApp.openById(Config.SPREADSHEET_ID);
  },
  
  /**
   * Get a sheet by name, creating it if it doesn't exist
   * @param {string} sheetName - Name of the sheet
   * @returns {Sheet} Sheet object
   */
  getSheet(sheetName) {
    const ss = SheetsService.getSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      SheetsService.initializeSheet(sheetName, sheet);
    }
    
    return sheet;
  },
  
  /**
   * Initialize a sheet with headers based on its type
   * @param {string} sheetName - Name of the sheet
   * @param {Sheet} sheet - Sheet object
   */
  initializeSheet(sheetName, sheet) {
    let headers = [];
    
    switch (sheetName) {
      case Config.SHEETS.REPORTS:
        headers = Config.REPORTS_SCHEMA;
        break;
      case Config.SHEETS.AUDIT_LOG:
        headers = Config.AUDIT_SCHEMA;
        break;
      case Config.SHEETS.USERS:
        headers = Config.USERS_SCHEMA;
        break;
      case Config.SHEETS.USER_PROCESS_ROLES:
        headers = Config.USER_PROCESS_ROLES_SCHEMA;
        break;
      // CBQ Module sheets
      case CBQConfig.SHEETS.DOCUMENTS:
        headers = CBQConfig.DOCUMENTS_SCHEMA;
        break;
      case CBQConfig.SHEETS.WORDS:
        headers = CBQConfig.WORDS_SCHEMA;
        break;
      case CBQConfig.SHEETS.WORD_STATS:
        headers = CBQConfig.WORD_STATS_SCHEMA;
        break;
      case CBQConfig.SHEETS.CONFIG:
        headers = CBQConfig.CONFIG_SCHEMA;
        break;
      case CBQConfig.SHEETS.PROCESSING_QUEUE:
        headers = CBQConfig.PROCESSING_QUEUE_SCHEMA;
        break;
      default:
        return;
    }
    
    if (headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#1a1a1a')
        .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
  },
  
  /**
   * Initialize all required sheets
   */
  initializeAllSheets() {
    const sheets = Object.values(Config.SHEETS);
    sheets.forEach(sheetName => {
      SheetsService.getSheet(sheetName);
    });
    Utils.log('INFO', 'SheetsService', 'All sheets initialized');
  },
  
  /**
   * Get all rows from a sheet as objects
   * @param {string} sheetName - Name of the sheet
   * @returns {Array} Array of row objects
   */
  getAllRows(sheetName) {
    const sheet = SheetsService.getSheet(sheetName);
    const data = sheet.getDataRange().getValues();
    
    if (data.length < 2) return [];
    
    const headers = data[0];
    const rows = data.slice(1);
    
    return rows.map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    });
  },
  
  /**
   * Get rows matching a filter function
   * @param {string} sheetName - Name of the sheet
   * @param {Function} filterFn - Filter function
   * @returns {Array} Filtered rows
   */
  getRows(sheetName, filterFn) {
    const allRows = SheetsService.getAllRows(sheetName);
    return filterFn ? allRows.filter(filterFn) : allRows;
  },
  
  /**
   * Get a single row by ID
   * @param {string} sheetName - Name of the sheet
   * @param {string} id - Row ID
   * @returns {Object|null} Row object or null
   */
  getRowById(sheetName, id) {
    const rows = SheetsService.getAllRows(sheetName);
    return rows.find(row => row.id === id) || null;
  },
  
  /**
   * Append a new row to a sheet
   * @param {string} sheetName - Name of the sheet
   * @param {Object} data - Row data object
   * @returns {Object} The appended row
   */
  appendRow(sheetName, data) {
    const sheet = SheetsService.getSheet(sheetName);
    const lastCol = sheet.getLastColumn();
    
    // Handle newly created sheets that may not have headers yet
    if (lastCol === 0) {
      throw new Error(`Sheet "${sheetName}" has no headers. Run initializeAllSheets() first.`);
    }
    
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    if (!headers || headers.length === 0 || !headers[0]) {
      throw new Error(`Sheet "${sheetName}" has no valid headers. Run initializeAllSheets() first.`);
    }
    
    const rowData = headers.map(header => data[header] !== undefined ? data[header] : '');
    sheet.appendRow(rowData);
    
    return data;
  },
  
  /**
   * Update a row by ID
   * @param {string} sheetName - Name of the sheet
   * @param {string} id - Row ID
   * @param {Object} data - Updated data
   * @returns {Object|null} Updated row or null if not found
   */
  updateRow(sheetName, id, data) {
    const sheet = SheetsService.getSheet(sheetName);
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    if (values.length < 2) return null;
    
    const headers = values[0];
    const idIndex = headers.indexOf('id');
    
    if (idIndex === -1) return null;
    
    // Find the row with matching ID
    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if (values[i][idIndex] === id) {
        rowIndex = i;
        break;
      }
    }
    
    if (rowIndex === -1) return null;
    
    // Update the row
    const updatedRow = headers.map((header, index) => {
      if (data.hasOwnProperty(header)) {
        return data[header];
      }
      return values[rowIndex][index];
    });
    
    sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([updatedRow]);
    
    // Return updated data as object
    const result = {};
    headers.forEach((header, index) => {
      result[header] = updatedRow[index];
    });
    
    return result;
  },
  
  /**
   * Delete a row by ID
   * @param {string} sheetName - Name of the sheet
   * @param {string} id - Row ID
   * @returns {boolean} True if deleted
   */
  deleteRow(sheetName, id) {
    const sheet = SheetsService.getSheet(sheetName);
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    if (values.length < 2) return false;
    
    const headers = values[0];
    const idIndex = headers.indexOf('id');
    
    if (idIndex === -1) return false;
    
    // Find and delete the row
    for (let i = 1; i < values.length; i++) {
      if (values[i][idIndex] === id) {
        sheet.deleteRow(i + 1);
        return true;
      }
    }
    
    return false;
  },
  
  /**
   * Get distinct values for a column
   * @param {string} sheetName - Name of the sheet
   * @param {string} columnName - Column name
   * @returns {Array} Distinct values
   */
  getDistinctValues(sheetName, columnName) {
    const rows = SheetsService.getAllRows(sheetName);
    const values = rows.map(row => row[columnName]).filter(v => v);
    return [...new Set(values)].sort();
  },
  
  /**
   * Bulk insert rows
   * @param {string} sheetName - Name of the sheet
   * @param {Array} dataArray - Array of row objects
   * @returns {number} Number of rows inserted
   */
  bulkInsert(sheetName, dataArray) {
    if (!dataArray || dataArray.length === 0) return 0;
    
    const sheet = SheetsService.getSheet(sheetName);
    const lastCol = sheet.getLastColumn();
    
    // Handle newly created sheets that may not have headers yet
    if (lastCol === 0) {
      throw new Error(`Sheet "${sheetName}" has no headers. Run initializeAllSheets() first.`);
    }
    
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    if (!headers || headers.length === 0 || !headers[0]) {
      throw new Error(`Sheet "${sheetName}" has no valid headers. Run initializeAllSheets() first.`);
    }
    
    const rowsData = dataArray.map(data => headers.map(header => data[header] !== undefined ? data[header] : ''));
    
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsData.length, headers.length).setValues(rowsData);
    
    return rowsData.length;
  },
  
  /**
   * Clear all data from a sheet (except headers)
   * @param {string} sheetName - Name of the sheet
   */
  clearSheet(sheetName) {
    const sheet = SheetsService.getSheet(sheetName);
    const lastRow = sheet.getLastRow();
    
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
  }
};

// Freeze to prevent modifications
Object.freeze(SheetsService);
/**
 * ═══════════════════════════════════════════════════════════════════════════
 * RCA MANAGER v3 - CASE MANAGEMENT (CORE LOGIC)
 * ═══════════════════════════════════════════════════════════════════════════
 * Core case management functions. Uses Script Properties for sheet IDs.
 */

// ═══════════════════════════════════════════════════════════════════════════
// SHEET CONNECTIONS (using Script Properties with fallbacks)
// ═══════════════════════════════════════════════════════════════════════════

function getSSSheet() {
  return SpreadsheetApp.openById(getSheetId('RCA_DATA')).getSheetByName("database");
}

function getDashboardSheet() {
  return SpreadsheetApp.openById(getSheetId('RCA_DATA')).getSheetByName("Dashboard");
}

function getControllerSheet() {
  return SpreadsheetApp.openById(getSheetId('CB_CONTROLLER')).getSheetByName("datasheet");
}

function getTempSheet() {
  return SpreadsheetApp.openById(getSheetId('CB_CONTROLLER')).getSheetByName("Temp");
}

function getRecoverySheet() {
  return SpreadsheetApp.openById(getSheetId('RECOVERY')).getSheetByName("data");
}

function getVbExceptionSheet() {
  return SpreadsheetApp.openById(getSheetId('VB_EXCEPTION')).getSheetByName("data");
}

function getEmConfirmationSheet() {
  return SpreadsheetApp.openById(getSheetId('EMAIL_CONFIRMATION'));
}

// Lazy-loaded sheet references
var _ss = null;
var _dashboard = null;
var _controllerSheet = null;
var _tempSheet = null;
var _recoverySheet = null;
var _vbExceptionSheet = null;
var _emConfirmationSheet = null;

function get_ss() { return _ss || (_ss = getSSSheet()); }
function get_dashboard() { return _dashboard || (_dashboard = getDashboardSheet()); }
function get_controllerSheet() { return _controllerSheet || (_controllerSheet = getControllerSheet()); }
function get_tempSheet() { return _tempSheet || (_tempSheet = getTempSheet()); }
function get_recoverySheet() { return _recoverySheet || (_recoverySheet = getRecoverySheet()); }
function get_vbExceptionSheet() { return _vbExceptionSheet || (_vbExceptionSheet = getVbExceptionSheet()); }
function get_emConfirmationSheet() { return _emConfirmationSheet || (_emConfirmationSheet = getEmConfirmationSheet()); }

// ═══════════════════════════════════════════════════════════════════════════
// AUTHORIZED USERS
// ═══════════════════════════════════════════════════════════════════════════

var users = [
  "naitik.joshi@priceline.com",
  "jeevan.bist@priceline.com",
  "arun.choubey@priceline.com",
  "anusha.shetty@priceline.com",
  "nitingopal.singh@priceline.com",
  "deb.mukherjee@priceline.com",
  "tapas.sharma@priceline.com",
  "avdhut.bidwe@priceline.com",
  "priyanka.thakkar@priceline.com",
  "viren.joshi@priceline.com",
  "shwetank.mishra@priceline.com",
  "zankhana.yagnik@priceline.com",
  "sandeep.agrawal@priceline.com",
  "zankhit.mehta@priceline.com",
  "krutika.patel@priceline.com",
  "jayesh.parmar@priceline.com",
  "priyanka.ratudi@priceline.com",
  "tirth.soni@priceline.com",
  "sanyam.sisodiya@priceline.com",
  "bharati.dash@priceline.com",
  "harshil.girnari@priceline.com",
  "kunal.jadaun@priceline.com",
  "mrunal.patel@priceline.com",
  "aniket.jha@priceline.com",
  "anagha.s@priceline.com",
  "easwaran.m@priceline.com",
  "anjali.raj@priceline.com"
];

// ═══════════════════════════════════════════════════════════════════════════
// CASE RETRIEVAL FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════

function NotClaimedCasesV4() {
  var ss = get_ss();
  var dashboard = get_dashboard();
  var Lastrow = ss.getLastRow();
  
  var user = Session.getActiveUser().getEmail();
  var rowAssigned = users.indexOf(user) + 7;
  var tempCell = dashboard.getRange("AC" + rowAssigned.toString());
  var row = 0;
  
  var filterFormula = `=Index(FILTER(ROW(database!Q:Q),(ISBLANK(database!Q:Q))+(database!Q:Q="Pending"), (ISBLANK(database!R:R))+(database!R:R="${user}"),(ISBLANK(database!S:S))),1)`;
  
  tempCell.setFormula(filterFormula);
  var result = tempCell.getValue();
  tempCell.clearContent();

  if (result <= Lastrow && result != "#N/A") {
    row = result;
  }
  return row;
}

function NotClaimedCasesV2(currency, portal) {
  var ss = get_ss();
  var Lastrow = ss.getLastRow();
  
  var user = Session.getActiveUser().getEmail();
  var valuesAnalyst = ss.getRange(2, 18, Lastrow).getValues();
  var valuesResolveDate = ss.getRange(2, 19, Lastrow).getValues();
  var valuesCurrency = ss.getRange(2, 12, Lastrow).getValues();
  var valuesPortal = ss.getRange(2, 33, Lastrow).getValues();
  var valuesRCA = ss.getRange(2, 17, Lastrow).getValues();

  if (currency == "Non USD") {
    for (i = 0; i < valuesAnalyst.length; i++) {
      if ((valuesAnalyst[i][0] == "" || valuesAnalyst[i][0] == user) && 
          (valuesRCA[i][0] == "Pending" || valuesRCA[i][0] == "") && 
          (valuesCurrency[i][0] != "USD") && 
          (valuesPortal[i][0] == portal) && 
          (valuesResolveDate[i][0] == "")) {
        return i + 2;
      }
    }
  } else {
    for (i = 0; i < valuesAnalyst.length; i++) {
      if ((valuesAnalyst[i][0] == "" || valuesAnalyst[i][0] == user) && 
          (valuesRCA[i][0] == "Pending" || valuesRCA[i][0] == "") && 
          (valuesCurrency[i][0] == "USD") && 
          (valuesPortal[i][0] == portal) && 
          (valuesResolveDate[i][0] == "")) {
        return i + 2;
      }
    }
  }

  if (i = valuesAnalyst.length) {
    return 0;
  }
}

function NotClaimedPriority() {
  var ss = get_ss();
  var Lastrow = ss.getLastRow();
  
  var user = Session.getActiveUser().getEmail();
  var valuesAnalyst = ss.getRange(2, 18, Lastrow).getValues();
  var valuesResolveDate = ss.getRange(2, 19, Lastrow).getValues();
  var valuesRCA = ss.getRange(2, 17, Lastrow).getValues();
  var valuesQueue = ss.getRange(2, 34, Lastrow).getValues();
  var valuesDueDate = ss.getRange(2, 6, Lastrow).getValues();
  var valuesReasonCode = ss.getRange(2, 4, Lastrow).getValues();

  var candidates = [];
  var today = new Date();
  
  // Collect all eligible cases with scoring
  for (var i = 0; i < valuesAnalyst.length; i++) {
    if ((valuesAnalyst[i][0] == "" || valuesAnalyst[i][0] == user) && 
        (valuesRCA[i][0] == "Pending" || valuesRCA[i][0] == "")) {
      
      // Calculate urgency score (higher = more urgent)
      var urgencyScore = 100;
      if (valuesDueDate[i][0]) {
        var dueDate = new Date(valuesDueDate[i][0]);
        var daysRemaining = Math.ceil((dueDate - today) / (1000 * 60 * 60 * 24));
        urgencyScore = Math.max(0, 100 - daysRemaining * 10); // 0-2 days = 80-100, 3-5 = 50-70
      }
      
      // Calculate type weight (fraud gets priority)
      var typeWeight = getCaseTypeWeight(valuesReasonCode[i][0]);
      
      // Combined score
      var totalScore = urgencyScore + typeWeight;
      
      candidates.push({
        row: i + 2,
        score: totalScore,
        queue: valuesQueue[i][0]
      });
    }
  }
  
  // Sort by score descending (highest priority first)
  candidates.sort(function(a, b) { return b.score - a.score; });
  
  // Return the highest priority case, or 0 if none
  return candidates.length > 0 ? candidates[0].row : 0;
}

/**
 * Get weight based on case type category for distribution
 */
function getCaseTypeWeight(reasonCode) {
  var code = String(reasonCode);
  
  // Fraud-related codes get highest priority
  var fraudCodes = ["10.4", "37", "AA", "F29", "4540", "127", "193", "176", "U02", "6006"];
  if (fraudCodes.some(function(f) { return code.indexOf(f) >= 0; })) {
    return 50; // Fraud priority
  }
  
  // Check against CONFIG categories if available
  if (typeof CONFIG !== 'undefined' && CONFIG.CASE_TYPES) {
    if (CONFIG.CASE_TYPES.SERVICE && CONFIG.CASE_TYPES.SERVICE.some(function(s) { 
      return code.indexOf(s) >= 0; 
    })) {
      return 40; // Service issues
    }
    if (CONFIG.CASE_TYPES.CUSTOMER && CONFIG.CASE_TYPES.CUSTOMER.some(function(c) { 
      return code.indexOf(c) >= 0; 
    })) {
      return 30; // Customer disputes
    }
    if (CONFIG.CASE_TYPES.HOTEL && CONFIG.CASE_TYPES.HOTEL.some(function(h) { 
      return code.indexOf(h) >= 0; 
    })) {
      return 20; // Hotel issues
    }
  }
  
  return 10; // Other/unknown
}

function ClaimCase(row) {
  var ss = get_ss();
  var user = Session.getActiveUser().getEmail();
  var lock = LockService.getScriptLock();
  
  try {
    // Wait up to 10 seconds for lock to prevent race condition
    lock.waitLock(10000);
    
    // Re-verify case is still unclaimed after acquiring lock
    var currentAnalyst = ss.getRange(row, 18).getValue();
    var currentRCA = ss.getRange(row, 17).getValue();
    
    if (currentAnalyst !== "" && currentAnalyst !== user) {
      // Case was claimed by someone else during our wait
      return { success: false, error: "Case already claimed by " + currentAnalyst };
    }
    
    if (currentRCA !== "" && currentRCA !== "Pending") {
      // Case was already resolved
      return { success: false, error: "Case already resolved" };
    }
    
    // Claim the case atomically
    ss.getRange(row, 18).setValue(user);
    ss.getRange(row, 17).setValue("Pending");
    ss.getRange(row, 32).setValue(new Date());
    SpreadsheetApp.flush(); // Force immediate write to prevent stale data
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function getCase(portal) {
  var ss = get_ss();
  var data = new Array();
  
  if (portal == "Priority") {
    var row = NotClaimedPriority();
  } else {
    var row = NotClaimedCasesV4();
  }

  if (users.includes(Session.getActiveUser().getEmail())) {
    if (row > 0) {
      var claimResult = ClaimCase(row);
      
      // Handle race condition - case was claimed by another analyst
      if (claimResult && claimResult.success === false) {
        data.push("Case already taken. Click Get Case again.", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        data.push(0);
        return data;
      }
      
      data.push("Case claimed.");
      data.push(row);
      data.push(ss.getRange(row, 1).getValue()); // caseid
      data.push(ss.getRange(row, 2).getValue()); // itn
      data.push(ss.getRange(row, 9).getValue()); // R#
      data.push(ss.getRange(row, 4).getValue()); // rCode
      data.push(ss.getRange(row, 7).getValue()); // dAmt
      data.push(ss.getRange(row, 13).getValue()); // dispute type
      data.push(ss.getRange(row, 27).getValue()); // refund
      data.push(ss.getRange(row, 10).getValue()); // tAmt
      data.push(ss.getRange(row, 12).getValue()); // currency

      data.push(Utilities.formatDate(ss.getRange(row, 5).getValue(), "IST", "MM/dd/yyyy")); // rDate
      data.push(Utilities.formatDate(ss.getRange(row, 6).getValue(), "IST", "MM/dd/yyyy")); // dDue
      data.push(ss.getRange(row, 33).getValue()); // portal
      data.push(ss.getRange(row, 29).getValue()); // policy
      data.push(ss.getRange(row, 22).getValue()); // internal notes
      data.push(ss.getRange(row, 35).getValue()); // affirm link
    } else {
      data.push("No Cases", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
      data.push(0);
    }
  } else {
    data.push("You are not Authorized", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
    data.push(0);
  }
  return data;
}

function searchCase(itn) {
  var ss = get_ss();
  var Lastrow = ss.getLastRow();
  
  var user = Session.getActiveUser().getEmail();
  var ITNs = ss.getRange(2, 2, Lastrow).getValues();
  var RCAanalysts = ss.getRange(2, 18, Lastrow).getValues();
  var row = 0;
  
  for (i = 0; i < ITNs.length; i++) {
    if (ITNs[i][0] == itn && RCAanalysts[i][0] == user) {
      row = i + 2;
      break;
    }
  }
  
  var data = new Array();
  if (row != 0) {
    data.push("Case retrieved.");
    data.push(row);
    data.push(ss.getRange(row, 1).getValue()); // caseid
    data.push(ss.getRange(row, 2).getValue()); // itn
    data.push(ss.getRange(row, 9).getValue()); // R#
    data.push(ss.getRange(row, 4).getValue()); // rCode
    data.push(ss.getRange(row, 7).getValue()); // dAmt
    data.push(ss.getRange(row, 13).getValue()); // dispute type
    data.push(ss.getRange(row, 27).getValue()); // refund
    data.push(ss.getRange(row, 10).getValue()); // tAmt
    data.push(ss.getRange(row, 12).getValue()); // currency

    data.push(Utilities.formatDate(ss.getRange(row, 5).getValue(), "IST", "MM/dd/yyyy")); // rDate
    data.push(Utilities.formatDate(ss.getRange(row, 6).getValue(), "IST", "MM/dd/yyyy")); // dDue
    data.push(ss.getRange(row, 33).getValue()); // portal
    data.push(ss.getRange(row, 29).getValue()); // policy
    data.push(ss.getRange(row, 20).getValue()); // issuer notes
    data.push(ss.getRange(row, 21).getValue()); // rebuttal
    data.push(ss.getRange(row, 22).getValue()); // internal notes
    data.push(ss.getRange(row, 35).getValue()); // affirm link
    data.push(ss.getRange(row, 17).getValue()); // rca
  } else {
    data.push("No Results", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
  }
  return data;
}

function getReservationData(itn) {
  var controllerSheet = get_controllerSheet();
  var tempSheet = get_tempSheet();
  
  const searchRange = 'datasheet!A:A';
  const searchValue = itn;
  const matchFormula = `=MATCH("${searchValue}", ${searchRange}, 0)`;

  var tempRowCell = users.indexOf(Session.getActiveUser().getEmail()) + 2;
  const tempCell = tempSheet.getRange('B' + tempRowCell.toString());
  tempCell.setFormula(matchFormula);

  const row = tempCell.getValue();
  tempCell.clearContent();

  var resinfo = new Array();
  resinfo.push("Reservation information fetched.");
  resinfo.push(row);
  resinfo.push(controllerSheet.getRange(row, 1).getValue()); // itn
  resinfo.push(Utilities.formatDate(controllerSheet.getRange(row, 4).getValue(), "CST", "MM/dd/yyyy")); // booking date
  resinfo.push(controllerSheet.getRange(row, 13).getValue()); // hotel
  resinfo.push("guest"); // guest
  resinfo.push(Utilities.formatDate(controllerSheet.getRange(row, 16).getValue(), "CST", "MM/dd/yyyy")); // check in
  resinfo.push(Utilities.formatDate(controllerSheet.getRange(row, 17).getValue(), "CST", "MM/dd/yyyy")); // checkout
  resinfo.push("test"); // email
  var rtype = controllerSheet.getRange(row, 18).getValue().trim() || "not available";
  resinfo.push(rtype); // rtype

  return resinfo;
}

function checkClaim(row) {
  var ss = get_ss();
  var user = Session.getActiveUser().getEmail();
  var claimed = ss.getRange(row, 18).getValue();

  if (user == claimed) {
    var response = ["Case currently claimed to " + claimed, "#2ea44f", "white"];
  } else {
    var response = ["Case currently claimed to " + claimed, "red", "white"];
  }

  return response;
}

function resolveRCA(row, rca, res, issuernotes, internalnotes) {
  var ss = get_ss();
  ss.getRange(row, 20).setValue(issuernotes);
  ss.getRange(row, 21).setValue(res);
  ss.getRange(row, 22).setValue(internalnotes);
  ss.getRange(row, 17).setValue(rca);
  ss.getRange(row, 19).setValue(new Date());
  ss.getRange(row, 24).clearContent();
  sendForEmailConfirmation(row);
  return "Case Updated.";
}

var Fraud = ["10.4", "37", 10.4, 37, "AA", "Card Member does not recognise Transaction or Transaction Amount-6014", 
             "Card not present fraud-F29", "Card not present-4540", "Does Not Recognize/ Remember/ No Knowledge-127", 
             "Fraud-193", "Fraudulent Transactions-193", "No knowledge-127", "U02", "Fraud", 
             "Card Member claims fraud-6006", "Does Not Recognize/ Remember/ No Knowledge-176", 1040, 
             "Fraud (No authorization)", "Fraud (Card not present)", "Not Recognized"];

function sendForEmailConfirmation(row) {
  var ss = get_ss();
  var emConfirmationSheet = get_emConfirmationSheet();
  
  var code = ss.getRange(row, 4).getValue();
  var interface = ss.getRange(row, 14).getValue();
  var tempSheet = emConfirmationSheet.getSheetByName('temp');
  var mainSheet = emConfirmationSheet.getSheetByName('data');

  var creditnot = ["Credit Not Processed", 13.7, 1370, "13.7", "13.70", "1370"];

  if (Fraud.includes(code) || (creditnot.includes(code) && interface == "call_center")) {
    var itn = ss.getRange(row, 2).getValue();
    var tempCell = tempSheet.getRange("B" + (Math.floor(Math.random() * 10)));
    tempCell.setFormula(`=COUNTIF(data!B:B,"${itn}")`);
    if (tempCell.getValue() == 0) {
      var lastrow_email = mainSheet.getLastRow();
      var user = Session.getActiveUser().getEmail();
      mainSheet.getRange(lastrow_email + 1, 1).setValue(user);
      mainSheet.getRange(lastrow_email + 1, 2).setValue(itn);
      var date = new Date();
      mainSheet.getRange(lastrow_email + 1, 3).setValue(Utilities.formatDate(date, "CST", "MM-dd-yyyy"));
      mainSheet.getRange(lastrow_email + 1, 4).setValue(ss.getRange(row, 6).getValue());
      mainSheet.getRange(lastrow_email + 1, 5).setValue("Pending");
    }
    tempCell.clearContent();
  }
}
/**
 * ═══════════════════════════════════════════════════════════════════════════
 * RCA MANAGER v3 - RECOVERY & VB EXCEPTIONS
 * ═══════════════════════════════════════════════════════════════════════════
 * Uses Script Properties for sheet IDs.
 */

function sendForRecovery(itn, rnum, amt, cin, cout, hotel, cbcode, notes) {
  var recoverySheet = get_recoverySheet();
  var lastrow_recovery = recoverySheet.getLastRow();
  
  var new_row = lastrow_recovery + 1;
  recoverySheet.getRange(new_row, 1).setValue(new Date());
  recoverySheet.getRange(new_row, 2).setValue(Session.getActiveUser().getEmail().split("@")[0]);
  recoverySheet.getRange(new_row, 3).setValue(itn);
  recoverySheet.getRange(new_row, 4).setValue(rnum);
  recoverySheet.getRange(new_row, 5).setValue(amt);
  recoverySheet.getRange(new_row, 6).setValue(cin);
  recoverySheet.getRange(new_row, 7).setValue(cout);
  recoverySheet.getRange(new_row, 8).setValue(hotel);
  recoverySheet.getRange(new_row, 9).setValue(cbcode);
  recoverySheet.getRange(new_row, 10).setValue(notes);

  return "Case sent for recovery.";
}

function sendForVbException(itn, rnum, amt, option, notes) {
  var vbExceptionSheet = get_vbExceptionSheet();
  var lastrow_vbException = vbExceptionSheet.getLastRow();
  
  var new_row = lastrow_vbException + 1;
  vbExceptionSheet.getRange(new_row, 1).setValue(new Date());
  vbExceptionSheet.getRange(new_row, 2).setValue(Session.getActiveUser().getEmail().split("@")[0]);
  vbExceptionSheet.getRange(new_row, 3).setValue(itn);
  vbExceptionSheet.getRange(new_row, 4).setValue(rnum);
  vbExceptionSheet.getRange(new_row, 5).setValue(amt);
  vbExceptionSheet.getRange(new_row, 6).setValue(option);
  vbExceptionSheet.getRange(new_row, 7).setValue(notes);
  
  return "Case sent for VB Exception.";
}

function updateRefundAmount(row, amount) {
  var ss = get_ss();
  ss.getRange(row, 27).setValue(amount);
}

function updateRoomType(row, roomtype) {
  var controllerSheet = get_controllerSheet();
  controllerSheet.getRange(row, 18).setValue(roomtype);
}

function getPendingCases() {
  var ss = get_ss();
  var Lastrow = ss.getLastRow();
  
  var user = Session.getActiveUser().getEmail();
  var RCAnalysts = ss.getRange(2, 18, Lastrow).getValues();
  var RCAValues = ss.getRange(2, 17, Lastrow).getValues();
  var rows = new Array();
  
  for (i = 0; i < RCAnalysts.length; i++) {
    if (RCAValues[i][0] == "Pending" && (RCAnalysts[i][0] == user)) {
      rows.push(i + 2);
    }
  }
  
  var disputes = new Array();

  if (rows.length > 0) {
    for (row of rows) {
      var dispute = new Array();
      dispute.push(row);
      dispute.push(ss.getRange(row, 1).getValue()); // caseid
      dispute.push(ss.getRange(row, 2).getValue()); // itn
      dispute.push(ss.getRange(row, 4).getValue()); // reasoncode
      dispute.push(Utilities.formatDate(ss.getRange(row, 5).getValue(), "CST", "MM/dd/yyyy")); // received
      dispute.push(Utilities.formatDate(ss.getRange(row, 6).getValue(), "CST", "MM/dd/yyyy")); // duedate
      dispute.push(ss.getRange(row, 22).getValue()); // notes
      disputes.push(dispute);
    }
  }

  return disputes;
}
/**
 * ═══════════════════════════════════════════════════════════════════════════
 * RCA MANAGER v3 - FILE UPLOADS
 * ═══════════════════════════════════════════════════════════════════════════
 * Uses Script Properties for folder IDs.
 */

// Function to upload image to Google Drive
function uploadImageToDrive(row, itn, inputId, base64Data) {
  const folder = DriveApp.getFolderById(getFolderId('IMAGES'));

  if (inputId == "wex") {
    var fileName = itn + " - Hotel Billed GAR.png";
    var col = 20;
  } else if (inputId == "ais") {
    var fileName = itn + " - Accertify Index Score.png";
    var col = 21;
  } else if (inputId == "rfs") {
    var fileName = itn + " - Refund screenshot.png";
    var col = 22;
  } else if (inputId == "ccontact") {
    var fileName = itn + " - Customer contact.png";
    var col = 23;
  } else if (inputId == "nonfraud") {
    var fileName = itn + " - Non Fraud Info.png";
    var col = 24;
  } else if (inputId == "nonfraud1") {
    var fileName = itn + " - Non Fraud Info1.png";
    var col = 25;
  } else {
    var fileName = itn + " - Non Fraud Info2.png";
    var col = 26;
  }

  try {
    // Decode the Base64 data and create a blob
    const decodedData = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decodedData, 'image/png', fileName);

    const uploadedFile = folder.createFile(blob);
    
    var controllerSheet = get_controllerSheet();
    controllerSheet.getRange(row, col).setValue(uploadedFile.getId());

    return {
      inputId: inputId,
      status: 'Uploaded',
      name: uploadedFile.getName(),
      url: uploadedFile.getUrl()
    };

  } catch (error) {
    Logger.log('Upload Error: ' + error.toString());
    throw new Error('Failed to upload the image.');
  }
}

function uploadAllFiles(formData) {
  const uploadedResults = [];
  
  // Get folder IDs from Script Properties
  const FOLDER_IDS = {
    confirmpage: getFolderId('CONFIRM_PAGE'),
    checkoutpage: getFolderId('CHECKOUT_PAGE'),
    confirmpages: getFolderId('CONFIRM_PAGE'),
    checkoutpages: getFolderId('CHECKOUT_PAGE'),
    wexandrefund: getFolderId('WEX_REFUND'),
    accproof: getFolderId('ACCERTIFY')
  };

  // Helper to upload a single file
  function uploadSingle(file, folderId, label, fileName) {
    const blob = Utilities.newBlob(Utilities.base64Decode(file.content), file.mimeType, fileName);
    DriveApp.getFolderById(folderId).createFile(blob);
    uploadedResults.push(`${label}: ${fileName}`);
  }

  // Option 1 (Required)
  if (formData.confirmpage) uploadSingle(formData.confirmpage, FOLDER_IDS.confirmpage, "Confirmation page", formData.itn + " - Reservation Confirmation.pdf");

  // Option 2 (Required)
  if (formData.checkoutpage) uploadSingle(formData.checkoutpage, FOLDER_IDS.checkoutpage, "Checkout page", formData.itn + " - Checkout Page.pdf");

  // Option 3 (Optional, merged)
  if (formData.confirmpages) uploadSingle(formData.confirmpages, FOLDER_IDS.confirmpages, "Confirmation pages", formData.itn + " - other reservations.pdf");

  // Option 4 (Optional, merged)
  if (formData.checkoutpages) uploadSingle(formData.checkoutpages, FOLDER_IDS.checkoutpages, "Checkout pages", formData.itn + " - other checkout pages.pdf");

  if (formData.wexandrefund) uploadSingle(formData.wexandrefund, FOLDER_IDS.wexandrefund, "Hotel billed/refund proofs", formData.itn + " - other hotel billed/refund proofs.pdf");
  if (formData.accproof) uploadSingle(formData.accproof, FOLDER_IDS.accproof, "Customer Proof", formData.itn + " - Customer Proof.pdf");

  return { status: "success", message: "Files uploaded successfully!", uploaded: uploadedResults };
}
/**
 * ═══════════════════════════════════════════════════════════════════════════
 * RCA MANAGER v3 - CANNED RESPONSES
 * ═══════════════════════════════════════════════════════════════════════════
 * Uses Script Properties for sheet ID.
 */

const SHEET_NAME = "Responses";

function getResponseSheet() {
  return SpreadsheetApp.openById(getSheetId('CANNED_RESPONSES'));
}

function getUserEmail() {
  return Session.getActiveUser().getEmail();
}

function getCannedResponses() {
  const email = getUserEmail();
  const sheet = getResponseSheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  return data
    .filter(row => row[3] === email)
    .map(row => ({ name: row[1], response: row[2], id: row[0] }));
}

function saveCannedResponse(name, response) {
  const email = getUserEmail();
  const sheet = getResponseSheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const now = new Date();

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === name && data[i][3] === email) {
      sheet.getRange(i + 1, 3).setValue(response);
      sheet.getRange(i + 1, 5).setValue(now);
      return "Updated";
    }
  }

  const id = new Date().getTime(); // Simple unique ID
  sheet.appendRow([id, name, response, email, now]);
  return "Saved";
}
/**
 * ═══════════════════════════════════════════════════════════════════════════
 * RCA MANAGER v3 - ADMIN PANEL SERVER FUNCTIONS
 * ═══════════════════════════════════════════════════════════════════════════
 * All admin-related server-side functions.
 * Isolated and feature-flagged for safe removal.
 */

// Session tracking sheet - lazily loaded
var _sessionSheet = null;

function getSessionSheet_() {
  if (!_sessionSheet) {
    var spreadsheet = SpreadsheetApp.openById(getSheetId('CB_CONTROLLER'));
    _sessionSheet = spreadsheet.getSheetByName("Sessions");
    
    // Initialize if it doesn't exist
    if (!_sessionSheet) {
      _sessionSheet = spreadsheet.insertSheet("Sessions");
      _sessionSheet.appendRow(["Email", "LastActive", "Status", "CurrentCaseRow", "SessionStart"]);
    }
  }
  return _sessionSheet;
}

/**
 * Record analyst heartbeat (called periodically from client)
 */
function recordHeartbeat() {
  if (!CONFIG.FEATURES.ENABLE_CASE_DISTRIBUTION) return { success: false };
  
  const sessionSheet = getSessionSheet_();
  const email = Session.getActiveUser().getEmail();
  const now = new Date();
  const data = sessionSheet.getDataRange().getValues();
  
  // Find existing session
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      sessionSheet.getRange(i + 1, 2).setValue(now);
      sessionSheet.getRange(i + 1, 3).setValue("Active");
      return { success: true, sessionRow: i + 1 };
    }
  }
  
  // Create new session
  sessionSheet.appendRow([email, now, "Active", "", now]);
  return { success: true, sessionRow: sessionSheet.getLastRow() };
}

/**
 * Get list of currently active analysts
 */
function getActiveAnalysts() {
  if (!CONFIG.FEATURES.ENABLE_CASE_DISTRIBUTION) return [];
  
  const sessionSheet = getSessionSheet_();
  const data = sessionSheet.getDataRange().getValues();
  const now = new Date();
  const staleMs = CONFIG.SESSION.STALE_MINUTES * 60 * 1000;
  const activeAnalysts = [];
  
  for (let i = 1; i < data.length; i++) {
    const lastActive = new Date(data[i][1]);
    if ((now - lastActive) < staleMs && data[i][2] === "Active") {
      activeAnalysts.push({
        email: data[i][0],
        lastActive: lastActive,
        currentCaseRow: data[i][3],
        sessionStart: data[i][4]
      });
    }
  }
  
  return activeAnalysts;
}

/**
 * Get all analysts with their enabled/disabled status
 */
function getAllAnalysts() {
  if (!isAdmin()) return { error: "Unauthorized" };
  
  const analysts = users.map(email => {
    return {
      email: email,
      name: email.split("@")[0].replace(".", " ").replace(/\b\w/g, c => c.toUpperCase()),
      enabled: true // All are enabled by default in current system
    };
  });
  
  return analysts;
}

/**
 * Get analyst statistics for admin dashboard
 */
function getAnalystStats() {
  if (!CONFIG.FEATURES.ENABLE_TIME_TRACKING) return [];
  
  const rcaSheet = get_ss();
  const data = rcaSheet.getDataRange().getValues();
  const stats = {};
  
  // Initialize stats for all analysts
  users.forEach(email => {
    stats[email] = {
      email: email,
      name: email.split("@")[0].replace(".", " ").replace(/\b\w/g, c => c.toUpperCase()),
      totalCases: 0,
      resolvedToday: 0,
      resolvedWeek: 0,
      avgTimeMinutes: 0,
      totalTimeMinutes: 0,
      todayTimeMinutes: 0,      // NEW: Time spent today only
      todayAvgTimeMinutes: 0,   // NEW: Average time per case today
      casesByType: { FRAUD: 0, SERVICE: 0, CUSTOMER: 0, HOTEL: 0, OTHER: 0 }
    };
  });
  
  const now = new Date();
  const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const weekStart = new Date(todayStart);
  weekStart.setDate(weekStart.getDate() - 7);
  
  // Process each case
  for (let i = 1; i < data.length; i++) {
    const analyst = data[i][17]; // Column R (18-1)
    const claimDate = data[i][31]; // Column AF (32-1)
    const resolveDate = data[i][18]; // Column S (19-1)
    const rcaValue = data[i][16]; // Column Q (17-1)
    
    if (!analyst || !stats[analyst]) continue;
    
    stats[analyst].totalCases++;
    
    // Categorize by case type
    const caseType = getCaseTypeCategory(rcaValue);
    stats[analyst].casesByType[caseType]++;
    
    // Calculate time if resolved
    if (claimDate && resolveDate) {
      const claimTime = new Date(claimDate);
      const resolveTime = new Date(resolveDate);
      const durationMinutes = (resolveTime - claimTime) / (1000 * 60);
      
      if (durationMinutes > 0 && durationMinutes < 1440) { // Ignore outliers > 24 hours
        stats[analyst].totalTimeMinutes += durationMinutes;
        
        // Track today's time separately
        if (resolveTime >= todayStart) {
          stats[analyst].todayTimeMinutes += durationMinutes;
        }
      }
      
      // Count resolved today/week
      if (resolveTime >= todayStart) {
        stats[analyst].resolvedToday++;
      }
      if (resolveTime >= weekStart) {
        stats[analyst].resolvedWeek++;
      }
    }
  }
  
  // Calculate averages
  Object.values(stats).forEach(s => {
    if (s.totalCases > 0) {
      s.avgTimeMinutes = Math.round(s.totalTimeMinutes / s.totalCases);
    }
    // NEW: Today's average
    if (s.resolvedToday > 0) {
      s.todayAvgTimeMinutes = Math.round(s.todayTimeMinutes / s.resolvedToday);
    }
  });
  
  return Object.values(stats).sort((a, b) => b.resolvedToday - a.resolvedToday);
}

/**
 * Get TODAY-ONLY analyst statistics for dedicated admin dashboard
 * Returns: { analyst, resolvedToday, avgTimeToday, totalTimeToday }
 */
function getTodayAnalystStats() {
  if (!CONFIG.FEATURES.ENABLE_TIME_TRACKING) return [];
  
  const rcaSheet = get_ss();
  const data = rcaSheet.getDataRange().getValues();
  const todayStats = {};
  
  const now = new Date();
  const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  
  // Initialize for all analysts
  users.forEach(email => {
    todayStats[email] = {
      email: email,
      name: email.split("@")[0].replace(".", " ").replace(/\b\w/g, c => c.toUpperCase()),
      resolvedToday: 0,
      totalTimeMinutes: 0,
      avgTimeMinutes: 0,
      casesByType: { FRAUD: 0, SERVICE: 0, CUSTOMER: 0, HOTEL: 0, OTHER: 0 }
    };
  });
  
  // Process only today's resolved cases
  for (let i = 1; i < data.length; i++) {
    const analyst = data[i][17]; // Column R
    const claimDate = data[i][31]; // Column AF
    const resolveDate = data[i][18]; // Column S
    const rcaValue = data[i][16]; // Column Q
    
    if (!analyst || !todayStats[analyst] || !resolveDate) continue;
    
    const resolveTime = new Date(resolveDate);
    
    // Only count if resolved TODAY
    if (resolveTime >= todayStart) {
      todayStats[analyst].resolvedToday++;
      
      const caseType = getCaseTypeCategory(rcaValue);
      todayStats[analyst].casesByType[caseType]++;
      
      if (claimDate) {
        const claimTime = new Date(claimDate);
        const durationMinutes = (resolveTime - claimTime) / (1000 * 60);
        
        if (durationMinutes > 0 && durationMinutes < 1440) {
          todayStats[analyst].totalTimeMinutes += durationMinutes;
        }
      }
    }
  }
  
  // Calculate averages
  Object.values(todayStats).forEach(s => {
    if (s.resolvedToday > 0) {
      s.avgTimeMinutes = Math.round(s.totalTimeMinutes / s.resolvedToday);
      s.totalTimeMinutes = Math.round(s.totalTimeMinutes);
    }
  });
  
  // Return sorted by resolved count, filter out inactive
  return Object.values(todayStats)
    .filter(s => s.resolvedToday > 0)
    .sort((a, b) => b.resolvedToday - a.resolvedToday);
}

/**
 * Get case distribution summary by urgency and type
 */
function getCaseDistributionSummary() {
  if (!CONFIG.FEATURES.ENABLE_QUEUE_DASHBOARD) return {};
  
  const rcaSheet = get_ss();
  const data = rcaSheet.getDataRange().getValues();
  const now = new Date();
  
  const summary = {
    byUrgency: {
      critical: 0,  // 0-2 days
      warning: 0,   // 3-5 days
      normal: 0     // 6+ days
    },
    byType: {
      FRAUD: 0,
      SERVICE: 0,
      CUSTOMER: 0,
      HOTEL: 0,
      OTHER: 0
    },
    byPortal: {
      Amex: 0,
      Affirm: 0,
      Braintree: 0,
      Chase: 0,
      Other: 0
    },
    total: 0
  };
  
  for (let i = 1; i < data.length; i++) {
    const rca = data[i][16]; // Column Q
    const analyst = data[i][17]; // Column R
    const resolveDate = data[i][18]; // Column S
    const dueDate = data[i][5]; // Column F
    const portal = data[i][32]; // Column AG
    
    // Only count pending/unclaimed cases
    if ((rca === "" || rca === "Pending") && !resolveDate) {
      summary.total++;
      
      // Urgency calculation
      if (dueDate) {
        const due = new Date(dueDate);
        const daysRemaining = Math.ceil((due - now) / (1000 * 60 * 60 * 24));
        
        if (daysRemaining <= 2) {
          summary.byUrgency.critical++;
        } else if (daysRemaining <= 5) {
          summary.byUrgency.warning++;
        } else {
          summary.byUrgency.normal++;
        }
      }
      
      // Case type (based on reason code patterns)
      const caseType = getCaseTypeCategory(rca);
      summary.byType[caseType]++;
      
      // Portal
      if (summary.byPortal.hasOwnProperty(portal)) {
        summary.byPortal[portal]++;
      } else {
        summary.byPortal.Other++;
      }
    }
  }
  
  return summary;
}

/**
 * Get agent's own queue count (enhanced version)
 */
function agentCount() {
  const user = Session.getActiveUser().getEmail();
  const rcaSheet = get_ss();
  const data = rcaSheet.getDataRange().getValues();
  
  let myCount = 0, myFraud = 0, priority = 0;
  let affirm = 0, amex = 0, braintree = 0, chase = 0;
  
  const fraudCodes = ["10.4", "37", "AA", "F29", "4540", "127", "193", "176", "U02", "6006"];
  
  // Today's date for resolved count
  const now = new Date();
  const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  
  for (let i = 1; i < data.length; i++) {
    const rca = data[i][16];
    const analyst = data[i][17];
    const resolveDate = data[i][18];
    const portal = data[i][32];
    const queue = data[i][33];
    const reasonCode = String(data[i][3]);
    
    // Count pending cases for portal stats
    if ((rca === "" || rca === "Pending") && !resolveDate) {
      // By portal
      if (portal === "Affirm") affirm++;
      else if (portal === "Amex") amex++;
      else if (portal === "Braintree") braintree++;
      else if (portal === "Chase") chase++;
      
      // Priority queue
      if (queue <= 2) priority++;
    }
    
    // My Cases = PENDING and ASSIGNED to this user (work queue)
    if ((rca === "" || rca === "Pending") && analyst === user) {
        myCount++;
        if (fraudCodes.some(code => reasonCode.includes(code))) {
          myFraud++;
        }
    }
  }
  
  return { myCount, myFraud, priority, affirm, amex, braintree, chase };
}

/**
 * Mark analyst as inactive (called on logout/close)
 */
function markInactive() {
  if (!CONFIG.FEATURES.ENABLE_CASE_DISTRIBUTION) return;
  
  const sessionSheet = getSessionSheet_();
  const email = Session.getActiveUser().getEmail();
  const data = sessionSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      sessionSheet.getRange(i + 1, 3).setValue("Inactive");
      sessionSheet.getRange(i + 1, 4).setValue(""); // Clear current case
      break;
    }
  }
}
/**
 * ═══════════════════════════════════════════════════════════════════════════
 * REPRESENTATION MANAGER MODULE
 * ═══════════════════════════════════════════════════════════════════════════
 * Handles chargeback case representation workflow.
 * Merged from: R Manager Code.gs + CaseManagement.gs
 * 
 * Sheet ID: 1LxZiB_TTCxLMa-xwHhOsC4C1D-jneBQAE2wF8f_YEQ4 (DO NOT CHANGE)
 */

// ═══════════════════════════════════════════════════════════════════════════
// CONFIGURATION
// ═══════════════════════════════════════════════════════════════════════════

const REP_CONFIG = {
  // Uses centralized ID from Core.gs CONFIG.DEFAULT_SHEET_IDS
  get SHEET_ID() { return getSheetId('REP_MANAGER'); },
  DATABASE_SHEET: 'database',
  DASHBOARD_SHEET: 'Dashboard',
  // Uses the consolidated users list from above
  get AUTHORIZED_USERS() { return users; }
};

// ═══════════════════════════════════════════════════════════════════════════
// SHEET ACCESS
// ═══════════════════════════════════════════════════════════════════════════

function getRepSpreadsheet_() {
  return SpreadsheetApp.openById(REP_CONFIG.SHEET_ID);
}

function getRepDatasheet_() {
  return getRepSpreadsheet_().getSheetByName(REP_CONFIG.DATABASE_SHEET);
}

function getRepDashboard_() {
  return getRepSpreadsheet_().getSheetByName(REP_CONFIG.DASHBOARD_SHEET);
}

// ═══════════════════════════════════════════════════════════════════════════
// AUTHORIZATION
// ═══════════════════════════════════════════════════════════════════════════

function isRepAuthorized_() {
  const user = Session.getActiveUser().getEmail();
  return REP_CONFIG.AUTHORIZED_USERS.includes(user);
}

// ═══════════════════════════════════════════════════════════════════════════
// CASE RETRIEVAL (V2 - Optimized with FILTER formula)
// ═══════════════════════════════════════════════════════════════════════════

function repNotClaimedCasesV2(portal, currency) {
  const user = Session.getActiveUser().getEmail();
  const dashboard = getRepDashboard_();
  const datasheet = getRepDatasheet_();
  const lastRow = datasheet.getLastRow();
  
  const rowAssigned = REP_CONFIG.AUTHORIZED_USERS.indexOf(user) + 25;
  const tempCell = dashboard.getRange("AC" + rowAssigned.toString());
  
  let filterFormula;
  if (currency === "USD") {
    filterFormula = `=Index(FILTER(ROW(database!A:A),(ISBLANK(database!Y:Y)),database!X:X="Ready",(ISBLANK(database!Z:Z))+(database!Z:Z="${user}"),(database!L:L="USD")+(database!L:L="CAD"),database!AG:AG="${portal}"),1)`;
  } else {
    filterFormula = `=Index(FILTER(ROW(database!A:A),(ISBLANK(database!Y:Y)),database!X:X="Ready",(ISBLANK(database!Z:Z))+(database!Z:Z="${user}"),(database!L:L<>"USD"),(database!L:L<>"CAD"),database!AG:AG="${portal}"),1)`;
  }
  
  tempCell.setFormula(filterFormula);
  SpreadsheetApp.flush();
  const result = tempCell.getValue();
  tempCell.clearContent();
  
  if (result && result <= lastRow && result !== "#N/A") {
    return result;
  }
  return 0;
}

// ═══════════════════════════════════════════════════════════════════════════
// MAIN API: Get Pending Uploads
// ═══════════════════════════════════════════════════════════════════════════

function repGetPendingUploads(portal, currency) {
  const data = [];
  
  if (!isRepAuthorized_()) {
    return [0, "Not Authorized", "", "", "", "", "", "", "", "", "", "", "", ""];
  }
  
  const datasheet = getRepDatasheet_();
  const row = repNotClaimedCasesV2(portal, currency);
  
  if (row > 0) {
    repClaimCase_(row);
    
    data.push(row);
    data.push(datasheet.getRange(row, 1).getValue());  // caseid
    data.push(datasheet.getRange(row, 2).getValue());  // itn
    data.push(datasheet.getRange(row, 4).getValue());  // reasoncode
    data.push(datasheet.getRange(row, 3).getValue());  // MOP
    data.push(datasheet.getRange(row, 9).getValue());  // Reservation#
    data.push(datasheet.getRange(row, 7).getValue());  // DisputeAmt
    data.push(datasheet.getRange(row, 10).getValue()); // TAmt
    data.push(datasheet.getRange(row, 27).getValue()); // RFAmt
    data.push(Utilities.formatDate(datasheet.getRange(row, 6).getValue(), "IST", "MM/dd/yyyy")); // Due Date
    data.push(datasheet.getRange(row, 13).getValue()); // Dispute Type
    data.push(datasheet.getRange(row, 17).getValue()); // RCA
    data.push(datasheet.getRange(row, 22).getValue()); // Internal Notes
    data.push(datasheet.getRange(row, 35).getValue()); // Affirm link
  } else {
    data.push(0, "No Cases", "", "", "", "", "", "", "", "", "", "", "", "");
  }
  
  return data;
}

// ═══════════════════════════════════════════════════════════════════════════
// CASE ACTIONS
// ═══════════════════════════════════════════════════════════════════════════

function repClaimCase_(row) {
  const user = Session.getActiveUser().getEmail();
  getRepDatasheet_().getRange(row, 26).setValue(user);
}

function repChallenge(row) {
  const datasheet = getRepDatasheet_();
  datasheet.getRange(row, 25).setValue("Challenged");
  datasheet.getRange(row, 30).setValue(new Date());
}

function repAccept(row) {
  const datasheet = getRepDatasheet_();
  datasheet.getRange(row, 25).setValue("Accepted");
  datasheet.getRange(row, 30).setValue(new Date());
}

function repReversal(row) {
  const datasheet = getRepDatasheet_();
  datasheet.getRange(row, 25).setValue("Reversal");
  datasheet.getRange(row, 30).setValue(new Date());
}

function repAutoclosed(row) {
  const datasheet = getRepDatasheet_();
  datasheet.getRange(row, 25).setValue("Auto-closed");
  datasheet.getRange(row, 30).setValue(new Date());
}

function repRecreate(row, res) {
  const datasheet = getRepDatasheet_();
  datasheet.getRange(row, 23).setValue(res);
  datasheet.getRange(row, 25).setValue("Recreate Request");
  datasheet.getRange(row, 30).setValue(new Date());
}

// ═══════════════════════════════════════════════════════════════════════════
// DASHBOARD COUNTS
// ═══════════════════════════════════════════════════════════════════════════

function repGetCounts() {
  const dashboard = getRepDashboard_();
  return [
    dashboard.getRange(17, 5).getValue(),
    dashboard.getRange(18, 5).getValue(),
    dashboard.getRange(19, 5).getValue(),
    dashboard.getRange(20, 5).getValue(),
    dashboard.getRange(21, 5).getValue()
  ];
}

// ═══════════════════════════════════════════════════════════════════════════
// FILE SEARCH (for Representation Manager)
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Search for files in Google Drive by ITN number
 * Returns array of [filename, fileId, lastUpdated] for PDF files
 * @param {string} ITN - The ITN/Merchant Order number to search for
 * @returns {Array} Array of file info arrays
 */
function repSearchFiles(ITN) {
  if (!ITN || ITN === '-') {
    return [];
  }
  
  var data = [];
  var now = new Date();
  var ninetyDaysAgo = new Date(now.getTime() - (90 * 24 * 60 * 60 * 1000)); // 90 days in milliseconds
  
  try {
    var searchFor = 'title contains "' + ITN + '"';
    var files = DriveApp.searchFiles(searchFor);
    
    while (files.hasNext()) {
      var file = files.next();
      var fileId = file.getId();
      var name = file.getName();
      var type = file.getMimeType();
      var lastUpdated = file.getLastUpdated();
      var formattedDate = Utilities.formatDate(lastUpdated, "CST", "MM/dd/yyyy h:mm");
      
      // Only include PDF files modified within the last 90 days
      if (type === 'application/pdf' && lastUpdated >= ninetyDaysAgo) {
        data.push([name, fileId, formattedDate, lastUpdated.getTime()]);
      }
    }
    
    // Sort by date descending (most recent first)
    data.sort(function(a, b) {
      return b[3] - a[3];
    });
    
    // Remove the timestamp column (only used for sorting)
    data = data.map(function(row) {
      return [row[0], row[1], row[2]];
    });
    
  } catch (e) {
    console.error('repSearchFiles: Error searching for ITN ' + ITN + ': ' + e.message);
  }
  
  return data;
}
/**
 * ═══════════════════════════════════════════════════════════════════════════
 * REBUTTAL GENERATOR MODULE
 * ═══════════════════════════════════════════════════════════════════════════
 * Generates PDF rebuttals from RCA data with screenshots.
 * Trigger-based workflow for background processing.
 * 
 * Sheet IDs (DO NOT CHANGE):
 * - RCA Database: 1LxZiB_TTCxLMa-xwHhOsC4C1D-jneBQAE2wF8f_YEQ4
 * - Datasheet: 1P4MLBYGsco4Wh1TnmnZjaxjNG7GBeZRXeIKf8oF36UQ
 * - Template: 12TYJf2Jc3-19oAcTZt4Fhi0TPgilADBKOasbiN5H5v8
 * - Temp Folder: 18tSuM9LKs1-Zdd2CTC-5zaM36e0RJkMd
 * - PDF Folder: 1sjsc-yS685Sl3f8WDXjklBZDWw6qm4Wq
 */

// ═══════════════════════════════════════════════════════════════════════════
// CONFIGURATION
// ═══════════════════════════════════════════════════════════════════════════

const REBUTTAL_CONFIG = {
  // Uses centralized IDs from Core.gs CONFIG.DEFAULT_SHEET_IDS / DEFAULT_FOLDER_IDS
  get RCA_SHEET_ID() { return getSheetId('REP_MANAGER'); },
  get DATASHEET_ID() { return getSheetId('REBUTTAL_DATASHEET'); },
  get TEMPLATE_ID() { return getSheetId('REBUTTAL_TEMPLATE'); },
  get TEMP_FOLDER_ID() { return getFolderId('REBUTTAL_TEMP'); },
  get PDF_FOLDER_ID() { return getFolderId('REBUTTAL_PDF'); },
  
  FRAUD_CODES: [
    "10.4", "37", 10.4, 37, "AA",
    "Card Member does not recognise Transaction or Transaction Amount-6014",
    "Card not present fraud-F29", "Card not present-4540",
    "Does Not Recognize/ Remember/ No Knowledge-127", "Fraud-193",
    "Fraudulent Transactions-193", "No knowledge-127", "U02", "Fraud",
    "Card Member claims fraud-6006",
    "Does Not Recognize/ Remember/ No Knowledge-176"
  ],
  
  MAX_BATCH_SIZE: 100
};

// ═══════════════════════════════════════════════════════════════════════════
// SHEET ACCESS (with null checks to prevent getRange errors)
// ═══════════════════════════════════════════════════════════════════════════

function getRebuttalRcaSheet_() {
  try {
    const sheetId = REBUTTAL_CONFIG.RCA_SHEET_ID;
    if (!sheetId) {
      console.error('REBUTTAL_CONFIG.RCA_SHEET_ID is empty or undefined');
      return null;
    }
    const ss = SpreadsheetApp.openById(sheetId);
    if (!ss) {
      console.error('Cannot open spreadsheet with ID: ' + sheetId);
      return null;
    }
    const sheet = ss.getSheetByName("database");
    if (!sheet) {
      console.error('Sheet "database" not found in spreadsheet: ' + sheetId + '. Available sheets: ' + ss.getSheets().map(s => s.getName()).join(', '));
      return null;
    }
    return sheet;
  } catch (e) {
    console.error('Error accessing RCA sheet: ' + e.message);
    return null;
  }
}

function getRebuttalRcaDashboard_() {
  try {
    const sheetId = REBUTTAL_CONFIG.RCA_SHEET_ID;
    if (!sheetId) {
      console.error('REBUTTAL_CONFIG.RCA_SHEET_ID is empty or undefined');
      return null;
    }
    const ss = SpreadsheetApp.openById(sheetId);
    if (!ss) {
      console.error('Cannot open spreadsheet with ID: ' + sheetId);
      return null;
    }
    const sheet = ss.getSheetByName("Dashboard");
    if (!sheet) {
      console.error('Sheet "Dashboard" not found in spreadsheet: ' + sheetId);
      return null;
    }
    return sheet;
  } catch (e) {
    console.error('Error accessing RCA Dashboard: ' + e.message);
    return null;
  }
}

function getRebuttalDatasheet_() {
  try {
    const sheetId = REBUTTAL_CONFIG.DATASHEET_ID;
    if (!sheetId) {
      console.error('REBUTTAL_CONFIG.DATASHEET_ID is empty or undefined');
      return null;
    }
    const ss = SpreadsheetApp.openById(sheetId);
    if (!ss) {
      console.error('Cannot open spreadsheet with ID: ' + sheetId);
      return null;
    }
    const sheet = ss.getSheetByName("datasheet");
    if (!sheet) {
      console.error('Sheet "datasheet" not found in spreadsheet: ' + sheetId + '. Available sheets: ' + ss.getSheets().map(s => s.getName()).join(', '));
      return null;
    }
    return sheet;
  } catch (e) {
    console.error('Error accessing Rebuttal Datasheet: ' + e.message);
    return null;
  }
}

function getRebuttalTempSheet_() {
  try {
    const sheetId = REBUTTAL_CONFIG.DATASHEET_ID;
    if (!sheetId) {
      console.error('REBUTTAL_CONFIG.DATASHEET_ID is empty or undefined');
      return null;
    }
    const ss = SpreadsheetApp.openById(sheetId);
    if (!ss) {
      console.error('Cannot open spreadsheet with ID: ' + sheetId);
      return null;
    }
    const sheet = ss.getSheetByName("Temp");
    if (!sheet) {
      console.error('Sheet "Temp" not found in spreadsheet: ' + sheetId);
      return null;
    }
    return sheet;
  } catch (e) {
    console.error('Error accessing Temp sheet: ' + e.message);
    return null;
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// MAIN: Create Rebuttals (Batch)
// ═══════════════════════════════════════════════════════════════════════════

function createRebuttals() {
  const rcasheet = getRebuttalRcaSheet_();
  
  // Guard clause: If sheet access failed, return early with error
  if (!rcasheet) {
    console.error('createRebuttals: Cannot access RCA sheet - check REBUTTAL_CONFIG.RCA_SHEET_ID and Script Properties');
    return { created: 0, error: 'Sheet access failed - check logs for details' };
  }
  
  const lastrowrca = rcasheet.getLastRow();
  
  // Statuses that should NOT get rebuttals
  const SKIP_STATUSES = ["Pending", "Resolved", "Chargeback Accepted", ""];
  // Statuses in column 24 (Templates) that indicate already processing/done
  const PROCESSING_STATUSES = ["Processing", "Ready", "Error"];
  
  const RCAValues = rcasheet.getRange(1, 17, lastrowrca).getValues();
  const OutcomeValues = rcasheet.getRange(1, 25, lastrowrca).getValues();
  const TemplatesValues = rcasheet.getRange(1, 24, lastrowrca).getValues();
  
  // Find rows that need rebuttals (and aren't already processing/done)
  const needed = [];
  for (let i = 0; i < RCAValues.length; i++) {
    const rcaValue = RCAValues[i].toString();
    const templateValue = TemplatesValues[i].toString().trim();
    
    if (SKIP_STATUSES.indexOf(rcaValue) < 0 &&
        RCAValues[i] &&
        OutcomeValues[i] == "" &&
        PROCESSING_STATUSES.indexOf(templateValue) < 0) {
      needed.push(i + 1);
    }
  }
  
  if (needed.length === 0) {
    console.log("No rebuttals to create");
    return { created: 0 };
  }
  
  const template = DriveApp.getFileById(REBUTTAL_CONFIG.TEMPLATE_ID);
  const tempFolder = DriveApp.getFolderById(REBUTTAL_CONFIG.TEMP_FOLDER_ID);
  const pdfFolder = DriveApp.getFolderById(REBUTTAL_CONFIG.PDF_FOLDER_ID);
  const dashboard = getRebuttalRcaDashboard_();
  
  let created = 0;
  const batchSize = Math.min(needed.length, REBUTTAL_CONFIG.MAX_BATCH_SIZE);
  
  for (let i = 0; i < batchSize; i++) {
    try {
      const row = needed[i];
      
      // IMMEDIATELY mark as Processing to prevent race condition
      rcasheet.getRange(row, 24).setValue("Processing");
      SpreadsheetApp.flush(); // Force write to prevent duplicates
      
      const result = createSingleRebuttal_(row, rcasheet, template, tempFolder, pdfFolder, dashboard);
      if (result) {
        created++;
      } else {
        // If creation failed, mark as Error so we can retry manually
        rcasheet.getRange(row, 24).setValue("Error");
      }
    } catch (e) {
      console.error(`Error creating rebuttal for row ${needed[i]}: ${e.message}`);
      // Mark as Error so it doesn't get retried automatically
      try {
        rcasheet.getRange(needed[i], 24).setValue("Error - " + e.message.substring(0, 50));
      } catch (e2) {}
    }
  }
  
  return { created, total: needed.length };
}

// ═══════════════════════════════════════════════════════════════════════════
// CREATE SINGLE REBUTTAL
// ═══════════════════════════════════════════════════════════════════════════

function createSingleRebuttal_(row, rcasheet, template, tempFolder, pdfFolder, dashboard) {
  // Create temp copy of template
  const tempFile = template.makeCopy(tempFolder).setName("Rebuttal_Response");
  const tempDoc = DocumentApp.openById(tempFile.getId());
  
  // Get case info
  const caseinfo = getRebuttalCaseInfo_(row, rcasheet);
  const caseid = caseinfo[0];
  const itn = caseinfo[1];
  const bookinginterface = caseinfo[13];
  const code = caseinfo[3];
  const dAmt = caseinfo[6];
  const issuernotes = caseinfo[19];
  const rnum = caseinfo[8];
  const response = caseinfo[20];
  const policy = caseinfo[28];
  
  // Get reservation data
  const resinfo = getRebuttalReservationData_(itn);
  const wex = resinfo[18];
  const aix = resinfo[19];
  const rfsht = resinfo[20];
  const ccontact = resinfo[21];
  const ip = resinfo[22];
  const nonfraud1 = resinfo[23];
  const nonfraud2 = resinfo[24];
  
  // Replace placeholders
  const body = tempDoc.getBody();
  body.replaceText("{Case Number}", caseid);
  body.replaceText("{ITN Number}", itn);
  body.replaceText("{Code}", code);
  body.replaceText("{Dispute Amt}", dAmt);
  body.replaceText("{IssuerNotes}", issuernotes || "No dispute declaration document received.");
  body.replaceText("{Reservation Number}", rnum);
  body.replaceText("{Cancel Policy}", policy);
  body.replaceText("{NotesforBank}", response);
  
  // Add screenshots
  let docMsg = "Screenshots included: ";
  const hasScreenshots = [rfsht, wex, aix, ccontact, ip, nonfraud1, nonfraud2]
    .some(s => s && s.toString().length > 4);
  
  if (hasScreenshots) {
    if (rfsht && rfsht.toString().length > 4) {
      insertRebuttalImage_(rfsht, "Proof of refund", tempDoc);
      docMsg += "Proof of refund, ";
    }
    if (wex && wex.toString().length > 4) {
      insertRebuttalImage_(wex, "Hotel billed us wholesale rates for this reservation.", tempDoc);
      docMsg += "Wex, ";
    }
    if (aix && aix.toString().length > 4) {
      insertRebuttalImage_(aix, "Accertify Index Score", tempDoc);
      docMsg += "index score, ";
    }
    if (ccontact && ccontact.toString().length > 4) {
      insertRebuttalImage_(ccontact, "Customer contact", tempDoc);
      docMsg += "salesforce, ";
    }
    if (ip && ip.toString().length > 4) {
      insertRebuttalImage_(ip, "Proof of Non Fraud", tempDoc);
      docMsg += "non fraud, ";
    }
    if (nonfraud1 && nonfraud1.toString().length > 4) {
      insertRebuttalImage_(nonfraud1, "Proof of Non Fraud #1", tempDoc);
      docMsg += "non fraud 1, ";
    }
    if (nonfraud2 && nonfraud2.toString().length > 4) {
      insertRebuttalImage_(nonfraud2, "Proof of Non Fraud #2", tempDoc);
      docMsg += "non fraud 2, ";
    }
  } else {
    body.appendParagraph("").appendText("Documents attached separately.");
  }
  
  tempDoc.saveAndClose();
  
  // Create PDF
  const pdfContent = tempDoc.getAs(MimeType.PDF);
  pdfFolder.createFile(pdfContent).setName(`${caseid}_${itn}_Rebuttal`);
  
  // Update dashboard
  const createdDt = new Date();
  dashboard.getRange("I7").setValue(Utilities.formatDate(createdDt, "IST", "MM/dd/yyyy h:mm"));
  
  // Cleanup temp doc
  DriveApp.getFileById(tempFile.getId()).setTrashed(true);
  
  // Check if fraud case
  if (REBUTTAL_CONFIG.FRAUD_CODES.includes(code) || 
      (code == 13.7 && bookinginterface === "call_center")) {
    docMsg += "\r\nEmail Confirmation Required";
  }
  
  // Update RCA sheet
  rcasheet.getRange(row, 24).setValue("Ready");
  const internalnote = rcasheet.getRange(row, 22).getValue();
  rcasheet.getRange(row, 22).setValue(internalnote + "\r\n" + docMsg);
  
  console.log(`Rebuttal created: ${itn}`);
  return true;
}

// ═══════════════════════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════

function getRebuttalCaseInfo_(row, rcasheet) {
  return rcasheet.getRange(row, 1, 1, 29).getValues()[0];
}

function getRebuttalReservationData_(ITN) {
  const ss = getRebuttalDatasheet_();
  const tempSheet = getRebuttalTempSheet_();
  
  // Guard clause: Return empty array if sheets can't be accessed
  if (!ss || !tempSheet) {
    console.error('getRebuttalReservationData_: Could not access datasheet or temp sheet for ITN: ' + ITN);
    return new Array(25).fill("");
  }
  
  try {
    tempSheet.getRange("B2").setFormula(`=match("${ITN}",datasheet!A:A,0)`);
    SpreadsheetApp.flush();
    const row = tempSheet.getRange("B2").getValue();
    tempSheet.getRange("B2").clearContent();
    
    if (!row || row === "#N/A") return new Array(25).fill("");
    
    return ss.getRange(row, 2, 1, 25).getValues()[0];
  } catch (e) {
    console.error('getRebuttalReservationData_: Error for ITN ' + ITN + ': ' + e.message);
    return new Array(25).fill("");
  }
}

function insertRebuttalImage_(imageFileId, text, doc) {
  const MAX_WIDTH = 680;
  const MAX_HEIGHT = 700;
  
  try {
    const body = doc.getBody();
    const file = DriveApp.getFileById(imageFileId);
    const imgBlob = file.getBlob();
    
    if (text) {
      body.appendParagraph("").appendText(text);
    }
    
    const tempImage = body.appendParagraph("").appendInlineImage(imgBlob);
    const originalWidth = tempImage.getWidth();
    const originalHeight = tempImage.getHeight();
    
    const widthScale = MAX_WIDTH / originalWidth;
    const heightScale = MAX_HEIGHT / originalHeight;
    const scale = Math.min(widthScale, heightScale);
    
    tempImage.setWidth(originalWidth * scale).setHeight(originalHeight * scale);
    body.appendPageBreak();
  } catch (e) {
    console.error(`Error inserting image ${imageFileId}: ${e.message}`);
  }
}
/**
 * CBQ Text Analytics Engine - Configuration
 * 
 * AI-Powered lexical intelligence system for chargeback rebuttal analysis.
 * Uses ChatGPT for classification, with customizable word filtering.
 * 
 * @module CBQConfig
 */

const CBQConfig = {
  // =====================================================
  // Sheet Names
  // =====================================================
  
  SHEETS: {
    DOCUMENTS: 'CBQ_Documents',
    WORDS: 'CBQ_Words',
    WORD_STATS: 'CBQ_WordStats',
    CONFIG: 'CBQ_Config',
    PROCESSING_QUEUE: 'CBQ_ProcessingQueue',
    FILTERED_WORDS: 'CBQ_FilteredWords'  // NEW: User-defined word exclusions
  },
  
  // =====================================================
  // AI Configuration
  // =====================================================
  
  AI: {
    ENABLED: true,
    MODEL: 'gpt-4o-mini',
    MAX_TOKENS: 500,
    TEMPERATURE: 0.1,
    API_KEY_PROPERTY: 'OPENAI_API_KEY',  // Script Property name
    
    // Rate limiting
    DELAY_BETWEEN_CALLS_MS: 200,
    MAX_RETRIES: 3
  },
  
  // =====================================================
  // Outcome Labels (3-Batch Classification)
  // =====================================================
  
  OUTCOME_LABELS: {
    AUTO_WIN_72H: { 
      id: 'AUTO_WIN_72H', 
      name: '72h Perfect', 
      weight: 1.0, 
      corpus: 'STRONG_POSITIVE',
      description: 'Strong evidence - auto-win within 72 hours'
    },
    WIN_30_90D: { 
      id: 'WIN_30_90D', 
      name: '90d Good', 
      weight: 0.6, 
      corpus: 'WEAK_POSITIVE',
      description: 'Good evidence - win after 30-90 day review'
    },
    AUTO_REJECT_72H: { 
      id: 'AUTO_REJECT_72H', 
      name: '72h Negative', 
      weight: -1.0, 
      corpus: 'NEGATIVE',
      description: 'Weak evidence - words to avoid'
    }
  },
  
  // =====================================================
  // Card Types (Auto-detected by AI)
  // =====================================================
  
  CARD_TYPES: ['VI', 'MC', 'DI', 'AX'],
  
  // =====================================================
  // Schema Definitions
  // =====================================================
  
  DOCUMENTS_SCHEMA: [
    'id',
    'sourceType',        // 'PDF' or 'TEXT'
    'driveFileId',       // Google Drive file ID (if PDF)
    'filename',          // Original filename
    'reasonCode',        // AI-detected reason code
    'cardType',          // AI-detected card type (VI, MC, DI, AX)
    'outcomeLabel',      // AUTO_WIN_72H, WIN_30_90D, AUTO_REJECT_72H
    'confidence',        // AI confidence score (0-1)
    'issuer',            // Issuer bank (optional)
    'dateRange',         // Date range coverage
    'rawText',           // Extracted text
    'wordCount',         // Total words extracted
    'keyTerms',          // AI-extracted key terms (JSON array)
    'status',            // 'QUEUED', 'PROCESSING', 'COMPLETED', 'ERROR'
    'errorMessage',      // Error details if failed
    'processedAt',       // When processing completed
    'uploadedBy',        // User email
    'uploadedAt',        // Upload timestamp
    'batchId'            // For batch processing tracking
  ],
  
  WORDS_SCHEMA: [
    'id',
    'word',              // Normalized word (lowercase, stemmed)
    'documentId',        // FK to CBQ_Documents
    'reasonCode',        // Denormalized for query speed
    'outcomeLabel',      // Denormalized
    'termFrequency',     // Count in this document
    'positions'          // JSON array of positions
  ],
  
  WORD_STATS_SCHEMA: [
    'word',              // Unique word (PK)
    'reasonCode',        // Reason code or 'GLOBAL'
    'totalOccurrences',  // Total times word appears
    'docFreqTotal',      // Total documents containing word
    'docFreqAutoWin',    // Docs in AUTO_WIN_72H
    'docFreqWin',        // Docs in WIN_30_90D
    'docFreqReject',     // Docs in AUTO_REJECT_72H
    'tfidfAutoWin',      // TF-IDF for AUTO_WIN_72H
    'tfidfWin',          // TF-IDF for WIN_30_90D
    'tfidfReject',       // TF-IDF for AUTO_REJECT_72H
    'positiveWeight',    // Weighted positive score
    'negativeWeight',    // Weighted negative score
    'netSignalScore',    // Positive - Negative
    'cooccurrences',     // JSON: top co-occurring words
    'updatedAt'          // Last calculation time
  ],
  
  /**
   * CBQ_FilteredWords schema - user-defined word exclusions
   */
  FILTERED_WORDS_SCHEMA: [
    'word',              // Word to exclude (lowercase)
    'category',          // 'COMPANY', 'STOPWORD', 'DATE', 'CUSTOM'
    'addedBy',           // User email
    'addedAt'            // Timestamp
  ],
  
  CONFIG_SCHEMA: [
    'key',
    'value',
    'description',
    'updatedBy',
    'updatedAt'
  ],
  
  PROCESSING_QUEUE_SCHEMA: [
    'id',
    'batchId',
    'documentId',
    'status',            // 'PENDING', 'IN_PROGRESS', 'COMPLETED', 'FAILED'
    'startedAt',
    'completedAt',
    'errorMessage'
  ],
  
  // =====================================================
  // Default Configuration Values
  // =====================================================
  
  DEFAULTS: {
    MIN_WORD_LENGTH: 3,
    MAX_WORD_LENGTH: 50,
    BATCH_SIZE: 10,
    WORD_BATCH_SIZE: 500,
    
    WEIGHT_AUTO_WIN_72H: 1.0,
    WEIGHT_WIN_30_90D: 0.6,
    WEIGHT_AUTO_REJECT_72H: 1.0,
    
    TOP_WORDS_LIMIT: 100,
    COOCCURRENCE_LIMIT: 20
  },
  
  // =====================================================
  // Default Filtered Words (Pre-loaded exclusions)
  // =====================================================
  
  DEFAULT_FILTERED_WORDS: {
    COMPANY: [
      'priceline', 'pricelinecom', 'pps', 'ppsnl'
    ],
    STOPWORD: [
      'the', 'a', 'an', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for',
      'of', 'with', 'by', 'from', 'as', 'is', 'was', 'are', 'were', 'been',
      'be', 'have', 'has', 'had', 'do', 'does', 'did', 'will', 'would',
      'could', 'should', 'may', 'might', 'must', 'shall', 'can', 'need',
      'this', 'that', 'these', 'those', 'it', 'its', 'they', 'them',
      'their', 'we', 'our', 'you', 'your', 'he', 'she', 'him', 'her',
      'i', 'me', 'my', 'not', 'no', 'yes', 'if', 'then', 'else', 'when',
      'where', 'which', 'who', 'what', 'how', 'why', 'all', 'each',
      'every', 'both', 'few', 'more', 'most', 'other', 'some', 'such',
      'only', 'own', 'same', 'than', 'too', 'very', 'just', 'also',
      'customer', 'cancellation', 'booking', 'reservation', 'hotel'
    ],
    DATE: [
      'january', 'february', 'march', 'april', 'may', 'june',
      'july', 'august', 'september', 'october', 'november', 'december',
      'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'
    ]
  },
  
  // =====================================================
  // Document Source Types
  // =====================================================
  
  SOURCE_TYPES: {
    PDF: 'PDF',
    TEXT: 'TEXT'
  },
  
  DOCUMENT_STATUS: {
    QUEUED: 'QUEUED',
    PROCESSING: 'PROCESSING',
    COMPLETED: 'COMPLETED',
    ERROR: 'ERROR'
  },
  
  // =====================================================
  // Audit Actions
  // =====================================================
  
  AUDIT_ACTIONS: {
    DOCUMENT_UPLOAD: 'CBQ_DOCUMENT_UPLOAD',
    DOCUMENT_PROCESS: 'CBQ_DOCUMENT_PROCESS',
    STATS_CALCULATE: 'CBQ_STATS_CALCULATE',
    CONFIG_UPDATE: 'CBQ_CONFIG_UPDATE',
    FILTER_ADD: 'CBQ_FILTER_ADD',
    FILTER_REMOVE: 'CBQ_FILTER_REMOVE',
    EXPORT: 'CBQ_EXPORT'
  }
};

// Freeze to prevent modifications
Object.freeze(CBQConfig);
Object.freeze(CBQConfig.SHEETS);
Object.freeze(CBQConfig.AI);
Object.freeze(CBQConfig.OUTCOME_LABELS);
Object.freeze(CBQConfig.DEFAULTS);
Object.freeze(CBQConfig.SOURCE_TYPES);
Object.freeze(CBQConfig.DOCUMENT_STATUS);
Object.freeze(CBQConfig.AUDIT_ACTIONS);
/**
 * CBQ Controller
 * 
 * AI-Powered API endpoints for CBQ module.
 * Uses ChatGPT for document classification and keyword extraction.
 * 
 * @module CBQController
 */

const CBQController = {
  // =====================================================
  // AI-Powered Document Upload
  // =====================================================
  
  /**
   * Upload and process text document with AI
   * @param {Object} data - Document data
   * @returns {Object} Created and classified document
   */
  uploadTextWithAI(data) {
    const user = Auth.getCurrentUser();
    
    if (!data.text || data.text.trim().length < 50) {
      throw new Error('Text content is required (minimum 50 characters)');
    }
    if (!data.batchType) {
      throw new Error('Batch type is required (AUTO_WIN_72H, WIN_30_90D, AUTO_REJECT_72H)');
    }
    
    // Validate batch type
    if (!CBQConfig.OUTCOME_LABELS[data.batchType]) {
      throw new Error(`Invalid batch type: ${data.batchType}`);
    }
    
    // Classify with AI
    const classification = AIProcessor.classifyDocument(data.text);
    
    if (!classification.success) {
      throw new Error(`AI classification failed: ${classification.error}`);
    }
    
    const document = CBQDocumentsDAO.create({
      sourceType: CBQConfig.SOURCE_TYPES.TEXT,
      filename: data.filename || 'Pasted Text',
      reasonCode: classification.data.reasonCode,
      cardType: classification.data.cardType,
      outcomeLabel: data.batchType,  // User-specified batch
      confidence: classification.data.confidence,
      issuer: data.issuer || '',
      dateRange: data.dateRange || '',
      rawText: data.text,
      keyTerms: JSON.stringify(classification.data.keyTerms || [])
    }, user);
    
    return {
      document,
      classification: classification.data
    };
  },
  
  /**
   * Upload and process PDF with AI
   * @param {Object} data - PDF data
   * @returns {Object} Created and classified document
   */
  uploadPDFWithAI(data) {
    const user = Auth.getCurrentUser();
    
    if (!data.driveFileId) {
      throw new Error('Drive file ID is required');
    }
    if (!data.batchType) {
      throw new Error('Batch type is required (AUTO_WIN_72H, WIN_30_90D, AUTO_REJECT_72H)');
    }
    
    // Validate batch type
    if (!CBQConfig.OUTCOME_LABELS[data.batchType]) {
      throw new Error(`Invalid batch type: ${data.batchType}`);
    }
    
    try {
      const file = DriveApp.getFileById(data.driveFileId);
      const filename = file.getName();
      
      // Extract text from PDF
      const extraction = TextExtractor.extractFromPDF(data.driveFileId);
      if (!extraction.success) {
        throw new Error(`PDF extraction failed: ${extraction.error}`);
      }
      
      // Classify with AI
      const classification = AIProcessor.classifyDocument(extraction.text);
      
      if (!classification.success) {
        throw new Error(`AI classification failed: ${classification.error}`);
      }
      
      const document = CBQDocumentsDAO.create({
        sourceType: CBQConfig.SOURCE_TYPES.PDF,
        driveFileId: data.driveFileId,
        filename: filename,
        reasonCode: classification.data.reasonCode,
        cardType: classification.data.cardType,
        outcomeLabel: data.batchType,
        confidence: classification.data.confidence,
        issuer: data.issuer || '',
        dateRange: data.dateRange || '',
        rawText: extraction.text,
        keyTerms: JSON.stringify(classification.data.keyTerms || [])
      }, user);
      
      return {
        document,
        classification: classification.data,
        extractionMethod: extraction.method
      };
      
    } catch (e) {
      throw new Error(`Cannot process file: ${e.message}`);
    }
  },
  
  /**
   * Bulk upload PDFs from a Drive folder with AI processing
   * @param {Object} data - Folder data with batchType
   * @returns {Object} Upload and processing result
   */
  uploadFromFolderWithAI(data) {
    const user = Auth.getCurrentUser();
    
    if (!data.folderId) {
      throw new Error('Folder ID is required');
    }
    if (!data.batchType) {
      throw new Error('Batch type is required (AUTO_WIN_72H, WIN_30_90D, AUTO_REJECT_72H)');
    }
    
    if (!CBQConfig.OUTCOME_LABELS[data.batchType]) {
      throw new Error(`Invalid batch type: ${data.batchType}`);
    }
    
    // Get PDF files from folder
    const files = TextExtractor.getFilesFromFolder(data.folderId);
    
    if (files.length === 0) {
      throw new Error('No PDF files found in folder');
    }
    
    const batchId = Utils.generateUUID();
    const results = {
      batchId,
      batchType: data.batchType,
      total: files.length,
      processed: 0,
      failed: 0,
      documents: []
    };
    
    // Process each file
    for (const file of files) {
      try {
        // Extract text
        const extraction = TextExtractor.extractFromPDF(file.id);
        if (!extraction.success) {
          results.documents.push({
            filename: file.name,
            success: false,
            error: extraction.error
          });
          results.failed++;
          continue;
        }
        
        // Classify with AI
        const classification = AIProcessor.classifyDocument(extraction.text);
        
        if (!classification.success) {
          results.documents.push({
            filename: file.name,
            success: false,
            error: classification.error
          });
          results.failed++;
          continue;
        }
        
        // Create document record
        const doc = CBQDocumentsDAO.create({
          sourceType: CBQConfig.SOURCE_TYPES.PDF,
          driveFileId: file.id,
          filename: file.name,
          reasonCode: classification.data.reasonCode,
          cardType: classification.data.cardType,
          outcomeLabel: data.batchType,
          confidence: classification.data.confidence,
          issuer: data.issuer || '',
          dateRange: data.dateRange || '',
          rawText: extraction.text,
          keyTerms: JSON.stringify(classification.data.keyTerms || []),
          batchId: batchId
        }, user);
        
        results.documents.push({
          id: doc.id,
          filename: file.name,
          success: true,
          reasonCode: classification.data.reasonCode,
          cardType: classification.data.cardType,
          confidence: classification.data.confidence
        });
        results.processed++;
        
        // Brief pause for rate limiting
        Utilities.sleep(CBQConfig.AI.DELAY_BETWEEN_CALLS_MS);
        
      } catch (e) {
        results.documents.push({
          filename: file.name,
          success: false,
          error: e.message
        });
        results.failed++;
      }
    }
    
    return results;
  },
  
  // =====================================================
  // Filtered Words Management
  // =====================================================
  
  /**
   * Get all filtered words
   * @returns {Array} Filtered words with metadata
   */
  getFilteredWords() {
    return CBQFilteredWordsDAO.getAll();
  },
  
  /**
   * Add a word to filter list
   * @param {string} word - Word to filter
   * @param {string} category - Category (COMPANY, STOPWORD, DATE, CUSTOM)
   * @returns {Object} Created filter entry
   */
  addFilteredWord(word, category = 'CUSTOM') {
    const user = Auth.getCurrentUser();
    
    if (!word || word.trim().length < 2) {
      throw new Error('Word is required (minimum 2 characters)');
    }
    
    const normalizedWord = word.toLowerCase().trim();
    
    // Check if already exists
    const existing = CBQFilteredWordsDAO.getByWord(normalizedWord);
    if (existing) {
      throw new Error(`Word "${normalizedWord}" is already filtered`);
    }
    
    const entry = CBQFilteredWordsDAO.create({
      word: normalizedWord,
      category: category
    }, user);
    
    AuditDAO.log({
      action: CBQConfig.AUDIT_ACTIONS.FILTER_ADD,
      entityType: 'CBQ_FilteredWords',
      entityId: normalizedWord,
      user: user,
      module: 'CBQ',
      details: `Added filtered word: ${normalizedWord} (${category})`
    });
    
    return entry;
  },
  
  /**
   * Remove a word from filter list
   * @param {string} word - Word to unfilter
   * @returns {boolean} Success
   */
  removeFilteredWord(word) {
    const user = Auth.getCurrentUser();
    
    const normalizedWord = word.toLowerCase().trim();
    const result = CBQFilteredWordsDAO.delete(normalizedWord);
    
    if (result) {
      AuditDAO.log({
        action: CBQConfig.AUDIT_ACTIONS.FILTER_REMOVE,
        entityType: 'CBQ_FilteredWords',
        entityId: normalizedWord,
        user: user,
        module: 'CBQ',
        details: `Removed filtered word: ${normalizedWord}`
      });
    }
    
    return result;
  },
  
  /**
   * Initialize default filtered words
   * @returns {Object} Initialization result
   */
  initializeFilteredWords() {
    const user = Auth.getCurrentUser();
    let added = 0;
    
    for (const [category, words] of Object.entries(CBQConfig.DEFAULT_FILTERED_WORDS)) {
      for (const word of words) {
        try {
          const existing = CBQFilteredWordsDAO.getByWord(word);
          if (!existing) {
            CBQFilteredWordsDAO.create({ word, category }, user);
            added++;
          }
        } catch (e) {
          // Skip duplicates
        }
      }
    }
    
    return { added, message: `Initialized ${added} filtered words` };
  },
  
  // =====================================================
  // Legacy Upload (without AI - manual reason code)
  // =====================================================
  
  uploadText(data) {
    const user = Auth.getCurrentUser();
    
    if (!data.text || data.text.trim().length < 10) {
      throw new Error('Text content is required (minimum 10 characters)');
    }
    if (!data.outcomeLabel) {
      throw new Error('Outcome label is required');
    }
    
    const document = CBQDocumentsDAO.create({
      sourceType: CBQConfig.SOURCE_TYPES.TEXT,
      filename: data.filename || 'Pasted Text',
      reasonCode: data.reasonCode || 'UNKNOWN',
      cardType: data.cardType || 'UNKNOWN',
      outcomeLabel: data.outcomeLabel,
      confidence: 0,
      rawText: data.text
    }, user);
    
    return document;
  },
  
  // =====================================================
  // Analytics (Word Cloud with Filtering)
  // =====================================================
  
  /**
   * Get word cloud data with filtered words excluded
   * @param {Object} options - Options (mode, limit, reasonCode)
   * @returns {Array} Word cloud data
   */
  getWordCloudData(options = {}) {
    const reasonCode = options.reasonCode || 'GLOBAL';
    const mode = options.mode || 'signal';
    const limit = options.limit || 100;
    
    // Get filtered words set
    const filteredWords = new Set(
      CBQFilteredWordsDAO.getAll().map(f => f.word.toLowerCase())
    );
    
    const statsMap = CBQStatsDAO.getAllAsMap(reasonCode);
    
    // Filter out excluded words
    const filteredStatsMap = new Map();
    for (const [word, stats] of statsMap) {
      if (!filteredWords.has(word.toLowerCase())) {
        filteredStatsMap.set(word, stats);
      }
    }
    
    return MetricsEngine.prepareWordCloudData(filteredStatsMap, mode, limit);
  },
  
  /**
   * Get bubble chart data with filtered words excluded
   * @param {Object} options - Options
   * @returns {Array} Bubble chart data
   */
  getBubbleChartData(options = {}) {
    const reasonCode = options.reasonCode || 'GLOBAL';
    const limit = options.limit || 100;
    
    // Get filtered words set
    const filteredWords = new Set(
      CBQFilteredWordsDAO.getAll().map(f => f.word.toLowerCase())
    );
    
    const statsMap = CBQStatsDAO.getAllAsMap(reasonCode);
    
    // Filter out excluded words
    const filteredStatsMap = new Map();
    for (const [word, stats] of statsMap) {
      if (!filteredWords.has(word.toLowerCase())) {
        filteredStatsMap.set(word, stats);
      }
    }
    
    return MetricsEngine.prepareBubbleChartData(filteredStatsMap, limit);
  },
  
  getDocuments(filters = {}) {
    return CBQDocumentsDAO.getAll(filters);
  },
  
  getQueueStatus() {
    return BatchProcessor.getQueueStatus();
  },
  
  /**
   * Process queued documents
   * @param {number} limit - Max documents to process
   * @returns {Object} Processing result
   */
  processQueue(limit = 10) {
    return BatchProcessor.processBatch(limit);
  },
  
  /**
   * Recalculate word statistics
   * @param {string} reasonCode - Reason code to recalculate (or 'GLOBAL')
   * @returns {Object} Calculation result
   */
  recalculateStats(reasonCode = 'GLOBAL') {
    return BatchProcessor.recalculateStats(reasonCode);
  },
  
  getStatsSummary() {
    return CBQStatsDAO.getSummary();
  },
  
  getTopWords(options = {}) {
    const reasonCode = options.reasonCode || 'GLOBAL';
    const limit = options.limit || 50;
    const direction = options.direction || 'positive';
    
    return CBQStatsDAO.getTopWords(reasonCode, limit, direction);
  },
  
  getDocumentCounts() {
    return CBQDocumentsDAO.getCountsByOutcome();
  },
  
  // =====================================================
  // AI Utilities
  // =====================================================
  
  /**
   * Test OpenAI API connection
   * @returns {Object} Connection test result
   */
  testAIConnection() {
    return AIProcessor.testConnection();
  },
  
  /**
   * Estimate cost for batch processing
   * @param {number} documentCount - Number of documents
   * @returns {Object} Cost estimate
   */
  estimateAICost(documentCount) {
    return AIProcessor.estimateCost(documentCount);
  },
  
  // =====================================================
  // Configuration
  // =====================================================
  
  /**
   * Get CBQ configuration options
   * @returns {Object} Config options
   */
  getConfig() {
    return {
      outcomeLabels: Object.values(CBQConfig.OUTCOME_LABELS),
      cardTypes: CBQConfig.CARD_TYPES,
      defaults: CBQConfig.DEFAULTS,
      aiEnabled: CBQConfig.AI.ENABLED,
      aiModel: CBQConfig.AI.MODEL
    };
  },
  
  exportStats(reasonCode = 'GLOBAL') {
    const user = Auth.getCurrentUser();
    
    AuditDAO.log({
      action: CBQConfig.AUDIT_ACTIONS.EXPORT,
      entityType: 'CBQ_WordStats',
      entityId: reasonCode,
      user: user,
      module: 'CBQ',
      details: `Exported word statistics for ${reasonCode}`
    });
    
    const stats = CBQStatsDAO.getByReasonCode(reasonCode);
    
    const headers = [
      'Word', 'Reason Code', 'Total Occurrences',
      'Doc Freq (Auto Win)', 'Doc Freq (Win)', 'Doc Freq (Reject)',
      'TF-IDF (Auto Win)', 'TF-IDF (Win)', 'TF-IDF (Reject)',
      'Positive Weight', 'Negative Weight', 'Net Signal Score'
    ];
    
    const rows = stats.map(s => [
      s.word, s.reasonCode, s.totalOccurrences,
      s.docFreqAutoWin, s.docFreqWin, s.docFreqReject,
      s.tfidfAutoWin, s.tfidfWin, s.tfidfReject,
      s.positiveWeight, s.negativeWeight, s.netSignalScore
    ]);
    
    return { headers, rows };
  }
};

Object.freeze(CBQController);
/**
 * CBQ Documents Data Access Object
 * 
 * CRUD operations for CBQ_Documents sheet.
 * Handles document metadata and status tracking.
 * 
 * @module CBQDocumentsDAO
 */

const CBQDocumentsDAO = {
  /**
   * Create a new document record
   * @param {Object} data - Document data
   * @param {string} uploadedBy - User email
   * @returns {Object} Created document
   */
  create(data, uploadedBy) {
    const now = Utils.now();
    const document = {
      id: Utils.generateUUID(),
      sourceType: data.sourceType || CBQConfig.SOURCE_TYPES.TEXT,
      driveFileId: data.driveFileId || '',
      filename: data.filename || 'Untitled',
      reasonCode: data.reasonCode || 'GLOBAL',
      outcomeLabel: data.outcomeLabel || '',
      issuer: data.issuer || '',
      dateRange: data.dateRange || '',
      rawText: data.rawText || '',
      wordCount: 0,
      status: CBQConfig.DOCUMENT_STATUS.QUEUED,
      errorMessage: '',
      processedAt: '',
      uploadedBy: uploadedBy,
      uploadedAt: now,
      batchId: data.batchId || ''
    };
    
    SheetsService.appendRow(CBQConfig.SHEETS.DOCUMENTS, document);
    
    AuditDAO.log({
      action: CBQConfig.AUDIT_ACTIONS.DOCUMENT_UPLOAD,
      entityType: 'CBQ_Document',
      entityId: document.id,
      user: uploadedBy,
      module: 'CBQ',
      details: `Uploaded: ${document.filename} (${document.sourceType})`
    });
    
    return document;
  },
  
  /**
   * Get document by ID
   * @param {string} id - Document ID
   * @returns {Object|null} Document or null
   */
  getById(id) {
    const rows = SheetsService.getAllRows(CBQConfig.SHEETS.DOCUMENTS);
    return rows.find(r => r.id === id) || null;
  },
  
  /**
   * Get documents by status
   * @param {string} status - Document status
   * @param {number} limit - Max documents to return
   * @returns {Array} Documents
   */
  getByStatus(status, limit = 100) {
    const rows = SheetsService.getAllRows(CBQConfig.SHEETS.DOCUMENTS);
    return rows.filter(r => r.status === status).slice(0, limit);
  },
  
  /**
   * Get documents pending processing
   * @param {number} limit - Max documents
   * @returns {Array} Queued documents
   */
  getQueued(limit = 10) {
    return this.getByStatus(CBQConfig.DOCUMENT_STATUS.QUEUED, limit);
  },
  
  /**
   * Get documents by outcome label
   * @param {string} outcomeLabel - Outcome label
   * @returns {Array} Documents
   */
  getByOutcome(outcomeLabel) {
    const rows = SheetsService.getAllRows(CBQConfig.SHEETS.DOCUMENTS);
    return rows.filter(r => r.outcomeLabel === outcomeLabel);
  },
  
  /**
   * Get documents by reason code
   * @param {string} reasonCode - Reason code
   * @returns {Array} Documents
   */
  getByReasonCode(reasonCode) {
    const rows = SheetsService.getAllRows(CBQConfig.SHEETS.DOCUMENTS);
    return rows.filter(r => r.reasonCode === reasonCode);
  },
  
  /**
   * Update document status
   * @param {string} id - Document ID
   * @param {string} status - New status
   * @param {Object} extras - Additional fields to update
   * @returns {Object} Updated document
   */
  updateStatus(id, status, extras = {}) {
    const updates = {
      status: status,
      ...extras
    };
    
    if (status === CBQConfig.DOCUMENT_STATUS.COMPLETED) {
      updates.processedAt = Utils.now();
    }
    
    return SheetsService.updateRow(CBQConfig.SHEETS.DOCUMENTS, id, updates);
  },
  
  /**
   * Mark document as processing
   * @param {string} id - Document ID
   * @returns {Object} Updated document
   */
  markProcessing(id) {
    return this.updateStatus(id, CBQConfig.DOCUMENT_STATUS.PROCESSING);
  },
  
  /**
   * Mark document as completed
   * @param {string} id - Document ID
   * @param {number} wordCount - Number of words extracted
   * @returns {Object} Updated document
   */
  markCompleted(id, wordCount) {
    return this.updateStatus(id, CBQConfig.DOCUMENT_STATUS.COMPLETED, {
      wordCount: wordCount
    });
  },
  
  /**
   * Mark document as error
   * @param {string} id - Document ID
   * @param {string} errorMessage - Error details
   * @returns {Object} Updated document
   */
  markError(id, errorMessage) {
    return this.updateStatus(id, CBQConfig.DOCUMENT_STATUS.ERROR, {
      errorMessage: errorMessage
    });
  },
  
  /**
   * Get document counts by outcome
   * @returns {Object} Counts per outcome
   */
  getCountsByOutcome() {
    const rows = SheetsService.getAllRows(CBQConfig.SHEETS.DOCUMENTS);
    const completed = rows.filter(r => r.status === CBQConfig.DOCUMENT_STATUS.COMPLETED);
    
    const counts = {
      total: completed.length,
      byOutcome: {},
      byReasonCode: {}
    };
    
    completed.forEach(doc => {
      // By outcome
      const outcome = doc.outcomeLabel || 'UNKNOWN';
      counts.byOutcome[outcome] = (counts.byOutcome[outcome] || 0) + 1;
      
      // By reason code
      const rc = doc.reasonCode || 'UNKNOWN';
      counts.byReasonCode[rc] = (counts.byReasonCode[rc] || 0) + 1;
    });
    
    return counts;
  },
  
  /**
   * Get all documents with pagination
   * @param {Object} options - Filter options
   * @returns {Array} Documents
   */
  getAll(options = {}) {
    let rows = SheetsService.getAllRows(CBQConfig.SHEETS.DOCUMENTS);
    
    if (options.outcomeLabel) {
      rows = rows.filter(r => r.outcomeLabel === options.outcomeLabel);
    }
    if (options.reasonCode) {
      rows = rows.filter(r => r.reasonCode === options.reasonCode);
    }
    if (options.status) {
      rows = rows.filter(r => r.status === options.status);
    }
    
    // Sort by upload date descending
    rows.sort((a, b) => new Date(b.uploadedAt) - new Date(a.uploadedAt));
    
    const limit = options.limit || 100;
    const offset = options.offset || 0;
    
    return rows.slice(offset, offset + limit);
  }
};

Object.freeze(CBQDocumentsDAO);
/**
 * CBQ Words Data Access Object
 * 
 * CRUD operations for CBQ_Words sheet.
 * Stores individual word occurrences per document.
 * 
 * @module CBQWordsDAO
 */

const CBQWordsDAO = {
  /**
   * Bulk insert words for a document
   * @param {string} documentId - Document ID
   * @param {string} reasonCode - Reason code
   * @param {string} outcomeLabel - Outcome label
   * @param {Array} tokens - Tokenized words from Tokenizer
   * @returns {number} Number of words inserted
   */
  bulkInsert(documentId, reasonCode, outcomeLabel, tokens) {
    if (!tokens || tokens.length === 0) return 0;
    
    const rows = tokens.map(token => ({
      id: Utils.generateUUID(),
      word: token.word,
      documentId: documentId,
      reasonCode: reasonCode,
      outcomeLabel: outcomeLabel,
      termFrequency: token.frequency,
      positions: Utils.safeStringify(token.positions)
    }));
    
    // Batch insert in chunks
    const batchSize = CBQConfig.DEFAULTS.WORD_BATCH_SIZE;
    for (let i = 0; i < rows.length; i += batchSize) {
      const batch = rows.slice(i, i + batchSize);
      SheetsService.appendRows(CBQConfig.SHEETS.WORDS, batch);
    }
    
    return rows.length;
  },
  
  /**
   * Get words by document ID
   * @param {string} documentId - Document ID
   * @returns {Array} Word entries
   */
  getByDocument(documentId) {
    const allRows = SheetsService.getAllRows(CBQConfig.SHEETS.WORDS);
    return allRows.filter(r => r.documentId === documentId);
  },
  
  /**
   * Get all words for a reason code
   * @param {string} reasonCode - Reason code
   * @returns {Array} Word entries
   */
  getByReasonCode(reasonCode) {
    const allRows = SheetsService.getAllRows(CBQConfig.SHEETS.WORDS);
    return allRows.filter(r => r.reasonCode === reasonCode);
  },
  
  /**
   * Get all words for an outcome label
   * @param {string} outcomeLabel - Outcome label
   * @returns {Array} Word entries
   */
  getByOutcome(outcomeLabel) {
    const allRows = SheetsService.getAllRows(CBQConfig.SHEETS.WORDS);
    return allRows.filter(r => r.outcomeLabel === outcomeLabel);
  },
  
  /**
   * Get all words (with optional filters)
   * @param {Object} filters - Optional filters
   * @returns {Array} Word entries
   */
  getAll(filters = {}) {
    let rows = SheetsService.getAllRows(CBQConfig.SHEETS.WORDS);
    
    if (filters.reasonCode) {
      rows = rows.filter(r => r.reasonCode === filters.reasonCode);
    }
    if (filters.outcomeLabel) {
      rows = rows.filter(r => r.outcomeLabel === filters.outcomeLabel);
    }
    
    return rows;
  },
  
  /**
   * Delete words for a document
   * @param {string} documentId - Document ID
   * @returns {number} Number of rows deleted
   */
  deleteByDocument(documentId) {
    const allRows = SheetsService.getAllRows(CBQConfig.SHEETS.WORDS);
    const toDelete = allRows.filter(r => r.documentId === documentId);
    
    toDelete.forEach(row => {
      SheetsService.deleteRow(CBQConfig.SHEETS.WORDS, row.id);
    });
    
    return toDelete.length;
  },
  
  /**
   * Get unique words with their document frequencies
   * @returns {Map} Word to frequency map
   */
  getUniqueWordCounts() {
    const allRows = SheetsService.getAllRows(CBQConfig.SHEETS.WORDS);
    const wordCounts = new Map();
    
    allRows.forEach(row => {
      const word = row.word;
      if (!wordCounts.has(word)) {
        wordCounts.set(word, { word, count: 0, documents: new Set() });
      }
      const entry = wordCounts.get(word);
      entry.count += row.termFrequency;
      entry.documents.add(row.documentId);
    });
    
    return wordCounts;
  }
};

Object.freeze(CBQWordsDAO);
/**
 * CBQ Stats Data Access Object
 * 
 * CRUD operations for CBQ_WordStats sheet.
 * Stores aggregated word statistics and signals.
 * 
 * @module CBQStatsDAO
 */

const CBQStatsDAO = {
  /**
   * Update or insert word statistics
   * @param {string} word - The word
   * @param {string} reasonCode - Reason code or 'GLOBAL'
   * @param {Object} stats - Statistics to update
   * @returns {Object} Updated stats
   */
  upsert(word, reasonCode, stats) {
    const existing = this.getByWord(word, reasonCode);
    const now = Utils.now();
    
    if (existing) {
      // Update existing
      return SheetsService.updateRowByCompositeKey(
        CBQConfig.SHEETS.WORD_STATS,
        { word, reasonCode },
        {
          ...stats,
          updatedAt: now
        }
      );
    } else {
      // Insert new
      const row = {
        word,
        reasonCode,
        totalOccurrences: stats.totalOccurrences || 0,
        docFreqTotal: stats.docFreqTotal || 0,
        docFreqAutoWin: stats.docFreqAutoWin || 0,
        docFreqWin: stats.docFreqWin || 0,
        docFreqReject: stats.docFreqReject || 0,
        tfidfAutoWin: stats.tfidfAutoWin || 0,
        tfidfWin: stats.tfidfWin || 0,
        tfidfReject: stats.tfidfReject || 0,
        positiveWeight: stats.positiveWeight || 0,
        negativeWeight: stats.negativeWeight || 0,
        netSignalScore: stats.netSignalScore || 0,
        cooccurrences: stats.cooccurrences || '[]',
        updatedAt: now
      };
      
      SheetsService.appendRow(CBQConfig.SHEETS.WORD_STATS, row);
      return row;
    }
  },
  
  /**
   * Bulk upsert word statistics
   * @param {Array} statsArray - Array of stats objects
   * @param {string} reasonCode - Reason code
   * @returns {number} Number of stats updated
   */
  bulkUpsert(statsArray, reasonCode) {
    let count = 0;
    
    statsArray.forEach(stats => {
      this.upsert(stats.word, reasonCode, stats);
      count++;
    });
    
    return count;
  },
  
  /**
   * Get statistics for a word
   * @param {string} word - The word
   * @param {string} reasonCode - Reason code (default: 'GLOBAL')
   * @returns {Object|null} Word statistics or null
   */
  getByWord(word, reasonCode = 'GLOBAL') {
    const allRows = SheetsService.getAllRows(CBQConfig.SHEETS.WORD_STATS);
    return allRows.find(r => r.word === word && r.reasonCode === reasonCode) || null;
  },
  
  /**
   * Get all statistics for a reason code
   * @param {string} reasonCode - Reason code
   * @returns {Array} Word statistics
   */
  getByReasonCode(reasonCode) {
    const allRows = SheetsService.getAllRows(CBQConfig.SHEETS.WORD_STATS);
    return allRows.filter(r => r.reasonCode === reasonCode);
  },
  
  /**
   * Get top words by net signal score
   * @param {string} reasonCode - Reason code
   * @param {number} limit - Max words
   * @param {string} direction - 'positive', 'negative', or 'absolute'
   * @returns {Array} Top words
   */
  getTopWords(reasonCode, limit = 50, direction = 'positive') {
    const rows = this.getByReasonCode(reasonCode);
    
    switch (direction) {
      case 'positive':
        return rows
          .filter(r => r.netSignalScore > 0)
          .sort((a, b) => b.netSignalScore - a.netSignalScore)
          .slice(0, limit);
      case 'negative':
        return rows
          .filter(r => r.netSignalScore < 0)
          .sort((a, b) => a.netSignalScore - b.netSignalScore)
          .slice(0, limit);
      default: // absolute
        return rows
          .sort((a, b) => Math.abs(b.netSignalScore) - Math.abs(a.netSignalScore))
          .slice(0, limit);
    }
  },
  
  /**
   * Get all statistics as a Map for quick lookup
   * @param {string} reasonCode - Reason code
   * @returns {Map} Word to stats map
   */
  getAllAsMap(reasonCode) {
    const rows = this.getByReasonCode(reasonCode);
    const map = new Map();
    rows.forEach(r => map.set(r.word, r));
    return map;
  },
  
  /**
   * Delete all statistics for a reason code
   * @param {string} reasonCode - Reason code
   * @returns {number} Rows deleted
   */
  deleteByReasonCode(reasonCode) {
    const rows = this.getByReasonCode(reasonCode);
    rows.forEach(r => {
      SheetsService.deleteRowByCompositeKey(CBQConfig.SHEETS.WORD_STATS, { word: r.word, reasonCode });
    });
    return rows.length;
  },
  
  /**
   * Get summary for dashboard
   * @returns {Object} Summary stats
   */
  getSummary() {
    const globalStats = this.getByReasonCode('GLOBAL');
    
    return {
      totalWords: globalStats.length,
      positiveWords: globalStats.filter(r => r.netSignalScore > 0).length,
      negativeWords: globalStats.filter(r => r.netSignalScore < 0).length,
      neutralWords: globalStats.filter(r => r.netSignalScore === 0).length,
      topPositive: this.getTopWords('GLOBAL', 10, 'positive'),
      topNegative: this.getTopWords('GLOBAL', 10, 'negative')
    };
  }
};

Object.freeze(CBQStatsDAO);
/**
 * CBQ Filtered Words DAO
 * 
 * Data Access Object for managing filtered/excluded words.
 * Stores words that should be excluded from word clouds and analytics.
 * 
 * @module CBQFilteredWordsDAO
 */

const CBQFilteredWordsDAO = {
  
  /**
   * Get the filtered words sheet
   * @returns {Sheet} The CBQ_FilteredWords sheet
   */
  getSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CBQConfig.SHEETS.FILTERED_WORDS);
    
    if (!sheet) {
      // Create the sheet if it doesn't exist
      sheet = ss.insertSheet(CBQConfig.SHEETS.FILTERED_WORDS);
      sheet.appendRow(CBQConfig.FILTERED_WORDS_SCHEMA);
      sheet.getRange(1, 1, 1, CBQConfig.FILTERED_WORDS_SCHEMA.length)
        .setFontWeight('bold')
        .setBackground('#1a1a2e');
    }
    
    return sheet;
  },
  
  /**
   * Get all filtered words
   * @returns {Array} Array of filtered word objects
   */
  getAll() {
    const sheet = this.getSheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return [];
    
    const headers = data[0];
    const results = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      
      results.push({
        word: row[0],
        category: row[1] || 'CUSTOM',
        addedBy: row[2] || '',
        addedAt: row[3] || ''
      });
    }
    
    return results;
  },
  
  /**
   * Get a filtered word by word
   * @param {string} word - The word to find
   * @returns {Object|null} Filtered word object or null
   */
  getByWord(word) {
    const sheet = this.getSheet();
    const data = sheet.getDataRange().getValues();
    const normalizedWord = word.toLowerCase().trim();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toLowerCase() === normalizedWord) {
        return {
          word: data[i][0],
          category: data[i][1] || 'CUSTOM',
          addedBy: data[i][2] || '',
          addedAt: data[i][3] || '',
          row: i + 1
        };
      }
    }
    
    return null;
  },
  
  /**
   * Create a new filtered word entry
   * @param {Object} data - Word data {word, category}
   * @param {string} user - User email
   * @returns {Object} Created entry
   */
  create(data, user) {
    const sheet = this.getSheet();
    const normalizedWord = data.word.toLowerCase().trim();
    
    // Check for duplicates
    const existing = this.getByWord(normalizedWord);
    if (existing) {
      throw new Error(`Word "${normalizedWord}" already exists`);
    }
    
    const now = new Date();
    const entry = {
      word: normalizedWord,
      category: data.category || 'CUSTOM',
      addedBy: user,
      addedAt: now
    };
    
    sheet.appendRow([
      entry.word,
      entry.category,
      entry.addedBy,
      entry.addedAt
    ]);
    
    return entry;
  },
  
  /**
   * Delete a filtered word
   * @param {string} word - The word to delete
   * @returns {boolean} True if deleted, false if not found
   */
  delete(word) {
    const sheet = this.getSheet();
    const data = sheet.getDataRange().getValues();
    const normalizedWord = word.toLowerCase().trim();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toLowerCase() === normalizedWord) {
        sheet.deleteRow(i + 1);
        return true;
      }
    }
    
    return false;
  },
  
  /**
   * Get all words as a Set (for fast lookup)
   * @returns {Set} Set of filtered words
   */
  getAllAsSet() {
    const words = this.getAll();
    return new Set(words.map(w => w.word.toLowerCase()));
  },
  
  /**
   * Bulk create filtered words
   * @param {Array} words - Array of {word, category} objects
   * @param {string} user - User email
   * @returns {Object} Result with added count
   */
  bulkCreate(words, user) {
    const sheet = this.getSheet();
    const existingSet = this.getAllAsSet();
    const now = new Date();
    let added = 0;
    
    const newRows = [];
    
    for (const item of words) {
      const normalizedWord = item.word.toLowerCase().trim();
      
      if (!existingSet.has(normalizedWord) && normalizedWord.length >= 2) {
        newRows.push([
          normalizedWord,
          item.category || 'CUSTOM',
          user,
          now
        ]);
        existingSet.add(normalizedWord);
        added++;
      }
    }
    
    if (newRows.length > 0) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, newRows.length, 4).setValues(newRows);
    }
    
    return { added };
  },
  
  /**
   * Get words by category
   * @param {string} category - Category to filter by
   * @returns {Array} Filtered words in category
   */
  getByCategory(category) {
    return this.getAll().filter(w => w.category === category);
  },
  
  /**
   * Clear all filtered words (admin only)
   * @returns {number} Number of words cleared
   */
  clearAll() {
    const sheet = this.getSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow > 1) {
      const count = lastRow - 1;
      sheet.deleteRows(2, count);
      return count;
    }
    
    return 0;
  }
};

Object.freeze(CBQFilteredWordsDAO);
/**
 * AI Processor - OpenAI GPT Integration for Chargeback Analytics
 * 
 * Uses GPT-4o-mini for cost-effective document classification.
 * Extracts reason codes, card types, and outcome classification.
 * 
 * API Key: Set in Script Properties as 'OPENAI_API_KEY'
 * 
 * @module AIProcessor
 */

const AIProcessor = {
  
  // =====================================================
  // Configuration
  // =====================================================
  
  CONFIG: {
    MODEL: 'gpt-4o-mini',
    MAX_TOKENS: 500,
    TEMPERATURE: 0.1,  // Low temperature for consistent classification
    API_ENDPOINT: 'https://api.openai.com/v1/chat/completions',
    
    // Tight system prompt for classification
    SYSTEM_PROMPT: `You are a chargeback analyst assistant. Analyze rebuttal documents and extract structured data.

RULES:
1. Extract the REASON CODE from the document (e.g., 13.7, 10.4, 4540, F29, etc.)
2. Identify CARD TYPE from content or transaction details:
   - VI = Visa
   - MC = MasterCard  
   - DI = Discover
   - AX = American Express
3. Classify the case outcome based on evidence quality:
   - AUTO_WIN_72H: Strong evidence, clear policy violation, likely auto-win
   - WIN_30_90D: Good evidence but needs review, eventual win expected
   - AUTO_REJECT_72H: Weak/missing evidence, high chance of loss
4. Extract up to 20 key terms that characterize this document

Respond ONLY in valid JSON format:
{
  "reasonCode": "13.7",
  "cardType": "VI", 
  "outcome": "AUTO_WIN_72H",
  "confidence": 0.95,
  "keyTerms": ["cancellation", "policy", "agreed", "terms"]
}

If data cannot be determined, use null for that field.`
  },
  
  // =====================================================
  // Core Classification
  // =====================================================
  
  /**
   * Get API key from Script Properties
   * @returns {string} API key or empty string
   */
  getApiKey() {
    const key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    if (!key) {
      throw new Error('OPENAI_API_KEY not found in Script Properties. Set it via: File → Project Properties → Script Properties');
    }
    return key;
  },
  
  /**
   * Classify a document using AI
   * @param {string} text - Document text content
   * @returns {Object} Classification result
   */
  classifyDocument(text) {
    try {
      if (!text || text.trim().length < 50) {
        return {
          success: false,
          error: 'Document text too short for classification',
          data: null
        };
      }
      
      // Truncate to avoid token limits (approx 4 chars per token)
      const maxChars = 12000;
      const truncatedText = text.length > maxChars 
        ? text.substring(0, maxChars) + '\n[...truncated...]' 
        : text;
      
      const userPrompt = `Analyze this chargeback rebuttal document:\n\n${truncatedText}`;
      
      const response = this.callOpenAI(userPrompt);
      
      if (!response.success) {
        return response;
      }
      
      // Parse the JSON response
      try {
        const data = JSON.parse(response.content);
        return {
          success: true,
          data: {
            reasonCode: data.reasonCode || 'UNKNOWN',
            cardType: data.cardType || 'UNKNOWN',
            outcome: data.outcome || 'WIN_30_90D',
            confidence: data.confidence || 0.5,
            keyTerms: data.keyTerms || []
          },
          error: null
        };
      } catch (parseError) {
        // Try to extract JSON from response
        const jsonMatch = response.content.match(/\{[\s\S]*\}/);
        if (jsonMatch) {
          try {
            const data = JSON.parse(jsonMatch[0]);
            return {
              success: true,
              data: {
                reasonCode: data.reasonCode || 'UNKNOWN',
                cardType: data.cardType || 'UNKNOWN',
                outcome: data.outcome || 'WIN_30_90D',
                confidence: data.confidence || 0.5,
                keyTerms: data.keyTerms || []
              },
              error: null
            };
          } catch (e) {
            // Fall through to error
          }
        }
        
        return {
          success: false,
          error: `Failed to parse AI response: ${parseError.message}`,
          data: null
        };
      }
      
    } catch (e) {
      return {
        success: false,
        error: e.message,
        data: null
      };
    }
  },
  
  /**
   * Make API call to OpenAI
   * @param {string} userPrompt - User message
   * @returns {Object} API response
   */
  callOpenAI(userPrompt) {
    try {
      const apiKey = this.getApiKey();
      
      const payload = {
        model: this.CONFIG.MODEL,
        messages: [
          { role: 'system', content: this.CONFIG.SYSTEM_PROMPT },
          { role: 'user', content: userPrompt }
        ],
        max_tokens: this.CONFIG.MAX_TOKENS,
        temperature: this.CONFIG.TEMPERATURE
      };
      
      const options = {
        method: 'post',
        contentType: 'application/json',
        headers: {
          'Authorization': `Bearer ${apiKey}`
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(this.CONFIG.API_ENDPOINT, options);
      const responseCode = response.getResponseCode();
      const responseBody = JSON.parse(response.getContentText());
      
      if (responseCode !== 200) {
        const errorMsg = responseBody.error?.message || 'Unknown API error';
        return {
          success: false,
          error: `OpenAI API Error (${responseCode}): ${errorMsg}`,
          content: null
        };
      }
      
      const content = responseBody.choices?.[0]?.message?.content;
      if (!content) {
        return {
          success: false,
          error: 'No content in API response',
          content: null
        };
      }
      
      return {
        success: true,
        content: content,
        error: null,
        usage: responseBody.usage
      };
      
    } catch (e) {
      return {
        success: false,
        error: `API call failed: ${e.message}`,
        content: null
      };
    }
  },
  
  // =====================================================
  // Keyword Extraction
  // =====================================================
  
  /**
   * Extract key terms from document using AI
   * @param {string} text - Document text
   * @param {number} limit - Max terms to extract
   * @returns {Object} {success, terms[], error}
   */
  extractKeyTerms(text, limit = 30) {
    try {
      const maxChars = 8000;
      const truncatedText = text.length > maxChars 
        ? text.substring(0, maxChars) 
        : text;
      
      const prompt = `Extract the ${limit} most important keywords and phrases from this chargeback document. Focus on:
- Legal/policy terms
- Action verbs
- Key concepts related to dispute outcome

Return ONLY a JSON array of strings, no other text:
["term1", "term2", "term3"]

Document:
${truncatedText}`;
      
      const response = this.callOpenAI(prompt);
      
      if (!response.success) {
        return { success: false, terms: [], error: response.error };
      }
      
      try {
        const terms = JSON.parse(response.content);
        if (Array.isArray(terms)) {
          return { success: true, terms: terms.slice(0, limit), error: null };
        }
      } catch (e) {
        // Try to extract array
        const arrayMatch = response.content.match(/\[[\s\S]*\]/);
        if (arrayMatch) {
          const terms = JSON.parse(arrayMatch[0]);
          return { success: true, terms: terms.slice(0, limit), error: null };
        }
      }
      
      return { success: false, terms: [], error: 'Could not parse terms from response' };
      
    } catch (e) {
      return { success: false, terms: [], error: e.message };
    }
  },
  
  // =====================================================
  // Batch Processing
  // =====================================================
  
  /**
   * Process multiple documents in batch
   * @param {Array} documents - Array of {id, text} objects
   * @param {string} batchType - 'AUTO_WIN_72H', 'WIN_30_90D', or 'AUTO_REJECT_72H'
   * @returns {Object} Batch results
   */
  processBatch(documents, batchType) {
    const results = {
      batchType: batchType,
      processed: 0,
      failed: 0,
      results: []
    };
    
    for (const doc of documents) {
      try {
        const classification = this.classifyDocument(doc.text);
        
        if (classification.success) {
          // Override outcome with user-specified batch type
          classification.data.outcome = batchType;
          results.results.push({
            id: doc.id,
            success: true,
            data: classification.data
          });
          results.processed++;
        } else {
          results.results.push({
            id: doc.id,
            success: false,
            error: classification.error
          });
          results.failed++;
        }
        
        // Brief pause to avoid rate limiting
        Utilities.sleep(200);
        
      } catch (e) {
        results.results.push({
          id: doc.id,
          success: false,
          error: e.message
        });
        results.failed++;
      }
    }
    
    return results;
  },
  
  // =====================================================
  // Utilities
  // =====================================================
  
  /**
   * Test API connection
   * @returns {Object} Test result
   */
  testConnection() {
    try {
      const response = this.callOpenAI('Say "Connection successful" in JSON: {"status": "ok"}');
      return {
        success: response.success,
        message: response.success ? 'API connection successful' : response.error,
        model: this.CONFIG.MODEL
      };
    } catch (e) {
      return {
        success: false,
        message: e.message,
        model: null
      };
    }
  },
  
  /**
   * Estimate API cost for batch
   * @param {number} documentCount - Number of documents
   * @param {number} avgCharsPerDoc - Average characters per document
   * @returns {Object} Cost estimate
   */
  estimateCost(documentCount, avgCharsPerDoc = 5000) {
    // GPT-4o-mini pricing: $0.15/1M input, $0.60/1M output
    const inputTokensPerDoc = Math.ceil(avgCharsPerDoc / 4);
    const outputTokens = 200; // Fixed output size
    const systemPromptTokens = 400;
    
    const totalInputTokens = documentCount * (inputTokensPerDoc + systemPromptTokens);
    const totalOutputTokens = documentCount * outputTokens;
    
    const inputCost = (totalInputTokens / 1000000) * 0.15;
    const outputCost = (totalOutputTokens / 1000000) * 0.60;
    
    return {
      documents: documentCount,
      estimatedInputTokens: totalInputTokens,
      estimatedOutputTokens: totalOutputTokens,
      estimatedCostUSD: (inputCost + outputCost).toFixed(4)
    };
  }
};

// Freeze to prevent modifications
Object.freeze(AIProcessor);
Object.freeze(AIProcessor.CONFIG);
/**
 * CBQ Text Extractor
 * 
 * Extracts text from PDFs and handles OCR fallback.
 * Uses Google Drive and Document APIs.
 * 
 * @module TextExtractor
 */

const TextExtractor = {
  /**
   * Extract text from a Google Drive PDF file
   * @param {string} fileId - Google Drive file ID
   * @returns {Object} {text, method, success, error}
   */
  extractFromPDF(fileId) {
    try {
      const file = DriveApp.getFileById(fileId);
      const blob = file.getBlob();
      const mimeType = blob.getContentType();
      
      if (mimeType !== 'application/pdf') {
        return {
          text: '',
          method: 'NONE',
          success: false,
          error: `Not a PDF file: ${mimeType}`
        };
      }
      
      // Try native text extraction first (for text-based PDFs)
      let text = this.extractNativeText(blob);
      
      if (text && text.trim().length > 50) {
        return {
          text: text,
          method: 'NATIVE',
          success: true,
          error: null
        };
      }
      
      // Fallback to OCR via Google Drive conversion
      text = this.extractViaOCR(file);
      
      if (text && text.trim().length > 0) {
        return {
          text: text,
          method: 'OCR',
          success: true,
          error: null
        };
      }
      
      return {
        text: '',
        method: 'FAILED',
        success: false,
        error: 'Could not extract text from PDF'
      };
      
    } catch (e) {
      Utils.log('ERROR', 'TextExtractor', `PDF extraction failed: ${e.message}`);
      return {
        text: '',
        method: 'ERROR',
        success: false,
        error: e.message
      };
    }
  },
  
  /**
   * Try to extract native text from PDF blob
   * @param {Blob} blob - PDF blob
   * @returns {string} Extracted text
   */
  extractNativeText(blob) {
    try {
      // Google Apps Script doesn't have native PDF text extraction
      // We'll use the OCR method which works for both
      return '';
    } catch (e) {
      return '';
    }
  },
  
  /**
   * Extract text via Google Drive OCR (convert to Google Doc)
   * @param {File} file - Drive file
   * @returns {string} Extracted text
   */
  extractViaOCR(file) {
    try {
      // Create a temporary Google Doc from the PDF (triggers OCR)
      const tempDoc = Drive.Files.copy(
        { title: 'TEMP_OCR_' + Date.now(), mimeType: MimeType.GOOGLE_DOCS },
        file.getId(),
        { ocr: true }
      );
      
      // Get the text content
      const doc = DocumentApp.openById(tempDoc.id);
      const text = doc.getBody().getText();
      
      // Clean up temp file
      DriveApp.getFileById(tempDoc.id).setTrashed(true);
      
      return text;
      
    } catch (e) {
      Utils.log('WARN', 'TextExtractor', `OCR extraction failed: ${e.message}`);
      
      // Alternative: Use Drive API v3 with OCR
      try {
        return this.extractViaAdvancedOCR(file);
      } catch (e2) {
        return '';
      }
    }
  },
  
  /**
   * Alternative OCR using Drive API
   * @param {File} file - Drive file
   * @returns {string} Extracted text
   */
  extractViaAdvancedOCR(file) {
    try {
      const blob = file.getBlob();
      
      // Create resource for Drive API
      const resource = {
        title: 'TEMP_OCR_' + Date.now(),
        mimeType: MimeType.GOOGLE_DOCS
      };
      
      // Insert with OCR
      const created = Drive.Files.insert(resource, blob, {
        ocr: true,
        convert: true
      });
      
      // Get text
      const doc = DocumentApp.openById(created.id);
      const text = doc.getBody().getText();
      
      // Cleanup
      DriveApp.getFileById(created.id).setTrashed(true);
      
      return text;
      
    } catch (e) {
      Utils.log('ERROR', 'TextExtractor', `Advanced OCR failed: ${e.message}`);
      return '';
    }
  },
  
  /**
   * Process raw text input (for paste mode)
   * @param {string} text - Raw text
   * @returns {Object} {text, success, error}
   */
  processRawText(text) {
    if (!text || typeof text !== 'string') {
      return {
        text: '',
        success: false,
        error: 'No text provided'
      };
    }
    
    // Basic cleaning
    const cleaned = text
      .replace(/\r\n/g, '\n')
      .replace(/\r/g, '\n')
      .replace(/\n{3,}/g, '\n\n')
      .trim();
    
    if (cleaned.length < 10) {
      return {
        text: '',
        success: false,
        error: 'Text too short'
      };
    }
    
    return {
      text: cleaned,
      success: true,
      error: null
    };
  },
  
  /**
   * Get files from a Drive folder
   * @param {string} folderId - Drive folder ID
   * @param {string} mimeType - Optional MIME type filter
   * @returns {Array} Array of file metadata
   */
  getFilesFromFolder(folderId, mimeType = 'application/pdf') {
    try {
      const folder = DriveApp.getFolderById(folderId);
      const files = folder.getFilesByType(mimeType);
      const result = [];
      
      while (files.hasNext()) {
        const file = files.next();
        result.push({
          id: file.getId(),
          name: file.getName(),
          size: file.getSize(),
          created: file.getDateCreated(),
          mimeType: file.getMimeType()
        });
      }
      
      return result;
      
    } catch (e) {
      Utils.log('ERROR', 'TextExtractor', `Folder access failed: ${e.message}`);
      return [];
    }
  },
  
  /**
   * Estimate processing time for a batch
   * @param {number} fileCount - Number of files
   * @returns {Object} Estimated seconds and quota usage
   */
  estimateProcessingTime(fileCount) {
    const secondsPerFile = 5; // Average with OCR
    const apiCallsPerFile = 3; // Create, read, delete
    
    return {
      estimatedSeconds: fileCount * secondsPerFile,
      estimatedMinutes: Math.ceil((fileCount * secondsPerFile) / 60),
      apiCalls: fileCount * apiCallsPerFile,
      withinQuota: (fileCount * apiCallsPerFile) < 10000 // Daily limit
    };
  }
};

Object.freeze(TextExtractor);
/**
 * CBQ Tokenizer
 * 
 * Text tokenization and normalization for lexical analysis.
 * Handles word extraction, stop word filtering, and optional stemming.
 * 
 * @module Tokenizer
 */

const Tokenizer = {
  /**
   * Tokenize text into normalized words
   * @param {string} text - Raw text to tokenize
   * @param {Object} options - Tokenization options
   * @returns {Array} Array of {word, position, frequency} objects
   */
  tokenize(text, options = {}) {
    if (!text || typeof text !== 'string') {
      return [];
    }
    
    const minLength = options.minLength || CBQConfig.DEFAULTS.MIN_WORD_LENGTH;
    const maxLength = options.maxLength || CBQConfig.DEFAULTS.MAX_WORD_LENGTH;
    const stopWords = new Set(options.stopWords || CBQConfig.STOP_WORDS);
    const useStemming = options.useStemming !== false;
    
    // Normalize: lowercase, remove special chars
    let normalized = text.toLowerCase();
    normalized = normalized.replace(/[^a-z0-9\s]/g, ' ');
    normalized = normalized.replace(/\s+/g, ' ').trim();
    
    // Split into words
    const rawWords = normalized.split(' ');
    
    // Track word frequencies and first positions
    const wordMap = new Map();
    
    rawWords.forEach((word, position) => {
      // Skip empty, too short, too long, numbers only
      if (!word || word.length < minLength || word.length > maxLength) {
        return;
      }
      if (/^\d+$/.test(word)) {
        return;
      }
      
      // Skip stop words
      if (stopWords.has(word)) {
        return;
      }
      
      // Optional stemming (simple Porter-like suffix removal)
      const stemmed = useStemming ? this.stem(word) : word;
      
      if (wordMap.has(stemmed)) {
        const entry = wordMap.get(stemmed);
        entry.frequency++;
        entry.positions.push(position);
      } else {
        wordMap.set(stemmed, {
          word: stemmed,
          frequency: 1,
          positions: [position]
        });
      }
    });
    
    // Convert to array
    return Array.from(wordMap.values());
  },
  
  /**
   * Simple Porter-like stemming
   * Removes common suffixes for English words
   * @param {string} word - Word to stem
   * @returns {string} Stemmed word
   */
  stem(word) {
    if (word.length < 4) return word;
    
    // Common suffix replacements
    const suffixes = [
      { pattern: /ies$/, replacement: 'y' },
      { pattern: /ied$/, replacement: 'y' },
      { pattern: /ing$/, replacement: '' },
      { pattern: /ed$/, replacement: '' },
      { pattern: /es$/, replacement: '' },
      { pattern: /s$/, replacement: '' },
      { pattern: /tion$/, replacement: 't' },
      { pattern: /ment$/, replacement: '' },
      { pattern: /ness$/, replacement: '' },
      { pattern: /able$/, replacement: '' },
      { pattern: /ible$/, replacement: '' },
      { pattern: /ful$/, replacement: '' },
      { pattern: /less$/, replacement: '' },
      { pattern: /ly$/, replacement: '' }
    ];
    
    for (const suffix of suffixes) {
      if (suffix.pattern.test(word) && word.replace(suffix.pattern, suffix.replacement).length >= 3) {
        return word.replace(suffix.pattern, suffix.replacement);
      }
    }
    
    return word;
  },
  
  /**
   * Get word count statistics
   * @param {string} text - Text to analyze
   * @returns {Object} Stats
   */
  getStats(text) {
    const tokens = this.tokenize(text);
    const totalWords = tokens.reduce((sum, t) => sum + t.frequency, 0);
    const uniqueWords = tokens.length;
    
    return {
      totalWords,
      uniqueWords,
      avgFrequency: uniqueWords > 0 ? totalWords / uniqueWords : 0
    };
  },
  
  /**
   * Extract n-grams (word pairs, triplets)
   * @param {string} text - Text to analyze
   * @param {number} n - N-gram size (2 = bigrams, 3 = trigrams)
   * @returns {Array} N-grams with frequencies
   */
  extractNgrams(text, n = 2) {
    if (!text || n < 2) return [];
    
    const words = text.toLowerCase()
      .replace(/[^a-z0-9\s]/g, ' ')
      .split(/\s+/)
      .filter(w => w.length >= 3);
    
    const ngramMap = new Map();
    
    for (let i = 0; i <= words.length - n; i++) {
      const ngram = words.slice(i, i + n).join(' ');
      ngramMap.set(ngram, (ngramMap.get(ngram) || 0) + 1);
    }
    
    return Array.from(ngramMap.entries())
      .map(([ngram, frequency]) => ({ ngram, frequency }))
      .sort((a, b) => b.frequency - a.frequency);
  },
  
  /**
   * Calculate co-occurrences (words appearing together in same document)
   * @param {Array} tokens - Tokenized words from tokenize()
   * @param {number} windowSize - Words apart to consider co-occurring
   * @returns {Map} Word pairs to frequency
   */
  calculateCooccurrences(tokens, windowSize = 5) {
    const cooccurrences = new Map();
    const words = tokens.map(t => t.word);
    
    for (let i = 0; i < words.length; i++) {
      for (let j = i + 1; j < Math.min(i + windowSize, words.length); j++) {
        const pair = [words[i], words[j]].sort().join('|');
        cooccurrences.set(pair, (cooccurrences.get(pair) || 0) + 1);
      }
    }
    
    return cooccurrences;
  },
  
  /**
   * Custom stop words management - get from config sheet
   * @returns {Array} Stop words array
   */
  getStopWords() {
    try {
      const configValue = CBQConfigDAO.get('STOP_WORDS');
      if (configValue) {
        return configValue.split(',').map(w => w.trim().toLowerCase());
      }
    } catch (e) {
      // Config sheet may not exist yet
    }
    return CBQConfig.STOP_WORDS;
  },
  
  /**
   * Validate and normalize outcome label
   * @param {string} label - Input label
   * @returns {string|null} Normalized label or null if invalid
   */
  normalizeOutcomeLabel(label) {
    if (!label) return null;
    const upper = label.toUpperCase().replace(/\s+/g, '_');
    return CBQConfig.OUTCOME_LABELS[upper] ? upper : null;
  }
};

Object.freeze(Tokenizer);
/**
 * CBQ Metrics Engine
 * 
 * TF-IDF calculation, weighted scoring, and signal analysis.
 * All calculations are deterministic and auditable.
 * 
 * @module MetricsEngine
 */

const MetricsEngine = {
  /**
   * Calculate TF-IDF for a word
   * TF-IDF = Term Frequency × log(Total Docs / Doc Frequency)
   * 
   * @param {number} termFrequency - Word count in document
   * @param {number} docFrequency - Docs containing word
   * @param {number} totalDocs - Total documents in corpus
   * @returns {number} TF-IDF score
   */
  calculateTFIDF(termFrequency, docFrequency, totalDocs) {
    if (docFrequency === 0 || totalDocs === 0) return 0;
    
    const tf = termFrequency;
    const idf = Math.log(totalDocs / docFrequency);
    
    return tf * idf;
  },
  
  /**
   * Calculate weighted positive score for a word
   * @param {number} tfidfAutoWin - TF-IDF in AUTO_WIN_72H corpus
   * @param {number} tfidfWin - TF-IDF in WIN_30_90D corpus
   * @returns {number} Positive weight score
   */
  calculatePositiveWeight(tfidfAutoWin, tfidfWin) {
    const weightAutoWin = CBQConfig.DEFAULTS.WEIGHT_AUTO_WIN_72H;
    const weightWin = CBQConfig.DEFAULTS.WEIGHT_WIN_30_90D;
    
    return (tfidfAutoWin * weightAutoWin) + (tfidfWin * weightWin);
  },
  
  /**
   * Calculate weighted negative score for a word
   * @param {number} tfidfReject - TF-IDF in AUTO_REJECT_72H corpus
   * @returns {number} Negative weight score
   */
  calculateNegativeWeight(tfidfReject) {
    const weightReject = CBQConfig.DEFAULTS.WEIGHT_AUTO_REJECT_72H;
    return tfidfReject * weightReject;
  },
  
  /**
   * Calculate net signal score
   * Positive values indicate positive outcome correlation
   * Negative values indicate rejection correlation
   * 
   * @param {number} positiveWeight - Positive score
   * @param {number} negativeWeight - Negative score
   * @returns {number} Net signal score
   */
  calculateNetSignal(positiveWeight, negativeWeight) {
    return positiveWeight - negativeWeight;
  },
  
  /**
   * Calculate all metrics for a word
   * @param {Object} wordData - Word frequency data across corpora
   * @returns {Object} Complete metrics
   */
  calculateWordMetrics(wordData) {
    const { 
      word,
      totalOccurrences,
      docFreqAutoWin = 0,
      docFreqWin = 0,
      docFreqReject = 0,
      totalDocsAutoWin = 1,
      totalDocsWin = 1,
      totalDocsReject = 1,
      avgTfAutoWin = 0,
      avgTfWin = 0,
      avgTfReject = 0
    } = wordData;
    
    // Calculate TF-IDF for each corpus
    const tfidfAutoWin = this.calculateTFIDF(avgTfAutoWin, docFreqAutoWin, totalDocsAutoWin);
    const tfidfWin = this.calculateTFIDF(avgTfWin, docFreqWin, totalDocsWin);
    const tfidfReject = this.calculateTFIDF(avgTfReject, docFreqReject, totalDocsReject);
    
    // Calculate weighted scores
    const positiveWeight = this.calculatePositiveWeight(tfidfAutoWin, tfidfWin);
    const negativeWeight = this.calculateNegativeWeight(tfidfReject);
    const netSignalScore = this.calculateNetSignal(positiveWeight, negativeWeight);
    
    return {
      word,
      totalOccurrences,
      docFreqAutoWin,
      docFreqWin,
      docFreqReject,
      tfidfAutoWin: Math.round(tfidfAutoWin * 1000) / 1000,
      tfidfWin: Math.round(tfidfWin * 1000) / 1000,
      tfidfReject: Math.round(tfidfReject * 1000) / 1000,
      positiveWeight: Math.round(positiveWeight * 1000) / 1000,
      negativeWeight: Math.round(negativeWeight * 1000) / 1000,
      netSignalScore: Math.round(netSignalScore * 1000) / 1000
    };
  },
  
  /**
   * Aggregate word statistics from raw word data
   * @param {Array} wordEntries - Array of word entries from CBQ_Words
   * @param {Object} docCounts - Document counts per outcome
   * @returns {Map} Word to aggregated stats
   */
  aggregateWordStats(wordEntries, docCounts) {
    const wordStats = new Map();
    
    // Group by word
    wordEntries.forEach(entry => {
      const word = entry.word;
      
      if (!wordStats.has(word)) {
        wordStats.set(word, {
          word,
          totalOccurrences: 0,
          docIdsAutoWin: new Set(),
          docIdsWin: new Set(),
          docIdsReject: new Set(),
          totalTfAutoWin: 0,
          totalTfWin: 0,
          totalTfReject: 0
        });
      }
      
      const stats = wordStats.get(word);
      stats.totalOccurrences += entry.termFrequency;
      
      // Track document IDs per outcome
      const outcome = entry.outcomeLabel;
      if (outcome === 'AUTO_WIN_72H') {
        stats.docIdsAutoWin.add(entry.documentId);
        stats.totalTfAutoWin += entry.termFrequency;
      } else if (outcome === 'WIN_30_90D') {
        stats.docIdsWin.add(entry.documentId);
        stats.totalTfWin += entry.termFrequency;
      } else if (outcome === 'AUTO_REJECT_72H') {
        stats.docIdsReject.add(entry.documentId);
        stats.totalTfReject += entry.termFrequency;
      }
    });
    
    // Calculate final metrics
    const result = new Map();
    
    wordStats.forEach((stats, word) => {
      const docFreqAutoWin = stats.docIdsAutoWin.size;
      const docFreqWin = stats.docIdsWin.size;
      const docFreqReject = stats.docIdsReject.size;
      
      // Average TF per document
      const avgTfAutoWin = docFreqAutoWin > 0 ? stats.totalTfAutoWin / docFreqAutoWin : 0;
      const avgTfWin = docFreqWin > 0 ? stats.totalTfWin / docFreqWin : 0;
      const avgTfReject = docFreqReject > 0 ? stats.totalTfReject / docFreqReject : 0;
      
      const metrics = this.calculateWordMetrics({
        word,
        totalOccurrences: stats.totalOccurrences,
        docFreqAutoWin,
        docFreqWin,
        docFreqReject,
        totalDocsAutoWin: docCounts.AUTO_WIN_72H || 1,
        totalDocsWin: docCounts.WIN_30_90D || 1,
        totalDocsReject: docCounts.AUTO_REJECT_72H || 1,
        avgTfAutoWin,
        avgTfWin,
        avgTfReject
      });
      
      result.set(word, metrics);
    });
    
    return result;
  },
  
  /**
   * Get top positive signal words
   * @param {Map} wordStats - Word statistics map
   * @param {number} limit - Max words to return
   * @returns {Array} Top positive words
   */
  getTopPositiveWords(wordStats, limit = 50) {
    return Array.from(wordStats.values())
      .filter(w => w.netSignalScore > 0)
      .sort((a, b) => b.netSignalScore - a.netSignalScore)
      .slice(0, limit);
  },
  
  /**
   * Get top negative signal words
   * @param {Map} wordStats - Word statistics map
   * @param {number} limit - Max words to return
   * @returns {Array} Top negative words
   */
  getTopNegativeWords(wordStats, limit = 50) {
    return Array.from(wordStats.values())
      .filter(w => w.netSignalScore < 0)
      .sort((a, b) => a.netSignalScore - b.netSignalScore)
      .slice(0, limit);
  },
  
  /**
   * Get words exclusive to positive corpora (absent in reject)
   * @param {Map} wordStats - Word statistics map
   * @param {number} limit - Max words
   * @returns {Array} Positive-exclusive words
   */
  getPositiveExclusiveWords(wordStats, limit = 50) {
    return Array.from(wordStats.values())
      .filter(w => w.docFreqReject === 0 && (w.docFreqAutoWin > 0 || w.docFreqWin > 0))
      .sort((a, b) => b.positiveWeight - a.positiveWeight)
      .slice(0, limit);
  },
  
  /**
   * Get words exclusive to negative corpus (absent in positive)
   * @param {Map} wordStats - Word statistics map
   * @param {number} limit - Max words
   * @returns {Array} Negative-exclusive words
   */
  getNegativeExclusiveWords(wordStats, limit = 50) {
    return Array.from(wordStats.values())
      .filter(w => w.docFreqAutoWin === 0 && w.docFreqWin === 0 && w.docFreqReject > 0)
      .sort((a, b) => b.negativeWeight - a.negativeWeight)
      .slice(0, limit);
  },
  
  /**
   * Get contentious words (high in both positive and negative)
   * @param {Map} wordStats - Word statistics map
   * @param {number} limit - Max words
   * @returns {Array} Contentious words
   */
  getContentiousWords(wordStats, limit = 50) {
    return Array.from(wordStats.values())
      .filter(w => w.positiveWeight > 0.1 && w.negativeWeight > 0.1)
      .sort((a, b) => Math.min(b.positiveWeight, b.negativeWeight) - Math.min(a.positiveWeight, a.negativeWeight))
      .slice(0, limit);
  },
  
  /**
   * Prepare data for word cloud visualization
   * @param {Map} wordStats - Word statistics map
   * @param {string} mode - 'signal', 'positive', 'negative', 'frequency'
   * @param {number} limit - Max words
   * @returns {Array} Word cloud data [{word, size, color}]
   */
  prepareWordCloudData(wordStats, mode = 'signal', limit = 100) {
    const words = Array.from(wordStats.values());
    let sorted;
    
    switch (mode) {
      case 'positive':
        sorted = words.filter(w => w.netSignalScore > 0)
          .sort((a, b) => b.positiveWeight - a.positiveWeight);
        break;
      case 'negative':
        sorted = words.filter(w => w.netSignalScore < 0)
          .sort((a, b) => b.negativeWeight - a.negativeWeight);
        break;
      case 'frequency':
        sorted = words.sort((a, b) => b.totalOccurrences - a.totalOccurrences);
        break;
      default: // signal
        sorted = words.sort((a, b) => Math.abs(b.netSignalScore) - Math.abs(a.netSignalScore));
    }
    
    const topWords = sorted.slice(0, limit);
    const maxValue = Math.max(...topWords.map(w => Math.abs(w.netSignalScore) || 1));
    
    return topWords.map(w => ({
      word: w.word,
      size: Math.max(10, Math.round((Math.abs(w.netSignalScore) / maxValue) * 100)),
      value: w.netSignalScore,
      color: w.netSignalScore > 0 ? '#22c55e' : w.netSignalScore < 0 ? '#ef4444' : '#808080'
    }));
  },
  
  /**
   * Prepare data for heat map visualization
   * @param {Map} wordStats - Word statistics map
   * @param {number} limit - Max words
   * @returns {Object} Heat map data
   */
  prepareHeatMapData(wordStats, limit = 50) {
    const topWords = Array.from(wordStats.values())
      .sort((a, b) => b.totalOccurrences - a.totalOccurrences)
      .slice(0, limit);
    
    return {
      words: topWords.map(w => w.word),
      outcomes: ['AUTO_WIN_72H', 'WIN_30_90D', 'AUTO_REJECT_72H'],
      data: topWords.map(w => [
        w.tfidfAutoWin,
        w.tfidfWin,
        w.tfidfReject
      ])
    };
  },
  
  /**
   * Prepare data for bubble chart visualization
   * @param {Map} wordStats - Word statistics map
   * @param {number} limit - Max bubbles
   * @returns {Array} Bubble data [{word, x, y, size}]
   */
  prepareBubbleChartData(wordStats, limit = 100) {
    return Array.from(wordStats.values())
      .filter(w => w.totalOccurrences >= 3) // Min frequency
      .sort((a, b) => b.totalOccurrences - a.totalOccurrences)
      .slice(0, limit)
      .map(w => ({
        word: w.word,
        x: w.positiveWeight,
        y: w.negativeWeight,
        size: Math.log(w.totalOccurrences + 1) * 10,
        netSignal: w.netSignalScore
      }));
  }
};

Object.freeze(MetricsEngine);
/**
 * CBQ Batch Processor
 * 
 * Quota-aware batch processing for document ingestion.
 * Supports incremental processing and trigger-based execution.
 * 
 * @module BatchProcessor
 */

const BatchProcessor = {
  /**
   * Process a single document
   * @param {string} documentId - Document ID
   * @returns {Object} Processing result
   */
  processDocument(documentId) {
    const startTime = Date.now();
    
    try {
      // Get document
      const doc = CBQDocumentsDAO.getById(documentId);
      if (!doc) {
        throw new Error('Document not found');
      }
      
      // Mark as processing
      CBQDocumentsDAO.markProcessing(documentId);
      
      // Extract text
      let text = '';
      let extractionMethod = '';
      
      if (doc.sourceType === CBQConfig.SOURCE_TYPES.PDF && doc.driveFileId) {
        const result = TextExtractor.extractFromPDF(doc.driveFileId);
        if (!result.success) {
          throw new Error(result.error);
        }
        text = result.text;
        extractionMethod = result.method;
      } else {
        const result = TextExtractor.processRawText(doc.rawText);
        if (!result.success) {
          throw new Error(result.error);
        }
        text = result.text;
        extractionMethod = 'TEXT';
      }
      
      // Tokenize
      const tokens = Tokenizer.tokenize(text);
      
      if (tokens.length === 0) {
        throw new Error('No words extracted from document');
      }
      
      // Store words
      const wordCount = CBQWordsDAO.bulkInsert(
        documentId,
        doc.reasonCode,
        doc.outcomeLabel,
        tokens
      );
      
      // Mark as completed
      CBQDocumentsDAO.markCompleted(documentId, wordCount);
      
      const duration = Date.now() - startTime;
      Utils.log('INFO', 'BatchProcessor', `Processed ${doc.filename}: ${wordCount} words in ${duration}ms`);
      
      return {
        success: true,
        documentId,
        wordCount,
        extractionMethod,
        durationMs: duration
      };
      
    } catch (e) {
      CBQDocumentsDAO.markError(documentId, e.message);
      Utils.log('ERROR', 'BatchProcessor', `Failed to process ${documentId}: ${e.message}`);
      
      return {
        success: false,
        documentId,
        error: e.message,
        durationMs: Date.now() - startTime
      };
    }
  },
  
  /**
   * Process a batch of queued documents
   * @param {number} limit - Max documents to process
   * @returns {Object} Batch result
   */
  processBatch(limit) {
    const batchSize = limit || CBQConfig.DEFAULTS.BATCH_SIZE;
    const queuedDocs = CBQDocumentsDAO.getQueued(batchSize);
    
    if (queuedDocs.length === 0) {
      return {
        processed: 0,
        succeeded: 0,
        failed: 0,
        results: []
      };
    }
    
    const results = [];
    let succeeded = 0;
    let failed = 0;
    
    for (const doc of queuedDocs) {
      // Check execution time limit (5 min = 300000ms)
      if (this.getRemainingExecutionTime() < 30000) {
        Utils.log('WARN', 'BatchProcessor', 'Approaching execution limit, stopping batch');
        break;
      }
      
      const result = this.processDocument(doc.id);
      results.push(result);
      
      if (result.success) {
        succeeded++;
      } else {
        failed++;
      }
    }
    
    return {
      processed: results.length,
      succeeded,
      failed,
      results
    };
  },
  
  /**
   * Recalculate all word statistics
   * @param {string} reasonCode - Reason code to recalculate (or 'GLOBAL')
   * @returns {Object} Calculation result
   */
  recalculateStats(reasonCode = 'GLOBAL') {
    const startTime = Date.now();
    
    try {
      // Get document counts
      const docCounts = CBQDocumentsDAO.getCountsByOutcome();
      
      // Get all word entries
      const wordEntries = reasonCode === 'GLOBAL' 
        ? CBQWordsDAO.getAll()
        : CBQWordsDAO.getByReasonCode(reasonCode);
      
      if (wordEntries.length === 0) {
        return { success: true, wordsUpdated: 0, reason: 'No words to process' };
      }
      
      // Aggregate and calculate metrics
      const wordStats = MetricsEngine.aggregateWordStats(wordEntries, docCounts.byOutcome);
      
      // Store results
      const statsArray = Array.from(wordStats.values());
      const updated = CBQStatsDAO.bulkUpsert(statsArray, reasonCode);
      
      // Log audit
      AuditDAO.log({
        action: CBQConfig.AUDIT_ACTIONS.STATS_CALCULATE,
        entityType: 'CBQ_WordStats',
        entityId: reasonCode,
        user: 'SYSTEM',
        module: 'CBQ',
        details: `Recalculated ${updated} word stats for ${reasonCode}`
      });
      
      const duration = Date.now() - startTime;
      Utils.log('INFO', 'BatchProcessor', `Stats recalculation: ${updated} words in ${duration}ms`);
      
      return {
        success: true,
        wordsUpdated: updated,
        durationMs: duration
      };
      
    } catch (e) {
      Utils.log('ERROR', 'BatchProcessor', `Stats recalculation failed: ${e.message}`);
      return {
        success: false,
        error: e.message
      };
    }
  },
  
  /**
   * Get remaining execution time (rough estimate)
   * Apps Script has 6-minute limit for web apps, 30 for triggers
   * @returns {number} Milliseconds remaining
   */
  getRemainingExecutionTime() {
    // This is approximate - Apps Script doesn't expose exact remaining time
    // Assume 5-minute limit for safety
    const MAX_EXECUTION_MS = 5 * 60 * 1000;
    const scriptStart = Session.getActiveUser() ? Date.now() : 0;
    return MAX_EXECUTION_MS - (Date.now() - scriptStart);
  },
  
  /**
   * Get processing queue status
   * @returns {Object} Queue status
   */
  getQueueStatus() {
    const queued = CBQDocumentsDAO.getByStatus(CBQConfig.DOCUMENT_STATUS.QUEUED);
    const processing = CBQDocumentsDAO.getByStatus(CBQConfig.DOCUMENT_STATUS.PROCESSING);
    const completed = CBQDocumentsDAO.getByStatus(CBQConfig.DOCUMENT_STATUS.COMPLETED);
    const errors = CBQDocumentsDAO.getByStatus(CBQConfig.DOCUMENT_STATUS.ERROR);
    
    return {
      queued: queued.length,
      processing: processing.length,
      completed: completed.length,
      errors: errors.length,
      total: queued.length + processing.length + completed.length + errors.length
    };
  },
  
  /**
   * Clear error status and re-queue failed documents
   * @returns {number} Documents re-queued
   */
  retryFailedDocuments() {
    const errors = CBQDocumentsDAO.getByStatus(CBQConfig.DOCUMENT_STATUS.ERROR);
    let requeued = 0;
    
    for (const doc of errors) {
      CBQDocumentsDAO.updateStatus(doc.id, CBQConfig.DOCUMENT_STATUS.QUEUED, {
        errorMessage: ''
      });
      requeued++;
    }
    
    return requeued;
  }
};

Object.freeze(BatchProcessor);
/**
 * ØN Master System - Report Workflow Engine
 * State machine for report lifecycle management
 */

const ReportWorkflow = {
  /**
   * Required approvals for a report to be accepted
   * All three roles must approve
   */
  REQUIRED_APPROVALS: ['accountable', 'consulted', 'informed'],
  
  /**
   * Change report state with validation
   * @param {string} reportId - Report ID
   * @param {string} newState - New state (Pending, Accepted, Rejected)
   * @param {string} user - User making the change
   * @param {string} comment - Optional comment
   * @returns {Object} Updated report
   */
  changeState(reportId, newState, user, comment = '') {
    const report = ReportsDAO.getReportById(reportId);
    
    if (!report) {
      throw new Error('Report not found');
    }
    
    const oldState = report.state;
    
    // Validate state transition
    ReportWorkflow.validateStateTransition(report, newState, user);
    
    // Handle rejection - returns to Responsible's queue
    if (newState === Config.STATES.REJECTED) {
      return ReportWorkflow.rejectReport(report, user, comment);
    }
    
    // Handle approval
    if (newState === Config.STATES.ACCEPTED) {
      return ReportWorkflow.approveReport(report, user, comment);
    }
    
    // Direct state change (admin only)
    if (Authorization.isAdmin(user)) {
      return ReportWorkflow.forceStateChange(report, newState, user, comment);
    }
    
    throw new Error('Invalid state transition');
  },
  
  /**
   * Validate if a state transition is allowed
   * @param {Object} report - Report object
   * @param {string} newState - Target state
   * @param {string} user - User requesting transition
   */
  validateStateTransition(report, newState, user) {
    // Admins can do anything
    if (Authorization.isAdmin(user)) return;
    
    // Check if user has permission to approve/reject
    if (!Authorization.isInApprovalChain(user, report)) {
      throw new Error('You are not in the approval chain for this report');
    }
    
    // Can only transition pending reports
    if (report.state !== Config.STATES.PENDING && !Authorization.isAdmin(user)) {
      throw new Error('Only pending reports can be approved or rejected');
    }
  },
  
  /**
   * Process approval by a user
   * @param {Object} report - Report object
   * @param {string} user - Approving user
   * @param {string} comment - Optional comment
   * @returns {Object} Updated report
   */
  approveReport(report, user, comment = '') {
    const approvalHistory = Utils.safeParseJSON(report.approvalHistory, []);
    const pendingApprovals = Utils.safeParseJSON(report.pendingApprovals, ReportWorkflow.REQUIRED_APPROVALS);
    
    // Determine which role this user is approving as
    const approvalRole = ReportWorkflow.getUserApprovalRole(user, report);
    
    if (!approvalRole) {
      throw new Error('You do not have an approval role for this report');
    }
    
    // Check if already approved
    const alreadyApproved = approvalHistory.some(
      a => a.user === user && a.action === 'approve'
    );
    
    if (alreadyApproved) {
      throw new Error('You have already approved this report');
    }
    
    // Add approval
    approvalHistory.push({
      user: user,
      action: 'approve',
      role: approvalRole,
      timestamp: Utils.now(),
      comment: comment
    });
    
    // Remove from pending
    const remainingApprovals = pendingApprovals.filter(r => r !== approvalRole);
    
    // Determine new state
    let newState = Config.STATES.PENDING;
    if (remainingApprovals.length === 0) {
      newState = Config.STATES.ACCEPTED;
    }
    
    // Update report
    const updates = {
      state: newState,
      pendingApprovals: Utils.safeStringify(remainingApprovals),
      approvalHistory: Utils.safeStringify(approvalHistory),
      updatedBy: user,
      updatedAt: Utils.now()
    };
    
    const updated = SheetsService.updateRow(Config.SHEETS.REPORTS, report.id, updates);
    
    // Log the approval
    AuditDAO.logApproval(report.id, user, approvalRole);
    
    if (newState === Config.STATES.ACCEPTED) {
      AuditDAO.logStateChange('Report', report.id, Config.STATES.PENDING, Config.STATES.ACCEPTED, user, 'All approvals received');
    }
    
    // Merge for return
    return { ...report, ...updates };
  },
  
  /**
   * Process rejection by a user
   * @param {Object} report - Report object
   * @param {string} user - Rejecting user
   * @param {string} comment - Rejection reason
   * @returns {Object} Updated report
   */
  rejectReport(report, user, comment = '') {
    const approvalHistory = Utils.safeParseJSON(report.approvalHistory, []);
    
    // Add rejection to history
    approvalHistory.push({
      user: user,
      action: 'reject',
      role: ReportWorkflow.getUserApprovalRole(user, report) || 'reviewer',
      timestamp: Utils.now(),
      comment: comment
    });
    
    // Reset to pending with all approvals required again
    const updates = {
      state: Config.STATES.REJECTED,
      pendingApprovals: Utils.safeStringify(ReportWorkflow.REQUIRED_APPROVALS),
      approvalHistory: Utils.safeStringify(approvalHistory),
      updatedBy: user,
      updatedAt: Utils.now()
    };
    
    const updated = SheetsService.updateRow(Config.SHEETS.REPORTS, report.id, updates);
    
    // Log the rejection
    AuditDAO.logRejection(report.id, user, comment);
    AuditDAO.logStateChange('Report', report.id, report.state, Config.STATES.REJECTED, user, comment);
    
    return { ...report, ...updates };
  },
  
  /**
   * Resubmit a rejected report
   * @param {string} reportId - Report ID
   * @param {string} user - User resubmitting
   * @returns {Object} Updated report
   */
  resubmitReport(reportId, user) {
    const report = ReportsDAO.getReportById(reportId);
    
    if (!report) {
      throw new Error('Report not found');
    }
    
    if (report.state !== Config.STATES.REJECTED) {
      throw new Error('Only rejected reports can be resubmitted');
    }
    
    // Check if user is responsible
    if (!Authorization.isResponsibleFor(user, report) && !Authorization.isAdmin(user)) {
      throw new Error('Only the responsible person can resubmit');
    }
    
    const approvalHistory = Utils.safeParseJSON(report.approvalHistory, []);
    approvalHistory.push({
      user: user,
      action: 'resubmit',
      role: 'responsible',
      timestamp: Utils.now(),
      comment: 'Report resubmitted for review'
    });
    
    const updates = {
      state: Config.STATES.PENDING,
      pendingApprovals: Utils.safeStringify(ReportWorkflow.REQUIRED_APPROVALS),
      approvalHistory: Utils.safeStringify(approvalHistory),
      updatedBy: user,
      updatedAt: Utils.now()
    };
    
    SheetsService.updateRow(Config.SHEETS.REPORTS, report.id, updates);
    
    AuditDAO.logStateChange('Report', report.id, Config.STATES.REJECTED, Config.STATES.PENDING, user, 'Resubmitted for review');
    
    return { ...report, ...updates };
  },
  
  /**
   * Force state change (admin only)
   * @param {Object} report - Report object
   * @param {string} newState - New state
   * @param {string} user - Admin user
   * @param {string} comment - Reason
   * @returns {Object} Updated report
   */
  forceStateChange(report, newState, user, comment = '') {
    if (!Authorization.isAdmin(user)) {
      throw new Error('Admin access required for force state change');
    }
    
    const oldState = report.state;
    
    const updates = {
      state: newState,
      updatedBy: user,
      updatedAt: Utils.now()
    };
    
    if (newState === Config.STATES.PENDING) {
      updates.pendingApprovals = Utils.safeStringify(ReportWorkflow.REQUIRED_APPROVALS);
    }
    
    SheetsService.updateRow(Config.SHEETS.REPORTS, report.id, updates);
    
    AuditDAO.logStateChange('Report', report.id, oldState, newState, user, `Admin force change: ${comment}`);
    
    return { ...report, ...updates };
  },
  
  /**
   * Get user's approval role for a report
   * @param {string} user - User email
   * @param {Object} report - Report object
   * @returns {string|null} Role name or null
   */
  getUserApprovalRole(user, report) {
    const userName = Utils.extractNameFromEmail(user).toLowerCase();
    const userEmail = user.toLowerCase();
    
    // Check each RACI role
    const roleChecks = [
      { field: 'accountable', role: 'accountable' },
      { field: 'consulted', role: 'consulted' },
      { field: 'informed', role: 'informed' },
      { field: 'qaResponsible', role: 'accountable' }
    ];
    
    for (const check of roleChecks) {
      const value = (report[check.field] || '').toLowerCase();
      if (value.includes(userEmail) || value.includes(userName)) {
        return check.role;
      }
    }
    
    // Admin can approve as any role
    if (Authorization.isAdmin(user)) {
      return 'admin';
    }
    
    return null;
  },
  
  /**
   * Get approval status for a report
   * @param {string} reportId - Report ID
   * @returns {Object} Approval status
   */
  getApprovalStatus(reportId) {
    const report = ReportsDAO.getReportById(reportId);
    
    if (!report) {
      return null;
    }
    
    const approvalHistory = Utils.safeParseJSON(report.approvalHistory, []);
    const pendingApprovals = Utils.safeParseJSON(report.pendingApprovals, []);
    
    const approvals = approvalHistory.filter(a => a.action === 'approve');
    const rejections = approvalHistory.filter(a => a.action === 'reject');
    
    return {
      state: report.state,
      pending: pendingApprovals,
      approved: approvals.map(a => ({ role: a.role, user: a.user, timestamp: a.timestamp })),
      rejected: rejections.map(r => ({ user: r.user, comment: r.comment, timestamp: r.timestamp })),
      history: approvalHistory
    };
  }
};

// Freeze to prevent modifications
Object.freeze(ReportWorkflow);
Object.freeze(ReportWorkflow.REQUIRED_APPROVALS);
/**
 * ØN Master System - Data Importer
 * Import reports from WorkFlow.txt data
 */

const DataImporter = {
  /**
   * Report data from WorkFlow.txt (39 reports + details from Reporting Tasks.txt)
   * Combined and normalized for import
   */
  REPORTS_DATA: [
    { name: "Performance Report", classification: "Master", frequency: "Daily", team: "Reporting Team", responsible: "Data Analyst", accountable: "Fraud Prevention QA", consulted: "Fraud Prevention Leader", informed: "Leadership / Client", owner: "Jatin S.", trainedResources: "Krutika P.", description: "Consolidated daily productivity and performance report using production hours and resolves data.", dataSource: "Accertify | CB Portals | AWS | SF | Recovery & QA Database", storageLocation: "One Drive > GAR Reporting > Documents > General > Reporting > Year > Performance Report > Month > Performance Report" },
    { name: "Daily Sales Report", classification: "Derived", frequency: "Daily", team: "Reporting Team", responsible: "Data Analyst", accountable: "Fraud Prevention QA", consulted: "Fraud Prevention Leader", informed: "Leadership", owner: "Krutika P.", description: "Daily sales volume and amount report from Chase, Braintree, and FiServ portals.", dataSource: "Chase | BT | FiServ", storageLocation: "Google Drive > GAR_Risk_Mgt > Financials > Sales" },
    { name: "KPI Dashboard", classification: "Master", frequency: "Daily", team: "Reporting Team", responsible: "Data Analyst", accountable: "Fraud Prevention QA", consulted: "Fraud Prevention Leader", informed: "All Team Leaders", owner: "Jatin S.", trainedResources: "Krutika P.", description: "Daily update of staffing, coverage, and attendance metrics on the KPI dashboard.", dataSource: "Monthly Schedule | Agent request forms", storageLocation: "One Drive > GAR Reporting > Documents > General > Schedules > Year > KPI - Dashboard - GAR RM Year" },
    { name: "PPS Master File", classification: "Master", frequency: "Weekly", team: "Reporting Team", responsible: "Data Analyst", accountable: "Fraud Prevention QA", consulted: "Fraud Prevention Leader", informed: "All Team Leaders", owner: "Jatin S.", trainedResources: "Krutika P.", description: "Central master file capturing Accertify, Chargeback resolves, and AWS activity for the campaign.", dataSource: "Accertify | CB Portals", storageLocation: "One Drive > GAR OPS LEADS > Documents > General > PPS - Master FIle" },
    { name: "Daily Revenue Update", classification: "Master", frequency: "Daily", team: "Reporting Team", responsible: "Data Analyst", accountable: "Fraud Prevention QA", consulted: "Fraud Prevention Leader", informed: "Leadership", owner: "Jatin S.", trainedResources: "Krutika P.", description: "Daily update of revenue pace and actuals based on team production hours.", dataSource: "AWS Data" },
    { name: "Daily Accertify Resolves", classification: "Derived", frequency: "Daily", team: "Reporting Team", responsible: "Data Analyst", accountable: "Fraud Prevention QA", consulted: "Fraud Prevention Leader", informed: "SME", owner: "Jatin S.", trainedResources: "Krutika P.", description: "Daily extract of all cases marked as Resolved in Accertify.", dataSource: "Accertify", storageLocation: "One Drive > GAR Reporting > Documents > General > Reporting > Year > Daily Resolves > Month > Accertify Resolves" },
    { name: "Daily AWS Data", classification: "Derived", frequency: "Daily", team: "Reporting Team", responsible: "Data Analyst", accountable: "Fraud Prevention QA", consulted: "Fraud Prevention Leader", informed: "SME", owner: "Jatin S.", trainedResources: "Krutika P.", description: "Daily login, logout, and productive hours report from AWS.", dataSource: "AWS", storageLocation: "One Drive > GAR Reporting > Documents > General > Reporting > Year > Daily Resolves > Month > Date daily report" },
    { name: "Daily Salesforce Resolves", classification: "Derived", frequency: "Daily", team: "Reporting Team", responsible: "Data Analyst", accountable: "Fraud Prevention QA", consulted: "Fraud Prevention Leader", informed: "SME", owner: "Jatin S.", trainedResources: "Krutika P.", description: "Daily extract of all cases resolved in Salesforce.", dataSource: "SalesForce", storageLocation: "One Drive > GAR Reporting > Documents > General > Reporting > Year > Daily Resolves > Month > Date daily report" },
    { name: "Fraud/CB Hourly Productivity", classification: "Derived", frequency: "Hourly / End of Day Summary", team: "Reporting Team", responsible: "Data Analyst", accountable: "Fraud Prevention QA", consulted: "Fraud Prevention Leader", informed: "All Team Leader", owner: "Jatin S.", trainedResources: "Krutika P.", description: "Hourly productivity tracking report for FA and CB agents.", dataSource: "Accertify | CB Portals", storageLocation: "One Drive > GAR OPS LEADS > Documents > General > PPS - Master FIle" },
    { name: "WTD Report", classification: "Master", frequency: "Monday", team: "Reporting Team", responsible: "Data Analyst", accountable: "Fraud Prevention QA", consulted: "Fraud Prevention Leader", informed: "Leadership", owner: "Jatin S.", trainedResources: "Krutika P.", description: "Weekly revenue projection update including pace, shrinkage, and attainment metrics.", dataSource: "All productivity portals & Revenue", storageLocation: "One Drive > GAR Reporting > Documents > General > Reporting > Year > YTD > WTD Summary Report" },
    { name: "Weekly Fraud Missed Ratio", classification: "Derived", frequency: "Monday", team: "Reporting Team", responsible: "Data Analyst", accountable: "Fraud Prevention QA", consulted: "Fraud Prevention Leader", informed: "SME", owner: "Jatin S.", trainedResources: "Krutika P.", description: "Weekly analysis of fraud missed across Alerts, RDR, and Chargebacks.", dataSource: "CB, RDR & Alerts portals", storageLocation: "One Drive > GAR Reporting > Documents > General > Reporting > Year > Performance Report > Month > Performance Report" },
    { name: "Queue Volume & Reject Rates", classification: "Master", frequency: "Weekly", team: "Fraud Prevention Team", responsible: "Fraud Prevention Team", accountable: "Fraud Prevention QA", consulted: "Fraud Prevention Leader", informed: "Leadership", owner: "Jeevan B.", description: "Transaction volumes for everyday based on recommendation code and booking status and rejected ratio.", dataSource: "Accertify" },
    { name: "Accertify All Transaction Activity", classification: "Master", frequency: "Daily (Raw Extract)", team: "Fraud Prevention Team", responsible: "Fraud Prevention Team", accountable: "Fraud Prevention QA", consulted: "Fraud Prevention Leader", informed: "SME", owner: "Jatin S.", trainedResources: "Krutika P.", description: "Daily extract of all transaction and resolution data from Accertify.", dataSource: "Accertify", storageLocation: "Local Drive > Getaroom (E:) > Krutika > Reporting > All Reservations Data > Year > Month" },
    { name: "Chargeback Queue", classification: "Master", frequency: "Daily", team: "Chargeback Team", responsible: "Chargeback Team", accountable: "Chargeback QA", consulted: "SME", informed: "Chargeback Leader", owner: "Jatin S.", trainedResources: "Krutika P.", description: "Daily working queue of incoming chargebacks prepared for Chargeback team action.", dataSource: "CB Portals", storageLocation: "Chargeback Management > RCA Files > Database > RCA Data" },
    { name: "Daily CB Monitoring", classification: "Master", frequency: "Daily", team: "Chargeback Team", responsible: "Chargeback Team", accountable: "Chargeback QA", consulted: "SME", informed: "Chargeback Leader", owner: "Krutika P.", description: "Daily update of sales, chargeback, and fraud CB metrics in the CB monitor sheet.", dataSource: "Chase | BT | FiServ", storageLocation: "Google Drive > GAR_Risk_Mgt > Financials > CB Ratio Monitor" },
    { name: "Daily CB Ratio", classification: "Master", frequency: "Daily", team: "Chargeback Team", responsible: "Chargeback Team", accountable: "Chargeback QA", consulted: "SME", informed: "Chargeback Leader", owner: "Krutika P.", description: "Daily calculation of chargeback ratios using sales and CB volumes.", dataSource: "Chase | BT | FiServ", storageLocation: "Google Drive > GAR_Risk_Mgt > Cgargeback Management > VIAS > Refund Thresholds" },
    { name: "CB Resolved Case Audit", classification: "Derived", frequency: "Daily", team: "Chargeback Team", responsible: "Chargeback Team", accountable: "Chargeback QA", consulted: "SME", informed: "Chargeback Leader", owner: "Jatin S.", trainedResources: "Krutika P.", description: "Audit of resolved chargeback cases to ensure documentation is correctly uploaded and completed.", dataSource: "CB portals" },
    { name: "Fraud Reason Code Mapping - CB", classification: "Derived", frequency: "As Triggered", team: "Chargeback Team", responsible: "Chargeback Team", accountable: "Chargeback QA", consulted: "SME", informed: "Chargeback Leader", owner: "Jatin S.", trainedResources: "Krutika P.", description: "Mapping fraud reason codes in FA QA files to calculate CB fraud missed ratios.", dataSource: "CB portals", storageLocation: "One Drive > GAR Reporting > Documents > General > QA Reports > Year > Fraud CB QA File - FA YTD Year" },
    { name: "Fraud Reason Code Mapping - Ethoca", classification: "Derived", frequency: "As Triggered", team: "Chargeback Team", responsible: "Chargeback Team", accountable: "Chargeback QA", consulted: "SME", informed: "Chargeback Leader", owner: "Jatin S.", trainedResources: "Krutika P.", description: "Mapping fraud reason codes in Alert QA files to calculate alert fraud missed ratios.", dataSource: "CB portals", storageLocation: "One Drive > Krutika > QA > DataBase_QC" },
    { name: "Fraud Reason Code Mapping - RDR", classification: "Derived", frequency: "As Triggered", team: "Chargeback Team", responsible: "Chargeback Team", accountable: "Chargeback QA", consulted: "SME", informed: "Chargeback Leader", owner: "Jatin S.", trainedResources: "Krutika P.", description: "Mapping fraud reason codes in RDR QA files to calculate RDR fraud missed ratios.", dataSource: "CB portals", storageLocation: "One Drive > GAR Reporting > Documents > General > QA Reports > Year > Fraud RDR QA File - FA YTD Year" },
    { name: "Guest Reservations CB Report", classification: "Master", frequency: "Friday", team: "Chargeback Team", responsible: "Chargeback Team", accountable: "Chargeback QA", consulted: "SME", informed: "Chargeback Leader", owner: "Krutika P.", description: "Weekly report of chargebacks related to guest reservations shared with clients.", dataSource: "All Chargebacks | Accertify | Supplychain", targetRecipients: "Aaron Choe & Bryan Hernandez - Priceline" },
    { name: "Top 3 Affiliates Dispute", classification: "Master", frequency: "Monthly", team: "Chargeback Team", responsible: "Chargeback Team", accountable: "Chargeback QA", consulted: "SME", informed: "Chargeback Leader", owner: "Jeevan B.", trainedResources: "Krutika P.", description: "Top 3 affiliates dispute report based on CB received date.", dataSource: "CB Financials | Tableau", targetRecipients: "Commissions team | FIN-Tech Team" },
    { name: "Top 5 Affiliates Data", classification: "Master", frequency: "Monthly", team: "Chargeback Team", responsible: "Chargeback Team", accountable: "Chargeback QA", consulted: "SME", informed: "Chargeback Leader", owner: "Jeevan B.", trainedResources: "Krutika P.", description: "Top 5 affiliates dispute report based on transaction date.", dataSource: "CB Financials | Tableau", targetRecipients: "PPS Operations" },
    { name: "Risk VB Report", classification: "Master", frequency: "Monthly", team: "Chargeback Team", responsible: "Chargeback Team", accountable: "Chargeback QA", consulted: "SME", informed: "Chargeback Leader", owner: "Jeevan B.", trainedResources: "Krutika P.", description: "Report on waiver request closed in the previous month.", dataSource: "SalesForce", targetRecipients: "Elizabeth - Priceline" },
    { name: "Excess Loss", classification: "Master", frequency: "Monthly", team: "Chargeback Team", responsible: "Chargeback Team", accountable: "Chargeback QA", consulted: "SME", informed: "Chargeback Leader", owner: "Jeevan B.", description: "The Chargebacks that we have lost & also provided manual refund causing us loss more than transaction amount.", dataSource: "CB Finacials | Supply Chain", targetRecipients: "Respective PSP" },
    { name: "Monthly CB Win/Loss Report", classification: "Master", frequency: "Monthly", team: "Chargeback Team", responsible: "Chargeback Team", accountable: "Chargeback QA", consulted: "SME", informed: "Chargeback Leader", owner: "Jeevan B.", description: "Win - Loss report of all the CB received in m-3.", dataSource: "CB Financials | RCA Report", targetRecipients: "PPS CB team | PPS Leadership Team | Client" },
    { name: "Ethoca Monthly Credit Request", classification: "Master", frequency: "Monthly", team: "Chargeback Team", responsible: "Chargeback Team", accountable: "Chargeback QA", consulted: "SME", informed: "Chargeback Leader", owner: "Jeevan B.", trainedResources: "Krutika P.", description: "Monthly credit request submission to Ethoca for non-impactful alerts.", dataSource: "Ethoca", targetRecipients: "Ethoca Team - customerservice@ethoca.com" },
    { name: "RDR Working Queue", classification: "Master", frequency: "Daily/As Triggered", team: "Chargeback Team", responsible: "Chargeback Team", accountable: "Chargeback QA", consulted: "SME", informed: "Chargeback Leader", owner: "Krutika P.", trainedResources: "Avdhut B. & Kunal J.", description: "Preparation of RDR working queue from Verifi for agent processing.", dataSource: "Verifi", storageLocation: "Google Drive > GAR_Risk_Mgt > RDR Alerts > Year > RDR Alerts - Month Year" },
    { name: "Alerts Financial Validation", classification: "Derived", frequency: "Twice a Week", team: "Chargeback Team", responsible: "Chargeback Team", accountable: "Chargeback QA", consulted: "SME", informed: "Chargeback Leader", owner: "Krutika P.", trainedResources: "Jatin S.", description: "Validation to ensure all alert-related refunds are processed correctly.", dataSource: "Ethoca", storageLocation: "Local Drive > Getaroom (E:) > Krutika > Reporting > Alerts > Alerts_MasterFile_Year" },
    { name: "Alerts Master File Update", classification: "Derived", frequency: "Twice a Week", team: "Chargeback Team", responsible: "Chargeback Team", accountable: "Chargeback QA", consulted: "SME", informed: "Chargeback Leader", owner: "Krutika P.", trainedResources: "Jatin S.", description: "Update of resolved Ethoca alerts into the Alerts Master File.", dataSource: "Ethoca", storageLocation: "Local Drive > Getaroom (E:) > Krutika > Reporting > Alerts > Alerts_MasterFile_Year" },
    { name: "Manual Refunds", classification: "Master", frequency: "Daily/As Triggered", team: "SME (Ops)", responsible: "SME (Ops)", accountable: "Chargeback QA", consulted: "Chargeback Leader", informed: "Leadership", owner: "Sandeep A.", trainedResources: "Krutika P.", description: "Processing and tracking of manual refund transactions (non-reporting task).", dataSource: "Salesforce", storageLocation: "Salesforce > Risk Refund Needed" },
    { name: "RDR Over Refunds Monitoring", classification: "Master", frequency: "Twice a Month", team: "Chargeback Team", responsible: "Chargeback Team", accountable: "Chargeback QA", consulted: "SME", informed: "Chargeback Leader", owner: "Jeevan B.", trainedResources: "Krutika P.", description: "Report of refund cases processed outside the RDR flow.", dataSource: "Verifi", targetRecipients: "Respective PSP" },
    { name: "BHI Reporting", classification: "Master", frequency: "Monthly", team: "SME", responsible: "SME", accountable: "SME", consulted: "Process Leader", informed: "Client", owner: "Jeevan B.", description: "Company level KPI report for PPS CB, recoveries, OPEX, attempts.", dataSource: "CB FInancials | Tableau", targetRecipients: "Michael & Brandon - Priceline" },
    { name: "Monthly Billing", classification: "Master", frequency: "Monthly", team: "SME", responsible: "SME", accountable: "SME", consulted: "Process Leader", informed: "Leadership", owner: "Jeevan B.", description: "Monthly attendance and billing report for invoicing purposes.", dataSource: "Attendance", targetRecipients: "PPS Leadership | Etech Billing Team" },
    { name: "CB FIN - Data Download (All networks)", classification: "Derived", frequency: "Daily / As Required", team: "SME", responsible: "SME", accountable: "SME", consulted: "Chargeback Leader", informed: "Leadership", owner: "Krutika P.", trainedResources: "Jatin S.", description: "Download of Chase and Amex adjustment reports for financial preparation.", dataSource: "Chase | AMEX", storageLocation: "Local Drive > Getaroom (E:) > Krutika > Reporting > CB Financials > Year > Month" },
    { name: "CB FIN Reconciliation (All networks)", classification: "Derived", frequency: "Biweekly", team: "SME", responsible: "SME", accountable: "SME", consulted: "Chargeback Leader", informed: "Leadership", owner: "Krutika P.", description: "Updating and reconciling Chase and Amex financial data in the master finance file.", dataSource: "Chase | AMEX", storageLocation: "Google Drive > GAR_Risk_Mgt > Financials > All Chargeback Financials Year MASTER" },
    { name: "FX Rates & Source File Update", classification: "Derived", frequency: "Biweekly", team: "SME", responsible: "SME", accountable: "SME", consulted: "Process Leader", informed: "Leadership", owner: "Krutika P.", description: "Update of FX rates in finance files and validation of currency discrepancies.", storageLocation: "Google Drive > GAR_Risk_Mgt > Financials > All Chargeback Financials Year MASTER & Source File Data" },
    { name: "CB Financials", classification: "Master", frequency: "Monthly/As Required", team: "SME", responsible: "SME", accountable: "SME", consulted: "Process Leader", informed: "Leadership/Client", owner: "Krutika P.", description: "Comprehensive chargeback financials report.", dataSource: "All Sources", storageLocation: "Google Drive > GAR_Risk_Mgt > Financials > All Chargeback Financials Year MASTER", targetRecipients: "PPS client operations | Accounting | Commissions team | FIN-Tech Team" }
  ],
  
  /**
   * Import all reports into the sheet
   * @returns {number} Number of reports imported
   */
  importAllReports() {
    const user = Auth.getCurrentUser() || 'system@init.com';
    const now = Utils.now();
    
    // Clear existing data first
    SheetsService.clearSheet(Config.SHEETS.REPORTS);
    
    const reportsToInsert = DataImporter.REPORTS_DATA.map(data => ({
      id: Utils.generateUUID(),
      name: data.name,
      classification: data.classification,
      description: data.description || '',
      frequency: data.frequency,
      team: data.team,
      responsible: data.responsible,
      accountable: data.accountable,
      consulted: data.consulted,
      informed: data.informed,
      qaResponsible: data.accountable, // QA is same as accountable in most cases
      finalAccountability: data.accountable,
      owner: data.owner || '',
      trainedResources: data.trainedResources || '',
      requestedBy: data.requestedBy || '',
      dataSource: data.dataSource || '',
      upstreamDeps: data.upstreamDeps || '',
      downstreamDeps: data.downstreamDeps || '',
      targetRecipients: data.targetRecipients || '',
      storageLocation: data.storageLocation || '',
      storageLinkType: Utils.detectStorageLinkType(data.storageLocation),
      state: Config.STATES.ACCEPTED, // Import as already accepted
      pendingApprovals: '[]',
      approvalHistory: Utils.safeStringify([{
        user: 'system',
        action: 'import',
        role: 'system',
        timestamp: now,
        comment: 'Imported from workflow data'
      }]),
      createdBy: user,
      createdAt: now,
      updatedBy: user,
      updatedAt: now
    }));
    
    const count = SheetsService.bulkInsert(Config.SHEETS.REPORTS, reportsToInsert);
    
    // Log the import
    AuditDAO.log({
      action: Config.AUDIT_ACTIONS.CREATE,
      entityType: 'System',
      entityId: 'import',
      user: user,
      details: `Bulk imported ${count} reports from workflow data`
    });
    
    Utils.log('INFO', 'DataImporter', `Imported ${count} reports`);
    
    return count;
  },
  
  /**
   * Get count of reports to import
   * @returns {number} Count
   */
  getImportCount() {
    return DataImporter.REPORTS_DATA.length;
  },
  
  /**
   * Team members for the system (27 users)
   */
  TEAM_MEMBERS: [
    { email: 'naitik.joshi@priceline.com', firstName: 'Naitik', lastName: 'Joshi' },
    { email: 'jeevan.bist@priceline.com', firstName: 'Jeevan', lastName: 'Bist' },
    { email: 'arun.choubey@priceline.com', firstName: 'Arun', lastName: 'Choubey' },
    { email: 'anusha.shetty@priceline.com', firstName: 'Anusha', lastName: 'Shetty' },
    { email: 'nitingopal.singh@priceline.com', firstName: 'Nitingopal', lastName: 'Singh' },
    { email: 'deb.mukherjee@priceline.com', firstName: 'Deb', lastName: 'Mukherjee' },
    { email: 'tapas.sharma@priceline.com', firstName: 'Tapas', lastName: 'Sharma' },
    { email: 'avdhut.bidwe@priceline.com', firstName: 'Avdhut', lastName: 'Bidwe' },
    { email: 'priyanka.thakkar@priceline.com', firstName: 'Priyanka', lastName: 'Thakkar' },
    { email: 'viren.joshi@priceline.com', firstName: 'Viren', lastName: 'Joshi' },
    { email: 'shwetank.mishra@priceline.com', firstName: 'Shwetank', lastName: 'Mishra' },
    { email: 'zankhana.yagnik@priceline.com', firstName: 'Zankhana', lastName: 'Yagnik' },
    { email: 'sandeep.agrawal@priceline.com', firstName: 'Sandeep', lastName: 'Agrawal' },
    { email: 'zankhit.mehta@priceline.com', firstName: 'Zankhit', lastName: 'Mehta' },
    { email: 'krutika.patel@priceline.com', firstName: 'Krutika', lastName: 'Patel' },
    { email: 'jayesh.parmar@priceline.com', firstName: 'Jayesh', lastName: 'Parmar' },
    { email: 'priyanka.ratudi@priceline.com', firstName: 'Priyanka', lastName: 'Ratudi' },
    { email: 'tirth.soni@priceline.com', firstName: 'Tirth', lastName: 'Soni' },
    { email: 'sanyam.sisodiya@priceline.com', firstName: 'Sanyam', lastName: 'Sisodiya' },
    { email: 'bharati.dash@priceline.com', firstName: 'Bharati', lastName: 'Dash' },
    { email: 'harshil.girnari@priceline.com', firstName: 'Harshil', lastName: 'Girnari' },
    { email: 'kunal.jadaun@priceline.com', firstName: 'Kunal', lastName: 'Jadaun' },
    { email: 'mrunal.patel@priceline.com', firstName: 'Mrunal', lastName: 'Patel' },
    { email: 'aniket.jha@priceline.com', firstName: 'Aniket', lastName: 'Jha' },
    { email: 'anagha.s@priceline.com', firstName: 'Anagha', lastName: 'S' },
    { email: 'easwaran.m@priceline.com', firstName: 'Easwaran', lastName: 'M' },
    { email: 'anjali.raj@priceline.com', firstName: 'Anjali', lastName: 'Raj' }
  ],
  
  // Backward compatibility alias
  get SAMPLE_USERS() { return this.TEAM_MEMBERS; },
  
  /**
   * Import sample users into RT_Users sheet
   * @returns {number} Number of users imported
   */
  importSampleUsers() {
    const user = Auth.getCurrentUser() || 'system@init.com';
    return UsersDAO.bulkImportUsers(DataImporter.SAMPLE_USERS, user);
  }
};

// Freeze to prevent modifications
Object.freeze(DataImporter);
/**
 * ØN Master System - Reports Data Access Object
 * CRUD operations for Reports data
 */

const ReportsDAO = {
  /**
   * Get all reports with optional filtering
   * @param {Object} filters - Filter criteria
   * @returns {Array} Filtered reports
   */
  getReports(filters = {}) {
    let reports = SheetsService.getAllRows(Config.SHEETS.REPORTS);
    
    // Apply filters
    if (filters.state) {
      reports = reports.filter(r => r.state === filters.state);
    }
    
    if (filters.team) {
      reports = reports.filter(r => r.team === filters.team);
    }
    
    if (filters.frequency) {
      reports = reports.filter(r => r.frequency === filters.frequency);
    }
    
    if (filters.classification) {
      reports = reports.filter(r => r.classification === filters.classification);
    }
    
    if (filters.responsible) {
      reports = reports.filter(r => 
        r.responsible && r.responsible.toLowerCase().includes(filters.responsible.toLowerCase())
      );
    }
    
    if (filters.search) {
      const searchLower = filters.search.toLowerCase();
      reports = reports.filter(r => 
        (r.name && r.name.toLowerCase().includes(searchLower)) ||
        (r.description && r.description.toLowerCase().includes(searchLower))
      );
    }
    
    if (filters.dateFrom) {
      const fromDate = new Date(filters.dateFrom);
      reports = reports.filter(r => new Date(r.createdAt) >= fromDate);
    }
    
    if (filters.dateTo) {
      const toDate = new Date(filters.dateTo);
      reports = reports.filter(r => new Date(r.createdAt) <= toDate);
    }
    
    // Sort by updated date descending
    reports.sort((a, b) => new Date(b.updatedAt) - new Date(a.updatedAt));
    
    return reports;
  },
  
  /**
   * Get a single report by ID
   * @param {string} reportId - Report ID
   * @returns {Object|null} Report or null
   */
  getReportById(reportId) {
    return SheetsService.getRowById(Config.SHEETS.REPORTS, reportId);
  },
  
  /**
   * Get reports for a specific user's personal queue
   * @param {string} email - User email
   * @returns {Array} User's reports
   */
  getMyReports(email) {
    const reports = SheetsService.getAllRows(Config.SHEETS.REPORTS);
    return reports.filter(r => 
      r.createdBy === email || 
      (r.responsible && r.responsible.toLowerCase().includes(email.split('@')[0].toLowerCase()))
    );
  },
  
  /**
   * Get reports pending user's approval
   * @param {string} email - User email
   * @returns {Array} Reports requiring approval
   */
  getReviewQueue(email) {
    const reports = SheetsService.getAllRows(Config.SHEETS.REPORTS);
    return Authorization.getReportsRequiringApproval(email, reports);
  },
  
  /**
   * Create a new report
   * Responsible (R) is auto-derived from creator's email.
   * Only admins can create reports for someone else.
   * @param {Object} data - Report data
   * @param {string} createdBy - Creator's email
   * @param {boolean} isAdmin - Whether creator is admin
   * @returns {Object} Created report
   */
  createReport(data, createdBy, isAdmin = false) {
    const now = Utils.now();
    
    // R (Responsible) is auto-derived from creator's email
    // Only admins can set a different responsible
    let responsible = createdBy;
    if (data.responsible && data.responsible !== createdBy) {
      if (!isAdmin) {
        throw new Error('Only admins can create reports for other users');
      }
      responsible = data.responsible;
    }
    
    const report = {
      id: Utils.generateUUID(),
      name: data.name || 'Untitled Report',
      classification: data.classification || 'Master',
      description: data.description || '',
      frequency: data.frequency || 'Daily',
      team: data.team || '',
      responsible: responsible,  // Email of responsible person
      accountable: data.accountable || '',  // Email
      consulted: data.consulted || '',      // Email (comma-separated for multiple)
      informed: data.informed || '',        // Email (comma-separated for multiple)
      qaResponsible: data.qaResponsible || '',
      finalAccountability: data.finalAccountability || '',
      owner: data.owner || '',
      trainedResources: data.trainedResources || '',
      requestedBy: data.requestedBy || '',
      dataSource: data.dataSource || '',
      upstreamDeps: data.upstreamDeps || '',
      downstreamDeps: data.downstreamDeps || '',
      targetRecipients: data.targetRecipients || '',
      storageLocation: data.storageLocation || '',
      storageLinkType: Utils.detectStorageLinkType(data.storageLocation),
      state: Config.STATES.PENDING,
      pendingApprovals: Utils.safeStringify(['accountable', 'consulted', 'informed']),
      approvalHistory: Utils.safeStringify([]),
      createdBy: createdBy,
      createdAt: now,
      updatedBy: createdBy,
      updatedAt: now
    };
    
    SheetsService.appendRow(Config.SHEETS.REPORTS, report);
    
    // Log the creation
    AuditDAO.log({
      action: Config.AUDIT_ACTIONS.CREATE,
      entityType: 'Report',
      entityId: report.id,
      user: createdBy,
      details: `Created report: ${report.name}`
    });
    
    return report;
  },
  
  /**
   * Update an existing report
   * @param {string} reportId - Report ID
   * @param {Object} data - Updated data
   * @param {string} updatedBy - Updater's email
   * @returns {Object} Updated report
   */
  updateReport(reportId, data, updatedBy) {
    const existing = ReportsDAO.getReportById(reportId);
    if (!existing) {
      throw new Error('Report not found');
    }
    
    // Update storage link type if storage location changed
    if (data.storageLocation && data.storageLocation !== existing.storageLocation) {
      data.storageLinkType = Utils.detectStorageLinkType(data.storageLocation);
    }
    
    data.updatedBy = updatedBy;
    data.updatedAt = Utils.now();
    
    const updated = SheetsService.updateRow(Config.SHEETS.REPORTS, reportId, data);
    
    // Log the update
    AuditDAO.log({
      action: Config.AUDIT_ACTIONS.UPDATE,
      entityType: 'Report',
      entityId: reportId,
      user: updatedBy,
      details: `Updated report: ${existing.name}`,
      previousValue: Utils.safeStringify(Utils.pick(existing, Object.keys(data))),
      newValue: Utils.safeStringify(Utils.pick(data, Object.keys(data)))
    });
    
    return updated;
  },
  
  /**
   * Delete a report
   * @param {string} reportId - Report ID
   * @param {string} deletedBy - Deleter's email
   * @returns {boolean} Success
   */
  deleteReport(reportId, deletedBy) {
    const existing = ReportsDAO.getReportById(reportId);
    if (!existing) {
      throw new Error('Report not found');
    }
    
    const deleted = SheetsService.deleteRow(Config.SHEETS.REPORTS, reportId);
    
    if (deleted) {
      AuditDAO.log({
        action: Config.AUDIT_ACTIONS.DELETE,
        entityType: 'Report',
        entityId: reportId,
        user: deletedBy,
        details: `Deleted report: ${existing.name}`
      });
    }
    
    return deleted;
  },
  
  /**
   * Get available filter options
   * @returns {Object} Filter options
   */
  getFilterOptions() {
    return {
      states: Object.values(Config.STATES),
      teams: Config.TEAMS,
      frequencies: Config.FREQUENCIES,
      classifications: Config.CLASSIFICATIONS,
      responsiblePersons: SheetsService.getDistinctValues(Config.SHEETS.REPORTS, 'responsible'),
      owners: SheetsService.getDistinctValues(Config.SHEETS.REPORTS, 'owner')
    };
  },
  
  /**
   * Get report counts by state
   * @returns {Object} Counts by state
   */
  getReportCounts() {
    const reports = SheetsService.getAllRows(Config.SHEETS.REPORTS);
    
    return {
      total: reports.length,
      pending: reports.filter(r => r.state === Config.STATES.PENDING).length,
      accepted: reports.filter(r => r.state === Config.STATES.ACCEPTED).length,
      rejected: reports.filter(r => r.state === Config.STATES.REJECTED).length
    };
  },
  
  /**
   * Add a comment to a report
   * @param {string} reportId - Report ID
   * @param {string} text - Comment text
   * @param {string} user - User email
   * @returns {Object} Updated report
   */
  addComment(reportId, text, user) {
    const report = ReportsDAO.getReportById(reportId);
    if (!report) throw new Error('Report not found');
    
    const comments = Utils.safeParse(report.comments) || [];
    comments.push({
      user: user,
      text: text,
      timestamp: Utils.now()
    });
    
    const updated = SheetsService.updateRow(Config.SHEETS.REPORTS, reportId, {
      comments: Utils.safeStringify(comments),
      updatedBy: user,
      updatedAt: Utils.now()
    });
    
    AuditDAO.log({
      action: Config.AUDIT_ACTIONS.COMMENT,
      entityType: 'Report',
      entityId: reportId,
      user: user,
      details: `Added comment: ${text.substring(0, 50)}...`
    });
    
    return updated;
  },
  
  /**
   * Deactivate (soft delete) a report
   * @param {string} reportId - Report ID
   * @param {string} user - User email
   * @returns {Object} Updated report
   */
  deactivateReport(reportId, user) {
    const report = ReportsDAO.getReportById(reportId);
    if (!report) throw new Error('Report not found');
    
    const updated = SheetsService.updateRow(Config.SHEETS.REPORTS, reportId, {
      isActive: false,
      updatedBy: user,
      updatedAt: Utils.now()
    });
    
    AuditDAO.log({
      action: Config.AUDIT_ACTIONS.DEACTIVATE,
      entityType: 'Report',
      entityId: reportId,
      user: user,
      details: `Deactivated report: ${report.name}`
    });
    
    return updated;
  },
  
  /**
   * Activate a report
   * @param {string} reportId - Report ID
   * @param {string} user - User email
   * @returns {Object} Updated report
   */
  activateReport(reportId, user) {
    const report = ReportsDAO.getReportById(reportId);
    if (!report) throw new Error('Report not found');
    
    const updated = SheetsService.updateRow(Config.SHEETS.REPORTS, reportId, {
      isActive: true,
      updatedBy: user,
      updatedAt: Utils.now()
    });
    
    AuditDAO.log({
      action: Config.AUDIT_ACTIONS.ACTIVATE,
      entityType: 'Report',
      entityId: reportId,
      user: user,
      details: `Activated report: ${report.name}`
    });
    
    return updated;
  },
  
  /**
   * Get reports formatted for CSV export
   * @param {Array} reportIds - Optional array of IDs to export (null = all)
   * @returns {Array} Reports for export
   */
  getReportsForExport(reportIds = null) {
    let reports = SheetsService.getAllRows(Config.SHEETS.REPORTS);
    
    if (reportIds && reportIds.length > 0) {
      reports = reports.filter(r => reportIds.includes(r.id));
    }
    
    return reports.map(r => ({
      Name: r.name,
      Classification: r.classification,
      Description: r.description,
      Frequency: r.frequency,
      Team: r.team,
      State: r.state,
      Responsible: r.responsible,
      Accountable: r.accountable,
      Consulted: r.consulted,
      Informed: r.informed,
      DataSource: r.dataSource,
      StorageLocation: r.storageLocation,
      CreatedBy: r.createdBy,
      CreatedAt: r.createdAt
    }));
  }
};

// Freeze to prevent modifications
Object.freeze(ReportsDAO);
/**
 * ØN Master System - Audit Data Access Object
 * RT-specific audit logging for compliance and traceability
 */

const AuditDAO = {
  /**
   * Log an audit entry
   * @param {Object} entry - Audit entry data
   */
  log(entry) {
    const auditEntry = {
      id: Utils.generateUUID(),
      timestamp: Utils.now(),
      user: entry.user || Auth.getCurrentUser(),
      action: entry.action || 'UNKNOWN',
      module: entry.module || 'RT',
      entityType: entry.entityType || '',
      entityId: entry.entityId || '',
      details: entry.details || '',
      previousValue: entry.previousValue || '',
      newValue: entry.newValue || ''
    };
    
    SheetsService.appendRow(Config.SHEETS.AUDIT_LOG, auditEntry);
    
    Utils.log('AUDIT', 'AuditDAO', `${auditEntry.action} on ${auditEntry.entityType}`, {
      user: auditEntry.user,
      entityId: auditEntry.entityId
    });
  },
  
  /**
   * Get audit entries with optional filtering
   * @param {Object} filters - Filter criteria
   * @returns {Array} Audit entries
   */
  getEntries(filters = {}) {
    let entries = SheetsService.getAllRows(Config.SHEETS.AUDIT_LOG);
    
    // Apply filters
    if (filters.user) {
      entries = entries.filter(e => e.user === filters.user);
    }
    
    if (filters.action) {
      entries = entries.filter(e => e.action === filters.action);
    }
    
    if (filters.entityType) {
      entries = entries.filter(e => e.entityType === filters.entityType);
    }
    
    if (filters.entityId) {
      entries = entries.filter(e => e.entityId === filters.entityId);
    }
    
    if (filters.module) {
      entries = entries.filter(e => e.module === filters.module);
    }
    
    if (filters.dateFrom) {
      const fromDate = new Date(filters.dateFrom);
      entries = entries.filter(e => new Date(e.timestamp) >= fromDate);
    }
    
    if (filters.dateTo) {
      const toDate = new Date(filters.dateTo);
      entries = entries.filter(e => new Date(e.timestamp) <= toDate);
    }
    
    // Limit results
    if (filters.limit) {
      entries = entries.slice(0, filters.limit);
    }
    
    // Sort by timestamp descending (most recent first)
    entries.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    
    return entries;
  },
  
  /**
   * Get audit entries for a specific entity
   * @param {string} entityType - Entity type
   * @param {string} entityId - Entity ID
   * @returns {Array} Audit entries for the entity
   */
  getEntityHistory(entityType, entityId) {
    return AuditDAO.getEntries({
      entityType: entityType,
      entityId: entityId
    });
  },
  
  /**
   * Get recent audit entries
   * @param {number} limit - Number of entries to return
   * @returns {Array} Recent audit entries
   */
  getRecent(limit = 50) {
    return AuditDAO.getEntries({ limit: limit });
  },
  
  /**
   * Get audit entries for a specific user
   * @param {string} email - User email
   * @param {number} limit - Number of entries to return
   * @returns {Array} User's audit entries
   */
  getUserActivity(email, limit = 100) {
    return AuditDAO.getEntries({
      user: email,
      limit: limit
    });
  },
  
  /**
   * Log a state change event
   * @param {string} entityType - Entity type
   * @param {string} entityId - Entity ID
   * @param {string} previousState - Previous state
   * @param {string} newState - New state
   * @param {string} user - User who made the change
   * @param {string} comment - Optional comment
   */
  logStateChange(entityType, entityId, previousState, newState, user, comment = '') {
    AuditDAO.log({
      action: Config.AUDIT_ACTIONS.STATE_CHANGE,
      entityType: entityType,
      entityId: entityId,
      user: user,
      details: comment || `State changed from ${previousState} to ${newState}`,
      previousValue: previousState,
      newValue: newState
    });
  },
  
  /**
   * Log an approval event
   * @param {string} entityId - Entity ID
   * @param {string} user - Approving user
   * @param {string} role - Role in approval chain
   */
  logApproval(entityId, user, role) {
    AuditDAO.log({
      action: Config.AUDIT_ACTIONS.APPROVE,
      entityType: 'Report',
      entityId: entityId,
      user: user,
      details: `Approved as ${role}`
    });
  },
  
  /**
   * Log a rejection event
   * @param {string} entityId - Entity ID
   * @param {string} user - Rejecting user
   * @param {string} reason - Rejection reason
   */
  logRejection(entityId, user, reason) {
    AuditDAO.log({
      action: Config.AUDIT_ACTIONS.REJECT,
      entityType: 'Report',
      entityId: entityId,
      user: user,
      details: reason || 'Rejected'
    });
  },
  
  /**
   * Get audit statistics
   * @returns {Object} Audit statistics
   */
  getStats() {
    const entries = SheetsService.getAllRows(Config.SHEETS.AUDIT_LOG);
    
    const actionCounts = {};
    entries.forEach(entry => {
      actionCounts[entry.action] = (actionCounts[entry.action] || 0) + 1;
    });
    
    // Get today's entries
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const todayEntries = entries.filter(e => new Date(e.timestamp) >= today);
    
    return {
      total: entries.length,
      today: todayEntries.length,
      byAction: actionCounts
    };
  }
};

// Freeze to prevent modifications
Object.freeze(AuditDAO);
/**
 * ØN Master System - Users Data Access Object
 * Handles user directory and smart display names
 */

const UsersDAO = {
  /**
   * Get all active users
   * @returns {Array} Array of user objects
   */
  getAllUsers() {
    return SheetsService.getAllRows(Config.SHEETS.USERS)
      .filter(u => u.active === true || u.active === 'true' || u.active === 'TRUE');
  },
  
  /**
   * Get all users (including inactive)
   * @returns {Array} Array of user objects
   */
  getAllUsersIncludingInactive() {
    return SheetsService.getAllRows(Config.SHEETS.USERS);
  },
  
  /**
   * Get user by email
   * @param {string} email - User email
   * @returns {Object|null} User object or null
   */
  getUserByEmail(email) {
    if (!email) return null;
    const users = SheetsService.getAllRows(Config.SHEETS.USERS);
    return users.find(u => u.email && u.email.toLowerCase() === email.toLowerCase()) || null;
  },
  
  /**
   * Create a new user
   * @param {Object} data - User data (email, firstName, lastName)
   * @param {string} createdBy - Email of creator
   * @returns {Object} Created user
   */
  createUser(data, createdBy) {
    if (!data.email) {
      throw new Error('Email is required');
    }
    
    // Validate email domain
    if (Config.VALID_EMAIL_DOMAIN && !data.email.toLowerCase().endsWith('@' + Config.VALID_EMAIL_DOMAIN)) {
      throw new Error(`Email must be from @${Config.VALID_EMAIL_DOMAIN} domain`);
    }
    
    // Check if user already exists
    const existing = UsersDAO.getUserByEmail(data.email);
    if (existing) {
      throw new Error('User with this email already exists');
    }
    
    const now = Utils.now();
    const user = {
      id: Utils.generateUUID(),
      email: data.email.toLowerCase(),
      firstName: data.firstName || UsersDAO.extractFirstName(data.email),
      lastName: data.lastName || '',
      displayName: '', // Will be computed
      active: true,
      createdAt: now,
      updatedAt: now
    };
    
    SheetsService.appendRow(Config.SHEETS.USERS, user);
    
    // Recompute display names for all users (to handle duplicates)
    UsersDAO.recomputeDisplayNames();
    
    // Log the action
    AuditDAO.log({
      action: Config.AUDIT_ACTIONS.USER_CREATE,
      entityType: 'User',
      entityId: user.email,
      user: createdBy,
      details: `Created user: ${user.firstName} ${user.lastName} (${user.email})`
    });
    
    return UsersDAO.getUserByEmail(user.email);
  },
  
  /**
   * Update a user
   * @param {string} email - User email
   * @param {Object} data - Updated data
   * @param {string} updatedBy - Email of updater
   * @returns {Object} Updated user
   */
  updateUser(email, data, updatedBy) {
    const user = UsersDAO.getUserByEmail(email);
    if (!user) {
      throw new Error('User not found');
    }
    
    const updates = {
      firstName: data.firstName !== undefined ? data.firstName : user.firstName,
      lastName: data.lastName !== undefined ? data.lastName : user.lastName,
      active: data.active !== undefined ? data.active : user.active,
      updatedAt: Utils.now()
    };
    
    SheetsService.updateRow(Config.SHEETS.USERS, email, updates);
    
    // Recompute display names
    UsersDAO.recomputeDisplayNames();
    
    AuditDAO.log({
      action: Config.AUDIT_ACTIONS.USER_UPDATE,
      entityType: 'User',
      entityId: email,
      user: updatedBy,
      details: `Updated user: ${email}`
    });
    
    return UsersDAO.getUserByEmail(email);
  },
  
  /**
   * Extract first name from email
   * @param {string} email - Email address
   * @returns {string} First name
   */
  extractFirstName(email) {
    if (!email) return '';
    const localPart = email.split('@')[0];
    const nameParts = localPart.split('.');
    if (nameParts.length > 0) {
      return Utils.capitalize(nameParts[0]);
    }
    return Utils.capitalize(localPart);
  },
  
  /**
   * Recompute display names for all users
   * Handles duplicate first names by adding last initial
   */
  recomputeDisplayNames() {
    const users = UsersDAO.getAllUsersIncludingInactive();
    if (users.length === 0) return;
    
    // Count first names
    const firstNameCounts = {};
    users.forEach(u => {
      const firstName = (u.firstName || '').toLowerCase();
      firstNameCounts[firstName] = (firstNameCounts[firstName] || 0) + 1;
    });
    
    // Update display names
    const sheet = SheetsService.getSheet(Config.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const displayNameIndex = headers.indexOf('displayName');
    const firstNameIndex = headers.indexOf('firstName');
    const lastNameIndex = headers.indexOf('lastName');
    const emailIndex = headers.indexOf('email');
    
    if (displayNameIndex === -1) return;
    
    for (let i = 1; i < data.length; i++) {
      const firstName = data[i][firstNameIndex] || '';
      const lastName = data[i][lastNameIndex] || '';
      const firstNameLower = firstName.toLowerCase();
      
      let displayName;
      if (firstNameCounts[firstNameLower] > 1 && lastName) {
        // Duplicate first name - add last initial
        displayName = `${firstName} ${lastName.charAt(0).toUpperCase()}`;
      } else {
        displayName = firstName;
      }
      
      data[i][displayNameIndex] = displayName;
    }
    
    // Write back only the displayName column
    for (let i = 1; i < data.length; i++) {
      sheet.getRange(i + 1, displayNameIndex + 1).setValue(data[i][displayNameIndex]);
    }
  },
  
  /**
   * Get smart display name for an email
   * @param {string} email - User email
   * @returns {string} Display name
   */
  getDisplayName(email) {
    if (!email) return '';
    
    const user = UsersDAO.getUserByEmail(email);
    if (user && user.displayName) {
      return user.displayName;
    }
    
    // Fallback: extract from email
    return UsersDAO.extractFirstName(email);
  },
  
  /**
   * Get users formatted for picker/autocomplete
   * @returns {Array} Array of {email, displayLabel} objects
   */
  getUsersForPicker() {
    const users = UsersDAO.getAllUsers();
    return users.map(u => ({
      email: u.email,
      displayLabel: u.displayName || u.firstName || u.email,
      firstName: u.firstName,
      lastName: u.lastName
    })).sort((a, b) => a.displayLabel.localeCompare(b.displayLabel));
  },
  
  /**
   * Search users by name or email
   * @param {string} query - Search query
   * @returns {Array} Matching users
   */
  searchUsers(query) {
    if (!query) return UsersDAO.getUsersForPicker();
    
    const q = query.toLowerCase();
    return UsersDAO.getUsersForPicker().filter(u => 
      u.displayLabel.toLowerCase().includes(q) ||
      u.email.toLowerCase().includes(q) ||
      (u.firstName && u.firstName.toLowerCase().includes(q)) ||
      (u.lastName && u.lastName.toLowerCase().includes(q))
    );
  },
  
  /**
   * Validate that an email exists in the Users directory
   * @param {string} email - Email to validate
   * @returns {boolean} True if valid
   */
  validateEmail(email) {
    if (!email) return false;
    return UsersDAO.getUserByEmail(email) !== null;
  },
  
  /**
   * Resolve multiple emails to display names
   * @param {string} emails - Comma-separated email addresses
   * @returns {string} Comma-separated display names
   */
  resolveEmailsToNames(emails) {
    if (!emails) return '';
    return emails.split(',')
      .map(e => e.trim())
      .filter(e => e)
      .map(e => UsersDAO.getDisplayName(e))
      .join(', ');
  },
  
  /**
   * Bulk import users from an array
   * @param {Array} usersData - Array of {email, firstName, lastName}
   * @param {string} importedBy - Email of importer
   * @returns {number} Number of users imported
   */
  bulkImportUsers(usersData, importedBy) {
    let count = 0;
    usersData.forEach(data => {
      try {
        // Skip if user already exists
        if (!UsersDAO.getUserByEmail(data.email)) {
          const now = Utils.now();
          const user = {
            id: Utils.generateUUID(),
            email: data.email.toLowerCase(),
            firstName: data.firstName || UsersDAO.extractFirstName(data.email),
            lastName: data.lastName || '',
            displayName: '',
            active: true,
            createdAt: now,
            updatedAt: now
          };
          SheetsService.appendRow(Config.SHEETS.USERS, user);
          count++;
        }
      } catch (e) {
        // Skip invalid emails
        Utils.log('WARN', 'UsersDAO', `Failed to import user: ${data.email}`, e);
      }
    });
    
    // Recompute display names once at the end
    if (count > 0) {
      UsersDAO.recomputeDisplayNames();
      
      AuditDAO.log({
        action: Config.AUDIT_ACTIONS.USER_CREATE,
        entityType: 'User',
        entityId: 'bulk',
        user: importedBy,
        details: `Bulk imported ${count} users`
      });
    }
    
    return count;
  }
};

Object.freeze(UsersDAO);
/**
 * ØN Master System - User Process Roles Data Access Object
 * Handles process membership and role-based access
 */

const UserProcessRolesDAO = {
  /**
   * Get all role assignments
   * @returns {Array} All role assignments
   */
  getAllRoles() {
    return SheetsService.getAllRows(Config.SHEETS.USER_PROCESS_ROLES);
  },
  
  /**
   * Get role assignments for a user
   * @param {string} email - User email
   * @returns {Array} User's role assignments
   */
  getUserRoles(email) {
    if (!email) return [];
    const roles = UserProcessRolesDAO.getAllRoles();
    return roles.filter(r => r.email && r.email.toLowerCase() === email.toLowerCase());
  },
  
  /**
   * Get role for a specific user in a specific process
   * @param {string} email - User email
   * @param {string} processId - Process ID (RT, ON, CBQ, etc.)
   * @returns {Object|null} Role assignment or null
   */
  getUserRoleForProcess(email, processId) {
    if (!email || !processId) return null;
    const roles = UserProcessRolesDAO.getAllRoles();
    return roles.find(r => 
      r.email && r.email.toLowerCase() === email.toLowerCase() &&
      r.processId && r.processId.toUpperCase() === processId.toUpperCase()
    ) || null;
  },
  
  /**
   * Check if user is assigned to a process
   * @param {string} email - User email
   * @param {string} processId - Process ID
   * @returns {boolean} True if user is in process
   */
  isUserInProcess(email, processId) {
    return UserProcessRolesDAO.getUserRoleForProcess(email, processId) !== null;
  },
  
  /**
   * Get all users in a process
   * @param {string} processId - Process ID
   * @returns {Array} Users in the process
   */
  getProcessUsers(processId) {
    if (!processId) return [];
    const roles = UserProcessRolesDAO.getAllRoles();
    return roles.filter(r => r.processId && r.processId.toUpperCase() === processId.toUpperCase());
  },
  
  /**
   * Assign user to a process with a role
   * @param {string} email - User email
   * @param {string} processId - Process ID
   * @param {string} role - Role (Viewer, Analyst, Manager, Admin)
   * @param {string} assignedBy - Email of assigner
   * @returns {Object} Role assignment
   */
  assignUserToProcess(email, processId, role, assignedBy) {
    if (!email || !processId || !role) {
      throw new Error('Email, processId, and role are required');
    }
    
    // Validate role
    const validRoles = Object.values(Config.PROCESS_ROLES);
    if (!validRoles.includes(role)) {
      throw new Error(`Invalid role. Must be one of: ${validRoles.join(', ')}`);
    }
    
    // Validate process
    const validProcesses = Object.keys(Config.MODULES);
    if (!validProcesses.includes(processId.toUpperCase())) {
      throw new Error(`Invalid process. Must be one of: ${validProcesses.join(', ')}`);
    }
    
    const now = Utils.now();
    const existing = UserProcessRolesDAO.getUserRoleForProcess(email, processId);
    
    if (existing) {
      // Update existing
      const updates = {
        role: role,
        updatedAt: now
      };
      SheetsService.updateRow(Config.SHEETS.USER_PROCESS_ROLES, existing.id, updates);
      
      AuditDAO.log({
        action: Config.AUDIT_ACTIONS.ROLE_ASSIGN,
        entityType: 'UserProcessRole',
        entityId: existing.id,
        user: assignedBy,
        details: `Updated role for ${email} in ${processId}: ${existing.role} → ${role}`
      });
      
      return { ...existing, ...updates };
    } else {
      // Create new
      const assignment = {
        id: Utils.generateUUID(),
        email: email.toLowerCase(),
        processId: processId.toUpperCase(),
        role: role,
        createdAt: now,
        updatedAt: now
      };
      
      SheetsService.appendRow(Config.SHEETS.USER_PROCESS_ROLES, assignment);
      
      AuditDAO.log({
        action: Config.AUDIT_ACTIONS.ROLE_ASSIGN,
        entityType: 'UserProcessRole',
        entityId: assignment.id,
        user: assignedBy,
        details: `Assigned ${email} to ${processId} as ${role}`
      });
      
      return assignment;
    }
  },
  
  /**
   * Remove user from a process
   * @param {string} email - User email
   * @param {string} processId - Process ID
   * @param {string} removedBy - Email of remover
   * @returns {boolean} True if removed
   */
  removeUserFromProcess(email, processId, removedBy) {
    const existing = UserProcessRolesDAO.getUserRoleForProcess(email, processId);
    if (!existing) return false;
    
    SheetsService.deleteRow(Config.SHEETS.USER_PROCESS_ROLES, existing.id);
    
    AuditDAO.log({
      action: Config.AUDIT_ACTIONS.ROLE_ASSIGN,
      entityType: 'UserProcessRole',
      entityId: existing.id,
      user: removedBy,
      details: `Removed ${email} from ${processId}`
    });
    
    return true;
  },
  
  /**
   * Check if user can perform an action in a process
   * @param {string} email - User email
   * @param {string} processId - Process ID
   * @param {string} action - Action to check (view, create, approve, admin)
   * @returns {boolean} True if allowed
   */
  canPerformAction(email, processId, action) {
    const roleAssignment = UserProcessRolesDAO.getUserRoleForProcess(email, processId);
    if (!roleAssignment) return false;
    
    const role = roleAssignment.role;
    
    // Define permissions per role
    const permissions = {
      [Config.PROCESS_ROLES.VIEWER]: ['view'],
      [Config.PROCESS_ROLES.ANALYST]: ['view', 'create', 'edit'],
      [Config.PROCESS_ROLES.MANAGER]: ['view', 'create', 'edit', 'approve', 'reject'],
      [Config.PROCESS_ROLES.ADMIN]: ['view', 'create', 'edit', 'approve', 'reject', 'delete', 'admin']
    };
    
    const allowedActions = permissions[role] || [];
    return allowedActions.includes(action.toLowerCase());
  },
  
  /**
   * Bulk assign users to a process
   * @param {Array} emails - Array of emails
   * @param {string} processId - Process ID
   * @param {string} role - Role to assign
   * @param {string} assignedBy - Email of assigner
   * @returns {number} Number of assignments
   */
  bulkAssignToProcess(emails, processId, role, assignedBy) {
    let count = 0;
    emails.forEach(email => {
      try {
        if (!UserProcessRolesDAO.isUserInProcess(email, processId)) {
          UserProcessRolesDAO.assignUserToProcess(email, processId, role, assignedBy);
          count++;
        }
      } catch (e) {
        Utils.log('WARN', 'UserProcessRolesDAO', `Failed to assign ${email} to ${processId}`, e);
      }
    });
    return count;
  }
};

Object.freeze(UserProcessRolesDAO);

/**
 * ═══════════════════════════════════════════════════════════════════════════
 * ØN Master System - Auto Report Creator
 * Handles automated creation of report tasks based on frequency
 * ═══════════════════════════════════════════════════════════════════════════
 */

const AutoReportCreator = {
  // Default Role Mapping (fallback)
  DEFAULT_ROLE_MAPPING: {
    'Data Analyst': 'naitik.joshi@priceline.com', 
    'Fraud Prevention Team': 'jeevan.bist@priceline.com',
    'Chargeback Team': 'krutika.patel@priceline.com',
    'SME': 'sanyam.sisodiya@priceline.com',
    'SME (Ops)': 'sandeep.agrawal@priceline.com',
    'Reporting Team': 'naitik.joshi@priceline.com' 
  },

  /**
   * Get role mapping from Script Properties or defaults
   * @returns {Object} map of Generic Role -> Email
   */
  getRoleMapping() {
    try {
      const props = PropertiesService.getScriptProperties();
      const stored = props.getProperty('RT_ROLE_MAPPING');
      if (stored) {
        return JSON.parse(stored);
      }
    } catch (e) {
      console.error('Error fetching role mapping', e);
    }
    return AutoReportCreator.DEFAULT_ROLE_MAPPING;
  },

  /**
   * Update role mapping
   * @param {Object} newMapping
   */
  updateRoleMapping(newMapping) {
    PropertiesService.getScriptProperties().setProperty('RT_ROLE_MAPPING', JSON.stringify(newMapping));
  },

  /**
   * Get mapped user for a generic role
   * @param {string} role
   * @returns {string} email
   */
  getMappedUser(role) {
    const mapping = AutoReportCreator.getRoleMapping();
    // Try exact match, then case-insensitive lookup
    if (mapping[role]) return mapping[role];
    
    const roleLower = (role || '').toLowerCase();
    for (const [key, value] of Object.entries(mapping)) {
      if (key.toLowerCase() === roleLower) return value;
    }
    
    return role; // Return original if no mapping found
  },

  /**
   * Initialize triggers for auto-creation
   * @returns {Object} result
   */
  initializeTriggers() {
    // Delete existing triggers first to avoid duplicates
    const triggers = ScriptApp.getProjectTriggers();
    let deleted = 0;
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === 'runDailyReportCreation') {
        ScriptApp.deleteTrigger(trigger);
        deleted++;
      }
    }

    // Create new daily trigger running between 1 AM and 2 AM
    ScriptApp.newTrigger('runDailyReportCreation')
      .timeBased()
      .everyDays(1)
      .atHour(1)
      .create();

    return { 
      success: true, 
      message: `Initialized daily trigger (1 AM). Deleted ${deleted} old triggers.` 
    };
  },

  /**
   * Main function to run daily and create due reports
   * @returns {number} count of created reports
   */
  runDailyReportCreation() {
    const today = new Date();
    console.log(`Starting AutoReportCreator for ${today.toISOString()}`);
    
    const reports = DataImporter.REPORTS_DATA; 
    let createdCount = 0;

    for (const reportTemplate of reports) {
      if (AutoReportCreator.shouldCreateReport(reportTemplate.frequency, today)) {
        try {
          AutoReportCreator.createReportFromTemplate(reportTemplate, today);
          createdCount++;
        } catch (e) {
          console.error(`Failed to create auto-report for ${reportTemplate.name}`, e);
        }
      }
    }
    
    console.log(`Auto-creation completed. Created ${createdCount} reports.`);
    return createdCount;
  },

  /**
   * Check if report should be created today based on frequency
   * @param {string} frequency
   * @param {Date} date
   * @returns {boolean}
   */
  shouldCreateReport(frequency, date) {
    if (!frequency) return false;
    
    const dayOfWeek = date.getDay(); // 0 = Sun, 1 = Mon, ...
    const dayOfMonth = date.getDate();
    const freq = frequency.toLowerCase();

    // Daily logic
    if (freq.includes('daily') || freq.includes('hourly')) {
      // Optional: Skip weekends if needed, but 'Daily' usually implies 7 days
      return true; 
    }

    // Weekly logic
    if (freq === 'weekly' || freq === 'monday') return dayOfWeek === 1; // Monday
    if (freq === 'friday') return dayOfWeek === 5;
    
    // Custom Weekly
    if (freq === 'twice a week') return dayOfWeek === 1 || dayOfWeek === 4; // Mon & Thu
    
    // Monthly/Bi-monthly
    if (freq === 'monthly' || freq.includes('monthly')) return dayOfMonth === 1;
    if (freq === 'twice a month') return dayOfMonth === 1 || dayOfMonth === 15;
    if (freq === 'biweekly') return dayOfMonth === 1 || dayOfMonth === 15; 
    
    // Manual only
    if (freq.includes('as triggered') || freq.includes('manual')) return false;

    return false;
  },

  /**
   * Create a single report instance from template
   * @param {Object} template
   * @param {Date} date
   * @returns {Object} created report
   */
  createReportFromTemplate(template, date) {
    const responsibleEmail = AutoReportCreator.getMappedUser(template.responsible);
    const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    
    // Construct report name
    let name = template.name;
    // Append date if it's not a master file or generic name
    if (!name.toLowerCase().includes('master')) {
        name = `${name} (${dateStr})`;
    }

    const reportData = {
      name: name,
      classification: template.classification,
      description: template.description || '',
      frequency: template.frequency,
      team: template.team,
      responsible: responsibleEmail,
      accountable: template.accountable,
      consulted: template.consulted,
      informed: template.informed,
      qaResponsible: template.accountable, 
      finalAccountability: template.accountable,
      owner: template.owner,
      trainedResources: template.trainedResources,
      dataSource: template.dataSource,
      targetRecipients: template.targetRecipients,
      storageLocation: template.storageLocation
    };
    
    // Create as 'system' admin
    return ReportsDAO.createReport(reportData, 'system@automation.bot', true); 
  },

  /**
   * Generate 1 month of demo data
   * @returns {number} count
   */
  generateDemoData() {
    const today = new Date();
    let count = 0;
    
    // Generate for last 30 days
    for (let i = 30; i >= 0; i--) {
        const d = new Date();
        d.setDate(today.getDate() - i);
        
        const reports = DataImporter.REPORTS_DATA;
        for (const reportTemplate of reports) {
            if (AutoReportCreator.shouldCreateReport(reportTemplate.frequency, d)) {
                
                // Create report
                const r = AutoReportCreator.createReportFromTemplate(reportTemplate, d);
                
                // For past reports, randomly mark as ACCEPTED to simulate completed work
                if (i > 0 && Math.random() > 0.3) { 
                     
                     // Create update object
                     const updates = {
                         state: Config.STATES.ACCEPTED,
                         pendingApprovals: '[]',
                         approvalHistory: Utils.safeStringify([{
                            user: 'system@demo.data',
                            action: 'approve',
                            role: 'admin',
                            timestamp: r.createdAt,
                            comment: 'Auto-generated demo completion'
                         }]),
                         updatedAt: r.createdAt
                     };
                     
                     // Update in sheet
                     SheetsService.updateRow(Config.SHEETS.REPORTS, r.id, updates);
                }
                count++;
            }
        }
    }
    return count;
  }
};

Object.freeze(AutoReportCreator);

/**
 * ═══════════════════════════════════════════════════════════════════════════
 * GLOBAL TRIGGERS & API
 * ═══════════════════════════════════════════════════════════════════════════
 */

/**
 * Trigger handler function (must be top-level)
 */
function runDailyReportCreation() {
  AutoReportCreator.runDailyReportCreation();
}

/**
 * API: Initialize auto-creation triggers
 */
function initializeAutoCreation() {
  if (!Authorization.isAdmin(Auth.getCurrentUser())) {
    throw new Error('Unauthorized');
  }
  return AutoReportCreator.initializeTriggers();
}

/**
 * API: Generate demo data
 */
function generateDemoData() {
  if (!Authorization.isAdmin(Auth.getCurrentUser())) {
    throw new Error('Unauthorized');
  }
  return AutoReportCreator.generateDemoData();
}

/**
 * API: Update role mappings
 */
function updateRoleMap(mapping) {
  if (!Authorization.isAdmin(Auth.getCurrentUser())) {
    throw new Error('Unauthorized');
  }
  AutoReportCreator.updateRoleMapping(mapping);
  return { success: true };
}

/**
 * API: Get role mappings
 */
function getRoleMap() {
    return AutoReportCreator.getRoleMapping();
}
/**
 * ═══════════════════════════════════════════════════════════════════════════
 * ORION ENTERPRISE PLATFORM - FA COACHING MODULE
 * ═══════════════════════════════════════════════════════════════════════════
 * 
 * Backend logic for the Fraud Analysis Coaching module.
 * Reads coaching case data from Google Sheets (QA-populated) and allows
 * SMEs to review cases and add coaching notes.
 * 
 * Sheet Schema (source):
 *   A: Coaching Date | B: Analyst Name | C: ITN | D: Date | E: Reso | F: Notes
 * 
 * LAST UPDATED: 2025-02-10
 */

// ═══════════════════════════════════════════════════════════════════════════
// FA COACHING CONFIG
// ═══════════════════════════════════════════════════════════════════════════

const FA_COACHING_CONFIG = {
  SHEET_ID: '1qvDu_adBcaP__9sCxZ63He349g7ERrReKfmKlOdaNsM',
  SHEET_NAME: 'Sheet1',
  
  // Column indices (0-based)
  COLS: {
    COACHING_DATE: 0,  // A
    ANALYST_NAME: 1,   // B
    ITN: 2,            // C
    DATE: 3,           // D
    RESO: 4,           // E
    NOTES: 5,          // F
    SME_NOTES: 6       // G (we will use this for SME notes)
  }
};

Object.freeze(FA_COACHING_CONFIG);

// ═══════════════════════════════════════════════════════════════════════════
// FA COACHING API
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Get all unique coaching dates from the sheet
 * @returns {string[]} Array of date strings sorted descending (newest first)
 */
function faGetCoachingDates() {
  try {
    const sheet = SpreadsheetApp.openById(FA_COACHING_CONFIG.SHEET_ID)
      .getSheetByName(FA_COACHING_CONFIG.SHEET_NAME);
    
    if (!sheet) return [];
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    const dateCol = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const dates = new Set();
    
    dateCol.forEach(row => {
      const val = row[0];
      if (val) {
        // Normalize: if it's a Date object, format it; if string, keep as-is
        let dateStr;
        if (val instanceof Date) {
          dateStr = Utilities.formatDate(val, Session.getScriptTimeZone(), 'MM/dd/yy');
        } else {
          dateStr = String(val).trim();
        }
        if (dateStr) dates.add(dateStr);
      }
    });
    
    // Sort descending by parsing dates
    return Array.from(dates).sort((a, b) => {
      const da = parseDateStr_(a);
      const db = parseDateStr_(b);
      return db - da;
    });
    
  } catch (e) {
    console.error('faGetCoachingDates error:', e);
    return [];
  }
}

/**
 * Get all cases for a given coaching date
 * @param {string} dateStr - The coaching date string to filter by
 * @returns {Object[]} Array of case objects {name, itn, date, reso, notes, smeNotes, rowIndex}
 */
function faGetCasesByDate(dateStr) {
  try {
    const sheet = SpreadsheetApp.openById(FA_COACHING_CONFIG.SHEET_ID)
      .getSheetByName(FA_COACHING_CONFIG.SHEET_NAME);
    
    if (!sheet) return [];
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    // Read all columns including potential SME_NOTES column
    const lastCol = Math.max(sheet.getLastColumn(), 7);
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    
    const cases = [];
    const targetDate = dateStr.trim();
    
    data.forEach((row, idx) => {
      const cellDate = row[FA_COACHING_CONFIG.COLS.COACHING_DATE];
      let cellDateStr;
      
      if (cellDate instanceof Date) {
        cellDateStr = Utilities.formatDate(cellDate, Session.getScriptTimeZone(), 'MM/dd/yy');
      } else {
        cellDateStr = String(cellDate || '').trim();
      }
      
      if (cellDateStr === targetDate) {
        cases.push({
          name: String(row[FA_COACHING_CONFIG.COLS.ANALYST_NAME] || ''),
          itn: String(row[FA_COACHING_CONFIG.COLS.ITN] || ''),
          date: formatCellDate_(row[FA_COACHING_CONFIG.COLS.DATE]),
          reso: String(row[FA_COACHING_CONFIG.COLS.RESO] || ''),
          notes: String(row[FA_COACHING_CONFIG.COLS.NOTES] || ''),
          smeNotes: String(row[FA_COACHING_CONFIG.COLS.SME_NOTES] || ''),
          rowIndex: idx + 2  // 1-indexed sheet row (data starts at row 2)
        });
      }
    });
    
    return cases;
    
  } catch (e) {
    console.error('faGetCasesByDate error:', e);
    return [];
  }
}

/**
 * Save SME coaching notes for a specific case
 * @param {number} rowIndex - The 1-indexed sheet row of the case
 * @param {string} notes - The SME notes to save
 * @returns {Object} Result object {success, message}
 */
function faSaveSmeNotes(rowIndex, notes) {
  try {
    const sheet = SpreadsheetApp.openById(FA_COACHING_CONFIG.SHEET_ID)
      .getSheetByName(FA_COACHING_CONFIG.SHEET_NAME);
    
    if (!sheet) {
      return { success: false, message: 'Sheet not found' };
    }
    
    // Ensure header exists for SME_Notes column (column G = 7)
    const smeCol = FA_COACHING_CONFIG.COLS.SME_NOTES + 1; // 1-indexed = 7
    const header = sheet.getRange(1, smeCol).getValue();
    if (!header || String(header).trim() === '') {
      sheet.getRange(1, smeCol).setValue('SME Notes');
    }
    
    // Write the notes
    sheet.getRange(rowIndex, smeCol).setValue(notes);
    
    const user = Session.getActiveUser().getEmail();
    console.log(`FA Coaching: SME notes saved by ${user} for row ${rowIndex}`);
    
    return { success: true, message: 'Notes saved successfully' };
    
  } catch (e) {
    console.error('faSaveSmeNotes error:', e);
    return { success: false, message: 'Failed to save notes: ' + e.message };
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// INTERNAL HELPERS
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Parse a date string in MM/dd/yy format to a Date object
 * @param {string} str - Date string
 * @returns {Date} Parsed date
 */
function parseDateStr_(str) {
  if (!str) return new Date(0);
  
  const parts = str.split('/');
  if (parts.length === 3) {
    let year = parseInt(parts[2], 10);
    if (year < 100) year += 2000;
    return new Date(year, parseInt(parts[0], 10) - 1, parseInt(parts[1], 10));
  }
  
  return new Date(str);
}

/**
 * Format a cell date value to a display string
 * @param {*} val - Cell value (Date object or string)
 * @returns {string} Formatted date string
 */
function formatCellDate_(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'MM/dd/yy');
  }
  return String(val).trim();
}
