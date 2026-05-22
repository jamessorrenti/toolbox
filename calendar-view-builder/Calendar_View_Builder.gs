/**
 * Calendar View Generator
 *
 * Generates dynamic, view-only calendar tabs from a local source sheet.
 * Supports create/replace/refresh actions, custom date ranges,
 * per-calendar filters, refresh, event hover details, and overflow modals.
 *
 * Configure source, dates, filters, and refresh from the top rows.
 */

// Global sheet name variables. The Key tab can override these — see
// KEY_SETUP_OPTIONS (defaultDataSheetName).
const APP_CONFIG = {
  keySheetName: "Key",
  dataSheetName: "Events"
};

// Per-execution render session. Only active during multi-calendar refreshes;
// lets loadSourceData / readKeyConfig_ / applyKeyOverrides_ skip duplicate sheet reads.
const RENDER_SESSION = {
  active: false,
  sourceData: Object.create(null),
  keyConfig: null,
  keyOverridesApplied: false,
};

function withRenderSession_(fn) {
  const wasActive = RENDER_SESSION.active;
  if (!wasActive) {
    RENDER_SESSION.active = true;
    RENDER_SESSION.sourceData = Object.create(null);
    RENDER_SESSION.keyConfig = null;
    RENDER_SESSION.keyOverridesApplied = false;
  }
  try {
    return fn();
  } finally {
    if (!wasActive) {
      RENDER_SESSION.active = false;
      RENDER_SESSION.sourceData = Object.create(null);
      RENDER_SESSION.keyConfig = null;
      RENDER_SESSION.keyOverridesApplied = false;
    }
  }
}

// Menu
function onOpen(e) {
  buildCalendarMenu();

  /*
  To merge this into another file's onOpen:
  - comment out or remove this function
  - add the block below to the other onOpen

  try {
    buildCalendarMenu();
  } catch (err) {
    console.error('buildCalendarMenu error:', err);
  }
  */
}

function buildCalendarMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  applyKeyOverrides_(ss);

  const keySheet = ss.getSheetByName(CALENDAR.keySheetName);
  if (keySheet) {
    try {
      applyKeySetupValidations_(keySheet);
    } catch (err) {
      Logger.log("Could not refresh Key dropdowns on open: " + err.message);
    }
  }

  const ui = SpreadsheetApp.getUi();

  const showInitialMenu = getKeyBooleanOption_(
    ss,
    "showInitialMenu",
    CALENDAR.showInitialMenu
  );
  const showEventListMenu = getKeyBooleanOption_(
    ss,
    "showEventListMenu",
    CALENDAR.showEventListMenu
  );
  const showSetKeyFromEventListMenu = getKeyBooleanOption_(
    ss,
    "showSetKeyFromEventListMenu",
    CALENDAR.showSetKeyFromEventListMenu
  );
  // Show Import Theme when the Key toggle says so OR when any calendar tab
  // has the per-tab override band — once a tab has overrides, the user needs
  // a way to import a theme into it regardless of the global toggle.
  const showImportThemeMenu = getKeyBooleanOption_(
    ss,
    "showImportThemeMenu",
    CALENDAR.showImportThemeMenu
  ) || hasAnyPerTabOverride_(ss);
  const showKeyConfiguratorMenuItems = getKeyBooleanOption_(
    ss,
    "showKeyConfiguratorMenuItems",
    CALENDAR.showKeyConfiguratorMenuItems
  );

  const menu = ui.createMenu(CALENDAR.menuName);

  if (showInitialMenu) {
    menu.addItem("Initial Setup", "initialSetup")
      .addSeparator();
  }

  menu.addItem("New Calendar Sheet", "newCalendarSheet")
    .addItem("Replace with Calendar", "replaceWithCalendar")
    .addSeparator()
    .addItem("Refresh All Calendars", "refreshAllCalendars")
    .addItem("Open Selected", "openSelected")
    .addSeparator()
    .addItem("Add Q1-Q4", "addQ1Q4")
    .addItem("Add Jan-Dec", "addJanDec");

  if (showEventListMenu) {
    menu.addSeparator()
      .addItem("Create Event List", "createEventList");
  }

  if (showSetKeyFromEventListMenu) {
    if (!showEventListMenu) menu.addSeparator();
    menu.addItem("Set Key From Event List", "setKeyFromEventList");
  }

  if (showImportThemeMenu) {
    if (!showEventListMenu && !showSetKeyFromEventListMenu) menu.addSeparator();
    menu.addItem("Import Theme", "importTheme");

    // "Add Theme Override" appears when the active tab is a calendar without
    // an override band. Menu state is rebuilt at onOpen, so switching tabs
    // requires a reload to refresh the menu — handler also validates.
    const activeForMenu = ss.getActiveSheet();
    if (activeForMenu && isCalendarSheet(activeForMenu) && !hasPerTabOverride_(activeForMenu)) {
      menu.addItem("Add Theme Override", "addThemeOverride");
    }
  }

  if (showKeyConfiguratorMenuItems) {
    menu.addSeparator()
      .addItem("Run key configurator", "runKeyConfigurator")
      .addItem("Set key-based validation", "setKeyBasedDataValidation")
      .addItem("Set key-based colors", "setKeyBasedConditionalFormatting");
  }

  if (!ss.getSheetByName(CALENDAR.keySheetName)) {
    menu.addSeparator()
      .addItem("Create Key (and customize)", "createKeySheet");
  } else if (keyHasMissingFeatures_(ss)) {
    // Migration helper: surfaces only when the script knows about Key options
    // the user's Key tab doesn't have yet. Appends new rows without touching
    // existing values.
    menu.addSeparator()
      .addItem("Update Key (add missing features)", "updateKey");
  }

  menu.addToUi();
}



// Customization
const CALENDAR = {
  version: "13.15.0",
  menuName: "Calendar Tools",
  showInitialMenu: true,
  showEventListMenu: true,
  showSetKeyFromEventListMenu: true,
  showImportThemeMenu: true,
  showKeyConfiguratorMenuItems: true,

  // Auto-Refresh: when TRUE, source-list edits invalidate calendars and the
  // calendar re-renders when the user switches to it. Default ON.
  autoRefresh: true,

  // Public theme registry. The index file lists themes; each theme file lives
  // in the same directory. Fork the toolbox repo and point these URLs at your
  // fork to ship private/internal themes.
  themes: {
    indexUrl: "https://raw.githubusercontent.com/jamessorrenti/toolbox/main/calendar-view-builder/themes/index.json",
    baseUrl: "https://raw.githubusercontent.com/jamessorrenti/toolbox/main/calendar-view-builder/themes/"
  },

  calendarBaseName: "Calendar View",
  defaultDataSheetName: APP_CONFIG.dataSheetName,
  keySheetName: APP_CONFIG.keySheetName,

  defaultYear: 2026,

  periodOptions: [
    "Q1", "Q2", "Q3", "Q4",
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
    "Year",
    "Custom"
  ],

  yearOptions: ["2025", "2026", "2027", "2028", "2029", "2030"],

  monthOptions: [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ],

  setup: {
    q1StartMonth: "January",
    q1StartDay: 1,
    startWeekOn: "Sunday",
    dayFormat: "EEEE",
    dateFormat: "d",
    monthFormat: "MMMM",
    fontFamily: "Inter",

    frozenWeekdayHeader: true,

    maxEvents: 4,
    filterDefaultLabel: "All Events",

    customAdditional: ["Owner"],
    customAdditionalLabels: true,
    customAdditionalLabelsStyle: "Bold",

    customDate: "Date",
    customTitle: "Title",
    customType: "Type",
    customCategory: "Category",
    customStatus: "Status",

    eventSortOrder: "Category ↓, Status ↓, Alphabetical ↓",
  },

  keyDefaults: {
    types: [
      ["Webinar", "🖥️"],
      ["Event", "📅"],
      ["Email", "📧"],
      ["Direct Mail", "📬"],
      ["Community", "👥"],
      ["Outbound", "📣"],
      ["Web", "🌐"],
      ["Paid", "💸"],
      ["Social", "📲"],
    ],

    categories: [
      ["Category 1", "#C8E6C9"],
      ["Category 2", "#BBDEFB"],
      ["Category 3", "#E1BEE7"],
    ],

    statuses: [
      ["Not Started", "⚪"],
      ["In Progress", "🔵"],
      ["In Review", "🟡"],
      ["Live", "🟢"],
      ["Paused", "🛑"],
      ["Complete", "✅"],
    ],
  },

  colors: {
    titleBackground: "#4A0039",
    titleFontColor: "#FFFFFF",
    titleRefreshColor: "#4A0039",
    titleAccentColor: "#7A005D",

    titleWeekBackground: "#EAB8F2",
    titleWeekFontColor: "#000000",

    month1HeaderBackground: "#7A005D",
    month1HeaderFontColor: "#FFFFFF",
    month1WeekBackground: "#EAB8F2",
    month1WeekFontColor: "#000000",
    month1DaysBackground: "#FFFFFF",
    month1DaysFontColor: "#000000",
    month1DaysInactiveBackgroundColor: "#F2EEEB",
    month1DaysInactiveFontColor: "#949494",

    month2HeaderBackground: "#0C674D",
    month2HeaderFontColor: "#FFFFFF",
    month2WeekBackground: "#BAE5D9",
    month2WeekFontColor: "#000000",
    month2DaysBackground: "#FFFFFF",
    month2DaysFontColor: "#000000",
    month2DaysInactiveBackgroundColor: "#F2EEEB",
    month2DaysInactiveFontColor: "#949494",

    month3HeaderBackground: "#1E4AA9",
    month3HeaderFontColor: "#FFFFFF",
    month3WeekBackground: "#B9DBF3",
    month3WeekFontColor: "#000000",
    month3DaysBackground: "#FFFFFF",
    month3DaysFontColor: "#000000",
    month3DaysInactiveBackgroundColor: "#F2EEEB",
    month3DaysInactiveFontColor: "#949494",

    eventDefaultBackground: "#FFFFFF",
    eventDefaultFontColor: "#000000",
    overflowBackground: "#F2EEEB",
    overflowFontColor: "#4A0039",

    todayBackground: "#FCE4EC",
    todayFontColor: "#4A0039",
  },

  layout: {
    frozenRows: 2,
    firstMonthRow: 5,
    columns: 7,
    renderRowCapacity: 700,
    dateRowHeight: 18,
    eventRowHeight: 22,
    spacerRowHeight: 10,
    bufferRowHeight: 5,
  },

  metadata: {
    overflowPrefix: "CALENDAR_OVERFLOW|",
  }
};

const KEY_SETUP_OPTIONS = [
  "defaultDataSheetName",
  "showInitialMenu",
  "showEventListMenu",
  "showSetKeyFromEventListMenu",
  "showImportThemeMenu",
  "showKeyConfiguratorMenuItems",
  "autoRefresh",
  "frozenWeekdayHeader",
  "customDate",
  "customTitle",
  "customType",
  "customCategory",
  "customStatus",
  "maxEvents",
  "customAdditional",
  "eventSortOrder",
  "customAdditionalLabels",
  "customAdditionalLabelsStyle",
  "q1StartMonth",
  "startWeekOn",
  "dayFormat",
  "dateFormat",
  "monthFormat",
  "fontFamily",
];

// Values written into the Key tab on creation when they should differ from the
// script defaults in CALENDAR / CALENDAR.setup. The script defaults still apply
// for spreadsheets that have no Key tab at all.
const KEY_INITIAL_VALUES = {
  showInitialMenu: false,
  showEventListMenu: false,
  showSetKeyFromEventListMenu: false,
  showImportThemeMenu: false,
  showKeyConfiguratorMenuItems: false,
  // autoRefresh defaults TRUE in the Key (different from the show* toggles).
  // It controls whether source-list edits trigger lazy calendar refresh on
  // tab switch.
  autoRefresh: true,
  // Blank by default so the Key cell starts empty — applyKeyOverrides_ then
  // leaves CALENDAR.setup.eventSortOrder at its script default
  // ("Category ↓, Status ↓, Alphabetical ↓"). Users can type a value to
  // override, or "Source ↓" to preserve source order without any sort.
  eventSortOrder: "",
  frozenWeekdayHeader: false,
};

const KEY_APPEARANCE_OPTIONS = [
  "titleBackground",
  "titleFontColor",
  "titleRefreshColor",
  "titleAccentColor",
  "titleWeekBackground",
  "titleWeekFontColor",
  "eventDefaultBackground",
  "eventDefaultFontColor",
  "overflowBackground",
  "overflowFontColor",
  "todayBackground",
  "todayFontColor",

  "month1HeaderBackground",
  "month1HeaderFontColor",
  "month1WeekBackground",
  "month1WeekFontColor",
  "month1DaysBackground",
  "month1DaysFontColor",
  "month1DaysInactiveBackgroundColor",
  "month1DaysInactiveFontColor",

  "month2HeaderBackground",
  "month2HeaderFontColor",
  "month2WeekBackground",
  "month2WeekFontColor",
  "month2DaysBackground",
  "month2DaysFontColor",
  "month2DaysInactiveBackgroundColor",
  "month2DaysInactiveFontColor",

  "month3HeaderBackground",
  "month3HeaderFontColor",
  "month3WeekBackground",
  "month3WeekFontColor",
  "month3DaysBackground",
  "month3DaysFontColor",
  "month3DaysInactiveBackgroundColor",
  "month3DaysInactiveFontColor",
];

// Options offered in the eventSortOrder dropdown.
//   "↓" = ascending  (A-Z, oldest first, lowest first)
//   "↑" = descending (Z-A, newest first, highest first)
// Multi-select on the Key cell stacks the order: primary, secondary, etc.
const EVENT_SORT_OPTIONS = [
  "Category ↓", "Category ↑",
  "Status ↓", "Status ↑",
  "Type ↓", "Type ↑",
  "Alphabetical ↓", "Alphabetical ↑",
  "Source ↓", "Source ↑",
];

// === Per-tab override =====================================================
// A calendar tab can optionally carry its own grouped column band (H–M) that
// overrides selected setup/appearance values just for that tab. Render path:
//   script defaults → Key tab overrides → per-tab overrides → rendered output
//
// The band is only present on tabs that have opted in (typically via
// Import Theme > "this tab only"). Tabs without it render as before.

// Spreadsheet-wide concerns that intentionally do NOT participate in per-tab
// overrides — they identify the source / control menu visibility.
const PER_TAB_OVERRIDE_EXCLUDED_OPTIONS = [
  "defaultDataSheetName",
  "showInitialMenu",
  "showEventListMenu",
  "showSetKeyFromEventListMenu",
  "showImportThemeMenu",
  "showKeyConfiguratorMenuItems",
];

// Boolean options get a 3-state dropdown ("", "TRUE", "FALSE") rather than a
// checkbox so users can distinguish "no override" (blank) from "override to
// FALSE" (explicit).
const PER_TAB_OVERRIDE_BOOLEAN_OPTIONS = [
  "frozenWeekdayHeader",
  "customAdditionalLabels",
];

const PER_TAB_OVERRIDE = {
  header: "Theme Override",
  importLabel: "Import Theme:",
  // Column layout (1-indexed): H I J K L M
  headerCol: 8,
  setupNameCol: 9,
  setupValueCol: 10,
  spacerCol: 11,
  appearanceNameCol: 12,
  appearanceValueCol: 13,
  // Row layout
  importRow: 2,
  subHeaderRow: 3,
  dataStartRow: 4,
};

function perTabOverrideSetupOptions_() {
  return KEY_SETUP_OPTIONS.filter(o => PER_TAB_OVERRIDE_EXCLUDED_OPTIONS.indexOf(o) < 0);
}

function hasPerTabOverride_(sheet) {
  if (!sheet) return false;
  try {
    const marker = sheet.getRange(1, PER_TAB_OVERRIDE.headerCol).getValue();
    return String(marker || "").trim() === PER_TAB_OVERRIDE.header;
  } catch (err) {
    return false;
  }
}

// True if any sheet in the workbook has the per-tab override band. Used to
// keep "Import Theme" reachable from the menu even when showImportThemeMenu
// is FALSE in the Key — once a tab has overrides, the user needs a way to
// import to it.
function hasAnyPerTabOverride_(ss) {
  try {
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      if (hasPerTabOverride_(sheets[i])) return true;
    }
  } catch (err) {}
  return false;
}

// Builds the per-tab override column band (H–M) on a calendar tab. Idempotent
// — returns immediately if the band already exists.
// Row offsets for the Category sub-section inside the override band, computed
// dynamically because they depend on perTabOverrideSetupOptions_() length.
function perTabOverrideCategoryHeaderRow_() {
  return PER_TAB_OVERRIDE.dataStartRow + perTabOverrideSetupOptions_().length + 1; // +1 = buffer row
}
function perTabOverrideCategoryDataStartRow_() {
  return perTabOverrideCategoryHeaderRow_() + 1;
}

function addPerTabOverride_(sheet) {
  if (hasPerTabOverride_(sheet)) return;

  const fontFamily = CALENDAR.setup.fontFamily || "Inter";
  const requiredCols = PER_TAB_OVERRIDE.appearanceValueCol;
  if (sheet.getMaxColumns() < requiredCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredCols - sheet.getMaxColumns());
  }

  const setupOptions = perTabOverrideSetupOptions_();
  const appearanceOptions = KEY_APPEARANCE_OPTIONS.slice();
  const headerColCount = requiredCols - PER_TAB_OVERRIDE.headerCol + 1;

  // Layout:
  //   I-J: setup rows | buffer row | "Category" sub-header | category rows
  //   L-M: appearance rows (run all the way down in parallel)
  const initialCategories = readKeyCategoriesDisplay_(sheet.getParent());
  const categoryHeaderRow = perTabOverrideCategoryHeaderRow_();
  const categoryDataStartRow = perTabOverrideCategoryDataStartRow_();
  const categoryRowCount = initialCategories.length;
  const lastSetupBlockRow = categoryDataStartRow + Math.max(categoryRowCount, 0) - 1;
  const lastAppearanceRow = PER_TAB_OVERRIDE.dataStartRow + appearanceOptions.length - 1;
  const lastDataRow = Math.max(lastSetupBlockRow, lastAppearanceRow);
  const dataRows = lastDataRow - PER_TAB_OVERRIDE.dataStartRow + 1;

  // Row 1: merged "Per-Tab Override" header
  const headerRange = sheet.getRange(1, PER_TAB_OVERRIDE.headerCol, 1, headerColCount);
  headerRange.merge()
    .setValue(PER_TAB_OVERRIDE.header)
    .setBackground(CALENDAR.colors.titleBackground)
    .setFontColor(CALENDAR.colors.titleFontColor)
    .setFontWeight("bold")
    .setFontSize(14)
    .setFontFamily(fontFamily)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // Row 2: informational pointer. We can't put a real "trigger import" control
  // here because simple onEdit triggers can't call UrlFetchApp or ui.prompt
  // (the authorization model forbids it). Instead, the presence of the
  // override band makes "Import Theme" appear in Calendar Tools regardless of
  // the showImportThemeMenu Key toggle (see onOpen).
  sheet.getRange(PER_TAB_OVERRIDE.importRow, PER_TAB_OVERRIDE.headerCol, 1, headerColCount)
    .merge()
    .setValue("Use Calendar Tools → Import Theme → apply to this tab")
    .setFontWeight("normal")
    .setFontStyle("italic")
    .setFontFamily(fontFamily)
    .setFontColor("#666666")
    .setBackground("#FAFAFA")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // Row 3: column sub-headers
  const subHeaders = [["", "Option", "Value", "", "Appearance", "Color"]];
  sheet.getRange(PER_TAB_OVERRIDE.subHeaderRow, PER_TAB_OVERRIDE.headerCol, 1, headerColCount)
    .setValues(subHeaders)
    .setFontWeight("bold")
    .setFontFamily(fontFamily)
    .setBackground("#EEEEEE")
    .setHorizontalAlignment("left");

  // Rows 4+: setup options (col I=name, col J=value, blank by default)
  if (setupOptions.length > 0) {
    const setupValues = setupOptions.map(o => [o, ""]);
    sheet.getRange(PER_TAB_OVERRIDE.dataStartRow, PER_TAB_OVERRIDE.setupNameCol, setupValues.length, 2)
      .setValues(setupValues);
  }

  // Category section header (col I "Category", col J "Category-Color")
  sheet.getRange(categoryHeaderRow, PER_TAB_OVERRIDE.setupNameCol, 1, 2)
    .setValues([["Category", "Category-Color"]])
    .setFontWeight("bold")
    .setFontFamily(fontFamily)
    .setBackground("#EEEEEE");

  // Pre-fill category names (no colors yet — blank color = inherit Key value)
  if (categoryRowCount > 0) {
    const categoryNameRows = initialCategories.map(p => [String(p[0] || "").trim(), ""]);
    sheet.getRange(categoryDataStartRow, PER_TAB_OVERRIDE.setupNameCol, categoryNameRows.length, 2)
      .setValues(categoryNameRows);
  }

  // Rows 4+: appearance options (col L=name, col M=color, blank by default)
  if (appearanceOptions.length > 0) {
    const appearanceValues = appearanceOptions.map(o => [o, ""]);
    sheet.getRange(PER_TAB_OVERRIDE.dataStartRow, PER_TAB_OVERRIDE.appearanceNameCol, appearanceValues.length, 2)
      .setValues(appearanceValues);
  }

  // Style data area
  if (dataRows > 0) {
    sheet.getRange(PER_TAB_OVERRIDE.dataStartRow, PER_TAB_OVERRIDE.headerCol, dataRows, headerColCount)
      .setFontFamily(fontFamily)
      .setFontSize(10)
      .setVerticalAlignment("middle");
    sheet.getRange(PER_TAB_OVERRIDE.dataStartRow, PER_TAB_OVERRIDE.setupNameCol, dataRows, 1)
      .setFontColor("#666666");
    sheet.getRange(PER_TAB_OVERRIDE.dataStartRow, PER_TAB_OVERRIDE.appearanceNameCol, dataRows, 1)
      .setFontColor("#666666");
    sheet.getRange(PER_TAB_OVERRIDE.dataStartRow, PER_TAB_OVERRIDE.headerCol, dataRows, 1)
      .setBackground("#F2F2F2");
    sheet.getRange(PER_TAB_OVERRIDE.dataStartRow, PER_TAB_OVERRIDE.spacerCol, dataRows, 1)
      .setBackground("#F2F2F2");
  }

  // Column widths
  sheet.setColumnWidth(PER_TAB_OVERRIDE.headerCol, 30);
  sheet.setColumnWidth(PER_TAB_OVERRIDE.setupNameCol, 180);
  sheet.setColumnWidth(PER_TAB_OVERRIDE.setupValueCol, 130);
  sheet.setColumnWidth(PER_TAB_OVERRIDE.spacerCol, 30);
  sheet.setColumnWidth(PER_TAB_OVERRIDE.appearanceNameCol, 260);
  sheet.setColumnWidth(PER_TAB_OVERRIDE.appearanceValueCol, 130);

  applyPerTabOverrideValidations_(sheet);

  // Group H–M and collapse
  try {
    sheet.getRange(1, PER_TAB_OVERRIDE.headerCol, 1, headerColCount).shiftColumnGroupDepth(1);
    const group = sheet.getColumnGroup(PER_TAB_OVERRIDE.headerCol, 1);
    if (group) group.collapse();
  } catch (err) {
    Logger.log("Could not group per-tab override columns: " + err.message);
  }
}

// Attach dropdowns to per-tab override value cells (column J). All use
// allow-invalid=true so leaving the cell blank still means "inherit Key value".
function applyPerTabOverrideValidations_(sheet) {
  if (!hasPerTabOverride_(sheet)) return;

  const setupOptions = perTabOverrideSetupOptions_();

  // Booleans → 3-state ("", "TRUE", "FALSE")
  PER_TAB_OVERRIDE_BOOLEAN_OPTIONS.forEach(opt => {
    const idx = setupOptions.indexOf(opt);
    if (idx < 0) return;
    setDropdownValidation(
      sheet.getRange(PER_TAB_OVERRIDE.dataStartRow + idx, PER_TAB_OVERRIDE.setupValueCol),
      ["TRUE", "FALSE"],
      true
    );
  });

  // q1StartMonth → month names
  const monthIdx = setupOptions.indexOf("q1StartMonth");
  if (monthIdx >= 0) {
    setDropdownValidation(
      sheet.getRange(PER_TAB_OVERRIDE.dataStartRow + monthIdx, PER_TAB_OVERRIDE.setupValueCol),
      CALENDAR.monthOptions.slice(),
      true
    );
  }

  // startWeekOn → day names + compressed weekend
  const dayIdx = setupOptions.indexOf("startWeekOn");
  if (dayIdx >= 0) {
    setDropdownValidation(
      sheet.getRange(PER_TAB_OVERRIDE.dataStartRow + dayIdx, PER_TAB_OVERRIDE.setupValueCol),
      [
        "Sunday", "Monday", "Tuesday", "Wednesday",
        "Thursday", "Friday", "Saturday",
        "Monday-CompressedWeekend"
      ],
      true
    );
  }

  // Source-header dropdowns for the columns that name a source field.
  const headers = readDefaultDataSheetHeaders_(sheet.getParent());
  [
    "customDate",
    "customTitle",
    "customType",
    "customCategory",
    "customStatus",
    "customAdditional",
  ].forEach(opt => {
    const idx = setupOptions.indexOf(opt);
    if (idx < 0) return;
    const cell = sheet.getRange(PER_TAB_OVERRIDE.dataStartRow + idx, PER_TAB_OVERRIDE.setupValueCol);
    if (headers.length) {
      setDropdownValidation(cell, headers, true);
    } else {
      cell.clearDataValidations();
    }
  });
}

// When a hex color is pasted into a color cell of the Theme Override band,
// apply it as the cell background. Affects: column M (appearance) anywhere
// from data row 4 down, and column J within the Category sub-section.
function formatEditedPerTabOverrideCell_(range) {
  const sheet = range.getSheet();
  if (!hasPerTabOverride_(sheet)) return;

  const col = range.getColumn();
  const row = range.getRow();

  const isAppearanceColor =
    col === PER_TAB_OVERRIDE.appearanceValueCol && row >= PER_TAB_OVERRIDE.dataStartRow;
  const isCategoryColor =
    col === PER_TAB_OVERRIDE.setupValueCol && row >= perTabOverrideCategoryDataStartRow_();

  if (!isAppearanceColor && !isCategoryColor) return;

  const color = String(range.getValue() || "").trim();
  if (looksLikeColor_(color)) {
    range.setBackground(color).setFontColor(readableTextColor_(color));
  } else {
    range.setBackground(null).setFontColor(null);
  }
}

// Reads non-blank override values from a calendar tab's H–M override section.
// Returns { setup, colors, categoryColors }. Empty maps if the tab has no band
// or has it with all blanks.
function readPerTabOverrides_(sheet) {
  const out = { setup: {}, colors: {}, categoryColors: {} };
  if (!hasPerTabOverride_(sheet)) return out;

  const setupOptions = perTabOverrideSetupOptions_();
  if (setupOptions.length > 0) {
    const values = sheet.getRange(
      PER_TAB_OVERRIDE.dataStartRow,
      PER_TAB_OVERRIDE.setupNameCol,
      setupOptions.length,
      2
    ).getValues();
    values.forEach(row => {
      const name = String(row[0] || "").trim();
      const rawValue = row[1];
      const hasValue = rawValue !== "" && rawValue !== null && rawValue !== undefined;
      if (name && hasValue) out.setup[name] = rawValue;
    });
  }

  // Category section: scan a window past the header row. Stop reading once we
  // get past the last row with content. Tolerates blank rows in between (a
  // user might have deleted a category row but kept the structure).
  const categoryDataStartRow = perTabOverrideCategoryDataStartRow_();
  const scanLastRow = Math.max(0, sheet.getLastRow() - categoryDataStartRow + 1);
  const maxScan = Math.min(50, scanLastRow);
  if (maxScan > 0) {
    const values = sheet.getRange(
      categoryDataStartRow,
      PER_TAB_OVERRIDE.setupNameCol,
      maxScan,
      2
    ).getValues();
    values.forEach(row => {
      const name = String(row[0] || "").trim();
      if (!name) return;
      const color = String(row[1] || "").trim();
      if (!color || !looksLikeColor_(color)) return;
      out.categoryColors[normalizeKey_(name)] = color;
    });
  }

  const appearanceOptions = KEY_APPEARANCE_OPTIONS;
  if (appearanceOptions.length > 0) {
    const values = sheet.getRange(
      PER_TAB_OVERRIDE.dataStartRow,
      PER_TAB_OVERRIDE.appearanceNameCol,
      appearanceOptions.length,
      2
    ).getValues();
    values.forEach(row => {
      const name = String(row[0] || "").trim();
      const rawValue = String(row[1] || "").trim();
      if (name && rawValue && looksLikeColor_(rawValue)) out.colors[name] = rawValue;
    });
  }

  return out;
}

// Wraps `fn` so any non-blank per-tab override on `sheet` is layered on top of
// CALENDAR.setup / CALENDAR.colors for the duration of the call, then restored.
// The overrides object (with categoryColors) is also passed to `fn` so the
// caller can layer category color overrides onto keyConfig.
// No-op (just calls fn) if the tab has no override band or all values blank.
function withPerTabOverridesApplied_(sheet, fn) {
  const overrides = readPerTabOverrides_(sheet);
  const setupKeys = Object.keys(overrides.setup);
  const colorKeys = Object.keys(overrides.colors);
  const categoryKeys = Object.keys(overrides.categoryColors);
  if (setupKeys.length === 0 && colorKeys.length === 0 && categoryKeys.length === 0) {
    return fn(overrides);
  }

  const setupSnapshot = {};
  Object.keys(CALENDAR.setup).forEach(k => { setupSnapshot[k] = CALENDAR.setup[k]; });
  const colorsSnapshot = {};
  Object.keys(CALENDAR.colors).forEach(k => { colorsSnapshot[k] = CALENDAR.colors[k]; });

  setupKeys.forEach(k => {
    if (Object.prototype.hasOwnProperty.call(CALENDAR.setup, k)) {
      CALENDAR.setup[k] = coerceKeyOverrideValue_(overrides.setup[k], CALENDAR.setup[k]);
    }
  });
  colorKeys.forEach(k => {
    if (Object.prototype.hasOwnProperty.call(CALENDAR.colors, k)) {
      CALENDAR.colors[k] = overrides.colors[k];
    }
  });

  try {
    return fn(overrides);
  } finally {
    Object.keys(setupSnapshot).forEach(k => { CALENDAR.setup[k] = setupSnapshot[k]; });
    Object.keys(colorsSnapshot).forEach(k => { CALENDAR.colors[k] = colorsSnapshot[k]; });
  }
}

// === End per-tab override =================================================

// === Auto-Refresh =========================================================
// Lazy refresh: when the user switches to a calendar tab, if any source-list
// edits have happened since that calendar was last refreshed, re-render it.
//
// Controlled by CALENDAR.autoRefresh (a Key-tab boolean, default TRUE).
//
// State (DocumentProperties / UserProperties so it survives across Apps Script
// executions — each onEdit / onSelectionChange runs in a fresh context):
//
//   DocumentProperties (shared across users):
//     autoRefreshLastSourceChangeAt    unix ms — last source edit
//     autoRefreshAt:<sheetId>          unix ms — last refresh of that calendar
//
//   UserProperties (per-user, so two editors don't fight):
//     autoRefreshLastActiveSheet       sheet id of the user's previous tab,
//                                      used by onSelectionChange to detect
//                                      a sheet change vs. a cell move.

const AUTO_REFRESH = {
  lastSourceChangeKey: "autoRefreshLastSourceChangeAt",
  perCalendarPrefix: "autoRefreshAt:",
  lastActiveSheetKey: "autoRefreshLastActiveSheet",
};

// Read the Key-tab `autoRefresh` toggle (CALENDAR.autoRefresh after
// applyKeyOverrides_ has run). Default TRUE in script; only an explicit
// FALSE in the Key tab turns it off.
function isAutoRefreshEnabled_() {
  return CALENDAR.autoRefresh !== false;
}

function calendarRefreshKey_(sheet) {
  return AUTO_REFRESH.perCalendarPrefix + sheet.getSheetId();
}

// Stamp a calendar as "just refreshed" so subsequent isCalendarStale_ checks
// know it's caught up. Called at the end of renderCalendarSheet.
function markCalendarRefreshed_(sheet) {
  try {
    PropertiesService.getDocumentProperties()
      .setProperty(calendarRefreshKey_(sheet), String(Date.now()));
  } catch (err) {}
}

function isCalendarStale_(sheet) {
  try {
    const docProps = PropertiesService.getDocumentProperties();
    const lastChangeAt = Number(docProps.getProperty(AUTO_REFRESH.lastSourceChangeKey) || 0);
    if (!lastChangeAt) return false; // No edits recorded yet
    const lastRefreshAt = Number(docProps.getProperty(calendarRefreshKey_(sheet)) || 0);
    return lastRefreshAt < lastChangeAt;
  } catch (err) {
    return false;
  }
}

// Kept so any installable trigger created by the previous v13.12.0 menu-based
// design keeps working harmlessly. New installs don't create it; existing
// users with a stale trigger can delete it from the Apps Script editor.
function scheduledAutoRefresh_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try { applyKeyOverrides_(ss); } catch (err) { return; }
  if (!isAutoRefreshEnabled_()) return;
  try {
    refreshAllCalendars({ silent: true });
  } catch (err) {
    Logger.log("scheduledAutoRefresh_ failed: " + err.message);
  }
}

// onEdit hook for edits NOT on the Key tab or a calendar tab — possibly the
// source data tab. Sets a shared "last source change" timestamp so that the
// next time anyone switches to a calendar tab, it knows whether to refresh.
function handleSourceEdit_(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const ss = sheet.getParent();
  const sheetName = sheet.getName();

  try { applyKeyOverrides_(ss); } catch (err) { return; }
  if (!isAutoRefreshEnabled_()) return;

  if (!isSourceDataSheet_(sheet)) return;

  let headers = [];
  try { headers = readHeaders_(sheet); } catch (err) {}
  if (!isRelevantSourceEdit_(e.range, headers)) return;

  try {
    PropertiesService.getDocumentProperties()
      .setProperty(AUTO_REFRESH.lastSourceChangeKey, String(Date.now()));
    Logger.log("[AR] stamped lastSourceChangeAt for edit on '" + sheetName + "' col " + e.range.getColumn());
  } catch (err) {}
}

// A sheet counts as "source data" if it matches the Key-configured default
// OR if any existing calendar tab's G1 (Source Data dropdown) names it.
// The second branch is essential for spreadsheets where the Key still has
// the script default ("Events") but the calendars actually point at a
// differently-named tab.
function isSourceDataSheet_(sheet) {
  const name = sheet.getName();
  if (!name) return false;
  if (name === CALENDAR.defaultDataSheetName) return true;

  const ss = sheet.getParent();
  try {
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      if (!isCalendarSheet(sheets[i])) continue;
      const g1 = String(sheets[i].getRange("G1").getDisplayValue() || "").trim();
      if (!g1) continue;
      // G1 may be multi-select (comma/semicolon separated) — match membership
      // rather than the whole string.
      const tokens = parseFilterValues_(g1);
      if (tokens.indexOf(name) >= 0) return true;
    }
  } catch (err) {}
  return false;
}

// Filter: only count edits to columns the renderer actually consumes. Other
// columns (notes, side data) don't invalidate the calendar.
function isRelevantSourceEdit_(range, headers) {
  if (!headers || !headers.length) return false;

  const relevant = [];
  if (CALENDAR.setup.customDate) relevant.push(CALENDAR.setup.customDate);
  if (CALENDAR.setup.customTitle) relevant.push(CALENDAR.setup.customTitle);
  if (CALENDAR.setup.customType) relevant.push(CALENDAR.setup.customType);
  if (CALENDAR.setup.customCategory) relevant.push(CALENDAR.setup.customCategory);
  if (CALENDAR.setup.customStatus) relevant.push(CALENDAR.setup.customStatus);
  relevant.push("Date", "Start Date", "Event Date", "Date/Time", "Date Time", "MMDD");
  relevant.push("Title", "Event Title", "Name", "Initiative", "Topic");
  relevant.push("Type", "Event Type", "Channel", "Tactic Type");
  relevant.push("Category", "Event Category", "Theme", "Product", "Pillar");
  relevant.push("Status", "Event Status");
  relevant.push("Description", "Desc", "Event Description", "Details");
  (CALENDAR.setup.customAdditional || []).forEach(f => {
    const name = String(f || "").split(",")[0].trim();
    if (name) relevant.push(name);
  });

  const normRelevant = relevant.map(normalizeHeader_).filter(n => n);
  const startCol = range.getColumn();
  const numCols = range.getNumColumns();

  for (let i = 0; i < numCols; i++) {
    const idx = startCol - 1 + i;
    if (idx < 0 || idx >= headers.length) continue;
    const headerName = String(headers[idx] || "").trim();
    if (!headerName) continue;
    if (normRelevant.indexOf(normalizeHeader_(headerName)) >= 0) return true;
  }
  return false;
}

// === End Auto-Refresh =====================================================

function newCalendarSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  applyKeyOverrides_(ss);
  const period = currentMonthName();
  const year = currentYear();
  const sheet = ss.insertSheet(uniqueSheetName(ss, CALENDAR.calendarBaseName));

  initializeCalendarSheet(sheet, period, year, getDefaultSourceSpec(ss));
  renderCalendarSheet(sheet);

  ss.setActiveSheet(sheet);
  ss.toast("New calendar sheet created.", CALENDAR.menuName, 4);
}

// Used by Add Q1-Q4 and Add Jan-Dec, which name tabs after their period.
// New Calendar Sheet intentionally uses the generic base name so the user can
// pick a period after creation.
function buildCalendarSheetName_(period, year) {
  const p = String(period || "").trim();
  const y = String(year || "").trim();

  if (!p) return CALENDAR.calendarBaseName;
  if (isCustom(p)) return CALENDAR.calendarBaseName;
  if (isYear(p)) return y ? ("Year " + y) : "Year";
  return y ? (p + " " + y) : p;
}

function replaceWithCalendar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const isDataSheet = sheet.getName() === CALENDAR.defaultDataSheetName;
  const hasContent = sheetHasContent(sheet);

  if (isDataSheet) {
    const response = SpreadsheetApp.getUi().alert(
      "Warning",
      'You are about to overwrite the sheet named "Data". If that sheet is live, add a new calendar tab instead.',
      SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
    );
    if (response !== SpreadsheetApp.getUi().Button.OK) return;
  } else if (hasContent) {
    const response = SpreadsheetApp.getUi().alert(
      "Replace current sheet?",
      "This sheet already has content. Replacing it will clear the current tab and draw the calendar on top.",
      SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
    );
    if (response !== SpreadsheetApp.getUi().Button.OK) return;
  }

  const controls = readControls(sheet);

  showWorkingModal("Generating calendar...");
  initializeCalendarSheet(sheet, controls.period, controls.year, controls.sourceSpec);
  renderCalendarSheet(sheet, controls);

  ss.setActiveSheet(sheet);
  ss.toast("Calendar replaced.", CALENDAR.menuName, 4);
}

function refreshAllCalendars(opts) {
  opts = opts || {};
  const silent = !!opts.silent;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const calendarSheets = ss.getSheets().filter(isCalendarSheet);

  if (!calendarSheets.length) {
    if (!silent) ss.toast("No calendar sheets found.", CALENDAR.menuName, 4);
    return;
  }

  applyKeyOverrides_(ss);
  if (!silent) showWorkingModal("Refreshing all calendars...");

  let refreshed = 0;
  const failures = [];

  withRenderSession_(() => {
    calendarSheets.forEach((sheet) => {
      try {
        const controls = getLiveControls_(sheet);
        renderCalendarSheet(sheet, controls);
        refreshed++;
      } catch (err) {
        failures.push(sheet.getName() + ": " + err.message);
      }
    });
  });

  try {
    if (activeSheet) ss.setActiveSheet(activeSheet);
  } catch (err) {}

  if (!silent) {
    if (failures.length) {
      ss.toast("Refreshed " + refreshed + " calendar(s). " + failures.length + " failed.", CALENDAR.menuName, 6);
    } else {
      ss.toast("Refreshed " + refreshed + " calendar(s).", CALENDAR.menuName, 4);
    }
  }
}

function openSelectedMore() {
  openSelected();
}

function openSelected() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Apps Script runs each menu click in a fresh JS context. Apply the user's
  // Key-tab overrides so grid math (startWeekOn, maxEvents) matches what was
  // rendered, otherwise the date computed from the selected cell can be wrong.
  applyKeyOverrides_(ss);
  const sheet = ss.getActiveSheet();
  const range = sheet.getActiveRange();

  if (!range || range.getNumRows() !== 1 || range.getNumColumns() !== 1) {
    ss.toast("Select a single calendar day cell first.", CALENDAR.menuName, 4);
    return;
  }

  if (!isCalendarSheet(sheet)) {
    ss.toast("Select a cell on a calendar sheet.", CALENDAR.menuName, 4);
    return;
  }

  const dateKeyValue = getSelectedDateKeyFromRange_(range);

  if (!dateKeyValue) {
    ss.toast("Select a date, event, or More cell on the calendar.", CALENDAR.menuName, 4);
    return;
  }

  showOverflowModal_(sheet, dateKeyValue);
}

function getSelectedDateKeyFromGridPosition_(range) {
  const sheet = range.getSheet();
  const row = range.getRow();
  const col = range.getColumn();

  if (col < 1 || col > CALENDAR.layout.columns) return "";

  const controls = getLiveControls_(sheet);
  const renderControls = normalizeControlsForRender_(controls);
  const startDate = getStartDateForControls_(controls);
  const monthCount = getDisplayedMonthCount(renderControls.period, controls.monthCount);
  const monthInfos = buildMonthInfos(renderControls, startDate, monthCount);

  const maxEvents = getMaxEvents_();
  const rowsPerWeek = 1 + maxEvents + 1;
  const monthBlockRows = getMonthBlockRows_();
  const weekStartIndex = getWeekStartIndex_();

  for (let i = 0; i < monthInfos.length; i++) {
    const monthStartRow = CALENDAR.layout.firstMonthRow + (i * monthBlockRows);
    const bodyStartRow = monthStartRow + 3;
    const bodyEndRow = bodyStartRow + (6 * rowsPerWeek) - 1;

    if (row < bodyStartRow || row > bodyEndRow) continue;

    const rowOffset = row - bodyStartRow;
    const week = Math.floor(rowOffset / rowsPerWeek);
    const day = col - 1;

    const monthDate = monthInfos[i];
    const firstOfMonth = new Date(monthDate.getFullYear(), monthDate.getMonth(), 1);
    const gridStart = new Date(firstOfMonth);
    const leadingDays = (firstOfMonth.getDay() - weekStartIndex + 7) % 7;
    gridStart.setDate(firstOfMonth.getDate() - leadingDays);

    const selectedDate = new Date(gridStart);
    selectedDate.setDate(gridStart.getDate() + (week * 7) + day);

    return dateKey(selectedDate);
  }

  return "";
}

function getSelectedDateKeyFromRange_(range) {
  return getSelectedDateKeyFromGridPosition_(range);
}









function addQ1Q4() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  applyKeyOverrides_(ss);
  const year = currentYear();
  const sourceSpec = getDefaultSourceSpec(ss);

  showWorkingModal("Creating Q1-Q4 tabs...");

  ["Q1", "Q2", "Q3", "Q4"].forEach((period) => {
    const sheet = ss.insertSheet(uniqueSheetName(ss, buildCalendarSheetName_(period, year)));
    initializeCalendarSheet(sheet, period, year, sourceSpec);
    renderCalendarSheet(sheet);
  });

  ss.toast("Q1-Q4 tabs created.", CALENDAR.menuName, 4);
}

function addJanDec() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  applyKeyOverrides_(ss);
  const year = currentYear();
  const sourceSpec = getDefaultSourceSpec(ss);

  showWorkingModal("Creating Jan-Dec tabs...");

  CALENDAR.monthOptions.forEach((period) => {
    const sheet = ss.insertSheet(uniqueSheetName(ss, buildCalendarSheetName_(period, year)));
    initializeCalendarSheet(sheet, period, year, sourceSpec);
    renderCalendarSheet(sheet);
  });

  ss.toast("Jan-Dec tabs created.", CALENDAR.menuName, 4);
}

function createKeySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CALENDAR.keySheetName);

  if (!sheet) {
    sheet = ss.insertSheet(CALENDAR.keySheetName);
  }

  buildKeySheet_(sheet);
  ss.setActiveSheet(sheet);
  ss.toast("Key sheet created.", CALENDAR.menuName, 4);
}

// True if the Key tab is missing any setup or appearance options that the
// current script version knows about. Used to conditionally show the
// "Update Key" menu item — fast short-circuit check.
function keyHasMissingFeatures_(ss) {
  const sheet = ss.getSheetByName(CALENDAR.keySheetName);
  if (!sheet) return false;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return KEY_SETUP_OPTIONS.length > 0 || KEY_APPEARANCE_OPTIONS.length > 0;

  // Setup column = K (col 11). Appearance column = N (col 14).
  let existingSetup = [];
  let existingAppearance = [];
  try {
    existingSetup = sheet.getRange(2, 11, lastRow - 1, 1).getValues().map(r => String(r[0] || "").trim());
    existingAppearance = sheet.getRange(2, 14, lastRow - 1, 1).getValues().map(r => String(r[0] || "").trim());
  } catch (err) {
    return false;
  }

  for (let i = 0; i < KEY_SETUP_OPTIONS.length; i++) {
    if (existingSetup.indexOf(KEY_SETUP_OPTIONS[i]) < 0) return true;
  }
  for (let i = 0; i < KEY_APPEARANCE_OPTIONS.length; i++) {
    if (existingAppearance.indexOf(KEY_APPEARANCE_OPTIONS[i]) < 0) return true;
  }
  return false;
}

// Returns { setup: [missingNames], appearance: [missingNames] }, preserving
// the canonical order from KEY_SETUP_OPTIONS / KEY_APPEARANCE_OPTIONS.
function keyMissingOptions_(ss) {
  const out = { setup: [], appearance: [] };
  const sheet = ss.getSheetByName(CALENDAR.keySheetName);
  if (!sheet) {
    out.setup = KEY_SETUP_OPTIONS.slice();
    out.appearance = KEY_APPEARANCE_OPTIONS.slice();
    return out;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    out.setup = KEY_SETUP_OPTIONS.slice();
    out.appearance = KEY_APPEARANCE_OPTIONS.slice();
    return out;
  }

  let existingSetup = [];
  let existingAppearance = [];
  try {
    existingSetup = sheet.getRange(2, 11, lastRow - 1, 1).getValues().map(r => String(r[0] || "").trim());
    existingAppearance = sheet.getRange(2, 14, lastRow - 1, 1).getValues().map(r => String(r[0] || "").trim());
  } catch (err) {}

  out.setup = KEY_SETUP_OPTIONS.filter(o => existingSetup.indexOf(o) < 0);
  out.appearance = KEY_APPEARANCE_OPTIONS.filter(o => existingAppearance.indexOf(o) < 0);
  return out;
}

// Appends any newly-introduced Key setup or appearance options to the existing
// Key tab without touching values the user has already set. Each new row uses
// the same default-value resolution as buildKeySetupRows_ (KEY_INITIAL_VALUES
// → CALENDAR.setup → CALENDAR), and boolean options get checkboxes.
function updateKey() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName(CALENDAR.keySheetName);
  if (!sheet) {
    ui.alert("Update Key", "No Key tab found. Use Create Key (and customize) first.", ui.ButtonSet.OK);
    return;
  }

  const missing = keyMissingOptions_(ss);
  if (missing.setup.length === 0 && missing.appearance.length === 0) {
    ui.alert("Update Key", "Key already has every current option.", ui.ButtonSet.OK);
    return;
  }

  const fontFamily = CALENDAR.setup.fontFamily || "Inter";
  let added = { setup: 0, appearance: 0 };

  // Append missing setup rows after the last existing setup row in col K.
  if (missing.setup.length > 0) {
    let lastSetupRow = 1;
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const values = sheet.getRange(2, 11, lastRow - 1, 1).getValues();
      for (let i = 0; i < values.length; i++) {
        if (String(values[i][0] || "").trim()) lastSetupRow = i + 2;
      }
    }

    const rows = missing.setup.map(opt => {
      let value = "";
      if (Object.prototype.hasOwnProperty.call(KEY_INITIAL_VALUES, opt)) {
        value = KEY_INITIAL_VALUES[opt];
      } else if (Object.prototype.hasOwnProperty.call(CALENDAR.setup, opt)) {
        value = CALENDAR.setup[opt];
      } else if (Object.prototype.hasOwnProperty.call(CALENDAR, opt)) {
        value = CALENDAR[opt];
      }
      return ["", opt, serializeKeyValue_(value)];
    });

    const needRows = lastSetupRow + rows.length;
    if (sheet.getMaxRows() < needRows) {
      sheet.insertRowsAfter(sheet.getMaxRows(), needRows - sheet.getMaxRows());
    }

    sheet.getRange(lastSetupRow + 1, 10, rows.length, 3).setValues(rows);
    sheet.getRange(lastSetupRow + 1, 10, rows.length, 3).setFontFamily(fontFamily);
    added.setup = rows.length;

    // Checkboxes for boolean options.
    missing.setup.forEach((opt, i) => {
      const row = lastSetupRow + 1 + i;
      const looksBoolean =
        typeof CALENDAR.setup[opt] === "boolean" ||
        typeof CALENDAR[opt] === "boolean" ||
        (Object.prototype.hasOwnProperty.call(KEY_INITIAL_VALUES, opt) &&
          typeof KEY_INITIAL_VALUES[opt] === "boolean");
      if (looksBoolean) sheet.getRange(row, 12).insertCheckboxes();
    });
  }

  // Append missing appearance rows after the last existing appearance row in col N.
  if (missing.appearance.length > 0) {
    let lastAppRow = 1;
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const values = sheet.getRange(2, 14, lastRow - 1, 1).getValues();
      for (let i = 0; i < values.length; i++) {
        if (String(values[i][0] || "").trim()) lastAppRow = i + 2;
      }
    }

    const rows = missing.appearance.map(opt => {
      const v = CALENDAR.colors[opt];
      return [opt, v || ""];
    });

    const needRows = lastAppRow + rows.length;
    if (sheet.getMaxRows() < needRows) {
      sheet.insertRowsAfter(sheet.getMaxRows(), needRows - sheet.getMaxRows());
    }

    sheet.getRange(lastAppRow + 1, 14, rows.length, 2).setValues(rows);
    sheet.getRange(lastAppRow + 1, 14, rows.length, 2).setFontFamily(fontFamily);
    rows.forEach((row, i) => {
      const color = String(row[1] || "").trim();
      if (looksLikeColor_(color)) {
        sheet.getRange(lastAppRow + 1 + i, 15)
          .setBackground(color)
          .setFontColor(readableTextColor_(color));
      }
    });
    added.appearance = rows.length;
  }

  // Refresh validations + dropdowns for new (and existing) rows.
  applyKeyOptionNotes_(sheet);
  applyKeySetupValidations_(sheet);

  ss.toast(
    "Key updated: " + added.setup + " setup option(s), " +
      added.appearance + " appearance option(s) added.",
    CALENDAR.menuName,
    6
  );
}

// One-shot first-time setup: ensures both the Event List tab and the Key tab
// exist, then runs the Key Configurator so Type/Category/Status get
// validation dropdowns and Category-based row colors. Safe to re-run; existing
// tabs are left in place.
function initialSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  applyKeyOverrides_(ss);
  showWorkingModal("Running initial setup...");

  const eventsCreated = !ss.getSheetByName(CALENDAR.defaultDataSheetName);
  const eventsSheet = ensureEventsSheet_(ss);

  const keyCreated = !ss.getSheetByName(CALENDAR.keySheetName);
  let keySheet = ss.getSheetByName(CALENDAR.keySheetName);
  if (!keySheet) {
    keySheet = ss.insertSheet(CALENDAR.keySheetName);
    buildKeySheet_(keySheet);
  }

  try {
    runKeyConfigurator();
  } catch (err) {
    Logger.log("Key configurator skipped during initial setup: " + err.message);
  }

  ss.setActiveSheet(eventsSheet);

  const parts = [];
  if (eventsCreated) parts.push('created "' + CALENDAR.defaultDataSheetName + '"');
  if (keyCreated) parts.push('created "' + CALENDAR.keySheetName + '"');
  parts.push("ran key configurator");
  ss.toast("Initial setup: " + parts.join(", ") + ".", CALENDAR.menuName, 6);
}

function createEventList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  applyKeyOverrides_(ss);

  const existed = !!ss.getSheetByName(CALENDAR.defaultDataSheetName);
  const sheet = ensureEventsSheet_(ss);
  ss.setActiveSheet(sheet);

  // If the Key already exists, wire up validation + Category-based row colors
  // right away so the event list is ready to use.
  let configuratorRan = false;
  if (ss.getSheetByName(CALENDAR.keySheetName)) {
    try {
      runKeyConfigurator();
      configuratorRan = true;
    } catch (err) {
      Logger.log("Key configurator skipped during createEventList: " + err.message);
    }
  }

  const parts = [];
  parts.push(existed
    ? '"' + CALENDAR.defaultDataSheetName + '" already exists'
    : '"' + CALENDAR.defaultDataSheetName + '" created');
  if (configuratorRan) parts.push("ran key configurator");
  ss.toast(parts.join(", ") + ".", CALENDAR.menuName, 5);
}

function ensureEventsSheet_(ss) {
  const existing = ss.getSheetByName(CALENDAR.defaultDataSheetName);
  if (existing) return existing;

  const sheet = ss.insertSheet(CALENDAR.defaultDataSheetName);
  buildEventsSheet_(sheet);
  return sheet;
}

// Interactive flow: pick an event list, then pick which columns/values to use
// for Date, Title, Type, Category, Status. Writes the result into the Key tab.
// Each prompt accepts either a number from the offered list or a custom name.
function setKeyFromEventList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  applyKeyOverrides_(ss);

  const eventSheet = pickEventListSheet_(ui, ss);
  if (!eventSheet) return;

  const headers = readHeaders_(eventSheet);
  if (!headers.length) {
    ui.alert("No headers found in row 1 of \"" + eventSheet.getName() + "\".");
    return;
  }

  const dateColumn = pickHeaderOrCustom_(ui, headers, "Date column",
    "Which column has event dates?", detectHeader_(headers, [CALENDAR.setup.customDate, "Date", "Start Date", "Event Date"]));
  if (dateColumn === null) return;

  const titleColumn = pickHeaderOrCustom_(ui, headers, "Title column",
    "Which column has event titles?", detectHeader_(headers, [CALENDAR.setup.customTitle, "Title", "Name", "Event Title"]));
  if (titleColumn === null) return;

  const typeColumn = pickHeaderOrCustom_(ui, headers, "Type column",
    "Which column has event types? (leave blank to skip)", detectHeader_(headers, ["Type", "Event Type", "Channel", "Tactic Type"]), true);
  if (typeColumn === false) return;
  const types = typeColumn
    ? pickValuesOrCustom_(ui, eventSheet, headers, typeColumn, "Type values")
    : [];
  if (types === null) return;

  const categoryColumn = pickHeaderOrCustom_(ui, headers, "Category column",
    "Which column has event categories? (leave blank to skip)", detectHeader_(headers, ["Category", "Event Category", "Theme", "Product", "Pillar"]), true);
  if (categoryColumn === false) return;
  const categories = categoryColumn
    ? pickValuesOrCustom_(ui, eventSheet, headers, categoryColumn, "Category values")
    : [];
  if (categories === null) return;

  const statusColumn = pickHeaderOrCustom_(ui, headers, "Status column",
    "Which column has event statuses? (leave blank to skip)", detectHeader_(headers, ["Status", "Event Status"]), true);
  if (statusColumn === false) return;
  const statuses = statusColumn
    ? pickValuesOrCustom_(ui, eventSheet, headers, statusColumn, "Status values")
    : [];
  if (statuses === null) return;

  const summary =
    "About to update the Key tab from \"" + eventSheet.getName() + "\":\n\n" +
    "• Date column: " + dateColumn + "\n" +
    "• Title column: " + titleColumn + "\n" +
    "• Types (" + types.length + "): " + (types.join(", ") || "(none)") + "\n" +
    "• Categories (" + categories.length + "): " + (categories.join(", ") || "(none)") + "\n" +
    "• Statuses (" + statuses.length + "): " + (statuses.join(", ") || "(none)") + "\n\n" +
    "Continue?";

  const confirm = ui.alert("Set Key From Event List", summary, ui.ButtonSet.OK_CANCEL);
  if (confirm !== ui.Button.OK) return;

  let keySheet = ss.getSheetByName(CALENDAR.keySheetName);
  if (!keySheet) {
    keySheet = ss.insertSheet(CALENDAR.keySheetName);
    buildKeySheet_(keySheet);
  }

  applyKeyFromEventList_(keySheet, {
    eventListName: eventSheet.getName(),
    dateColumn,
    titleColumn,
    types,
    categories,
    statuses
  });
  ss.toast("Key updated from \"" + eventSheet.getName() + "\".", CALENDAR.menuName, 5);
}

function pickEventListSheet_(ui, ss) {
  const sheets = ss.getSheets().filter(s => s.getName() !== CALENDAR.keySheetName && !isCalendarSheet(s));
  if (!sheets.length) {
    ui.alert("No eligible event list tabs found. Create one with \"Create Event List\" first.");
    return null;
  }

  const choices = sheets.map((s, i) => (i + 1) + ". " + s.getName()).join("\n");
  const response = ui.prompt(
    "Set Key From Event List",
    "Pick an event list tab.\n\nEnter the number, or type a tab name:\n\n" + choices,
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return null;

  const text = String(response.getResponseText() || "").trim();
  if (!text) return null;

  const asNum = Number(text);
  if (Number.isInteger(asNum) && asNum >= 1 && asNum <= sheets.length) {
    return sheets[asNum - 1];
  }
  const byName = ss.getSheetByName(text);
  if (byName) return byName;
  ui.alert("Could not find a tab named \"" + text + "\".");
  return null;
}

function readHeaders_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(v => String(v || "").trim())
    .filter(v => v !== "");
}

function detectHeader_(headers, candidates) {
  const norm = headers.map(normalizeHeader_);
  for (const cand of candidates) {
    const i = norm.indexOf(normalizeHeader_(cand));
    if (i >= 0) return headers[i];
  }
  return "";
}

// Returns: a header name (string), null on cancel, or false when allowBlank
// is true and the user submitted nothing. Distinguishing "" / null / false
// lets the caller tell skip-this-column apart from cancel-the-flow.
function pickHeaderOrCustom_(ui, headers, title, message, suggested, allowBlank) {
  const choices = headers.map((h, i) => (i + 1) + ". " + h).join("\n");
  const suggestion = suggested ? "\n\nDetected: " + suggested : "";
  const skipHint = allowBlank ? "\n\nLeave blank to skip." : "";

  const response = ui.prompt(
    title,
    message + suggestion + skipHint + "\n\nEnter a number, type a custom column name, or press OK to accept the detected value.\n\n" + choices,
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return null;

  const text = String(response.getResponseText() || "").trim();

  if (!text) {
    if (allowBlank && !suggested) return false;
    if (suggested) return suggested;
    ui.alert("A value is required.");
    return pickHeaderOrCustom_(ui, headers, title, message, suggested, allowBlank);
  }

  const asNum = Number(text);
  if (Number.isInteger(asNum) && asNum >= 1 && asNum <= headers.length) {
    return headers[asNum - 1];
  }
  return text;
}

// Returns: string[] of selected values, or null on cancel.
// User input: "all" / comma-separated numbers / comma-separated custom names /
// mixed (numbers become indexed values, non-numeric tokens become custom names).
function pickValuesOrCustom_(ui, sheet, headers, columnName, title) {
  const colIndex = findColumnIndex(headers, [columnName]);
  let values = [];
  if (colIndex >= 0) {
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const rows = sheet.getRange(2, colIndex + 1, lastRow - 1, 1).getValues();
      values = uniqueValuesFromColumn_(rows, 0);
    }
  }

  const choices = values.length
    ? values.map((v, i) => (i + 1) + ". " + v).join("\n")
    : "(no values found in column \"" + columnName + "\")";

  const response = ui.prompt(
    title,
    "Pick values for the Key.\n\n" +
    'Enter "all", comma-separated numbers (e.g. 1,3,5), or comma-separated custom names.\n' +
    "Leave blank to skip.\n\n" + choices,
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return null;

  const text = String(response.getResponseText() || "").trim();
  if (!text) return [];

  if (text.toLowerCase() === "all") return values.slice();

  const tokens = text.split(",").map(t => t.trim()).filter(t => t !== "");
  const result = [];
  const seen = Object.create(null);
  tokens.forEach(token => {
    const n = Number(token);
    let pick = "";
    if (Number.isInteger(n) && n >= 1 && n <= values.length) {
      pick = values[n - 1];
    } else {
      pick = token;
    }
    const key = normalizeKey_(pick);
    if (!seen[key]) {
      seen[key] = true;
      result.push(pick);
    }
  });
  return result;
}

function applyKeyFromEventList_(keySheet, config) {
  const maxRows = Math.max(keySheet.getMaxRows() - 1, 0);

  if (config.types.length) {
    const iconMap = pairsToMap_(CALENDAR.keyDefaults.types);
    const rows = config.types.map(t => [t, iconMap[normalizeKey_(t)] || ""]);
    if (maxRows > 0) keySheet.getRange(2, 1, maxRows, 2).clearContent();
    keySheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }

  if (config.categories.length) {
    const colorMap = pairsToMap_(CALENDAR.keyDefaults.categories);
    const palette = CALENDAR.keyDefaults.categories.map(p => p[1]);
    const rows = config.categories.map((cat, i) => {
      const color = colorMap[normalizeKey_(cat)] || palette[i % Math.max(palette.length, 1)] || "#FFFFFF";
      return [cat, color];
    });
    if (maxRows > 0) {
      keySheet.getRange(2, 4, maxRows, 2)
        .clearContent()
        .setBackground(null)
        .setFontColor(null);
    }
    keySheet.getRange(2, 4, rows.length, 2).setValues(rows);
    rows.forEach((row, i) => {
      const color = row[1];
      if (looksLikeColor_(color)) {
        keySheet.getRange(i + 2, 4, 1, 2)
          .setBackground(color)
          .setFontColor(readableTextColor_(color));
      }
    });
  }

  if (config.statuses.length) {
    const iconMap = pairsToMap_(CALENDAR.keyDefaults.statuses);
    const rows = config.statuses.map(s => [s, iconMap[normalizeKey_(s)] || ""]);
    if (maxRows > 0) keySheet.getRange(2, 7, maxRows, 2).clearContent();
    keySheet.getRange(2, 7, rows.length, 2).setValues(rows);
  }

  if (config.eventListName) {
    updateKeyOptionValue_(keySheet, "defaultDataSheetName", config.eventListName);
    // Update the in-memory CALENDAR.defaultDataSheetName so the customDate /
    // customTitle dropdowns below populate from the right tab.
    applyKeyOverrides_(keySheet.getParent());
  }
  if (config.dateColumn) updateKeyOptionValue_(keySheet, "customDate", config.dateColumn);
  if (config.titleColumn) updateKeyOptionValue_(keySheet, "customTitle", config.titleColumn);

  applyKeySetupValidations_(keySheet);
}

function pairsToMap_(pairs) {
  const out = {};
  (pairs || []).forEach(p => { out[normalizeKey_(p[0])] = p[1]; });
  return out;
}

function updateKeyOptionValue_(keySheet, optionName, value) {
  const row = findKeyOptionRow_(keySheet, optionName);
  if (!row) return;
  keySheet.getRange(row, 12).setValue(value);
}

// Fetch a theme manifest from the configured public repo, prompt the user to
// pick one, fetch the theme JSON, and write it into the Key tab.
function importTheme() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  ss.toast("Fetching themes...", CALENDAR.menuName, 3);
  const manifest = fetchThemeManifest_();
  if (!manifest || !manifest.length) {
    ui.alert("Could not load the theme list.\n\nTried: " + CALENDAR.themes.indexUrl + "\n\nCheck your network connection, then try again.");
    return;
  }

  const list = manifest.map((t, i) => {
    const desc = t.description ? " — " + t.description : "";
    return (i + 1) + ". " + t.name + desc;
  }).join("\n");

  const response = ui.prompt(
    "Import Theme",
    "Pick a theme.\n\nEnter the number, or type the exact theme name:\n\n" + list,
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const text = String(response.getResponseText() || "").trim();
  if (!text) return;

  let picked = null;
  const asNum = Number(text);
  if (Number.isInteger(asNum) && asNum >= 1 && asNum <= manifest.length) {
    picked = manifest[asNum - 1];
  } else {
    picked = manifest.find(t => String(t.name || "").trim().toLowerCase() === text.toLowerCase());
  }
  if (!picked) {
    ui.alert('No theme matched "' + text + '".');
    return;
  }

  ss.toast('Fetching "' + picked.name + '"...', CALENDAR.menuName, 3);
  const theme = fetchTheme_(picked.file);
  if (!theme) {
    ui.alert("Could not load theme \"" + picked.name + "\" from " + CALENDAR.themes.baseUrl + picked.file);
    return;
  }

  // Target: prompt only when active sheet is a calendar tab. From the Key
  // tab or a regular data tab the only sensible target is the Key.
  const activeSheet = ss.getActiveSheet();
  let target = "key";
  if (isCalendarSheet(activeSheet)) {
    const choice = ui.alert(
      "Apply theme",
      'Apply "' + picked.name + '" to which target?\n\n' +
      '• YES — only the active calendar tab "' + activeSheet.getName() + '"\n' +
      '  (creates per-tab override columns if not present)\n\n' +
      '• NO — the Key tab (affects every calendar without its own override)\n\n' +
      '• CANCEL — abort',
      ui.ButtonSet.YES_NO_CANCEL
    );
    if (choice === ui.Button.YES) target = "tab";
    else if (choice === ui.Button.NO) target = "key";
    else return;
  }

  if (target === "tab") {
    if (!hasPerTabOverride_(activeSheet)) addPerTabOverride_(activeSheet);
    const changes = applyThemeToTabOverride_(activeSheet, theme);

    // Refresh the calendar so the theme takes effect immediately.
    let refreshed = false;
    try {
      renderCalendarSheet(activeSheet);
      refreshed = true;
    } catch (err) {
      Logger.log("Refresh after per-tab theme import failed: " + err.message);
    }

    ss.toast(
      'Theme "' + picked.name + '" applied to "' + activeSheet.getName() + '" — ' +
        themeChangeSummary_(changes) +
        (refreshed ? '. Calendar refreshed.' : '. Refresh the calendar to see it.'),
      CALENDAR.menuName,
      8
    );
    return;
  }

  let keySheet = ss.getSheetByName(CALENDAR.keySheetName);
  if (!keySheet) {
    keySheet = ss.insertSheet(CALENDAR.keySheetName);
    buildKeySheet_(keySheet);
  }

  const changes = applyThemeToKey_(keySheet, theme);
  ss.toast(
    'Theme "' + picked.name + '" applied to Key — ' +
      themeChangeSummary_(changes) + '. Refresh All Calendars to see it.',
    CALENDAR.menuName,
    8
  );
}

function themeChangeSummary_(changes) {
  const parts = [];
  parts.push(changes.colors + ' colors');
  parts.push(changes.setup + ' setup options');
  if (changes.categories) parts.push(changes.categories + ' categories');
  return parts.join(', ') + ' updated';
}

// Writes a theme into a calendar tab's per-tab override band (H–M). The band
// must already exist (caller is responsible for addPerTabOverride_).
// Skips theme.setup entries that aren't in perTabOverrideSetupOptions_(), so a
// theme can be shared with the Key-target flow even if it sets excluded
// options like showInitialMenu.
function applyThemeToTabOverride_(sheet, theme) {
  let colorChanges = 0;
  let setupChanges = 0;

  const setupOptions = perTabOverrideSetupOptions_();
  const appearanceOptions = KEY_APPEARANCE_OPTIONS;

  if (theme.colors && typeof theme.colors === "object" && appearanceOptions.length > 0) {
    const values = sheet.getRange(
      PER_TAB_OVERRIDE.dataStartRow,
      PER_TAB_OVERRIDE.appearanceNameCol,
      appearanceOptions.length,
      2
    ).getValues();
    for (let i = 0; i < values.length; i++) {
      const appearance = String(values[i][0] || "").trim();
      if (!appearance) continue;
      if (!Object.prototype.hasOwnProperty.call(theme.colors, appearance)) continue;
      const newColor = String(theme.colors[appearance] || "").trim();
      if (!looksLikeColor_(newColor)) continue;
      const row = PER_TAB_OVERRIDE.dataStartRow + i;
      sheet.getRange(row, PER_TAB_OVERRIDE.appearanceValueCol)
        .setValue(newColor)
        .setBackground(newColor)
        .setFontColor(readableTextColor_(newColor));
      colorChanges++;
    }
  }

  if (theme.setup && typeof theme.setup === "object") {
    Object.keys(theme.setup).forEach(option => {
      const idx = setupOptions.indexOf(option);
      if (idx < 0) return;
      const row = PER_TAB_OVERRIDE.dataStartRow + idx;
      sheet.getRange(row, PER_TAB_OVERRIDE.setupValueCol).setValue(theme.setup[option]);
      setupChanges++;
    });
  }

  // categoryPalette is positional: palette[0] → 1st named category in the
  // override band's Category section, palette[1] → 2nd, etc.
  let categoryChanges = 0;
  if (Array.isArray(theme.categoryPalette) && theme.categoryPalette.length > 0) {
    const categoryDataStartRow = perTabOverrideCategoryDataStartRow_();
    const lastRow = sheet.getLastRow();
    const maxScan = Math.min(50, Math.max(0, lastRow - categoryDataStartRow + 1));
    if (maxScan > 0) {
      const nameValues = sheet.getRange(
        categoryDataStartRow,
        PER_TAB_OVERRIDE.setupNameCol,
        maxScan,
        1
      ).getValues();
      let paletteIdx = 0;
      for (let i = 0; i < nameValues.length && paletteIdx < theme.categoryPalette.length; i++) {
        const name = String(nameValues[i][0] || "").trim();
        if (!name) continue;
        const color = String(theme.categoryPalette[paletteIdx] || "").trim();
        paletteIdx++;
        if (!looksLikeColor_(color)) continue;
        const row = categoryDataStartRow + i;
        sheet.getRange(row, PER_TAB_OVERRIDE.setupValueCol)
          .setValue(color)
          .setBackground(color)
          .setFontColor(readableTextColor_(color));
        categoryChanges++;
      }
    }
  }

  return { colors: colorChanges, setup: setupChanges, categories: categoryChanges };
}

// Adds the Theme Override band to the active calendar tab and pre-populates
// every cell with the currently-effective value (script default overlaid with
// Key overrides). Lets users see "what this tab is currently using" and then
// modify individual cells to override, or clear them to inherit from the Key.
function addThemeOverride() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  applyKeyOverrides_(ss);

  const sheet = ss.getActiveSheet();
  if (!isCalendarSheet(sheet)) {
    ui.alert("Add Theme Override", "Switch to a calendar tab first.", ui.ButtonSet.OK);
    return;
  }
  if (hasPerTabOverride_(sheet)) {
    ui.alert(
      "Add Theme Override",
      '"' + sheet.getName() + '" already has a Theme Override. Expand columns H–M to edit it.',
      ui.ButtonSet.OK
    );
    return;
  }

  addPerTabOverride_(sheet);
  populatePerTabOverrideFromCurrent_(sheet);

  ss.toast(
    'Theme Override added to "' + sheet.getName() + '". Expand columns H–M to edit.',
    CALENDAR.menuName,
    6
  );
}

// Pre-populate the override band with CALENDAR.setup / CALENDAR.colors values
// that are currently in effect (after applyKeyOverrides_ runs). Caller is
// responsible for ensuring the band exists and CALENDAR is up-to-date.
function populatePerTabOverrideFromCurrent_(sheet) {
  if (!hasPerTabOverride_(sheet)) return;

  const setupOptions = perTabOverrideSetupOptions_();
  if (setupOptions.length > 0) {
    const values = setupOptions.map(opt => {
      const v = Object.prototype.hasOwnProperty.call(CALENDAR.setup, opt)
        ? CALENDAR.setup[opt]
        : "";
      // Booleans use the "TRUE"/"FALSE" tokens the 3-state dropdown expects.
      if (PER_TAB_OVERRIDE_BOOLEAN_OPTIONS.indexOf(opt) >= 0) {
        return [v ? "TRUE" : "FALSE"];
      }
      return [serializeKeyValue_(v)];
    });
    sheet.getRange(
      PER_TAB_OVERRIDE.dataStartRow,
      PER_TAB_OVERRIDE.setupValueCol,
      values.length,
      1
    ).setValues(values);
  }

  const appearanceOptions = KEY_APPEARANCE_OPTIONS;
  if (appearanceOptions.length > 0) {
    const values = appearanceOptions.map(opt => {
      const v = Object.prototype.hasOwnProperty.call(CALENDAR.colors, opt)
        ? CALENDAR.colors[opt]
        : "";
      return [v];
    });
    sheet.getRange(
      PER_TAB_OVERRIDE.dataStartRow,
      PER_TAB_OVERRIDE.appearanceValueCol,
      values.length,
      1
    ).setValues(values);

    values.forEach((row, i) => {
      const color = String(row[0] || "").trim();
      if (looksLikeColor_(color)) {
        sheet.getRange(
          PER_TAB_OVERRIDE.dataStartRow + i,
          PER_TAB_OVERRIDE.appearanceValueCol
        )
          .setBackground(color)
          .setFontColor(readableTextColor_(color));
      }
    });
  }

  // Category section: rewrite rows with current Key category names + colors,
  // applying live background to color cells.
  const keyCategories = readKeyCategoriesDisplay_(sheet.getParent());
  if (keyCategories.length > 0) {
    const categoryDataStartRow = perTabOverrideCategoryDataStartRow_();
    const rows = keyCategories.map(p => [String(p[0] || "").trim(), String(p[1] || "").trim()]);
    sheet.getRange(categoryDataStartRow, PER_TAB_OVERRIDE.setupNameCol, rows.length, 2)
      .setValues(rows);
    rows.forEach((row, i) => {
      const color = row[1];
      if (looksLikeColor_(color)) {
        sheet.getRange(categoryDataStartRow + i, PER_TAB_OVERRIDE.setupValueCol)
          .setBackground(color)
          .setFontColor(readableTextColor_(color));
      }
    });
  }
}

function fetchThemeManifest_() {
  try {
    const response = UrlFetchApp.fetch(CALENDAR.themes.indexUrl, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      Logger.log("Theme manifest HTTP " + response.getResponseCode());
      return null;
    }
    const data = JSON.parse(response.getContentText());
    return Array.isArray(data.themes) ? data.themes : null;
  } catch (err) {
    Logger.log("fetchThemeManifest_ error: " + err.message);
    return null;
  }
}

function fetchTheme_(filename) {
  try {
    const response = UrlFetchApp.fetch(CALENDAR.themes.baseUrl + filename, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      Logger.log("Theme " + filename + " HTTP " + response.getResponseCode());
      return null;
    }
    return JSON.parse(response.getContentText());
  } catch (err) {
    Logger.log("fetchTheme_ error: " + err.message);
    return null;
  }
}

// Walk the Key tab's appearance section (N:O) and update matching color cells.
// Walk the setup section (K:L) and update matching option cells. Returns counts.
function applyThemeToKey_(keySheet, theme) {
  let colorChanges = 0;
  let setupChanges = 0;
  let categoryChanges = 0;

  if (theme.colors && typeof theme.colors === "object") {
    const lastRow = keySheet.getLastRow();
    if (lastRow >= 2) {
      const range = keySheet.getRange(2, 14, lastRow - 1, 2);
      const values = range.getValues();
      for (let i = 0; i < values.length; i++) {
        const appearance = String(values[i][0] || "").trim();
        if (!appearance) continue;
        if (!Object.prototype.hasOwnProperty.call(theme.colors, appearance)) continue;
        const newColor = String(theme.colors[appearance] || "").trim();
        if (!looksLikeColor_(newColor)) continue;

        const row = i + 2;
        keySheet.getRange(row, 15)
          .setValue(newColor)
          .setBackground(newColor)
          .setFontColor(readableTextColor_(newColor));
        colorChanges++;
      }
    }
  }

  if (theme.setup && typeof theme.setup === "object") {
    Object.keys(theme.setup).forEach(option => {
      const row = findKeyOptionRow_(keySheet, option);
      if (!row) return;
      keySheet.getRange(row, 12).setValue(theme.setup[option]);
      setupChanges++;
    });
  }

  // categoryPalette is positional: palette[0] → 1st named category in the
  // Key's Category section, palette[1] → 2nd, etc. Categories beyond the
  // palette length keep their existing colors.
  if (Array.isArray(theme.categoryPalette) && theme.categoryPalette.length > 0) {
    const lastRow = keySheet.getLastRow();
    if (lastRow >= 2) {
      const categoryValues = keySheet.getRange(2, 4, lastRow - 1, 1).getValues();
      let paletteIdx = 0;
      for (let i = 0; i < categoryValues.length && paletteIdx < theme.categoryPalette.length; i++) {
        const name = String(categoryValues[i][0] || "").trim();
        if (!name) continue;
        const color = String(theme.categoryPalette[paletteIdx] || "").trim();
        paletteIdx++;
        if (!looksLikeColor_(color)) continue;
        const row = i + 2;
        keySheet.getRange(row, 5)
          .setValue(color);
        // The Key tab styles both Category name and Color cells with the
        // background color — mirror what buildKeySheet_ does at creation.
        keySheet.getRange(row, 4, 1, 2)
          .setBackground(color)
          .setFontColor(readableTextColor_(color));
        categoryChanges++;
      }
    }
  }

  applyKeyOverrides_(keySheet.getParent());
  applyKeySetupValidations_(keySheet);

  return { colors: colorChanges, setup: setupChanges, categories: categoryChanges };
}

function buildEventsSheet_(sheet) {
  const headers = ["Title", "Date", "Category", "Type", "Status"];
  const fontFamily = CALENDAR.setup.fontFamily || "Inter";

  sheet.clear({ contentsOnly: false });
  sheet.setHiddenGridlines(false);

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground(CALENDAR.colors.titleBackground)
    .setFontColor(CALENDAR.colors.titleFontColor)
    .setFontWeight("bold")
    .setFontFamily(fontFamily)
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle");

  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 240);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 140);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 120);

  const maxRows = sheet.getMaxRows();
  if (maxRows > 1) {
    sheet.getRange(2, 2, maxRows - 1, 1).setNumberFormat("m/d/yyyy");
    sheet.getRange(2, 1, maxRows - 1, headers.length)
      .setFontFamily(fontFamily)
      .setVerticalAlignment("middle");
  }

  sheet.setRowHeight(1, 30);
}

function buildKeySheet_(sheet) {
  sheet.clear({ contentsOnly: false });
  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(1);

  const typeRows = [["Type", "Type-Icon"]].concat(CALENDAR.keyDefaults.types);
  const categoryRows = [["Category", "Category-Color"]].concat(CALENDAR.keyDefaults.categories);
  const statusRows = [["Status", "Status-Icon"]].concat(CALENDAR.keyDefaults.statuses);
  const setupRows = [["Additional Setup >", "Option", "Option-Value"]].concat(buildKeySetupRows_());
  const appearanceRows = [["Appearance", "Appearance-Color"]].concat(buildKeyAppearanceRows_());

  sheet.getRange(1, 1, typeRows.length, 2).setValues(typeRows);
  sheet.getRange(1, 4, categoryRows.length, 2).setValues(categoryRows);
  sheet.getRange(1, 7, statusRows.length, 2).setValues(statusRows);
  sheet.getRange(1, 10, setupRows.length, 3).setValues(setupRows);
  sheet.getRange(1, 14, appearanceRows.length, 2).setValues(appearanceRows);

  const maxRows = Math.max(
    typeRows.length,
    categoryRows.length,
    statusRows.length,
    setupRows.length,
    appearanceRows.length
  );

  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 30);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 30);
  sheet.setColumnWidth(7, 150);
  sheet.setColumnWidth(8, 90);
  sheet.setColumnWidth(9, 30);
  sheet.setColumnWidth(10, 150);
  sheet.setColumnWidth(11, 210);
  sheet.setColumnWidth(12, 130);
  sheet.setColumnWidth(13, 30);
  sheet.setColumnWidth(14, 260);
  sheet.setColumnWidth(15, 130);

  applyKeySheetFormatting_(sheet, maxRows);

  setCheckboxIfOption_(sheet, "showInitialMenu");
  setCheckboxIfOption_(sheet, "showEventListMenu");
  setCheckboxIfOption_(sheet, "showSetKeyFromEventListMenu");
  setCheckboxIfOption_(sheet, "showImportThemeMenu");
  setCheckboxIfOption_(sheet, "showKeyConfiguratorMenuItems");
  setCheckboxIfOption_(sheet, "autoRefresh");
  setCheckboxIfOption_(sheet, "frozenWeekdayHeader");
  setCheckboxIfOption_(sheet, "customAdditionalLabels");
  applyKeyOptionNotes_(sheet);

  for (let r = 2; r <= categoryRows.length; r++) {
    const color = String(sheet.getRange(r, 5).getValue() || "").trim();

    if (looksLikeColor_(color)) {
      sheet.getRange(r, 4, 1, 2)
        .setBackground(color)
        .setFontColor(readableTextColor_(color));
    }
  }

  for (let r = 2; r <= appearanceRows.length; r++) {
    const color = String(sheet.getRange(r, 15).getValue() || "").trim();

    if (looksLikeColor_(color)) {
      sheet.getRange(r, 15)
        .setBackground(color)
        .setFontColor(readableTextColor_(color));
    }
  }

  applyKeySetupConditionalFormatting_(sheet);
  applyKeySetupValidations_(sheet);

  try {
    sheet.getRange("K:O").shiftColumnGroupDepth(1);
    const group = sheet.getColumnGroup(11, 1);
    if (group) group.collapse();
  } catch (err) {
    Logger.log("Could not group/collapse K:O: " + err.message);
  }
}

// Attach dropdowns to the value cells (column L = 12) of specific setup
// options so users can pick from valid choices instead of typing.
//   q1StartMonth         — month names (strict)
//   startWeekOn          — day names (strict)
//   defaultDataSheetName — list of available tabs (allow-invalid)
//   customDate           — headers of the resolved defaultDataSheetName
//   customTitle          — headers of the resolved defaultDataSheetName
// Safe to call repeatedly; setDataValidation just replaces any prior rule.
function applyKeySetupValidations_(sheet) {
  const ss = sheet.getParent();

  const monthRow = findKeyOptionRow_(sheet, "q1StartMonth");
  if (monthRow) {
    setDropdownValidation(sheet.getRange(monthRow, 12), CALENDAR.monthOptions.slice(), false);
  }

  const dayRow = findKeyOptionRow_(sheet, "startWeekOn");
  if (dayRow) {
    setDropdownValidation(
      sheet.getRange(dayRow, 12),
      [
        "Sunday",
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
        "Monday-CompressedWeekend"
      ],
      false
    );
  }

  const tabRow = findKeyOptionRow_(sheet, "defaultDataSheetName");
  if (tabRow) {
    const names = buildSourceSheetNames_(ss);
    if (names.length) {
      setDropdownValidation(sheet.getRange(tabRow, 12), names, true);
    } else {
      sheet.getRange(tabRow, 12).clearDataValidations();
    }
  }

  const headers = readDefaultDataSheetHeaders_(ss);
  // customAdditional is a comma-separated list of source headers. Apply the
  // same headers-based dropdown so single-pick works; if the user manually
  // enables "Allow multiple selections" on this cell in the Data validation
  // dialog, picking multiple options writes them as comma-separated text,
  // which is exactly the format customAdditional already expects.
  [
    "customDate",
    "customTitle",
    "customType",
    "customCategory",
    "customStatus",
    "customAdditional",
  ].forEach(optionName => {
    const row = findKeyOptionRow_(sheet, optionName);
    if (!row) return;
    if (headers.length) {
      setDropdownValidation(sheet.getRange(row, 12), headers, true);
    } else {
      sheet.getRange(row, 12).clearDataValidations();
    }
  });

  // eventSortOrder: each option is a "<Field> <↓|↑>" token. Multi-select
  // (enable "Allow multiple selections" on the cell) lets you stack sorts —
  // commas separate primary, secondary, tertiary, etc.
  const sortRow = findKeyOptionRow_(sheet, "eventSortOrder");
  if (sortRow) {
    setDropdownValidation(
      sheet.getRange(sortRow, 12),
      EVENT_SORT_OPTIONS,
      true
    );
  }
}

function readDefaultDataSheetHeaders_(ss) {
  // First try the configured default.
  const name = CALENDAR.defaultDataSheetName;
  if (name) {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      const headers = readHeaders_(sheet);
      if (headers.length) return headers;
    }
  }

  // Fallback: the first source sheet referenced by any calendar's G1
  // Source Data dropdown. Handles the common case where the Key still says
  // the script default ("Events") but the actual source tab is named
  // something else (e.g. "Tactics List").
  try {
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      if (!isCalendarSheet(sheets[i])) continue;
      const g1 = String(sheets[i].getRange("G1").getDisplayValue() || "").trim();
      if (!g1) continue;
      // G1 may be a multi-select list — try each token, return the first
      // sheet that resolves and has headers.
      const tokens = parseFilterValues_(g1);
      for (let j = 0; j < tokens.length; j++) {
        const target = ss.getSheetByName(tokens[j]);
        if (!target) continue;
        const headers = readHeaders_(target);
        if (headers.length) return headers;
      }
    }
  } catch (err) {}

  return [];
}

// Reads Key categories preserving display case (readKeyConfig_ normalizes the
// keys for lookup). Returns [[name, color], ...]. Falls back to script
// defaults if the Key tab is missing or has no usable Category column.
function readKeyCategoriesDisplay_(ss) {
  const sheet = ss.getSheetByName(CALENDAR.keySheetName);
  if (!sheet) return CALENDAR.keyDefaults.categories.slice();

  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return CALENDAR.keyDefaults.categories.slice();

  const header = values[0].map(v => String(v || "").trim());
  const categoryCol = findColumnIndex(header, ["Category"]);
  const categoryColorCol = findColumnIndex(header, ["Category-Color", "Category Color", "Color"]);
  if (categoryCol < 0) return CALENDAR.keyDefaults.categories.slice();

  const out = [];
  values.slice(1).forEach(row => {
    const name = String(row[categoryCol] || "").trim();
    if (!name) return;
    const color = categoryColorCol >= 0 ? String(row[categoryColorCol] || "").trim() : "";
    out.push([name, color]);
  });
  return out.length ? out : CALENDAR.keyDefaults.categories.slice();
}



function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();

  if (sheet.getName() === CALENDAR.keySheetName) {
    formatEditedKeyCell_(e.range);
    refreshKeyHeaderDropdownsOnEdit_(e.range);
    return;
  }

  if (!isCalendarSheet(sheet)) {
    // Not a Key, not a calendar — could be the source data tab.
    // handleSourceEdit_ early-returns if Auto-Refresh is off or the edit
    // isn't on a column the renderer consumes.
    handleSourceEdit_(e);
    return;
  }

  formatEditedPerTabOverrideCell_(e.range);

  const cell = e.range.getA1Notation();

  if (cell === "A1" || cell === "B1") {
    syncModeDefaults_(sheet);
    syncModeStyles_(sheet);
    return;
  }

  if (cell === "G1") {
    sheet.getRange("A2").setValue(CALENDAR.setup.filterDefaultLabel);
    sheet.getRange("B2").clearContent();
    setFilterDropdowns_(sheet, true);
    syncFilterStyles_(sheet);
    renderCalendarSheet(sheet, getLiveControls_(sheet));
    return;
  }

  if (cell === "A2") {
    sheet.getRange("B2").clearContent();
    setFilterDropdowns_(sheet, true);
    syncFilterStyles_(sheet);
    renderCalendarSheet(sheet, getLiveControls_(sheet));
    return;
  }

  if (cell === "B2") {
    syncFilterStyles_(sheet);
    renderCalendarSheet(sheet, getLiveControls_(sheet));
    return;
  }

  if (cell !== "G2") return;
  if (String(e.value) !== "TRUE") return;

  let refreshed = false;

  try {
    const controls = getLiveControls_(sheet);
    renderCalendarSheet(sheet, controls);
    refreshed = true;
    sheet.getParent().toast("Calendar updated.", CALENDAR.menuName, 3);
  } catch (err) {
    sheet.getParent().toast("Calendar update failed: " + err.message, CALENDAR.menuName, 5);
  } finally {
    if (refreshed) {
      sheet.getRange("G2").setValue(false);
      SpreadsheetApp.flush();
    }
  }
}

function onSelectionChange(e) {
  if (!e || !e.range) return;

  let userProps;
  try { userProps = PropertiesService.getUserProperties(); } catch (err) { return; }

  const sheet = e.range.getSheet();
  const sheetId = String(sheet.getSheetId());
  const sheetName = sheet.getName();
  const prevSheetId = userProps.getProperty(AUTO_REFRESH.lastActiveSheetKey);
  userProps.setProperty(AUTO_REFRESH.lastActiveSheetKey, sheetId);

  if (prevSheetId === sheetId) return;
  if (!isCalendarSheet(sheet)) return;
  if (!isCalendarStale_(sheet)) return;

  const ss = sheet.getParent();
  try { applyKeyOverrides_(ss); } catch (err) { return; }
  if (!isAutoRefreshEnabled_()) return;

  try {
    renderCalendarSheet(sheet);
    Logger.log("[AR] auto-refreshed '" + sheetName + "' on tab switch");
  } catch (err) {
    Logger.log("[AR] auto-refresh on tab switch failed: " + err.message);
  }
}










// When the value cell next to defaultDataSheetName changes, the headers of the
// resolved source tab may have changed too, so refresh the customDate /
// customTitle dropdowns.
function refreshKeyHeaderDropdownsOnEdit_(range) {
  if (range.getColumn() !== 12) return;
  const row = range.getRow();
  if (row < 2) return;

  const sheet = range.getSheet();
  const optionName = String(sheet.getRange(row, 11).getValue() || "").trim();
  if (optionName !== "defaultDataSheetName") return;

  applyKeyOverrides_(sheet.getParent());
  applyKeySetupValidations_(sheet);
}

function formatEditedKeyCell_(range) {
  const row = range.getRow();
  const col = range.getColumn();

  if (row < 2) return;

  if (col === 5) {
    const color = String(range.getValue() || "").trim();

    if (looksLikeColor_(color)) {
      range.getSheet().getRange(row, 4, 1, 2)
        .setBackground(color)
        .setFontColor(readableTextColor_(color));
    } else {
      range.getSheet().getRange(row, 4, 1, 2)
        .setBackground(null)
        .setFontColor(null);
    }

    return;
  }

  if (col === 15) {
    const color = String(range.getValue() || "").trim();

    if (looksLikeColor_(color)) {
      range.getSheet().getRange(row, 15)
        .setBackground(color)
        .setFontColor(readableTextColor_(color));
    } else {
      range.getSheet().getRange(row, 15)
        .setBackground(null)
        .setFontColor(null);
    }
  }
}



function initializeCalendarSheet(sheet, period, year, sourceSpec) {
  clearFullSheet_(sheet);

  ensureCanvasSize_(sheet);
  sheet.setFrozenRows(getFrozenRowCount_());
  sheet.setHiddenGridlines(true);

  const initialMonthCount = getDisplayedMonthCount(period, 1);
  const initialStartDate = getStartDateForMode_(period, year);

  writeControlBlock_(sheet, {
    period: period,
    year: year,
    monthCount: initialMonthCount,
    startDate: initialStartDate,
    sourceSpec: sourceSpec || getDefaultSourceSpec(sheet.getParent()),
    filterField: CALENDAR.setup.filterDefaultLabel,
    filterValue: ""
  });

  setControlFormats_(sheet, period);

  setDropdownValidation(sheet.getRange("A1"), CALENDAR.periodOptions, false);
  setDropdownValidation(sheet.getRange("B1"), CALENDAR.yearOptions, false);
  setSourceSheetDropdown_(sheet);
  setFilterDropdowns_(sheet, true);

  // Column widths are set once at the end of renderCalendarSheet so we don't
  // briefly resize columns (especially the narrow control-row spacer in
  // column E) and then resize them back during refresh.

  sheet.getRange("A1:G2").setFontFamily(CALENDAR.setup.fontFamily);
  sheet.getRange("G2").setNote("Check to refresh this calendar after changing the controls.");
}

function renderCalendarSheet(sheet, controlsOverride) {
  const ss = sheet.getParent();
  applyKeyOverrides_(ss);

  // Per-tab overrides layer on top of Key for the duration of this render,
  // then are restored so the next calendar in Refresh All starts clean.
  withPerTabOverridesApplied_(sheet, (overrides) => {
    renderCalendarSheetBody_(sheet, controlsOverride, overrides);
  });

  // Auto-Refresh bookkeeping: mark this calendar as up-to-date so subsequent
  // tab switches know it doesn't need re-rendering until the next source edit.
  markCalendarRefreshed_(sheet);
}

function renderCalendarSheetBody_(sheet, controlsOverride, overrides) {
  const ss = sheet.getParent();
  const controls = controlsOverride || getLiveControls_(sheet);
  const source = loadSourceData(ss, controls.sourceSpec);
  const keyConfig = readKeyConfig_(ss);

  // Layer per-tab category color overrides on top of keyConfig's Key-derived
  // map. Setup + appearance overrides were already applied to CALENDAR.* by
  // withPerTabOverridesApplied_; category colors aren't part of CALENDAR.* so
  // they're merged here instead.
  if (overrides && overrides.categoryColors) {
    Object.keys(overrides.categoryColors).forEach(k => {
      keyConfig.categoryColors[k] = overrides.categoryColors[k];
    });
  }

  const renderControls = normalizeControlsForRender_(controls);
  rebuildCalendarCanvas(sheet, renderControls);

  const startDate = getStartDateForControls_(controls);
  const monthCount = getDisplayedMonthCount(renderControls.period, controls.monthCount);

  sheet.getRange("D2").setValue(startDate);
  sheet.getRange("D2").setNumberFormat("m/d/yy");
  sheet.getRange("D1").setValue(monthCount);

  applyModeStyles_(sheet, renderControls.period);

  const filteredRows = filterRowsForCalendar_(source.rows, source.headers, controls);
  const eventsByDate = indexEventsByDate(filteredRows, source.headers, source.sourceMeta, keyConfig);
  const monthInfos = buildMonthInfos(renderControls, startDate, monthCount);

  let row = CALENDAR.layout.firstMonthRow;

  monthInfos.forEach((monthInfo, index) => {
    renderMonthSection(sheet, row, monthInfo, eventsByDate, index);
    row += getMonthBlockRows_();
  });

  if (isCompressedWeekend_()) {
    // Mon-Fri at standard width, Sat + Sun at half width.
    sheet.setColumnWidths(1, 5, 145);
    sheet.setColumnWidths(6, 2, 72);
  } else {
    sheet.setColumnWidths(1, 7, 145);
  }
}

function rebuildCalendarCanvas(sheet, controls) {
  clearCalendarCanvas_(sheet);

  ensureCanvasSize_(sheet);
  sheet.setFrozenRows(getFrozenRowCount_());
  sheet.setHiddenGridlines(true);

  writeControlBlock_(sheet, controls);
  setControlFormats_(sheet, controls.period);

  setDropdownValidation(sheet.getRange("A1"), CALENDAR.periodOptions, false);
  setDropdownValidation(sheet.getRange("B1"), CALENDAR.yearOptions, false);
  setSourceSheetDropdown_(sheet);
  setFilterDropdowns_(sheet, false);

  // Column widths are set once at the end of renderCalendarSheet so we don't
  // briefly resize columns (especially the narrow control-row spacer in
  // column E) and then resize them back during refresh.

  sheet.getRange(3, 1, CALENDAR.layout.renderRowCapacity, 7).clear({ contentsOnly: false });
  renderFrozenWeekdayHeader_(sheet);
}

function clearFullSheet_(sheet) {
  sheet.clearConditionalFormatRules();
  sheet.getRange(1, 1, CALENDAR.layout.renderRowCapacity + 2, 7).clear({ contentsOnly: false });
}

function clearCalendarCanvas_(sheet) {
  sheet.clearConditionalFormatRules();
  sheet.getRange(3, 1, CALENDAR.layout.renderRowCapacity, 7).clear({ contentsOnly: false });
  renderFrozenWeekdayHeader_(sheet);
}


function getFrozenRowCount_() {
  return CALENDAR.setup.frozenWeekdayHeader ? 3 : CALENDAR.layout.frozenRows;
}

function renderFrozenWeekdayHeader_(sheet) {
  const bufferHeight = CALENDAR.layout.bufferRowHeight || 5;

  sheet.getRange(3, 1, 2, 7).clear({ contentsOnly: false });

  sheet.setRowHeight(3, bufferHeight);
  sheet.setRowHeight(4, bufferHeight);

  if (!CALENDAR.setup.frozenWeekdayHeader) return;

  const weekdays = getWeekdayLabels();

  sheet.getRange(3, 1, 1, 7)
    .setValues([weekdays])
    .setBackground(CALENDAR.colors.titleWeekBackground)
    .setFontColor(CALENDAR.colors.titleWeekFontColor)
    .setFontWeight("bold")
    .setFontFamily(CALENDAR.setup.fontFamily)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  sheet.setRowHeight(3, 26);
}

function getMonthBlockRows_() {
  return 2 + 1 + (6 * (1 + getMaxEvents_() + 1)) + 1;
}

function getMaxEvents_() {
  const n = Number(CALENDAR.setup.maxEvents);
  if (!n || isNaN(n) || n < 1) return 4;
  return Math.floor(n);
}

function setSourceSheetDropdown_(sheet) {
  const ss = sheet.getParent();
  const cell = sheet.getRange("G1");
  const names = buildSourceSheetNames_(ss);

  if (!names.length) {
    cell.clearDataValidations();
    cell.setValue("");
    cell.setNote('No source tabs are available yet. Use "Calendar Tools > New Calendar Sheet" to create a tab.');
    return;
  }

  // Allow-invalid so the cell tolerates a comma-separated value once the user
  // enables Data → Data validation → Allow multiple selections. With one tab
  // selected the cell behaves like the legacy single-source dropdown.
  setDropdownValidation(cell, names, true);

  const current = String(cell.getValue() || "").trim();
  if (!current) {
    cell.setValue(getDefaultSourceSpec(ss));
  }

  cell.setNote(
    "Choose a local source tab for this calendar. " +
    "To combine events from multiple tabs, open Data > Data validation on this cell " +
    "and enable Allow multiple selections."
  );
}

function setFilterDropdowns_(sheet, resetValueValidation) {
  const ss = sheet.getParent();
  const sourceSpec = String(sheet.getRange("G1").getDisplayValue() || getDefaultSourceSpec(ss)).trim();
  const source = loadSourceData(ss, sourceSpec);
  const headers = (source.headers || []).filter(h => String(h || "").trim() !== "");
  const defaultLabel = CALENDAR.setup.filterDefaultLabel;
  const filterCell = sheet.getRange("A2");
  const valueCell = sheet.getRange("B2");
  const filterOptions = [defaultLabel].concat(headers);

  // Allow-invalid so the user can enable "Allow multiple selections" on this
  // cell via Sheets' data validation dialog and pick more than one column at
  // once — filterRowsForCalendar_ handles the comma-separated multi value.
  setDropdownValidation(filterCell, filterOptions, true);

  let filterField = String(filterCell.getDisplayValue() || "").trim();
  // Treat a value containing only the default label or a single picked column
  // as known. Multi-select strings won't match indexOf — accept them as-is.
  const fieldTokens = parseFilterValues_(filterField).filter(t => t !== defaultLabel);
  if (!filterField || (fieldTokens.length === 0 && filterField !== defaultLabel)) {
    filterField = defaultLabel;
    filterCell.setValue(filterField);
  }

  filterCell.setNote("Choose a source column to filter. Enable Allow multiple selections via Data > Data validation to filter on several columns at once (OR across columns + values).");

  // Multi-column support: collect unique values across all picked columns.
  const filterColumns = parseFilterValues_(filterField).filter(c => c !== defaultLabel);
  if (filterColumns.length === 0) {
    if (resetValueValidation) valueCell.clearDataValidations();
    valueCell.setNote("Choose a filter field first.");
    syncFilterStyles_(sheet);
    return;
  }

  const seen = Object.create(null);
  const values = [];
  filterColumns.forEach(col => {
    const colIndex = findColumnIndex(headers, [col]);
    if (colIndex < 0) return;
    uniqueValuesFromColumn_(source.rows, colIndex).forEach(v => {
      const k = normalizeKey_(v);
      if (!seen[k]) { seen[k] = true; values.push(v); }
    });
  });
  values.sort((a, b) => String(a).localeCompare(String(b)));

  if (!values.length) {
    if (resetValueValidation) valueCell.clearDataValidations();
    valueCell.setNote("No values found for this filter field.");
    syncFilterStyles_(sheet);
    return;
  }

  if (resetValueValidation || !valueCell.getDataValidation()) {
    setDropdownValidation(valueCell, values, true);
  }

  valueCell.setNote("Choose one or more values. To use native multi-select, enable Allow multiple selections for this cell in Data validation. With multi-column A2 + multi-value B2, the filter is OR across both axes.");
  syncFilterStyles_(sheet);
}

function uniqueValuesFromColumn_(rows, colIndex) {
  if (colIndex < 0) return [];

  const seen = {};
  const out = [];

  (rows || []).forEach(row => {
    const value = String(row[colIndex] || "").trim();
    if (!value) return;
    const key = normalizeKey_(value);
    if (seen[key]) return;
    seen[key] = true;
    out.push(value);
  });

  return out.sort((a, b) => String(a).localeCompare(String(b)));
}

function buildSourceSheetNames_(ss) {
  return ss.getSheets()
    .map(s => s.getName())
    .filter(name => name !== CALENDAR.keySheetName);
}

function ensureCanvasSize_(sheet) {
  if (sheet.getMaxRows() < CALENDAR.layout.renderRowCapacity + 2) {
    sheet.insertRowsAfter(sheet.getMaxRows(), (CALENDAR.layout.renderRowCapacity + 2) - sheet.getMaxRows());
  }

  if (sheet.getMaxColumns() < 7) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 7 - sheet.getMaxColumns());
  }
}

function writeControlBlock_(sheet, controls) {
  sheet.getRange("A1").setValue(controls.period);
  sheet.getRange("B1").setValue(String(controls.year));
  sheet.getRange("C1").setValue("Months:");
  sheet.getRange("D1").setValue(controls.monthCount);
  sheet.getRange("F1").setValue("Source Data:");
  sheet.getRange("G1").setValue(controls.sourceSpec || getDefaultSourceSpec(sheet.getParent()));

  sheet.getRange("A2").setValue(controls.filterField || CALENDAR.setup.filterDefaultLabel);
  sheet.getRange("B2").setValue(controls.filterValue || "");
  sheet.getRange("C2").setValue("Starting:");
  sheet.getRange("D2").setValue(controls.startDate);

  sheet.getRange("F2").setValue("Refresh:");
  sheet.getRange("G2").insertCheckboxes();
  sheet.getRange("G2").setValue(false);
}

function setControlFormats_(sheet, period) {
  const custom = isCustom(period);
  const yearMode = isYear(period);
  const accent = CALENDAR.colors.titleAccentColor;

  sheet.getRange("A1:G2")
    .setBackground(CALENDAR.colors.titleBackground)
    .setFontColor(CALENDAR.colors.titleFontColor)
    .setFontFamily(CALENDAR.setup.fontFamily)
    .setVerticalAlignment("middle");

  sheet.getRange("A1:B1")
    .setHorizontalAlignment("center")
    .setFontWeight("bold")
    .setFontSize(18);

  sheet.getRange("A2")
    .setHorizontalAlignment("center")
    .setFontWeight("bold")
    .setFontSize(12);

  sheet.getRange("B2")
    .setHorizontalAlignment("center")
    .setFontWeight("normal")
    .setFontSize(12);

  sheet.getRange("C1:C2")
    .setHorizontalAlignment("right")
    .setFontWeight("bold")
    .setFontSize(12);

  sheet.getRange("D1")
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle")
    .setFontWeight(custom ? "bold" : "normal")
    .setFontSize(12);

  sheet.getRange("D2")
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle")
    .setFontWeight((custom || yearMode) ? "bold" : "normal")
    .setFontSize(12);

  sheet.getRange("F1:F2")
    .setHorizontalAlignment("right")
    .setFontWeight("bold")
    .setFontSize(12);

  sheet.getRange("G1")
    .setHorizontalAlignment("left")
    .setFontWeight("normal")
    .setFontSize(12);

  sheet.getRange("G2")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground(CALENDAR.colors.titleRefreshColor)
    .setFontColor("#FFFFFF")
    .setFontWeight("normal");

  if (custom) {
    sheet.getRange("B1")
      .setBackground(CALENDAR.colors.titleBackground)
      .setFontColor(CALENDAR.colors.titleBackground)
      .setFontWeight("bold");

    sheet.getRange("D1:D2")
      .setBackground(accent)
      .setFontColor("#FFFFFF")
      .setFontWeight("bold");
  } else if (yearMode) {
    sheet.getRange("B1")
      .setBackground(CALENDAR.colors.titleBackground)
      .setFontColor(CALENDAR.colors.titleFontColor)
      .setFontWeight("bold");

    sheet.getRange("D1")
      .setBackground(CALENDAR.colors.titleBackground)
      .setFontColor(CALENDAR.colors.titleFontColor)
      .setFontWeight("normal");

    sheet.getRange("D2")
      .setBackground(accent)
      .setFontColor("#FFFFFF")
      .setFontWeight("bold");
  } else {
    sheet.getRange("B1")
      .setBackground(CALENDAR.colors.titleBackground)
      .setFontColor(CALENDAR.colors.titleFontColor)
      .setFontWeight("bold");

    sheet.getRange("D1")
      .setBackground(CALENDAR.colors.titleBackground)
      .setFontColor(CALENDAR.colors.titleFontColor)
      .setFontWeight("normal");

    sheet.getRange("D2")
      .setBackground(CALENDAR.colors.titleBackground)
      .setFontColor(CALENDAR.colors.titleFontColor)
      .setFontWeight("normal");
  }

  syncFilterStyles_(sheet);

  sheet.setRowHeight(1, 30);
  sheet.setRowHeight(2, 24);
}

function syncFilterStyles_(sheet) {
  const filterField = String(sheet.getRange("A2").getDisplayValue() || "").trim();
  const filterValue = String(sheet.getRange("B2").getDisplayValue() || "").trim();
  const isActiveFilter =
    filterField &&
    filterField !== CALENDAR.setup.filterDefaultLabel &&
    filterValue;

  sheet.getRange("A2:B2")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  if (isActiveFilter) {
    sheet.getRange("B2")
      .setBackground(CALENDAR.colors.titleAccentColor)
      .setFontColor("#FFFFFF");
  } else {
    sheet.getRange("B2")
      .setBackground(CALENDAR.colors.titleBackground)
      .setFontColor(CALENDAR.colors.titleFontColor);
  }
}

function syncModeStyles_(sheet) {
  const period = String(sheet.getRange("A1").getValue() || "").trim();
  const custom = isCustom(period);
  const yearMode = isYear(period);
  const accent = CALENDAR.colors.titleAccentColor;

  if (custom) {
    sheet.getRange("B1")
      .setBackground(CALENDAR.colors.titleBackground)
      .setFontColor(CALENDAR.colors.titleBackground)
      .setFontWeight("bold");

    sheet.getRange("D1:D2")
      .setBackground(accent)
      .setFontColor("#FFFFFF")
      .setFontWeight("bold");
    return;
  }

  if (yearMode) {
    sheet.getRange("B1")
      .setBackground(CALENDAR.colors.titleBackground)
      .setFontColor(CALENDAR.colors.titleFontColor)
      .setFontWeight("bold");

    sheet.getRange("D1")
      .setBackground(CALENDAR.colors.titleBackground)
      .setFontColor(CALENDAR.colors.titleFontColor)
      .setFontWeight("normal");

    sheet.getRange("D2")
      .setBackground(accent)
      .setFontColor("#FFFFFF")
      .setFontWeight("bold");
    return;
  }

  sheet.getRange("B1")
    .setBackground(CALENDAR.colors.titleBackground)
    .setFontColor(CALENDAR.colors.titleFontColor)
    .setFontWeight("bold");

  sheet.getRange("D1")
    .setBackground(CALENDAR.colors.titleBackground)
    .setFontColor(CALENDAR.colors.titleFontColor)
    .setFontWeight("normal");

  sheet.getRange("D2")
    .setBackground(CALENDAR.colors.titleBackground)
    .setFontColor(CALENDAR.colors.titleFontColor)
    .setFontWeight("normal");
}

function applyModeStyles_(sheet, period) {
  setControlFormats_(sheet, period);
}

function readControls(sheet) {
  const period = String(sheet.getRange("A1").getValue() || currentMonthName()).trim();
  const year = Number(sheet.getRange("B1").getValue() || currentYear());
  const monthCountRaw = sheet.getRange("D1").getValue();
  const startDateRaw = sheet.getRange("D2").getValue();
  const sourceSpec = String(sheet.getRange("G1").getValue() || getDefaultSourceSpec(sheet.getParent())).trim();
  const filterField = String(sheet.getRange("A2").getValue() || CALENDAR.setup.filterDefaultLabel).trim();
  const filterValue = String(sheet.getRange("B2").getValue() || "").trim();

  if (!period) throw new Error("A1 period is missing.");
  if (!year || isNaN(year)) throw new Error("B1 year is missing or invalid.");

  const monthCount = Number(monthCountRaw);
  const startDate = startDateRaw instanceof Date && !isNaN(startDateRaw.getTime())
    ? normalizeDateToFirstOfMonth_(normalizeSpreadsheetDate_(startDateRaw))
    : getStartDateForMode_(period, year);

  return { period, year, monthCount, startDate, sourceSpec, filterField, filterValue };
}

function getLiveControls_(sheet) {
  const period = String(sheet.getRange("A1").getDisplayValue() || "").trim();
  const year = Number(sheet.getRange("B1").getDisplayValue() || currentYear());
  const monthCount = Number(sheet.getRange("D1").getDisplayValue() || 1);
  const startDateRaw = sheet.getRange("D2").getValue();
  const sourceSpec = String(sheet.getRange("G1").getDisplayValue() || getDefaultSourceSpec(sheet.getParent())).trim();
  const filterField = String(sheet.getRange("A2").getDisplayValue() || CALENDAR.setup.filterDefaultLabel).trim();
  const filterValue = String(sheet.getRange("B2").getDisplayValue() || "").trim();

  const startDate = startDateRaw instanceof Date && !isNaN(startDateRaw.getTime())
    ? normalizeDateToFirstOfMonth_(normalizeSpreadsheetDate_(startDateRaw))
    : getStartDateForMode_(period, year);

  return { period, year, monthCount, startDate, sourceSpec, filterField, filterValue };
}

function normalizeDateToFirstOfMonth_(date) {
  const d = new Date(date);
  d.setDate(1);
  d.setHours(0, 0, 0, 0);
  return d;
}

function getDisplayedMonthCount(period, customCount) {
  if (isYear(period)) return 12;

  if (isCustom(period)) {
    const n = Number(customCount);
    return Math.max(1, isNaN(n) ? 1 : Math.floor(n));
  }

  return isQuarter(period) ? 3 : 1;
}

function getStartDateForMode_(period, year) {
  if (isYear(period)) {
    return new Date(year, 0, 1);
  }

  if (isQuarter(period)) {
    const q1StartMonth = monthNameToNumber(CALENDAR.setup.q1StartMonth);
    const q1StartDay = CALENDAR.setup.q1StartDay;
    const qNum = Number(period.substring(1));
    const quarterOffsetMonths = (qNum - 1) * 3;
    const startDate = new Date(year, q1StartMonth - 1, q1StartDay);
    startDate.setMonth(startDate.getMonth() + quarterOffsetMonths);
    return normalizeDateToFirstOfMonth_(startDate);
  }

  if (isCustom(period)) {
    return new Date(year, 0, 1);
  }

  const monthNum = monthNameToNumber(period);
  return new Date(year, monthNum - 1, 1);
}

function getStartDateForControls_(controls) {
  if (isCustom(controls.period) || isYear(controls.period)) {
    return controls.startDate instanceof Date
      ? normalizeDateToFirstOfMonth_(controls.startDate)
      : getStartDateForMode_(controls.period, controls.year);
  }

  return getStartDateForMode_(controls.period, controls.year);
}

function normalizeControlsForRender_(controls) {
  const renderControls = {
    period: controls.period,
    year: controls.year,
    monthCount: controls.monthCount,
    startDate: controls.startDate,
    sourceSpec: controls.sourceSpec,
    filterField: controls.filterField || CALENDAR.setup.filterDefaultLabel,
    filterValue: controls.filterValue || ""
  };

  if (isMonth(renderControls.period)) {
    renderControls.period = Utilities.formatDate(
      renderControls.startDate,
      Session.getScriptTimeZone(),
      "MMMM"
    );
  }

  return renderControls;
}

function buildMonthInfos(controls, startDate, monthCount) {
  if (isYear(controls.period)) {
    const first = new Date(startDate);
    first.setDate(1);

    return Array.from({ length: 12 }, (_, i) => {
      const d = new Date(first);
      d.setMonth(d.getMonth() + i);
      d.setDate(1);
      return d;
    });
  }

  if (isCustom(controls.period)) {
    const count = Math.max(1, Number(monthCount) || 1);
    const first = new Date(startDate);
    first.setDate(1);

    return Array.from({ length: count }, (_, i) => {
      const d = new Date(first);
      d.setMonth(d.getMonth() + i);
      d.setDate(1);
      return d;
    });
  }

  if (isQuarter(controls.period)) {
    const first = new Date(startDate);
    first.setDate(1);

    return [0, 1, 2].map((offset) => {
      const d = new Date(first);
      d.setMonth(d.getMonth() + offset);
      d.setDate(1);
      return d;
    });
  }

  return [new Date(startDate)];
}

function renderMonthSection(sheet, startRow, monthDate, eventsByDate, monthIndexInView) {
  const tz = Session.getScriptTimeZone();
  const monthName = Utilities.formatDate(monthDate, tz, CALENDAR.setup.monthFormat);
  const year = monthDate.getFullYear();
  const theme = getMonthTheme(monthIndexInView);
  const maxEvents = getMaxEvents_();
  const fontFamily = CALENDAR.setup.fontFamily;
  const todayKey = dateKey(new Date());

  // --- Month title rows (startRow + 0, startRow + 1) ---
  const titleRange = sheet.getRange(startRow, 1, 2, 7);
  titleRange.setValues([
    [monthName, year, "", "", "", "", ""],
    ["",        "",   "", "", "", "", ""]
  ]);
  titleRange
    .setBackground(theme.monthBackground)
    .setFontColor(theme.monthFontColor)
    .setFontWeight("bold")
    .setFontFamily(fontFamily)
    .setVerticalAlignment("middle");
  sheet.getRange(startRow, 1, 1, 2)
    .setFontSize(18)
    .setHorizontalAlignment("left");

  // --- Weekday header (startRow + 2) ---
  const weekdays = getWeekdayLabels();
  sheet.getRange(startRow + 2, 1, 1, 7)
    .setValues([weekdays])
    .setBackground(theme.weekBackground)
    .setFontColor(theme.weekFontColor)
    .setFontWeight("bold")
    .setFontFamily(fontFamily)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // --- 6-week body grid ---
  const firstOfMonth = new Date(monthDate.getFullYear(), monthDate.getMonth(), 1);
  const weekStartIndex = getWeekStartIndex_();
  const leadingDays = (firstOfMonth.getDay() - weekStartIndex + 7) % 7;
  const gridStart = new Date(firstOfMonth);
  gridStart.setDate(firstOfMonth.getDate() - leadingDays);

  const bodyStartRow = startRow + 3;
  const rowsPerWeek = 1 + maxEvents + 1;
  const bodyRows = 6 * rowsPerWeek;

  // Font color + weight live inside the RichTextValues themselves. Cell-level
  // setFontColors / setFontWeights would override the per-character bolding of
  // additional-field labels, so we do not call them on the batched body.
  const blankRich = SpreadsheetApp.newRichTextValue().setText("").build();

  const richValues = [];
  const backgrounds = [];
  const horizontalAlignments = [];
  const verticalAlignments = [];
  const notes = [];

  for (let week = 0; week < 6; week++) {
    const dateRowRich = new Array(7);
    const dateRowBg = new Array(7);
    const dateRowHa = new Array(7);
    const dateRowVa = new Array(7);
    const dateRowNotes = new Array(7);

    const eventRowsRich = [];
    const eventRowsBg = [];
    const eventRowsHa = [];
    const eventRowsVa = [];
    const eventRowsNotes = [];

    for (let i = 0; i < maxEvents + 1; i++) {
      eventRowsRich.push(new Array(7));
      eventRowsBg.push(new Array(7));
      eventRowsHa.push(new Array(7));
      eventRowsVa.push(new Array(7));
      eventRowsNotes.push(new Array(7));
    }

    for (let day = 0; day < 7; day++) {
      const cellDate = new Date(gridStart);
      cellDate.setDate(gridStart.getDate() + week * 7 + day);

      const inMonth =
        cellDate.getMonth() === monthDate.getMonth() &&
        cellDate.getFullYear() === monthDate.getFullYear();

      const key = dateKey(cellDate);
      const events = eventsByDate[key] || [];
      const isToday = key === todayKey;

      // Date cell: bold + colored, encoded in rich text. Today gets its own
      // highlight color set on the date cell only — event cells underneath
      // keep their category-color backgrounds intact.
      const dateColor = isToday
        ? CALENDAR.colors.todayFontColor
        : (inMonth ? theme.daysFontColor : theme.inactiveDaysFontColor);
      const dateText = formatCalendarDateLabel_(cellDate);
      const dateStyle = SpreadsheetApp.newTextStyle()
        .setForegroundColor(dateColor)
        .setFontFamily(fontFamily)
        .setBold(true)
        .build();
      dateRowRich[day] = SpreadsheetApp.newRichTextValue()
        .setText(dateText)
        .setTextStyle(0, dateText.length, dateStyle)
        .build();
      dateRowBg[day] = isToday
        ? CALENDAR.colors.todayBackground
        : (inMonth ? theme.daysBackground : theme.inactiveDaysBackground);
      dateRowHa[day] = "left";
      dateRowVa[day] = "middle";
      dateRowNotes[day] = "";

      const visibleEvents = events.slice(0, maxEvents);
      const overflowEvents = events.slice(maxEvents);

      for (let i = 0; i < maxEvents; i++) {
        const event = visibleEvents[i];
        let rich = blankRich;
        let bg = inMonth ? CALENDAR.colors.eventDefaultBackground : theme.inactiveDaysBackground;

        if (event) {
          const built = buildEventCellRichText_(event, inMonth, theme);
          rich = built.rich;
          bg = built.background;
        }

        eventRowsRich[i][day] = rich;
        eventRowsBg[i][day] = bg;
        eventRowsHa[i][day] = "left";
        eventRowsVa[i][day] = "top";
        eventRowsNotes[i][day] = "";
      }

      // Overflow cell: bold + colored, encoded in rich text.
      const overflowIdx = maxEvents;
      if (overflowEvents.length > 0) {
        const text = overflowEvents.length === 1
          ? "1 More..."
          : (overflowEvents.length + " More...");
        const ofColor = inMonth ? CALENDAR.colors.overflowFontColor : theme.inactiveDaysFontColor;
        const ofStyle = SpreadsheetApp.newTextStyle()
          .setForegroundColor(ofColor)
          .setFontFamily(fontFamily)
          .setBold(true)
          .build();
        eventRowsRich[overflowIdx][day] = SpreadsheetApp.newRichTextValue()
          .setText(text)
          .setTextStyle(0, text.length, ofStyle)
          .build();
        eventRowsBg[overflowIdx][day] = inMonth ? CALENDAR.colors.overflowBackground : theme.inactiveDaysBackground;
        eventRowsHa[overflowIdx][day] = "left";
        eventRowsVa[overflowIdx][day] = "top";
        eventRowsNotes[overflowIdx][day] = "Use Calendar Tools > Open Selected to view this date.";
      } else {
        eventRowsRich[overflowIdx][day] = blankRich;
        eventRowsBg[overflowIdx][day] = inMonth ? CALENDAR.colors.eventDefaultBackground : theme.inactiveDaysBackground;
        eventRowsHa[overflowIdx][day] = "left";
        eventRowsVa[overflowIdx][day] = "top";
        eventRowsNotes[overflowIdx][day] = "";
      }
    }

    richValues.push(dateRowRich);
    backgrounds.push(dateRowBg);
    horizontalAlignments.push(dateRowHa);
    verticalAlignments.push(dateRowVa);
    notes.push(dateRowNotes);

    for (let i = 0; i < maxEvents + 1; i++) {
      richValues.push(eventRowsRich[i]);
      backgrounds.push(eventRowsBg[i]);
      horizontalAlignments.push(eventRowsHa[i]);
      verticalAlignments.push(eventRowsVa[i]);
      notes.push(eventRowsNotes[i]);
    }
  }

  const body = sheet.getRange(bodyStartRow, 1, bodyRows, 7);
  body.clear({ contentsOnly: false });
  body.setFontFamily(fontFamily);
  body.setFontSize(9);
  body.setWrap(true);
  body.setRichTextValues(richValues);
  body.setBackgrounds(backgrounds);
  body.setHorizontalAlignments(horizontalAlignments);
  body.setVerticalAlignments(verticalAlignments);
  body.setNotes(notes);

  // Borders: top on each week's date row, bottom on the overflow row, left+right throughout.
  for (let week = 0; week < 6; week++) {
    const dateRow = bodyStartRow + week * rowsPerWeek;
    sheet.getRange(dateRow, 1, 1, 7)
      .setBorder(true, true, false, true, false, false, "#d6d6d6", SpreadsheetApp.BorderStyle.SOLID);
    if (maxEvents > 0) {
      sheet.getRange(dateRow + 1, 1, maxEvents, 7)
        .setBorder(false, true, false, true, false, false, "#d6d6d6", SpreadsheetApp.BorderStyle.SOLID);
    }
    sheet.getRange(dateRow + maxEvents + 1, 1, 1, 7)
      .setBorder(false, true, true, true, false, false, "#d6d6d6", SpreadsheetApp.BorderStyle.SOLID);
  }

  // Row heights — batched per week.
  for (let week = 0; week < 6; week++) {
    const dateRow = bodyStartRow + week * rowsPerWeek;
    sheet.setRowHeights(dateRow, 1, CALENDAR.layout.dateRowHeight);
    sheet.setRowHeights(dateRow + 1, maxEvents + 1, CALENDAR.layout.eventRowHeight);
  }

  const spacerRow = bodyStartRow + bodyRows;
  sheet.getRange(spacerRow, 1, 1, 7).clear({ contentsOnly: false });
  sheet.setRowHeight(spacerRow, CALENDAR.layout.spacerRowHeight);

  sheet.setRowHeight(startRow, 34);
  sheet.setRowHeight(startRow + 1, 24);
  sheet.setRowHeight(startRow + 2, 26);
}

function buildEventCellRichText_(event, inMonth, theme) {
  const parts = buildEventDisplayParts_(event);
  const text = parts.text;
  const background = inMonth && looksLikeColor_(event.categoryColor)
    ? event.categoryColor
    : (inMonth ? CALENDAR.colors.eventDefaultBackground : theme.inactiveDaysBackground);
  const fontColor = inMonth ? readableTextColor_(background) : theme.inactiveDaysFontColor;

  const baseStyle = SpreadsheetApp.newTextStyle()
    .setForegroundColor(fontColor)
    .setFontFamily(CALENDAR.setup.fontFamily)
    .setBold(false)
    .setItalic(false)
    .setUnderline(false)
    .build();

  const builder = SpreadsheetApp.newRichTextValue()
    .setText(text)
    .setTextStyle(0, text.length, baseStyle);

  if (event.titleUrl || event.url) {
    builder.setLinkUrl(0, parts.titleText.length, event.titleUrl || event.url);
    builder.setTextStyle(0, parts.titleText.length, baseStyle);
  }

  let cursor = parts.titleText.length;
  parts.additionalParts.forEach(part => {
    cursor += 1;
    const lineStart = cursor;
    const lineEnd = lineStart + part.text.length;
    const labelEnd = lineStart + Number(part.labelLength || 0);

    if (lineEnd > lineStart) {
      builder.setTextStyle(lineStart, lineEnd, baseStyle);
    }

    if (labelEnd > lineStart) {
      builder.setTextStyle(lineStart, labelEnd, buildAdditionalLineTextStyle_(part, fontColor));
    }

    cursor = lineEnd;
  });

  return { rich: builder.build(), background, fontColor };
}


function formatCalendarDateLabel_(date) {
  const format = String(CALENDAR.setup.dateFormat || "F").trim() || "F";
  return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), format);
}












function buildEventDisplayParts_(event) {
  const iconPrefix = event.typeIcon ? `${event.typeIcon} ` : "";
  const iconSuffix = event.statusIcon ? ` ${event.statusIcon}` : "";
  const titleText = `${iconPrefix}${event.title || event.label || "Untitled"}${iconSuffix}`;
  const additionalParts = buildAdditionalEventParts_(event.additional);

  return {
    titleText,
    additionalParts,
    text: additionalParts.length
      ? titleText + "\n" + additionalParts.map(part => part.text).join("\n")
      : titleText
  };
}

function buildAdditionalEventParts_(additional) {
  return (additional || [])
    .filter(item => String(item.value || "").trim() !== "")
    .map(item => {
      const name = applyCase_(item.name, item.labelCase);
      const value = applyCase_(item.value, item.valueCase);
      const hasLabel = !!item.showLabel;
      // Leading single space indents the line slightly under the event title.
      // Included in labelText so any label styling (bold/color) covers it too.
      const labelText = hasLabel ? ` ${name}:` : "";
      const text = hasLabel ? `${labelText} ${value}` : ` ${value}`;

      return {
        text,
        labelLength: labelText.length,
        style: String(item.style || "").trim().toLowerCase(),
        color: String(item.color || "").trim()
      };
    });
}



function buildAdditionalLineTextStyle_(part, fallbackColor) {
  let builder = SpreadsheetApp.newTextStyle()
    .setForegroundColor(looksLikeColor_(part.color) ? part.color : fallbackColor)
    .setFontFamily(CALENDAR.setup.fontFamily)
    .setUnderline(false);

  if (part.style === "bold") {
    builder = builder.setBold(true).setItalic(false);
  } else if (part.style === "italic") {
    builder = builder.setBold(false).setItalic(true);
  } else {
    builder = builder.setBold(false).setItalic(false);
  }

  return builder.build();
}

function buildEventDisplayText_(event) {
  const parts = buildEventDisplayParts_(event);
  return parts.text;
}



function buildAdditionalEventLines_(additional) {
  return buildAdditionalEventParts_(additional).map(part => part.text);
}





function showOverflowModal_(sheet, dateKeyValue) {
  const ss = sheet.getParent();
  const controls = getLiveControls_(sheet);
  const source = loadSourceData(ss, controls.sourceSpec);
  const keyConfig = readKeyConfig_(ss);
  const filteredRows = filterRowsForCalendar_(source.rows, source.headers, controls);
  const eventsByDate = indexEventsByDate(filteredRows, source.headers, source.sourceMeta, keyConfig);
  const events = eventsByDate[dateKeyValue] || [];

  const dateLabel = formatDateKeyForDisplay_(dateKeyValue);

  const rowsHtml = events.map((event) => {
    const bg = looksLikeColor_(event.categoryColor) ? event.categoryColor : "#FFFFFF";
    const fg = readableTextColor_(bg);
    const label = escapeHtml_(buildEventDisplayText_(event));
    const type = escapeHtml_(event.type || "");
    const category = escapeHtml_(event.category || "");
    const status = escapeHtml_(event.status || "");
    const href = escapeHtml_(event.titleUrl || event.url || "");

    return (
      '<div class="event" style="background:' + bg + ';color:' + fg + ';">' +
        '<a class="title" style="color:' + fg + ';" href="' + href + '" target="_blank">' + label + '</a>' +
        '<div class="meta">' +
          (type ? '<span>Type: ' + type + '</span>' : '') +
          (category ? '<span>Category: ' + category + '</span>' : '') +
          (status ? '<span>Status: ' + status + '</span>' : '') +
        '</div>' +
      '</div>'
    );
  }).join("");

  const html = HtmlService.createHtmlOutput(
    '<!doctype html><html><head><base target="_top"><style>' +
      'body{font-family:Arial,sans-serif;margin:0;padding:18px;background:#fff;color:#222;}' +
      'h1{font-size:18px;margin:0 0 4px;color:#4A0039;}' +
      '.sub{font-size:12px;color:#555;margin-bottom:14px;}' +
      '.event{border-radius:10px;padding:10px 12px;margin:0 0 10px;border:1px solid rgba(0,0,0,.12);white-space:pre-line;}' +
      '.title{font-size:14px;font-weight:700;text-decoration:none;white-space:pre-line;}' +
      '.title:hover{text-decoration:none;}' +
      '.meta{display:flex;flex-wrap:wrap;gap:8px;margin-top:6px;font-size:11px;opacity:.86;}' +
      '.empty{padding:12px;border:1px dashed #bbb;border-radius:10px;color:#555;}' +
    '</style></head><body>' +
      '<h1>' + escapeHtml_(dateLabel) + '</h1>' +
      '<div class="sub">Events link back to the source title cell.</div>' +
      (rowsHtml || '<div class="empty">No events found for this date.</div>') +
    '</body></html>'
  ).setWidth(460).setHeight(420);

  SpreadsheetApp.getUi().showModalDialog(html, "Selected date");
}



function formatDateKeyForDisplay_(dateKeyValue) {
  const parts = String(dateKeyValue || "").split("-");
  if (parts.length !== 3) return dateKeyValue;

  const date = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "EEEE, MMMM d, yyyy");
}

function getMonthTheme(monthIndexInView) {
  const themeIndex = monthIndexInView % 3;

  if (themeIndex === 0) {
    return {
      monthBackground: CALENDAR.colors.month1HeaderBackground,
      monthFontColor: CALENDAR.colors.month1HeaderFontColor,
      weekBackground: CALENDAR.colors.month1WeekBackground,
      weekFontColor: CALENDAR.colors.month1WeekFontColor,
      daysBackground: CALENDAR.colors.month1DaysBackground,
      daysFontColor: CALENDAR.colors.month1DaysFontColor,
      inactiveDaysBackground: CALENDAR.colors.month1DaysInactiveBackgroundColor,
      inactiveDaysFontColor: CALENDAR.colors.month1DaysInactiveFontColor,
    };
  }

  if (themeIndex === 1) {
    return {
      monthBackground: CALENDAR.colors.month2HeaderBackground,
      monthFontColor: CALENDAR.colors.month2HeaderFontColor,
      weekBackground: CALENDAR.colors.month2WeekBackground,
      weekFontColor: CALENDAR.colors.month2WeekFontColor,
      daysBackground: CALENDAR.colors.month2DaysBackground,
      daysFontColor: CALENDAR.colors.month2DaysFontColor,
      inactiveDaysBackground: CALENDAR.colors.month2DaysInactiveBackgroundColor,
      inactiveDaysFontColor: CALENDAR.colors.month2DaysInactiveFontColor,
    };
  }

  return {
    monthBackground: CALENDAR.colors.month3HeaderBackground,
    monthFontColor: CALENDAR.colors.month3HeaderFontColor,
    weekBackground: CALENDAR.colors.month3WeekBackground,
    weekFontColor: CALENDAR.colors.month3WeekFontColor,
    daysBackground: CALENDAR.colors.month3DaysBackground,
    daysFontColor: CALENDAR.colors.month3DaysFontColor,
    inactiveDaysBackground: CALENDAR.colors.month3DaysInactiveBackgroundColor,
    inactiveDaysFontColor: CALENDAR.colors.month3DaysInactiveFontColor,
  };
}

function getWeekdayLabels() {
  const tz = Session.getScriptTimeZone();
  const base = new Date(2024, 0, 7); // Sunday
  const startIndex = getWeekStartIndex_();
  const out = [];

  for (let i = 0; i < 7; i++) {
    const d = new Date(base);
    d.setDate(base.getDate() + startIndex + i);
    out.push(Utilities.formatDate(d, tz, CALENDAR.setup.dayFormat));
  }

  return out;
}

function getWeekStartIndex_() {
  const days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  const configured = String(CALENDAR.setup.startWeekOn || "Sunday").trim();
  if (configured === "Monday-CompressedWeekend") return 1;
  const index = days.indexOf(configured);

  return index >= 0 ? index : 0;
}

function isCompressedWeekend_() {
  return String(CALENDAR.setup.startWeekOn || "").trim() === "Monday-CompressedWeekend";
}

// Resolves one or more source sheets from G1 (sourceSpec) and returns a
// unified rows/headers structure.
//
// Multi-source: when sourceSpec parses (via parseFilterValues_) to more than
// one sheet name, headers from each resolved sheet are merged in first-seen
// order, each row is realigned to that unioned header layout, and each row
// carries its own __sourceMeta so the title link in indexEventsByDate routes
// back to the correct sheet. Missing columns on a given source contribute
// blank cells for that row — no warnings.
//
// Empty spec falls back to defaultDataSheetName (then to the first sheet).
// Non-empty spec that resolves to zero existing sheets surfaces a single
// toast per render session.
function loadSourceData(ss, sourceSpec) {
  const cacheKey = String(sourceSpec || "");

  if (RENDER_SESSION.active && Object.prototype.hasOwnProperty.call(RENDER_SESSION.sourceData, cacheKey)) {
    return RENDER_SESSION.sourceData[cacheKey];
  }

  const resolved = resolveSources_(ss, sourceSpec);

  const headerOrder = [];
  const headerIndex = {};
  const aggregatedRows = [];
  let firstSourceMeta = null;

  resolved.sources.forEach(parsed => {
    if (!parsed.sheet) return;
    const values = parsed.sheet.getDataRange().getValues();
    if (!values || values.length === 0) return;

    const sourceHeaders = values[0].map(v => String(v || "").trim());
    const remap = sourceHeaders.map(h => {
      const norm = normalizeKey_(h);
      if (norm === "") return -1;
      if (!Object.prototype.hasOwnProperty.call(headerIndex, norm)) {
        headerIndex[norm] = headerOrder.length;
        headerOrder.push(h);
      }
      return headerIndex[norm];
    });

    const perSourceMeta = {
      spreadsheetName: parsed.spreadsheetName,
      spreadsheetId: parsed.spreadsheetId,
      sheetId: parsed.sheetId,
      sheetName: parsed.sheetName,
      sourceType: parsed.sourceType,
    };
    if (!firstSourceMeta) firstSourceMeta = perSourceMeta;

    values.slice(1).forEach((row, index) => {
      if (!row.some(cell => String(cell || "").trim() !== "")) return;
      const aligned = new Array(headerOrder.length).fill("");
      for (let i = 0; i < sourceHeaders.length; i++) {
        const target = remap[i];
        if (target >= 0) aligned[target] = row[i];
      }
      aligned.__sourceRowNumber = index + 2;
      aligned.__sourceMeta = perSourceMeta;
      aggregatedRows.push(aligned);
    });
  });

  // Earlier sources' rows may be shorter than the final header count if a
  // later source introduced new columns. Pad them so every row matches.
  const finalLen = headerOrder.length;
  if (finalLen > 0) {
    aggregatedRows.forEach(row => {
      while (row.length < finalLen) row.push("");
    });
  }

  const result = {
    spreadsheetName: firstSourceMeta ? firstSourceMeta.spreadsheetName : ss.getName(),
    sheetName: resolved.sources.map(s => s.sheetName).filter(Boolean).join(", "),
    headers: headerOrder,
    rows: aggregatedRows,
    sourceMeta: firstSourceMeta || {
      spreadsheetName: ss.getName(),
      spreadsheetId: ss.getId(),
      sheetName: "",
      sheetId: null,
      sourceType: "local",
    },
  };

  if (RENDER_SESSION.active) {
    RENDER_SESSION.sourceData[cacheKey] = result;
  }

  if (resolved.unresolvedSpec) {
    try {
      const seen = RENDER_SESSION.toastedSpecs || (RENDER_SESSION.toastedSpecs = {});
      if (!seen[cacheKey]) {
        seen[cacheKey] = true;
        ss.toast(
          'No source sheets matched "' + resolved.unresolvedSpec + '". Set G1 to a tab name.',
          CALENDAR.menuName,
          6
        );
      }
    } catch (err) {}
  }

  return result;
}

// Parses sourceSpec into 1+ resolved source descriptors. Empty spec falls
// back to defaultDataSheetName, then to the first sheet. Multi-value spec
// is parsed with parseFilterValues_ (comma / semicolon / newline) — same
// pattern as the A2/B2 OR/OR filter so users see one consistent rule.
function resolveSources_(activeSpreadsheet, sourceSpec) {
  const spec = String(sourceSpec || "").trim();

  if (!spec) {
    return { sources: [defaultSourceMeta_(activeSpreadsheet)], unresolvedSpec: "" };
  }

  const names = parseFilterValues_(spec);
  const sources = [];
  names.forEach(name => {
    const sheet = activeSpreadsheet.getSheetByName(name);
    if (!sheet) return;
    sources.push({
      spreadsheetName: activeSpreadsheet.getName(),
      spreadsheetId: activeSpreadsheet.getId(),
      sheet: sheet,
      sheetName: sheet.getName(),
      sheetId: sheet.getSheetId(),
      sourceType: "local",
    });
  });

  if (sources.length) return { sources: sources, unresolvedSpec: "" };
  return { sources: [], unresolvedSpec: spec };
}

function defaultSourceMeta_(activeSpreadsheet) {
  const dataSheet = activeSpreadsheet.getSheetByName(CALENDAR.defaultDataSheetName);
  if (dataSheet) {
    return {
      spreadsheetName: activeSpreadsheet.getName(),
      spreadsheetId: activeSpreadsheet.getId(),
      sheet: dataSheet,
      sheetName: dataSheet.getName(),
      sheetId: dataSheet.getSheetId(),
      sourceType: "local",
    };
  }

  const fallback = activeSpreadsheet.getSheets()[0];
  return {
    spreadsheetName: activeSpreadsheet.getName(),
    spreadsheetId: activeSpreadsheet.getId(),
    sheet: fallback,
    sheetName: fallback.getName(),
    sheetId: fallback.getSheetId(),
    sourceType: "local",
  };
}

function filterRowsForCalendar_(rows, headers, controls) {
  const defaultLabel = CALENDAR.setup.filterDefaultLabel;
  const filterField = String(controls.filterField || defaultLabel).trim();
  const filterValue = String(controls.filterValue || "").trim();

  if (!filterField || filterField === defaultLabel || !filterValue) return rows;

  // A2 supports multi-select: parse the field cell the same way as the value
  // cell. Default-label entries are filtered out so they don't act as columns.
  const filterColumns = parseFilterValues_(filterField).filter(c => c !== defaultLabel);
  if (!filterColumns.length) return rows;

  const colIndices = [];
  filterColumns.forEach(col => {
    const idx = findColumnIndex(headers, [col]);
    if (idx >= 0) colIndices.push(idx);
  });
  if (!colIndices.length) return rows;

  const selectedValues = parseFilterValues_(filterValue);
  if (!selectedValues.length) return rows;

  const selectedMap = {};
  selectedValues.forEach(v => { selectedMap[normalizeKey_(v)] = true; });

  // OR across columns AND across values: a row passes if any selected column
  // contains any selected value. Each row matches at most once regardless of
  // how many columns hit (the loop short-circuits on first true).
  return (rows || []).filter(row => {
    for (let i = 0; i < colIndices.length; i++) {
      const v = String(row[colIndices[i]] || "").trim();
      if (selectedMap[normalizeKey_(v)]) return true;
    }
    return false;
  });
}

function parseFilterValues_(value) {
  return String(value || "")
    .split(/[,;\n]/)
    .map(part => String(part || "").trim())
    .filter(part => part !== "");
}

function readKeyConfig_(ss) {
  if (RENDER_SESSION.active && RENDER_SESSION.keyConfig) {
    return RENDER_SESSION.keyConfig;
  }

  const config = {
    typeIcons: {},
    categoryColors: {},
    statusIcons: {}
  };

  CALENDAR.keyDefaults.types.forEach(row => {
    config.typeIcons[normalizeKey_(row[0])] = String(row[1] || "");
  });

  CALENDAR.keyDefaults.categories.forEach(row => {
    config.categoryColors[normalizeKey_(row[0])] = String(row[1] || "");
  });

  CALENDAR.keyDefaults.statuses.forEach(row => {
    config.statusIcons[normalizeKey_(row[0])] = String(row[1] || "");
  });

  const sheet = ss.getSheetByName(CALENDAR.keySheetName);
  if (!sheet) return config;

  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return config;

  const header = values[0].map(v => String(v || "").trim());

  const typeCol = findColumnIndex(header, ["Type"]);
  const typeIconCol = findColumnIndex(header, ["Type-Icon", "Type Icon", "Icon"]);
  const categoryCol = findColumnIndex(header, ["Category"]);
  const categoryColorCol = findColumnIndex(header, ["Category-Color", "Category Color", "Color"]);
  const statusCol = findColumnIndex(header, ["Status"]);
  const statusIconCol = findColumnIndex(header, ["Status-Icon", "Status Icon"]);

  values.slice(1).forEach(row => {
    if (typeCol >= 0 && typeIconCol >= 0) {
      const key = normalizeKey_(row[typeCol]);
      const value = String(row[typeIconCol] || "").trim();
      if (key && value) config.typeIcons[key] = value;
    }

    if (categoryCol >= 0 && categoryColorCol >= 0) {
      const key = normalizeKey_(row[categoryCol]);
      const value = String(row[categoryColorCol] || "").trim();
      if (key && looksLikeColor_(value)) config.categoryColors[key] = value;
    }

    if (statusCol >= 0 && statusIconCol >= 0) {
      const key = normalizeKey_(row[statusCol]);
      const value = String(row[statusIconCol] || "").trim();
      if (key && value) config.statusIcons[key] = value;
    }
  });

  if (RENDER_SESSION.active) RENDER_SESSION.keyConfig = config;
  return config;
}

function indexEventsByDate(rows, headers, sourceMeta, keyConfig) {
  const map = {};
  if (!rows || !rows.length || !headers || !headers.length) return map;

  const customDate = CALENDAR.setup.customDate;
  const customTitle = CALENDAR.setup.customTitle;
  const additionalConfig = parseAdditionalConfig_(CALENDAR.setup.customAdditional);

  const dateCol = findColumnIndex(headers, [
    customDate,
    "Date",
    "Start Date",
    "Event Date",
    "Date/Time",
    "Date Time",
    "Current MMDD",
    "Original MMDD",
    "MMDD",
  ]);

  const titleCol = findColumnIndex(headers, [
    customTitle,
    "Title",
    "Event Title",
    "Name",
    "Initiative",
    "Topic",
    "Engagement Type",
  ]);

  const customType = CALENDAR.setup.customType;
  const customCategory = CALENDAR.setup.customCategory;
  const customStatus = CALENDAR.setup.customStatus;

  const typeCol = findColumnIndex(headers, [
    customType,
    "Type",
    "Event Type",
    "Channel",
    "Tactic Type",
  ]);

  const categoryCol = findColumnIndex(headers, [
    customCategory,
    "Category",
    "Event Category",
    "Theme",
    "Product",
    "Pillar",
  ]);

  const statusCol = findColumnIndex(headers, [
    customStatus,
    "Status",
    "Event Status",
  ]);

  const descriptionCol = findColumnIndex(headers, [
    "Description",
    "Desc",
    "Event Description",
    "Details",
  ]);

  const additionalCols = additionalConfig
    .map((cfg) => ({
      ...cfg,
      index: findColumnIndex(headers, [cfg.name]),
    }))
    .filter((col) => col.index >= 0);

  const effectiveTitleCol = titleCol >= 0 ? titleCol : 0;
  const linkTargetCol = descriptionCol >= 0 ? descriptionCol : (headers.length > 1 ? 1 : effectiveTitleCol);

  if (dateCol < 0) return map;

  rows.forEach((row, index) => {
    const sourceRowNumber = row.__sourceRowNumber || (index + 2);
    const dates = parseDateValue(row[dateCol]);
    if (!dates.length) return;

    const title = effectiveTitleCol >= 0 ? String(row[effectiveTitleCol] || "").trim() : "";
    const type = typeCol >= 0 ? String(row[typeCol] || "").trim() : "";
    const category = categoryCol >= 0 ? String(row[categoryCol] || "").trim() : "";
    const status = statusCol >= 0 ? String(row[statusCol] || "").trim() : "";

    const label = title || "Untitled";
    const typeIcon = type ? keyConfig.typeIcons[normalizeKey_(type)] || "" : "";
    const statusIcon = status ? keyConfig.statusIcons[normalizeKey_(status)] || "" : "";
    const categoryColor = category ? keyConfig.categoryColors[normalizeKey_(category)] || "" : "";

    // Each row may carry its own __sourceMeta when sourceSpec resolves to
    // multiple sheets (v13.15.0+). Fall back to the calendar-wide sourceMeta
    // when the row didn't get tagged (single-source path).
    const rowSourceMeta = row.__sourceMeta || sourceMeta;
    const titleUrl = buildSourceCellUrl(rowSourceMeta, sourceRowNumber, linkTargetCol + 1);
    const sourceUrl = titleUrl;

    const additional = additionalCols.map((col) => ({
      name: col.name,
      value: String(row[col.index] || "").trim(),
      showLabel: col.showLabel,
      style: col.style,
      color: col.color,
      labelCase: col.labelCase,
      valueCase: col.valueCase,
    }));

    dates.forEach((d) => {
      const key = dateKey(d);
      if (!map[key]) map[key] = [];

      map[key].push({
        label,
        url: sourceUrl,
        titleUrl,
        title,
        type,
        category,
        status,
        typeIcon,
        statusIcon,
        categoryColor,
        detailDate: String(row[dateCol] || "").trim(),
        additional,
        sourceRow: sourceRowNumber,
      });
    });
  });

  // Sort each day's events per the CALENDAR.setup.eventSortOrder spec.
  // No-op if the spec is blank or all entries parse as unknown.
  const sortSpec = parseEventSortSpec_(CALENDAR.setup.eventSortOrder);
  if (sortSpec.length > 0) {
    Object.keys(map).forEach(key => {
      sortDayEvents_(map[key], sortSpec);
    });
  }

  return map;
}

// Parses a value like "Category ↓, Status ↑, Alphabetical ↓" into:
//   [{field: "Category", dir: "Asc"}, {field: "Status", dir: "Desc"}, ...]
// Accepts arrays (joined first), comma/semicolon/newline separators, and
// "Asc"/"Desc" or "↓"/"↑" direction tokens. Unknown segments are dropped.
function parseEventSortSpec_(raw) {
  if (Array.isArray(raw)) raw = raw.join(",");
  return String(raw || "")
    .split(/[,;\n]/)
    .map(s => String(s || "").trim())
    .filter(s => s !== "")
    .map(s => {
      let m = s.match(/^(.+?)\s*([↓↑])$/);
      if (m) return { field: m[1].trim(), dir: m[2] === "↓" ? "Asc" : "Desc" };
      m = s.match(/^(.+?)\s+(asc|desc)$/i);
      if (m) return { field: m[1].trim(), dir: m[2].toLowerCase() === "asc" ? "Asc" : "Desc" };
      return null;
    })
    .filter(p => p);
}

function sortDayEvents_(events, sortSpec) {
  if (!events || events.length < 2) return;
  // Stable-sort tie-breaker via the original index.
  events.forEach((ev, i) => { ev._sortIndex = i; });
  events.sort((a, b) => {
    for (let i = 0; i < sortSpec.length; i++) {
      const va = eventSortValue_(a, sortSpec[i].field);
      const vb = eventSortValue_(b, sortSpec[i].field);
      if (va < vb) return sortSpec[i].dir === "Asc" ? -1 : 1;
      if (va > vb) return sortSpec[i].dir === "Asc" ? 1 : -1;
    }
    return a._sortIndex - b._sortIndex;
  });
  events.forEach(ev => { delete ev._sortIndex; });
}

function eventSortValue_(event, field) {
  switch (String(field || "").toLowerCase()) {
    case "category":     return String(event.category || "").toLowerCase();
    case "status":       return String(event.status || "").toLowerCase();
    case "type":         return String(event.type || "").toLowerCase();
    case "alphabetical":
    case "title":        return String(event.title || "").toLowerCase();
    case "source":       return Number(event.sourceRow || 0);
    default:             return "";
  }
}

function parseAdditionalConfig_(arr) {
  const defaultShowLabel = !!CALENDAR.setup.customAdditionalLabels;
  const defaultStyle = String(CALENDAR.setup.customAdditionalLabelsStyle || "").trim().toLowerCase();

  return (arr || [])
    .map((entry) => String(entry || "").trim())
    .filter((entry) => entry !== "")
    .map((entry) => {
      const parts = entry.split(",").map((p) => String(p || "").trim()).filter((p) => p !== "");
      const name = parts[0] || "";

      const cfg = {
        name,
        showLabel: defaultShowLabel,
        style: defaultStyle,
        color: "",
        labelCase: "",
        valueCase: "",
      };

      for (let i = 1; i < parts.length; i++) {
        const token = parts[i];
        const lower = token.toLowerCase();

        if (lower === "true" || lower === "false") {
          cfg.showLabel = lower === "true";
          continue;
        }

        if (lower === "bold" || lower === "italic") {
          cfg.style = lower;
          continue;
        }

        if (lower === "uppercase" || lower === "lowercase") {
          cfg.labelCase = lower;
          cfg.valueCase = lower;
          continue;
        }

        if (lower === "label uppercase" || lower === "label lowercase") {
          cfg.labelCase = lower.replace("label ", "");
          continue;
        }

        if (lower === "value uppercase" || lower === "value lowercase") {
          cfg.valueCase = lower.replace("value ", "");
          continue;
        }

        if (looksLikeColor_(token)) {
          cfg.color = token;
          continue;
        }
      }

      return cfg;
    });
}



function looksLikeColor_(value) {
  const s = String(value || "").trim();
  if (!s) return false;
  if (/^#[0-9a-f]{3}([0-9a-f]{3})?$/i.test(s)) return true;
  if (/^(red|blue|green|black|white|gray|grey|purple|orange|yellow|pink|teal|navy|maroon|brown|gold|silver)$/i.test(s)) return true;
  return false;
}

function readableTextColor_(background) {
  const hex = colorNameToHex_(background);
  if (!hex) return CALENDAR.colors.eventDefaultFontColor;

  const normalized = normalizeHexColor_(hex);
  if (!normalized) return CALENDAR.colors.eventDefaultFontColor;

  const r = parseInt(normalized.substring(1, 3), 16);
  const g = parseInt(normalized.substring(3, 5), 16);
  const b = parseInt(normalized.substring(5, 7), 16);
  const luminance = (0.299 * r + 0.587 * g + 0.114 * b);

  return luminance > 150 ? "#000000" : "#FFFFFF";
}

function normalizeHexColor_(value) {
  const s = String(value || "").trim();

  let m = s.match(/^#([0-9a-f]{3})$/i);
  if (m) {
    return "#" + m[1].split("").map(ch => ch + ch).join("").toUpperCase();
  }

  m = s.match(/^#([0-9a-f]{6})$/i);
  if (m) {
    return "#" + m[1].toUpperCase();
  }

  return "";
}

function colorNameToHex_(value) {
  const s = String(value || "").trim().toLowerCase();
  const map = {
    red: "#FF0000",
    blue: "#0000FF",
    green: "#008000",
    black: "#000000",
    white: "#FFFFFF",
    gray: "#808080",
    grey: "#808080",
    purple: "#800080",
    orange: "#FFA500",
    yellow: "#FFFF00",
    pink: "#FFC0CB",
    teal: "#008080",
    navy: "#000080",
    maroon: "#800000",
    brown: "#A52A2A",
    gold: "#FFD700",
    silver: "#C0C0C0",
  };

  if (map[s]) return map[s];
  return value;
}

function applyCase_(value, mode) {
  const text = String(value || "");
  const lowerMode = String(mode || "").trim().toLowerCase();

  if (lowerMode === "uppercase") return text.toUpperCase();
  if (lowerMode === "lowercase") return text.toLowerCase();

  return text;
}

function buildSourceCellUrl(sourceMeta, rowNumber, columnNumber) {
  const colLetter = columnToLetter(columnNumber);
  return `https://docs.google.com/spreadsheets/d/${sourceMeta.spreadsheetId}/edit#gid=${sourceMeta.sheetId}&range=${colLetter}${rowNumber}`;
}

function columnToLetter(columnNumber) {
  let dividend = columnNumber;
  let columnLetter = "";

  while (dividend > 0) {
    const modulo = (dividend - 1) % 26;
    columnLetter = String.fromCharCode(65 + modulo) + columnLetter;
    dividend = Math.floor((dividend - modulo) / 26);
  }

  return columnLetter;
}

function parseDateValue(value) {
  const out = [];
  if (value === null || value === undefined || value === "") return out;

  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    out.push(normalizeSpreadsheetDate_(value));
    return out;
  }

  const text = String(value).trim();
  if (!text) return out;

  // Only treat as a range if the separator is surrounded by whitespace.
  // Prevents titles like "Roadshow - East" or ISO dates like "2026-05-04" from being mis-read.
  const rangeMatch = text.match(/^(.+?)\s+(?:-|to)\s+(.+)$/i);

  if (rangeMatch) {
    const start = parseSingleDate(rangeMatch[1]);
    const end = parseSingleDate(rangeMatch[2]);

    if (start && end) {
      const cursor = new Date(start);

      while (cursor <= end) {
        out.push(new Date(cursor));
        cursor.setDate(cursor.getDate() + 1);
      }

      return out;
    }
    // Partial / non-date range: fall through and try the whole text as a single date.
  }

  const single = parseSingleDate(text);
  if (single) out.push(single);

  return out;
}

function parseSingleDate(text) {
  const s = String(text || "").trim();
  if (!s) return null;

  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));

  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (m) {
    let year = Number(m[3]);
    if (year < 100) year += 2000;
    return new Date(year, Number(m[1]) - 1, Number(m[2]));
  }

  m = s.match(/^(\d{1,2})\/(\d{1,2})$/);
  if (m) {
    return new Date(currentYear(), Number(m[1]) - 1, Number(m[2]));
  }

  return null;
}

function dateKey(date) {
  return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

// The spreadsheet's timezone — used only when normalizing typed Date cells
// read from the sheet. Cached because each lookup is a service call. Always
// returns a non-empty string; falls back to the Apps Script project TZ if the
// spreadsheet lookup throws, returns null, or returns a non-string.
let __CACHED_CALENDAR_TZ = null;
function calendarTimeZone_() {
  if (__CACHED_CALENDAR_TZ) return __CACHED_CALENDAR_TZ;
  let tz = null;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) tz = ss.getSpreadsheetTimeZone();
  } catch (err) {
    tz = null;
  }
  if (!tz || typeof tz !== "string") {
    tz = Session.getScriptTimeZone();
  }
  __CACHED_CALENDAR_TZ = tz;
  return tz;
}

// Source date cells can land in two different shapes, both of which cause the
// "events shifted by one day" bug if not handled:
//
//   1. Midnight in the spreadsheet's TZ — the normal case when a date is
//      typed into the Sheets UI. Shifts when the script project's TZ differs
//      from the spreadsheet's TZ.
//   2. Midnight UTC regardless of spreadsheet TZ — the common case when dates
//      were imported from CSV / BigQuery / another system / an Apps Script
//      project running in UTC. Sheets displays them in UTC even when the
//      spreadsheet's configured TZ is something else, but Apps Script's
//      Utilities.formatDate uses the configured TZ and the day shifts.
//
// In both cases the fix is the same: extract the Y/M/D the spreadsheet
// actually shows for this instant, then reconstruct a Date at midnight in
// the script TZ. After that, the rest of the script can use plain JS Date
// math safely.
//
// Detection: an instant whose milliseconds are an exact multiple of one day
// (86,400,000 ms) is midnight UTC, so interpret it in UTC. Anything else is
// midnight in some named TZ, so interpret it in the spreadsheet's configured
// TZ.
//
// Defensive: if the format call fails for any reason, return the original
// Date so the calendar still renders (without the fix) instead of crashing.
function normalizeSpreadsheetDate_(value) {
  if (Object.prototype.toString.call(value) !== "[object Date]") return value;
  if (isNaN(value.getTime())) return value;

  const scriptTz = Session.getScriptTimeZone();
  const interpretTz = (value.getTime() % 86400000 === 0)
    ? "UTC"
    : calendarTimeZone_();

  if (!interpretTz || typeof interpretTz !== "string" || interpretTz === scriptTz) {
    return value;
  }

  try {
    const ymd = Utilities.formatDate(value, interpretTz, "yyyy-MM-dd").split("-");
    return new Date(Number(ymd[0]), Number(ymd[1]) - 1, Number(ymd[2]));
  } catch (err) {
    Logger.log("normalizeSpreadsheetDate_ skipped: " + err.message);
    return value;
  }
}

function findColumnIndex(headers, candidates) {
  const normHeaders = headers.map(normalizeHeader_);

  for (const cand of candidates) {
    const idx = normHeaders.indexOf(normalizeHeader_(cand));
    if (idx >= 0) return idx;
  }

  return -1;
}

function normalizeHeader_(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "");
}

function normalizeKey_(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function setDropdownValidation(range, values, allowInvalid) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(!!allowInvalid)
    .build();

  range.setDataValidation(rule);
}

function getDefaultSourceSpec(ss) {
  const dataSheet = ss.getSheetByName(CALENDAR.defaultDataSheetName);
  if (dataSheet) return CALENDAR.defaultDataSheetName;

  const sheets = ss.getSheets().filter(sheet => sheet.getName() !== CALENDAR.keySheetName);
  return sheets.length ? sheets[0].getName() : CALENDAR.defaultDataSheetName;
}

function uniqueSheetName(ss, baseName) {
  let name = baseName;
  let n = 2;

  while (ss.getSheetByName(name)) {
    name = baseName + " (" + n + ")";
    n++;
  }

  return name;
}

function isCalendarSheet(sheet) {
  return String(sheet.getRange("C2").getValue() || "") === "Starting:" ||
         String(sheet.getRange("F1").getValue() || "") === "Source Data:";
}

function sheetHasContent(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow === 0 || lastCol === 0) return false;

  const sample = sheet.getRange(1, 1, Math.min(lastRow, 20), Math.min(lastCol, 10)).getValues();
  return sample.some(row => row.some(cell => String(cell || "").trim() !== ""));
}

function showWorkingModal(message) {
  const html = HtmlService.createHtmlOutput(
    '<!doctype html><html><head><base target="_top"><style>' +
    'body{font-family:Arial,sans-serif;margin:0;padding:20px;text-align:center;}' +
    '.box{display:inline-block;padding:18px 22px;border:1px solid #d0d0d0;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.08);}' +
    '.spin{width:28px;height:28px;margin:0 auto 12px;border:4px solid #d9d9d9;border-top:4px solid #4A0039;border-radius:50%;animation:spin 1s linear infinite;}' +
    '@keyframes spin{0%{transform:rotate(0deg)}100%{transform:rotate(360deg)}}' +
    '.msg{font-size:14px;color:#333;}' +
    '</style></head><body><div class="box"><div class="spin"></div><div class="msg">' +
    escapeHtml_(message || "Working...") +
    '</div></div><script>setTimeout(function(){google.script.host.close();}, 1200);</script></body></html>'
  ).setWidth(270).setHeight(160);

  SpreadsheetApp.getUi().showModalDialog(html, "Please wait");
}

function escapeHtml_(text) {
  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function currentYear() {
  return CALENDAR.defaultYear || (new Date().getFullYear());
}

function currentMonthName() {
  const tz = Session.getScriptTimeZone();
  return Utilities.formatDate(new Date(), tz, "MMMM");
}

function monthNameToNumber(monthName) {
  const months = CALENDAR.monthOptions;
  const idx = months.indexOf(String(monthName || "").trim());

  if (idx < 0) throw new Error("Invalid month name: " + monthName);

  return idx + 1;
}

function isQuarter(period) {
  return /^Q[1-4]$/i.test(String(period || "").trim());
}

function isMonth(value) {
  return CALENDAR.monthOptions.indexOf(String(value || "").trim()) >= 0;
}

function isYear(period) {
  return String(period || "").trim() === "Year";
}

function isCustom(period) {
  return String(period || "").trim() === "Custom";
}

function syncModeDefaults_(sheet) {
  const period = String(sheet.getRange("A1").getValue() || "").trim();
  const year = Number(sheet.getRange("B1").getValue() || currentYear());

  if (isYear(period)) {
    sheet.getRange("D1").setValue(12);

    const currentStart = sheet.getRange("D2").getValue();
    if (!(currentStart instanceof Date) || isNaN(currentStart.getTime())) {
      sheet.getRange("D2").setValue(getStartDateForMode_(period, year));
    }

    return;
  }

  if (isCustom(period)) {
    const n = Number(sheet.getRange("D1").getValue());

    if (!n || isNaN(n) || n < 1) {
      sheet.getRange("D1").setValue(1);
    }

    return;
  }

  if (isQuarter(period) || isMonth(period)) {
    sheet.getRange("D1").setValue(getDisplayedMonthCount(period, 1));
    sheet.getRange("D2").setValue(getStartDateForMode_(period, year));
  }
}


function buildKeySetupRows_() {
  return KEY_SETUP_OPTIONS.map(option => {
    let value = "";

    if (Object.prototype.hasOwnProperty.call(KEY_INITIAL_VALUES, option)) {
      value = KEY_INITIAL_VALUES[option];
    } else if (Object.prototype.hasOwnProperty.call(CALENDAR.setup, option)) {
      value = CALENDAR.setup[option];
    } else if (Object.prototype.hasOwnProperty.call(CALENDAR, option)) {
      value = CALENDAR[option];
    }

    return ["", option, serializeKeyValue_(value)];
  });
}



function buildKeyAppearanceRows_() {
  return KEY_APPEARANCE_OPTIONS
    .filter(option => Object.prototype.hasOwnProperty.call(CALENDAR.colors, option))
    .map(option => [option, CALENDAR.colors[option]]);
}



function serializeKeyValue_(value) {
  if (Array.isArray(value)) return value.join(", ");
  return value;
}





function applyKeyOptionNotes_(sheet) {
  const notes = {
    dayFormat: "DateTime format for weekday labels. Examples: E = Tue, EEEE = Tuesday.",
    dateFormat: "DateTime format for calendar date labels. Examples: d = 4, dd = 04, EE d = Tue 4.",
    monthFormat: "DateTime format for month headers. Examples: M = 9, MM = 09, MMM = Sep, MMMM = September.",
  };

  Object.keys(notes).forEach(optionName => {
    const row = findKeyOptionRow_(sheet, optionName);
    if (!row) return;

    sheet.getRange(row, 12).setNote(notes[optionName]);
  });
}



function applyKeySheetFormatting_(sheet, maxRows) {
  const fontFamily = CALENDAR.setup.fontFamily || "Inter";

  sheet.getRange(1, 1, maxRows, 15)
    .setFontFamily(fontFamily)
    .setFontSize(10)
    .setVerticalAlignment("middle");

  sheet.getRange("A1:B1")
    .setBackground(CALENDAR.colors.titleBackground)
    .setFontColor(CALENDAR.colors.titleFontColor)
    .setFontWeight("bold")
    .setHorizontalAlignment("left");

  sheet.getRange("D1:E1")
    .setBackground(CALENDAR.colors.titleBackground)
    .setFontColor(CALENDAR.colors.titleFontColor)
    .setFontWeight("bold")
    .setHorizontalAlignment("left");

  sheet.getRange("G1:H1")
    .setBackground(CALENDAR.colors.titleBackground)
    .setFontColor(CALENDAR.colors.titleFontColor)
    .setFontWeight("bold")
    .setHorizontalAlignment("left");

  sheet.getRange("J1:L1")
    .setBackground(CALENDAR.colors.titleBackground)
    .setFontColor(CALENDAR.colors.titleFontColor)
    .setFontFamily(fontFamily)
    .setFontWeight("bold")
    .setHorizontalAlignment("left");

  sheet.getRange("N1:O1")
    .setBackground(CALENDAR.colors.titleBackground)
    .setFontColor(CALENDAR.colors.titleFontColor)
    .setFontFamily(fontFamily)
    .setFontWeight("bold")
    .setHorizontalAlignment("left");

  sheet.getRange("C:C").setBackground("#000000");
  sheet.getRange("F:F").setBackground("#000000");
  sheet.getRange("I:I").setBackground("#666666");
  sheet.getRange("M:M").setBackground("#666666");

  sheet.getRange("A:A").setHorizontalAlignment("left");
  sheet.getRange("B:B")
    .setHorizontalAlignment("center")
    .setFontSize(14);
  sheet.getRange("D:D").setHorizontalAlignment("left");
  sheet.getRange("E:E").setHorizontalAlignment("center");
  sheet.getRange("G:G").setHorizontalAlignment("left");
  sheet.getRange("H:H").setHorizontalAlignment("center");

  sheet.getRange("J:J")
    .setHorizontalAlignment("left")
    .setFontColor("#666666");

  sheet.getRange("J1")
    .setFontColor(CALENDAR.colors.titleFontColor);

  sheet.getRange("K:K")
    .setFontFamily("Courier New")
    .setHorizontalAlignment("left");

  sheet.getRange("K1")
    .setFontFamily(fontFamily);

  sheet.getRange("L:L")
    .setHorizontalAlignment("left");

  sheet.getRange("N:N")
    .setFontFamily("Courier New")
    .setHorizontalAlignment("left");

  sheet.getRange("N1")
    .setFontFamily(fontFamily);

  sheet.getRange("O:O")
    .setHorizontalAlignment("center");

  sheet.getRange("A1:O1")
    .setFontSize(10)
    .setFontWeight("bold");

  sheet.setRowHeight(1, 28);

  if (maxRows > 1) {
    sheet.getRange(2, 1, maxRows - 1, 15)
      .setFontSize(10)
      .setFontWeight("normal");

    sheet.getRange(2, 2, maxRows - 1, 1)
      .setFontSize(14);
  }
}



function setCheckboxIfOption_(sheet, optionName) {
  const row = findKeyOptionRow_(sheet, optionName);
  if (!row) return;

  sheet.getRange(row, 12).insertCheckboxes();
}



function findKeyOptionRow_(sheet, optionName) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const values = sheet.getRange(2, 11, lastRow - 1, 1).getValues();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || "").trim() === optionName) {
      return i + 2;
    }
  }

  return 0;
}



function applyKeySetupConditionalFormatting_(sheet) {
  // Reserved for future Key tab setup formatting rules.
}






function getKeyBooleanOption_(ss, optionName, defaultValue) {
  const sheet = ss.getSheetByName(CALENDAR.keySheetName);
  if (!sheet) return defaultValue;

  const row = findKeyOptionRow_(sheet, optionName);
  if (!row) return defaultValue;

  const value = sheet.getRange(row, 12).getValue();

  if (value === true || value === false) return value;

  const text = String(value).trim().toLowerCase();
  if (text === "true") return true;
  if (text === "false") return false;

  return defaultValue;
}

function applyKeyOverrides_(ss) {
  if (RENDER_SESSION.active && RENDER_SESSION.keyOverridesApplied) return;

  const sheet = ss.getSheetByName(CALENDAR.keySheetName);
  if (!sheet) {
    if (RENDER_SESSION.active) RENDER_SESSION.keyOverridesApplied = true;
    return;
  }

  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) {
    if (RENDER_SESSION.active) RENDER_SESSION.keyOverridesApplied = true;
    return;
  }

  values.slice(1).forEach(row => {
    const option = String(row[10] || "").trim();
    const rawOptionValue = row[11];
    const hasOptionValue =
      rawOptionValue !== "" &&
      rawOptionValue !== null &&
      rawOptionValue !== undefined;

    if (option && hasOptionValue) {
      if (Object.prototype.hasOwnProperty.call(CALENDAR.setup, option)) {
        CALENDAR.setup[option] = coerceKeyOverrideValue_(rawOptionValue, CALENDAR.setup[option]);
      } else if (Object.prototype.hasOwnProperty.call(CALENDAR, option)) {
        CALENDAR[option] = coerceKeyOverrideValue_(rawOptionValue, CALENDAR[option]);
      }
    }

    const appearance = String(row[13] || "").trim();
    const appearanceColor = String(row[14] || "").trim();

    if (
      appearance &&
      appearanceColor !== "" &&
      Object.prototype.hasOwnProperty.call(CALENDAR.colors, appearance) &&
      looksLikeColor_(appearanceColor)
    ) {
      CALENDAR.colors[appearance] = appearanceColor;
    }
  });

  // Keep the Key Configurator's target sheet name in sync with whatever
  // CALENDAR.defaultDataSheetName resolved to (script default or Key override).
  KEY_CONFIGURATOR_CONFIG.targetSheetName = CALENDAR.defaultDataSheetName;

  if (RENDER_SESSION.active) RENDER_SESSION.keyOverridesApplied = true;
}



function coerceKeyOverrideValue_(value, currentValue) {
  if (typeof currentValue === "boolean") {
    if (value === true || value === false) return value;

    const text = String(value || "").trim().toLowerCase();
    return text === "true";
  }

  if (typeof currentValue === "number") {
    const n = Number(value);
    return isNaN(n) ? currentValue : n;
  }

  if (Array.isArray(currentValue)) {
    return String(value || "")
      .split(/[,;\n]/)
      .map(part => String(part || "").trim())
      .filter(part => part !== "");
  }

  return String(value || "").trim();
}



const KEY_CONFIGURATOR_CONFIG = {
  targetSheetName: APP_CONFIG.dataSheetName,
  keySheetName: APP_CONFIG.keySheetName,

  // If targetSheetName is not found, prompt user to pick a target tab.
  // Key sheet must still exist.
  allowTargetSheetChooser: true,

  headerRow: 1,
  startRow: 2,

  validation: {
    enabled: true,

    // Empty = validate every matching header between target and Key.
    // Example restricted mode: ["Type", "Category", "Status"]
    headersToValidate: [],

    replaceExistingValidation: true,
    allowInvalid: true,
    showDropdown: true,

    // true  = use source ranges like Key!A2:A, effectively open-ended
    // false = use source ranges like Key!A2:A10, ending at last non-blank value
    useOpenEndedKeyRanges: true
  },

  colorizer: {
    enabled: true,

    // The target/key column used to match row category.
    targetMatchHeader: "Category",
    keyMatchHeader: "Category",

    // The Key sheet column containing hex colors.
    keyColorHeader: "Category-Color",

    // true  = apply color rules across all target columns that have headers
    // false = apply color rules only to the target match column
    applyToAllHeaderColumns: true,

    replaceExistingCategoryRules: true
  }
};

/**
 * Run both validation and conditional formatting.
 * Good menu function if you want one button.
 */
function runKeyConfigurator() {
  if (KEY_CONFIGURATOR_CONFIG.validation.enabled) {
    setKeyBasedDataValidation();
  }

  if (KEY_CONFIGURATOR_CONFIG.colorizer.enabled) {
    setKeyBasedConditionalFormatting();
  }
}

/**
 * Applies data validation to target columns whose headers
 * match headers in the Key sheet.
 */
function setKeyBasedDataValidation() {
  const ss = SpreadsheetApp.getActive();
  const config = KEY_CONFIGURATOR_CONFIG;
  const validationConfig = config.validation;

  const targetSheet = getTargetSheet_(ss, config);
  const keySheet = getRequiredSheet_(ss, config.keySheetName);

  const targetHeaders = getHeaders_(targetSheet, config.headerRow);
  const keyHeaders = getHeaders_(keySheet, config.headerRow);

  const targetHeaderMap = buildHeaderMap_(targetHeaders);
  const keyHeaderMap = buildHeaderMap_(keyHeaders);

  const headerPairs = getValidationHeaderPairs_(
    validationConfig,
    targetHeaders,
    keyHeaderMap,
    targetHeaderMap
  );

  if (!headerPairs.length) {
    Logger.log("No matching headers found for data validation.");
    return;
  }

  const validationRowCount = targetSheet.getMaxRows() - config.startRow + 1;

  if (validationRowCount <= 0) {
    Logger.log("No target rows available for data validation.");
    return;
  }

  headerPairs.forEach(pair => {
    const targetHeader = targetHeaderMap[normalizeValue_(pair.source)];
    const keyHeader = keyHeaderMap[normalizeValue_(pair.key)];

    if (!targetHeader || !keyHeader) {
      Logger.log(`Skipping validation for "${pair.source}" → "${pair.key}" — column missing on one side.`);
      return;
    }

    const keyValuesRange = getKeyValidationSourceRange_(
      keySheet,
      keyHeader.column,
      config.startRow,
      validationConfig.useOpenEndedKeyRanges
    );

    if (!keyValuesRange) {
      Logger.log(`Skipping validation for "${pair.source}" — Key column "${pair.key}" has no values.`);
      return;
    }

    const targetRange = targetSheet.getRange(
      config.startRow,
      targetHeader.column,
      validationRowCount,
      1
    );

    const rule = SpreadsheetApp
      .newDataValidation()
      .requireValueInRange(keyValuesRange, validationConfig.showDropdown)
      .setAllowInvalid(validationConfig.allowInvalid)
      .build();

    if (validationConfig.replaceExistingValidation) {
      targetRange.clearDataValidations();
    }

    targetRange.setDataValidation(rule);

    Logger.log(
      `Applied validation to ${targetSheet.getName()}!${columnToLetter_(targetHeader.column)}${config.startRow}:${columnToLetter_(targetHeader.column)} using ${keySheet.getName()}!${columnToLetter_(keyHeader.column)}${config.startRow}:${columnToLetter_(keyHeader.column)}`
    );
  });
}

/**
 * Creates conditional formatting rules based on:
 * Target[Category] matching Key[Category],
 * then using Key[Category-Color] as the fill color.
 */
function setKeyBasedConditionalFormatting() {
  const ss = SpreadsheetApp.getActive();
  const config = KEY_CONFIGURATOR_CONFIG;
  const colorConfig = config.colorizer;

  const targetSheet = getTargetSheet_(ss, config);
  const keySheet = getRequiredSheet_(ss, config.keySheetName);

  // Source-side category column may be renamed via CALENDAR.setup.customCategory.
  // The Key tab's own Category column stays at the fixed "Category" header.
  const targetMatchHeader = CALENDAR.setup.customCategory || colorConfig.targetMatchHeader;
  const targetMatchCol = getColumnByHeader_(
    targetSheet,
    targetMatchHeader,
    config.headerRow
  );

  const keyMatchCol = getColumnByHeader_(
    keySheet,
    colorConfig.keyMatchHeader,
    config.headerRow
  );

  const keyColorCol = getColumnByHeader_(
    keySheet,
    colorConfig.keyColorHeader,
    config.headerRow
  );

  const keyLastRow = keySheet.getLastRow();

  if (keyLastRow < config.startRow) {
    Logger.log("No Key rows available for conditional formatting.");
    return;
  }

  const keyValues = keySheet
    .getRange(config.startRow, keyMatchCol, keyLastRow - config.startRow + 1, 1)
    .getValues();

  const keyColors = keySheet
    .getRange(config.startRow, keyColorCol, keyLastRow - config.startRow + 1, 1)
    .getValues();

  const targetMatchLetter = columnToLetter_(targetMatchCol);

  const targetRange = getConditionalFormatApplyRange_(
    targetSheet,
    targetMatchCol,
    config,
    colorConfig
  );

  const newRules = [];

  keyValues.forEach((row, i) => {
    const matchValue = String(row[0] || "").trim();
    const color = String(keyColors[i][0] || "").trim();

    if (!matchValue || !isValidHexColor_(color)) return;

    const formula = `=$${targetMatchLetter}${config.startRow}="${escapeFormulaString_(matchValue)}"`;

    const rule = SpreadsheetApp
      .newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setBackground(color)
      .setRanges([targetRange])
      .build();

    newRules.push(rule);
  });

  let existingRules = targetSheet.getConditionalFormatRules();

  if (colorConfig.replaceExistingCategoryRules) {
    existingRules = existingRules.filter(rule => {
      return !isKeyConfiguratorColorRule_(rule, targetSheet, targetRange);
    });
  }

  targetSheet.setConditionalFormatRules([
    ...existingRules,
    ...newRules
  ]);

  Logger.log(`Created ${newRules.length} conditional formatting rules.`);
}

/**
 * Optional wrapper if your existing menu wants validation only.
 */
function refreshKeyValidationFromMenu() {
  setKeyBasedDataValidation();
}

/**
 * Optional wrapper if your existing menu wants colors only.
 */
function refreshKeyColorsFromMenu() {
  setKeyBasedConditionalFormatting();
}

/**
 * Optional wrapper if your existing menu wants both.
 */
function refreshKeyConfigurationFromMenu() {
  runKeyConfigurator();
}

/*******************************
 * Sheet resolution helpers
 *******************************/

function getTargetSheet_(ss, config) {
  const sheet = ss.getSheetByName(config.targetSheetName);

  if (sheet) return sheet;

  if (!config.allowTargetSheetChooser) {
    throw new Error(`Target sheet not found: ${config.targetSheetName}`);
  }

  return chooseTargetSheetFromUi_(ss, config);
}

function getRequiredSheet_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`Sheet not found: ${sheetName}`);
  }

  return sheet;
}

function chooseTargetSheetFromUi_(ss, config) {
  const ui = SpreadsheetApp.getUi();

  const sheets = ss
    .getSheets()
    .filter(sheet => sheet.getName() !== config.keySheetName);

  if (!sheets.length) {
    throw new Error("No eligible target sheets found.");
  }

  const choices = sheets
    .map((sheet, index) => `${index + 1}. ${sheet.getName()}`)
    .join("\n");

  const response = ui.prompt(
    "Choose target sheet",
    `Target sheet "${config.targetSheetName}" was not found.\n\nEnter the number of the target sheet:\n\n${choices}`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    throw new Error("Target sheet selection was cancelled.");
  }

  const selectedNumber = Number(response.getResponseText().trim());

  if (
    !Number.isInteger(selectedNumber) ||
    selectedNumber < 1 ||
    selectedNumber > sheets.length
  ) {
    throw new Error(`Invalid sheet selection: ${response.getResponseText()}`);
  }

  const selectedSheet = sheets[selectedNumber - 1];

  Logger.log(`Selected target sheet: ${selectedSheet.getName()}`);

  return selectedSheet;
}

/*******************************
 * Shared helpers
 *******************************/

function getHeaders_(sheet, headerRow) {
  const lastColumn = sheet.getLastColumn();

  if (lastColumn < 1) return [];

  const values = sheet
    .getRange(headerRow, 1, 1, lastColumn)
    .getValues()[0];

  return values
    .map((value, index) => ({
      name: String(value || "").trim(),
      normalizedName: normalizeValue_(value),
      column: index + 1
    }))
    .filter(header => header.name);
}

function buildHeaderMap_(headers) {
  const map = {};

  headers.forEach(header => {
    if (!map[header.normalizedName]) {
      map[header.normalizedName] = header;
    }
  });

  return map;
}

function getColumnByHeader_(sheet, headerName, headerRow) {
  const headers = getHeaders_(sheet, headerRow);
  const normalizedHeaderName = normalizeValue_(headerName);

  const header = headers.find(item => {
    return item.normalizedName === normalizedHeaderName;
  });

  if (!header) {
    throw new Error(`Header "${headerName}" not found on sheet "${sheet.getName()}"`);
  }

  return header.column;
}

// Returns [{source, key}] pairs naming a target column and the Key column
// to validate it against. The two names can differ — used when the source
// has a renamed column (e.g. customCategory = "Theme") but the Key tab's
// section is still called "Category".
function getValidationHeaderPairs_(validationConfig, targetHeaders, keyHeaderMap, targetHeaderMap) {
  const pairs = [];
  const seen = Object.create(null);

  function add(sourceName, keyName) {
    if (!sourceName || !keyName) return;
    const k = normalizeValue_(sourceName);
    if (seen[k]) return;
    seen[k] = true;
    pairs.push({ source: sourceName, key: keyName });
  }

  if (validationConfig.headersToValidate && validationConfig.headersToValidate.length) {
    validationConfig.headersToValidate.forEach(n => add(n, n));
  } else {
    // Auto-detect: any target header that also exists by name in the Key.
    targetHeaders.forEach(h => {
      if (keyHeaderMap[h.normalizedName]) add(h.name, h.name);
    });
  }

  // Custom mappings: source column named via CALENDAR.setup.customX is
  // validated against the Key's fixed Type / Category / Status column.
  [
    { setup: "customType", keyName: "Type" },
    { setup: "customCategory", keyName: "Category" },
    { setup: "customStatus", keyName: "Status" },
  ].forEach(m => {
    const sourceName = CALENDAR.setup[m.setup];
    if (!sourceName) return;
    if (sourceName === m.keyName) return; // already covered by auto-detect
    if (!keyHeaderMap[normalizeValue_(m.keyName)]) return;
    if (!targetHeaderMap[normalizeValue_(sourceName)]) return;
    add(sourceName, m.keyName);
  });

  return pairs;
}

function getKeyValidationSourceRange_(sheet, column, startRow, useOpenEndedRange) {
  if (useOpenEndedRange) {
    const maxRows = sheet.getMaxRows();

    if (maxRows < startRow) return null;

    return sheet.getRange(
      startRow,
      column,
      maxRows - startRow + 1,
      1
    );
  }

  return getNonBlankColumnRange_(
    sheet,
    column,
    startRow
  );
}

function getNonBlankColumnRange_(sheet, column, startRow) {
  const lastRow = sheet.getLastRow();

  if (lastRow < startRow) return null;

  const values = sheet
    .getRange(startRow, column, lastRow - startRow + 1, 1)
    .getValues();

  let lastNonBlankOffset = -1;

  values.forEach((row, index) => {
    if (String(row[0] || "").trim()) {
      lastNonBlankOffset = index;
    }
  });

  if (lastNonBlankOffset === -1) return null;

  return sheet.getRange(
    startRow,
    column,
    lastNonBlankOffset + 1,
    1
  );
}

function getConditionalFormatApplyRange_(targetSheet, targetMatchCol, config, colorConfig) {
  const rowCount = targetSheet.getMaxRows() - config.startRow + 1;

  if (rowCount <= 0) {
    throw new Error("No target rows available for conditional formatting.");
  }

  if (colorConfig.applyToAllHeaderColumns) {
    const lastHeaderColumn = getLastHeaderColumn_(targetSheet, config.headerRow);

    return targetSheet.getRange(
      config.startRow,
      1,
      rowCount,
      lastHeaderColumn
    );
  }

  return targetSheet.getRange(
    config.startRow,
    targetMatchCol,
    rowCount,
    1
  );
}

function getLastHeaderColumn_(sheet, headerRow) {
  const headers = sheet
    .getRange(headerRow, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  for (let i = headers.length - 1; i >= 0; i--) {
    if (String(headers[i] || "").trim()) {
      return i + 1;
    }
  }

  throw new Error(`No headers found on row ${headerRow} of sheet "${sheet.getName()}"`);
}

function isKeyConfiguratorColorRule_(rule, targetSheet, targetRange) {
  const ranges = rule.getRanges();

  return ranges.some(range => {
    return (
      range.getSheet().getName() === targetSheet.getName() &&
      range.getRow() === targetRange.getRow() &&
      range.getColumn() === targetRange.getColumn() &&
      range.getNumColumns() === targetRange.getNumColumns()
    );
  });
}

function normalizeValue_(value) {
  return String(value || "").trim().toLowerCase();
}

function isValidHexColor_(value) {
  return /^#[0-9A-Fa-f]{6}$/.test(String(value || "").trim());
}

function escapeFormulaString_(value) {
  return String(value).replace(/"/g, '""');
}

function columnToLetter_(column) {
  let letter = "";

  while (column > 0) {
    const remainder = (column - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    column = Math.floor((column - 1) / 26);
  }

  return letter;
}