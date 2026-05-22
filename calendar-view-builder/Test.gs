/**
 * Calendar View Builder — internal smoke tests.
 *
 * Paste this as a SECOND .gs file in your Apps Script project, alongside
 * Calendar_View_Builder.gs. The main script doesn't reference it, the menu
 * doesn't show it, and the README/CHANGELOG don't mention it. Run the
 * functions below manually from the Apps Script editor's Run button when
 * you want to verify behavior.
 *
 * Each test logs to View > Logs (or the Executions panel). No assertions
 * — they print observable values so you can eyeball-check correctness.
 *
 * Conventions:
 * - testAll() runs everything in order.
 * - test* functions each focus on one area.
 * - All tests assume the active spreadsheet has a Key tab and at least one
 *   calendar tab (use Initial Setup if not).
 */

// ---------- entry ----------

function testAll() {
  Logger.log("=== testAll ===");
  testEnvironment();
  testKeyOverrides();
  testDateParsing();
  testSourceDetection();
  testEventSortSpec();
  testStaleDetection();
  testThemeFetch();
  testRenderActive();
  Logger.log("=== done ===");
}

// ---------- environment ----------

function testEnvironment() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("--- environment ---");
  Logger.log("script version: " + CALENDAR.version);
  Logger.log("spreadsheet TZ: " + ss.getSpreadsheetTimeZone());
  Logger.log("script TZ:      " + Session.getScriptTimeZone());
  Logger.log("active sheet:   " + ss.getActiveSheet().getName());
  Logger.log("sheets:         " + ss.getSheets().map(s => s.getName()).join(", "));
}

// ---------- Key overrides ----------

function testKeyOverrides() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  applyKeyOverrides_(ss);
  Logger.log("--- key overrides applied ---");
  Logger.log("defaultDataSheetName: " + CALENDAR.defaultDataSheetName);
  Logger.log("customDate:           " + CALENDAR.setup.customDate);
  Logger.log("customTitle:          " + CALENDAR.setup.customTitle);
  Logger.log("customType:           " + CALENDAR.setup.customType);
  Logger.log("customCategory:       " + CALENDAR.setup.customCategory);
  Logger.log("customStatus:         " + CALENDAR.setup.customStatus);
  Logger.log("customAdditional:     " + JSON.stringify(CALENDAR.setup.customAdditional));
  Logger.log("eventSortOrder:       " + CALENDAR.setup.eventSortOrder);
  Logger.log("startWeekOn:          " + CALENDAR.setup.startWeekOn);
  Logger.log("q1StartMonth:         " + CALENDAR.setup.q1StartMonth);
  Logger.log("autoRefresh:          " + CALENDAR.autoRefresh);
}

// ---------- date parsing ----------

function testDateParsing() {
  Logger.log("--- parseDateValue ---");
  const inputs = [
    "5/4/2026",
    "2026-05-04",
    "5/4",
    "5/4/2026 - 5/7/2026",
    "5/4/2026 to 5/7/2026",
    "Roadshow - East",            // not a range — title with dash
    "",                             // blank
    null,
    new Date(2026, 4, 4),
    new Date("2026-05-04T00:00:00Z"), // midnight UTC
  ];
  inputs.forEach(input => {
    try {
      const parsed = parseDateValue(input);
      const repr = (input instanceof Date) ? input.toISOString() : JSON.stringify(input);
      Logger.log("  " + repr + " → " + parsed.map(d => dateKey(d)).join(", "));
    } catch (err) {
      Logger.log("  threw on " + input + ": " + err.message);
    }
  });
}

// ---------- source detection ----------

function testSourceDetection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  applyKeyOverrides_(ss);
  Logger.log("--- isSourceDataSheet_ for every tab ---");
  ss.getSheets().forEach(s => {
    Logger.log("  " + s.getName() + " → " + isSourceDataSheet_(s));
  });
  Logger.log("--- readDefaultDataSheetHeaders_ ---");
  Logger.log("  " + JSON.stringify(readDefaultDataSheetHeaders_(ss)));
}

// ---------- sort spec ----------

function testEventSortSpec() {
  Logger.log("--- parseEventSortSpec_ ---");
  const inputs = [
    "Category ↓, Status ↓, Alphabetical ↓",
    "Source ↓",
    "Category ↑",
    "Title Asc, Category Desc",
    "",
    "garbage",
    "Foo ↓",
  ];
  inputs.forEach(input => {
    Logger.log("  " + JSON.stringify(input) + " → " + JSON.stringify(parseEventSortSpec_(input)));
  });

  Logger.log("--- sortDayEvents_ on sample events ---");
  const events = [
    { title: "Webinar A", category: "Marketing", status: "Live",       sourceRow: 10 },
    { title: "Email B",   category: "Sales",     status: "Draft",      sourceRow: 11 },
    { title: "Webinar C", category: "Marketing", status: "Draft",      sourceRow: 12 },
    { title: "Event D",   category: "Sales",     status: "Live",       sourceRow: 13 },
  ];
  const spec = parseEventSortSpec_("Category ↓, Status ↓, Alphabetical ↓");
  sortDayEvents_(events, spec);
  events.forEach(e => Logger.log("  " + e.category + " / " + e.status + " / " + e.title));
}

// ---------- stale detection ----------

function testStaleDetection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("--- isCalendarStale_ ---");
  const docProps = PropertiesService.getDocumentProperties();
  Logger.log("  lastSourceChangeAt: " + docProps.getProperty("autoRefreshLastSourceChangeAt"));
  ss.getSheets().filter(isCalendarSheet).forEach(s => {
    Logger.log("  " + s.getName() + " stale=" + isCalendarStale_(s) +
      " lastRefreshAt=" + docProps.getProperty("autoRefreshAt:" + s.getSheetId()));
  });
}

// ---------- theme fetch ----------

function testThemeFetch() {
  Logger.log("--- fetchThemeManifest_ ---");
  const themes = fetchThemeManifest_();
  if (!themes) {
    Logger.log("  fetch failed");
    return;
  }
  Logger.log("  count: " + themes.length);
  themes.forEach(t => Logger.log("    " + t.name + " (" + t.file + ")"));

  Logger.log("--- fetchTheme_('berry.json') ---");
  const berry = fetchTheme_("berry.json");
  if (!berry) {
    Logger.log("  fetch failed");
    return;
  }
  Logger.log("  name: " + berry.name);
  Logger.log("  colors keys: " + Object.keys(berry.colors || {}).length);
  Logger.log("  categoryPalette: " + JSON.stringify(berry.categoryPalette));
}

// ---------- end-to-end render ----------

function testRenderActive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  if (!isCalendarSheet(sheet)) {
    Logger.log("--- testRenderActive: active sheet is not a calendar — switch to one and re-run");
    return;
  }
  Logger.log("--- testRenderActive: " + sheet.getName());
  const start = Date.now();
  try {
    renderCalendarSheet(sheet);
    Logger.log("  rendered in " + (Date.now() - start) + " ms");
  } catch (err) {
    Logger.log("  render threw: " + err.message);
  }
}
