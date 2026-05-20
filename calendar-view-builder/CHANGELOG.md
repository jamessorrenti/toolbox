# Calendar View Builder

All notable changes to Calendar View Builder are documented here.


---

## v13.9.0

### Added

- New `startWeekOn` option: `Monday-CompressedWeekend`. Starts the week on Monday and renders Saturday + Sunday columns at half width (72px vs the standard 145px), giving Mon-Fri more horizontal real estate for events.

---

## v13.8.0

### Added

- Dropdown data validation on the Key tab for five setup options:
  - `q1StartMonth` — month names (strict).
  - `startWeekOn` — day names (strict).
  - `defaultDataSheetName` — available tabs in the spreadsheet (accepts custom names).
  - `customDate` — headers of the resolved `defaultDataSheetName` tab (accepts custom names).
  - `customTitle` — same as `customDate`.
- `customDate` / `customTitle` dropdowns automatically refresh when `defaultDataSheetName` changes (manual edit on the Key tab, or via **Set Key From Event List**). Dropdowns are also refreshed on every spreadsheet open so newly-added tabs and source columns appear without manual intervention.

### Changed

- **Set Key From Event List** now also sets `defaultDataSheetName` to the chosen event list tab, so the rest of the Calendar Tools menu targets the same sheet without a separate step.
- Default `q1StartMonth` is now `January` (was `February`). Override it in the Key tab for fiscal calendars that start elsewhere.

---

## v13.7.0

### Added

- **Set Key From Event List** menu item. Interactive walk-through that:
  - lets you pick an existing event list tab (by number or by name),
  - prompts for the Date and Title columns (accepts numbered options or custom names),
  - prompts for the Type, Category, and Status columns and the values within them (`all`, numbered picks, custom names, or any mix),
  - confirms, then writes the chosen values into the Key tab. Matching default icons and Category colors are reused; new Categories get colors from the default palette.
- New `defaultDataSheetName` setup option, default `Events`. Override it in the `Key` tab to point all Calendar Tools menu actions at a differently-named event list (e.g. `Tactics List`). Kept in sync with the Key Configurator's `targetSheetName`.
- New `showSetKeyFromEventListMenu` toggle to gate the new menu item.
- **Create Event List** now auto-runs the Key Configurator when the Key tab is present, so dropdowns + Category-based row colors are wired up immediately.

### Changed

- Default event list tab name is now `Events` (was `Tactics List`).
- Default `customDate` is now `Date` (was `MMDD`).
- Default `customTitle` is now `Title` (was `Name`).
- The four menu-visibility toggles + `frozenWeekdayHeader` now default to `FALSE` in the Key tab when it is created (script defaults remain `TRUE`). Net effect: a brand-new spreadsheet without a Key shows all menu items; once you have a Key, opt in to whichever you want by checking the corresponding box.
- README: rewrote the Optional first-time setup section, added a **Set Key From Event List** section, and split the setup options table into "Script default" and "Key default" columns.

---

## v13.6.0

### Added

- **Initial Setup** menu item. One-shot bootstrap for a brand-new spreadsheet:
  - Creates an event list tab with frozen `Title`, `Date`, `Category`, `Type`, `Status` headers.
  - Creates the `Key` tab if it does not exist.
  - Runs the Key Configurator so `Type`, `Category`, and `Status` get dropdown validation and rows get colored by `Category`.
  - Idempotent — re-running leaves existing tabs alone.
- **Create Event List** menu item. Creates the default event list tab if it is missing, or just navigates to it.
- Two new boolean toggles on the `Key` tab, both default `TRUE`:
  - `showInitialMenu` — controls visibility of **Initial Setup**.
  - `showEventListMenu` — controls visibility of **Create Event List**.
- README now has an **Optional first-time setup** section and a **Menu visibility toggles** subsection, plus a note that any pre-existing event list works without running Initial Setup.

### Changed

- The **Calendar Tools** menu order is now: *Initial Setup → calendar actions → Create Event List → Key Configurator → Create Key (and customize)*. Each toggleable section is gated by its `show*` option in the Key tab.

---

## v13.5.1

### Fixed

- **Additional-field labels are bold again.** The batched renderer in v13.5.0 was applying cell-level `setFontWeights("normal")` and `setFontColors(...)` after `setRichTextValues(...)`, which clobbered the per-character bold and color encoded in each event's rich text. Font weight and color now live entirely inside the `RichTextValue`s, including for date cells and overflow ("More...") cells.
- **Open Selected modal now finds events.** Apps Script starts each menu click in a fresh JS context, so `CALENDAR.setup` had script defaults instead of the user's Key-tab overrides. `openSelected` now calls `applyKeyOverrides_` before computing the selected date, so `maxEvents`, `startWeekOn`, and friends match what was actually rendered.

### Changed

- **New Calendar Sheet** is back to creating a tab named `Calendar View` (or `Calendar View (N)`), so the user can choose a period after the tab is created. **Add Q1-Q4** and **Add Jan-Dec** still name tabs after their period (`Q2 2026`, `May 2026`).

---

## v13.5.0

### Added

- Single source of truth for the version string: `CALENDAR.version`.
- New calendar tabs are now named after the selected period (e.g. `May 2026`, `Q2 2026`, `Year 2026`) instead of `Calendar View`. **Add Q1-Q4** and **Add Jan-Dec** now share the same naming logic.

### Changed

- Calendar rendering is batched. The per-day `setValue` / `setBackground` / `setFontColor` / `setBorder` chain in `renderMonthSection` is replaced with a single `setRichTextValues` / `setBackgrounds` / `setFontColors` / `setFontWeights` / `setHorizontalAlignments` / `setVerticalAlignments` / `setNotes` write per month, plus 18 batched border calls per month. Refresh times are significantly faster, particularly in `Year` mode.
- `Refresh All Calendars` now reuses source data, key config, and key overrides across the run. Each source sheet and the Key sheet are read once per refresh instead of once per calendar tab.
- Calendar column widths are set once per refresh instead of once per month.
- Date parsing is stricter:
  - Range expansion only triggers when the separator (`-` or `to`) is surrounded by whitespace. Titles like `Roadshow - East` and ISO dates like `2026-05-04` are no longer mis-read as ranges.
  - Removed the loose `new Date(string)` fallback in `parseSingleDate`. Date values must match one of the explicit patterns (`yyyy-mm-dd`, `m/d/yyyy`, `m/d`).

### Fixed

- Removed a duplicate definition of `getWeekStartIndex_`.

### Removed

- The per-cell helpers `renderEventRowsForDay_`, `clearEventCell_`, `setEventCell_`, and `setOverflowCell_` are gone. Their behavior is now inlined in the batched `renderMonthSection` and the new `buildEventCellRichText_` helper.

---

## v13.4.4

### Changed

- Renamed the calendar menu item from **Open Selected More...** to **Open Selected**.
- Simplified the More-cell note to display only:

  ```text
  Use Calendar Tools > Open Selected to view this date.
  ```

- Removed the hidden date payload from More-cell notes.
- Updated Open Selected behavior so it derives the selected date from the selected cell’s grid position instead of reading a date key from a note.

### Fixed

- Fixed lingering menu text that still displayed **Open Selected More...**.
- Removed dependency on note metadata for determining selected dates.

## v13.4.3

### Changed

- Removed the bottom-right toast that appeared when selecting calendar cells.
- Kept the instructional note only on More cells.
- Updated the Open Selected modal so it shows **all events for the selected day**, not only overflow events.
- Updated modal helper text to clarify that events link back to the source title cell.

### Fixed

- Fixed modal behavior so visible events and overflow events are both included.

---

## v13.4.2

### Changed

- Replaced the attempted cell-level developer metadata approach with date calculation from the calendar grid position.
- Kept notes only on actual More cells.
- Allowed Open Selected to work from:
  - date cells
  - event cells
  - empty cells inside a day block
  - More cells

### Fixed

- Fixed Apps Script error:

  ```text
  Exception: Adding developer metadata to arbitrary ranges is not currently supported.
  ```

- Removed unsupported arbitrary-range developer metadata usage.

---

## v13.4.1

### Changed

- Attempted to move selected-date metadata off visible notes and into hidden developer metadata.
- More-cell notes remained visible.
- Date cells and event cells were intended to support Open Selected without visible notes.

### Removed

- Removed visible selected-date notes from non-More cells.

### Known issue

- This version used unsupported cell-level developer metadata and was replaced by v13.4.2.

---

## v13.4

### Added

- Added **Open Selected** behavior for any cell in a calendar day block.
- Added support for selecting a date cell, event cell, empty day-block cell, or More cell and opening the day modal.
- Added selected-date metadata note support.
- Added More-cell note:

  ```text
  Use Calendar Tools > Open Selected to view this date.
  ```

### Changed

- Renamed the modal from **More events** to **Selected date**.
- Updated Key option notes so they appear only on Option-Value cells.
- Simplified date-format note examples.

---

## v13.3

### Added

- Added note help for Key tab format options:
  - `dayFormat`
  - `dateFormat`
  - `monthFormat`

### Changed

- Changed default `dateFormat` from `F` to `d`.
- Updated Key tab notes:
  - `dayFormat`: examples include `E = Tue`, `EEEE = Tuesday`
  - `dateFormat`: examples include `d = 4`, `dd = 04`, `EE d = Tue 4`
  - `monthFormat`: examples include `M = 9`, `MM = 09`, `MMM = Sep`, `MMMM = September`

### Fixed

- Reapplied non-underlined text styling after setting event title links.
- Fixed event titles appearing underlined because of rich text link styling.

---

## v13.2

### Fixed

- Fixed date labels rendering literally as `F`.
- Replaced spreadsheet number formatting with Apps Script date formatting via `Utilities.formatDate`.

### Changed

- Calendar date labels now render through:

  ```js
  formatCalendarDateLabel_(cellDate)
  ```

- `dateFormat` now supports Apps Script / Java date-time format patterns such as:
  - `d`
  - `dd`
  - `F`
  - `FF`
  - `EEEE`
  - `EE d`

---

## v13.1

### Added

- Added `dateFormat` setup option.
- Added `dateFormat` to Key tab setup options below `dayFormat`.
- Set initial default:

  ```js
  dateFormat: "F"
  ```

### Changed

- Date cells were updated to explicitly format date labels through a configurable format.

### Known issue

- The first implementation used spreadsheet number formatting and caused date labels to display literally as `F`.
- Replaced in v13.2.

---

## v13

### Added

- Added new rich text event renderer.
- Added partial rich-text styling for custom additional event labels.
- Styled only the custom additional field label and colon, not the value.

### Changed

- Custom additional fields now render like:

  ```text
  Campaign Launch
  Owner: James
  ```

- Only `Owner:` receives configured label styling.
- The additional field value remains in the normal event text style.

---

## v13b

### Added

- Added beta event-cell renderer.
- Added support for rendering custom additional fields below the event title.
- Added rich text styling for additional event lines.
- Added parser support for custom additional field tokens:
  - `bold`
  - `italic`
  - `uppercase`
  - `lowercase`
  - `label uppercase`
  - `label lowercase`
  - `value uppercase`
  - `value lowercase`
  - hex colors

### Changed

- Event title links remained attached to the title text.
- Custom additional fields were displayed inside the same event cell.

### Later adjusted

- Full-line styling was changed in v13 so only labels and colons are styled.

---

## v12.5.2

### Fixed

- Fixed Calendar Tools menu failing to appear after deleting the Key tab.
- Added missing `getKeyBooleanOption_()` helper.
- Menu now falls back to script defaults when the Key tab does not exist.

---

## v12.5.1

### Changed

- Reordered Key tab setup options:

  ```text
  customAdditional
  customAdditionalLabels
  customAdditionalLabelsStyle
  ```

- Removed grey conditional formatting for custom additional options on the Key tab.
- Set `customAdditionalLabels` default to `TRUE`.

### Fixed

- Fixed `Key!J1` text color so it stays white.
- Fixed `K1` and `N1` so header cells use the project/default font instead of Courier New.
- Set `B2:B` on the Key tab to font size `14`.

---

## v12.5

### Added

- Added Key tab setup/appearance override system.
- Added support for reading setup overrides from the Key tab.
- Added support for reading appearance/color overrides from the Key tab.
- Added `showKeyConfiguratorMenuItems` to Key tab options.
- Added fixed boolean override handling so `FALSE` values are respected.
- Added custom additional field display below event title.
- Added support for `customAdditional`, `customAdditionalLabels`, and `customAdditionalLabelsStyle`.
- Added `startWeekOn` support for calendar grid rendering.
- Added grouped/collapsed Key tab setup and appearance columns.

### Changed

- Removed `q1StartDay` from exposed Key tab setup options.
- Kept `q1StartMonth` exposed.
- Key tab setup values are generated from script defaults instead of duplicating hard-coded default values.
- Blank Key tab values fall back to script defaults.
- Unknown setup and appearance keys are ignored.

### Fixed

- Fixed boolean `FALSE` values being treated as blank.
- Fixed `frozenWeekdayHeader = FALSE` override behavior.
- Fixed `showKeyConfiguratorMenuItems` by reading it directly during `onOpen()`.

---

## v12

### Added

- Added `frozenWeekdayHeader` setup toggle.
- Added title-week colors:

  ```js
  titleWeekBackground: "#EAB8F2"
  titleWeekFontColor: "#000000"
  ```

- Added frozen weekday header row in row `3`.
- Added buffer row behavior so row `4` remains a visual spacer.
- Added configurable `bufferRowHeight`.

### Changed

- Calendar canvas starts at row `5`.
- Rows `3` and `4` use small buffer heights when frozen weekday header is disabled.
- Row `3` uses normal height when frozen weekday header is enabled.
- Frozen row count changes dynamically:
  - `2` rows frozen when weekday header is disabled
  - `3` rows frozen when weekday header is enabled

---

## v12 Key Settings / Formatted Builds

### Added

- Added new Key tab setup section in columns `J:L`:
  - `J`: Additional Setup
  - `K`: Option
  - `L`: Option-Value
- Added new Key tab appearance section in columns `N:O`:
  - `N`: Appearance
  - `O`: Appearance-Color
- Grouped columns `K:O` by default.
- Collapsed grouped settings columns by default.
- Added checkboxes for true/false setup values.
- Added Key tab appearance overrides.
- Added auto-coloring for Appearance-Color cells in column `O`.

### Changed

- Used script defaults as the source of truth for generated Key tab values.
- Kept Option and Appearance values as actual variable names.
- Used `Courier New` for variable-name columns:
  - `K:K`
  - `N:N`
- Kept project default font for the rest of the Key tab.
- Improved Key tab formatting to more closely match the reference sheet.
- Added dark/grey spacer columns.
- Kept `D:E` paired auto-coloring from Category-Color.
- Changed `N:O` behavior so only `O` auto-colors, not `N`.

---

## v11.7

### Added

- Added support for preserving `B2` data validation during normal refreshes.
- Added manual workflow support for Google Sheets native **Allow multiple selections** on `B2`.
- Added `syncFilterStyles_(sheet)` helper for `A2/B2` styling.
- Added `clearFullSheet_(sheet)` to distinguish first-time initialization from calendar canvas refresh.
- Added safer refresh behavior so only the calendar canvas is cleared during normal rebuilds.

### Changed

- `A2` and `B2` are centered.
- `B2` uses normal header background when blank or when `A2` is `All Events`.
- `B2` only uses accent background when an active filter is applied.
- `B2` validation is reset when `A2` or `G1` changes, but not during ordinary calendar refreshes.
- Updated `B2` helper note to instruct users to enable native **Allow multiple selections** manually in Google Sheets data validation.

### Fixed

- Fixed refresh behavior that could overwrite `B2` validation and remove manually enabled native multi-select settings.
- Fixed incorrect source-cell links when a calendar filter is applied by preserving the original source row number.

---

## v11.6

### Added

- Added Calendar Tools menu item:

  ```text
  Refresh All Calendars
  ```

- Kept the `G2` checkbox refresh for the current calendar tab.
- Added `A2` filter field control.
- Added `B2` filter value control.
- Added `All Events` default filter state.
- Added filter field dropdown populated from the selected source sheet headers.
- Added filter value dropdown populated from unique values in the selected source column.
- Added source-change handling:
  - changing `G1` resets `A2` to `All Events`
  - clears `B2`
  - rebuilds filter dropdowns
- Added per-calendar filtering during render.
- Added `Open Selected More...` menu item for opening overflow modals from selected More cells.

### Changed

- Calendar refreshes preserve current tab-specific controls.
- Overflow logic changed so `maxEvents: 4` shows 4 actual events and then uses a separate overflow row for `x More...`.
- Overflow modal links continue to point back to the source event row.
- B2 filter value initially supported multiple values via comma-separated text.

---

## v11.5

### Added

- Reworked the calendar body into a stacked week grid.
- Added category fill/background color for event cells.
- Added `maxEvents` setup option with default value:

  ```js
  maxEvents: 4
  ```

- Added event rows beneath each date row.
- Added overflow display using `x More...`.
- Added hidden overflow lookup note format:

  ```text
  CALENDAR_OVERFLOW|yyyy-mm-dd
  ```

- Added overflow modal helper functions.
- Added `onSelectionChange(e)` behavior to detect More cells.
- Added source links for event cells.
- Added modal links back to the source Title/Name cell.

### Changed

- Replaced multi-event rich text in one cell with one event per row/cell.
- Changed Category color behavior from text color to event cell background fill.
- Event cells were top-aligned for consistent visual layout.
- `maxEvents` was clarified to mean visible event rows before overflow.

### Known limitation

- Opening a modal directly from `onSelectionChange(e)` was unreliable in Google Sheets simple triggers.
- Later versions used a menu action instead.

---

## v11

### Added

- Added a `Key` tab concept for calendar customization.
- Added Calendar Tools menu option **Create Key (and customize)** when the `Key` tab does not exist.
- Added default Key values for:
  - `Type` / `Type-Icon`
  - `Category` / `Category-Color`
  - `Status` / `Status-Icon`
- Added event display support for:
  - prepending Type icons before event names
  - appending Status icons after event names
  - applying Category color styling to events
- Added fallback defaults when no Key tab exists.
- Added support for custom source headers through:

  ```js
  customDate: "Date"
  customTitle: "Title"
  ```

### Changed

- Kept `Date` and `Title` as the generic default source headers.
- Supported sheet-specific overrides such as:

  ```js
  customDate: "MMDD"
  customTitle: "Name"
  ```

---

## v10.1

### Changed

- Removed external spreadsheet `openById()` source loading flow.
- Simplified source handling to local spreadsheet tabs only.
- Added support for using a hidden `External` sheet with `IMPORTRANGE()` formulas instead of Apps Script external access.
- Preserved `H1:H2` during calendar refreshes and rebuilds.
- Added source-sheet dropdown behavior for `G1`.
- Added automatic defaulting to the `Data` tab when available.
- Added fallback messaging when no source tabs exist.

### Fixed

- Fixed stale quarter-month blocks remaining after switching from quarter views to single-month views.
- Fixed year-only refreshes not updating rendered month/year output.
- Fixed refresh behavior where changing only `D1` in Custom mode did not redraw month counts correctly.
- Fixed refresh checkbox behavior so failed refreshes do not automatically uncheck `G2`.

---

## v10

### Added

- Added richer custom additional field configuration syntax.
- Added support for additional-field:
  - label visibility
  - label style
  - color override
  - uppercase/lowercase transforms
- Added compact custom additional configuration support, for example:

  ```js
  customAdditional: [
    "Status, true, bold, red, uppercase",
    "DRI"
  ]
  ```

- Added parsing helpers for:
  - `bold`
  - `italic`
  - `uppercase`
  - `lowercase`
  - hex colors
  - named colors

### Changed

- Renamed:

  ```js
  refreshControlColor
  ```

  to:

  ```js
  titleRefreshColor
  ```

- Reordered setup options so `customDate` appears before `customTitle`.
- Improved comments for `customAdditional`.

### Fixed

- Fixed rich text styling so additional-field labels and values can be styled independently.
- Fixed event links so only main event titles are clickable.
- Preserved plain-text styling for additional metadata lines.

---

## v9.9

### Added

- Added support for rendering additional metadata beneath event titles.
- Added setup options:

  ```js
  customAdditional
  customAdditionalLabels
  customAdditionalLabelsStyle
  ```

### Changed

- Additional metadata renders as indented lines below the title.
- Main event title remains the linked line.
- Additional lines remain plain text.

---

## v9.8

### Added

- Added setup-level source-header overrides:

  ```js
  customDate: "Date"
  customTitle: "Title"
  ```

- Added priority matching so configured custom headers are checked before default candidates.

### Changed

- Source-column matching continues to use ordered candidate priority.

---

## v9.7.2.1

### Fixed

- Fixed refresh behavior so changing only `D1` in Custom mode correctly redraws the calendar.
- Added live control reads during refresh execution.

---

## v9.7.2

### Fixed

- Fixed refresh logic so Custom month counts are respected during redraw.
- Fixed stale state usage during refresh operations.

---

## v9.7.1

### Fixed

- Fixed Custom-mode styling behavior for `B1`.
- `B1` now hides text by matching the font color to the background instead of changing the background color itself.

---

## v9.7

### Added

- Added:

  ```js
  titleAccentColor
  ```

- Added mode-specific header styling:
  - Custom mode highlights `D1:D2`.
  - Year mode highlights `D2`.
- Added improved mode synchronization helpers.

### Changed

- Custom mode hides `B1` by matching its font color to the background.
- Year mode preserves the year display but highlights the start date.

---

## v9.6

### Added

- Added `Year` mode.
- Added support for 12-month rolling views starting from `D2`.
- Added behavior where Year mode can display January–December or a shifted 12-month range such as April–March.

### Changed

- `D2` became the start-date anchor for Custom and Year modes.
- Fixed modes such as month/quarter can still reseed `D2` to their default start date.

### Fixed

- Fixed year-only refresh behavior so the chosen year affects fixed month views.
- Fixed quarter/month start-date behavior when changing `A1` and `B1`.

---

## v9.5

### Added

- Added `Custom` mode to the period dropdown.
- Added top-control redesign:
  - `C1` = Months label
  - `D1` = month count
  - `C2` = Starting label
  - `D2` = start date
  - `F2` = Refresh label
  - `G2` = refresh checkbox
- Added support for arbitrary custom month counts.
- Added support for custom start dates.
- Added theme cycling for custom ranges longer than three months.

### Changed

- `D2` became the start-date anchor for Custom mode.
- Refresh remains explicit through the checkbox instead of immediately redrawing on every control edit.

### Fixed

- Fixed stale render behavior when switching from a quarter view to a month view.
- Fixed full canvas rebuild behavior for Custom mode.

---

## v9.3

### Added

- Added manual refresh controls in the calendar header:
  - `D1` = Refresh label
  - `D2` = refresh checkbox
- Added full canvas rebuild behavior through `rebuildCalendarCanvas()`.
- Added `refreshControlColor`.
- Added `renderRowCapacity` for defining the redraw canvas size.

### Changed

- Calendar redraws now rebuild the controls and body together.
- Refresh became explicit via checkbox instead of relying only on control edits.
- The sheet is cleared/rebuilt to avoid stale quarter/month artifacts.

### Fixed

- Fixed lingering quarter blocks after switching to a month view.
- Fixed stale formatting and stale rendered calendar blocks after refresh.
- Fixed checkbox refresh so it can reset after a successful update.

---

## v9.1

### Fixed

- Added missing `monthOptions` configuration.
- Fixed `Cannot read properties of undefined (reading 'indexOf')` caused by missing month options.
- Stabilized `isMonth()` and `monthNameToNumber()` behavior.

---

## v9

### Added

- Added date-number rich text styling so only date numbers remain bold.
- Added non-bold linked event text.
- Added explicit left/top alignment for date cells with links.
- Added support for link-like columns such as:
  - `Landing Page`
  - `Link`
  - `URL`

### Changed

- Event links are styled as normal text rather than blue/underlined links.
- Rich text handling was refined for better visual consistency.

### Fixed

- Fixed alignment shifts caused by links in calendar cells.
- Fixed event text inheriting bold styling from date numbers.

---

## v8

### Added

- Added explicit view-only calendar behavior.
- Added aggressive render-area reset logic for redraws.
- Added rich text event bullets.
- Added links from calendar event bullets back to the source row/cell.
- Added blank-line spacing between multiple events on the same date.
- Added source-cell URL construction helpers.
- Added source metadata tracking.

### Changed

- Event bullets are linked back to their source row while still appearing visually like normal text.
- Date cells render using rich text so dates and event lines can be styled separately.
- Quarter-to-month cleanup became more aggressive.

### Known issue

- Early aggressive row deletion could hit Google Sheets limitations around deleting all non-frozen rows.
- Later versions moved toward full canvas rebuild / clear behavior instead of unsafe deletion.

---

## v7

### Fixed

- Fixed month mapping bug where month names were offset because `monthNameToNumber()` used the wrong option list.
- Fixed issue where choosing January could render a later month.
- Ensured the month shown in `A1` matches the rendered calendar month.

### Changed

- Locked the month-name mapping to the 12-month `monthOptions` list.
- Kept the v6 layout and rendering structure.

---

## v6

### Changed

- Renamed helpers to a cleaner non-underscore naming convention.
- Returned to a single `CALENDAR` configuration object.
- Expanded in-code comments and structure documentation.
- Standardized helper naming and calendar rendering function names.

### Fixed

- Cleaned up helper-name inconsistencies from earlier iterations.
- Improved maintainability of the full script replacement.

---

## v5

### Changed

- Split weekday-header colors from calendar-day cell colors.
- Replaced broad day-color settings with clearer theme settings:
  - `month1/week/day/inactive`
  - `month2/week/day/inactive`
  - `month3/week/day/inactive`
- Clarified inactive-day definition as spillover days from the previous/next month.
- Added explicit inactive-day font colors.

### Fixed

- Fixed inactive day styling so only out-of-month spillover days receive inactive formatting.
- Fixed weekday header coloring so it does not also color normal day cells.

---

## v4

### Fixed

- Fixed inactive-day behavior so blank inactive background settings render as white.
- Fixed quarter-to-month switching so old month blocks are cleared before redraw.
- Fixed stale March/April blocks remaining after switching from Q1 to January.

### Changed

- Preserved the v3 layout, font, and alignment updates while making rendering cleanup more aggressive.

---

## v3

### Changed

- Renamed the global configuration object from versioned naming to the cleaner `CALENDAR`.
- Standardized the font to `Inter`.
- Changed weekday rendering to full day names using `EEEE`.
- Tightened control-row alignment and formatting.

### Fixed

- Added the missing date-key helper used during event indexing.
- Improved date grouping reliability for calendar event placement.

---

## v2 — Calendar View Builder

### Changed

- Reworked the prototype into a reusable **Calendar View Builder**.
- Replaced the single **Create/Refresh Calendar** flow with a fuller Calendar Tools menu:
  - **New Calendar Sheet**
  - **Replace with Calendar**
  - **Add Q1-Q4**
  - **Add Jan-Dec**

### Added

- Added new standardized top-control layout:
  - `A1` = Period dropdown
  - `B1` = Year dropdown
  - `F1` = Source Data label
  - `G1` = source sheet name / spreadsheet URL / spreadsheet ID
  - `A2` = Starting label
  - `B2` = first date for selected period
- Added support for creating multiple generated tabs:
  - one tab per quarter
  - one tab per month
- Added configurable defaults:
  - default source sheet name
  - default year
  - fiscal Q1 start month/day
  - weekday/month formatting
- Added initial multi-month color themes:
  - main month
  - second month
  - third month
  - inactive days
- Added warning behavior before replacing existing/nonblank tabs.
- Continued supporting local sheets and external spreadsheet URLs/IDs.

---

## v1 — Calendar View Generator

### Added

- Initial Google Sheets calendar generator prototype.
- Added custom menu:

  ```text
  Calendar Tools → Create/Refresh Calendar
  ```

- Added single-sheet calendar rendering workflow.
- Added overwrite confirmation before replacing a rendered calendar tab.
- Added simple `onEdit(e)` auto-refresh behavior for control changes.
- Added support for:
  - month views
  - quarter views
  - configurable Q1 start month/day
  - local source sheets
  - external spreadsheet URLs
  - external spreadsheet IDs

### Added control row

- `A1` = Period dropdown
- `B1` = Year dropdown
- `C1:D1` = Source Data label/source input
- `H1:I1` = Q1 Start Month controls
- `J1:K1` = Q1 Start Day controls

### Added calendar rendering behavior

- Added big top header block.
- Added Sunday-first weekday headers.
- Added date-grid calendar layout.
- Added event indexing by date.
- Added quarter rendering support.

### Defaults

- Default calendar tab:

  ```text
  Calendar View
  ```

- Default source fallback tab:

  ```text
  Previous
  ```

- Default year:

  ```text
  2026
  ```

- Default Q1 start:

  ```text
  February 1
  ```

---