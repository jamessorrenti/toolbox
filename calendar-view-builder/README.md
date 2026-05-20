# Calendar View Builder

A Google Apps Script tool for generating dynamic, view-only calendar tabs inside Google Sheets.

Calendar View Builder reads event data from a local source sheet, then builds formatted calendar views that can be refreshed, filtered, customized, and styled from a generated `Key` tab.

---

## Features

- Generate calendar tabs from a source sheet in the same spreadsheet
- Create a new calendar tab
- Replace the active tab with a calendar view
- Refresh one calendar or all calendar tabs
- Create monthly, quarterly, yearly, and custom calendar views
- Filter each calendar tab by source-sheet column values
- Use a `Key` tab to define event types, categories, statuses, icons, and colors
- Customize calendar colors, fonts, date formats, and display settings from the `Key` tab
- Display additional event fields below each event title
- Open a modal showing all events for a selected calendar day
- Link event labels back to their source row/cell
- Keep calendar views read-only and rebuildable

---

## Files

```text
calendar-view-builder/
  Calendar_View_Builder.gs
  README.md
  CHANGELOG.md
```

`Calendar_View_Builder.gs` is the Apps Script file to paste into a Google Sheets-bound Apps Script project.

---

## Requirements

- A Google account
- A Google Sheet
- Access to Google Apps Script through **Extensions > Apps Script**
- A source data tab in the same spreadsheet

This script is designed as a **bound Apps Script**, meaning it runs inside the spreadsheet where it is installed.

---

## Setup

1. Open or create a Google Sheet.
2. Add a source data tab named `Events`.
3. Add event headers in row 1.
4. Add event rows starting in row 2.
5. Go to **Extensions > Apps Script**.
6. Create a new `.gs` file named `Calendar_View_Builder`.
7. Paste in the contents of `Calendar_View_Builder.gs`.
8. Save the Apps Script project.
9. Reload the spreadsheet.
10. Use the **Calendar Tools** menu.

The first time you run a menu action, Google may ask you to authorize the script.

---

## Source sheet format

By default, the script expects a local source tab named:

```text
Events
```

You can change the default source tab in the script:

```js
const APP_CONFIG = {
  keySheetName: "Key",
  dataSheetName: "Events"
};
```

The source sheet should use row 1 for headers and row 2 and below for event rows.

### Required columns

| Purpose | Default column name |
|---|---|
| Event date | `Date` |
| Event title | `Title` |

These are controlled by setup options:

```js
customDate: "Date",
customTitle: "Title",
```

They can also be overridden from the generated `Key` tab.

### Optional columns

The script can use these optional columns when present:

| Purpose | Common column names |
|---|---|
| Type | `Type`, `Event Type`, `Channel`, `Tactic Type` |
| Category | `Category`, `Event Category`, `Theme`, `Product`, `Pillar` |
| Status | `Status`, `Event Status` |
| Description / link target | `Description`, `Desc`, `Event Description`, `Details` |

Type values can display icons. Category values can control event background colors. Status values can display icons.

### Date values

The source date column can contain:

| Input type | Example |
|---|---|
| Google Sheets date | `5/4/2026` |
| Text date | `5/4/2026` |
| Month/day text | `5/4` |
| ISO date | `2026-05-04` |
| Date range | `5/4/2026 - 5/7/2026` |
| Date range with `to` | `5/4/2026 to 5/7/2026` |

Date ranges are expanded so the event appears on each date in the range.

---

## Calendar Tools menu

After reloading the spreadsheet, the script adds a custom menu called **Calendar Tools**.

| Action | Description |
|---|---|
| New Calendar Sheet | Creates a new calendar tab |
| Replace with Calendar | Replaces the active tab with a calendar view |
| Refresh All Calendars | Rebuilds all calendar tabs |
| Open Selected | Opens a modal for the selected calendar day |
| Add Q1-Q4 | Creates one tab for each quarter |
| Add Jan-Dec | Creates one tab for each month |
| Create Key (and customize) | Creates the Key tab if it does not exist |

Additional Key Configurator actions can be shown or hidden from the `Key` tab.

---

## Calendar controls

Each calendar tab includes controls in the top rows.

### Row 1

| Cell | Purpose |
|---|---|
| A1 | Period |
| B1 | Year |
| C1 | Month count label |
| D1 | Month count |
| F1 | Source data label |
| G1 | Source sheet |

### Row 2

| Cell | Purpose |
|---|---|
| A2 | Filter field |
| B2 | Filter value |
| C2 | Starting label |
| D2 | Start date |
| F2 | Refresh label |
| G2 | Refresh checkbox |

Check the refresh box in `G2` to refresh only that calendar tab.

---

## Calendar modes

Supported period options include quarters, individual months, `Year`, and `Custom`.

| Mode | Description |
|---|---|
| Month | Shows one month |
| Quarter | Shows three months |
| Year | Shows twelve months starting from the selected start date |
| Custom | Shows a custom number of months starting from the selected start date |

---

## Key tab customization

The `Key` tab controls event icons, category colors, status icons, setup options, and appearance options.

| Columns | Section |
|---|---|
| A:B | Type and Type-Icon |
| D:E | Category and Category-Color |
| G:H | Status and Status-Icon |
| J:L | Additional setup options |
| N:O | Appearance colors |

Spacer columns are used for visual separation. Columns `K:O` are grouped and collapsed by default when the Key tab is created.

### Type icons

The Type section maps source Type values to icons.

| Type | Type-Icon |
|---|---|
| Webinar | 🖥️ |
| Event | 📅 |
| Email | 📧 |

When an event has a matching Type, the icon appears before the event title.

### Category colors

The Category section maps source Category values to event background colors.

| Category | Category-Color |
|---|---|
| Category 1 | `#C8E6C9` |
| Category 2 | `#BBDEFB` |
| Category 3 | `#E1BEE7` |

The `Category` and `Category-Color` cells are automatically colorized from the hex value in `Category-Color`.

### Status icons

The Status section maps source Status values to icons.

| Status | Status-Icon |
|---|---|
| Not Started | ⚪ |
| In Progress | 🔵 |
| Live | 🟢 |
| Complete | ✅ |

When an event has a matching Status, the icon appears after the event title.

---

## Setup options

Setup options are generated from the script defaults. The `Key` tab can override them.

| Option | Default | Description |
|---|---|---|
| `showKeyConfiguratorMenuItems` | `TRUE` | Show extra Key Configurator menu items |
| `frozenWeekdayHeader` | `TRUE` | Show a frozen weekday header row |
| `customDate` | `Date` | Source column used for event dates |
| `customTitle` | `Title` | Source column used for event titles |
| `maxEvents` | `4` | Maximum visible events per day before showing More |
| `customAdditional` | `Owner` | Additional source fields to display below event title |
| `customAdditionalLabels` | `TRUE` | Show labels for additional fields |
| `customAdditionalLabelsStyle` | `Bold` | Style applied to additional field labels |
| `q1StartMonth` | `February` | Month used as the start of Q1 |
| `startWeekOn` | `Sunday` | First day of the week |
| `dayFormat` | `EEEE` | Format for weekday labels |
| `dateFormat` | `d` | Format for date labels inside calendar cells |
| `monthFormat` | `MMMM` | Format for month headers |
| `fontFamily` | `Inter` | Default font family |

Blank setup values fall back to the script defaults. Unknown setup options are ignored.

---

## Appearance options

Appearance options are generated from `CALENDAR.colors`.

Examples:

| Appearance | Appearance-Color |
|---|---|
| `titleBackground` | `#4A0039` |
| `titleFontColor` | `#FFFFFF` |
| `titleRefreshColor` | `#4A0039` |
| `titleAccentColor` | `#7A005D` |
| `titleWeekBackground` | `#EAB8F2` |
| `titleWeekFontColor` | `#000000` |
| `eventDefaultBackground` | `#FFFFFF` |
| `eventDefaultFontColor` | `#000000` |
| `overflowBackground` | `#F2EEEB` |
| `overflowFontColor` | `#4A0039` |

Column `O` is automatically colorized from the hex value in `Appearance-Color`.

Unknown appearance variables are ignored. Blank appearance values fall back to the script defaults.

---

## Date formatting

The script uses Apps Script / Java date-time format patterns for several options.

### `dayFormat`

Controls weekday labels.

| Format | Output |
|---|---|
| `E` | `Tue` |
| `EEEE` | `Tuesday` |

### `dateFormat`

Controls the date label shown inside each calendar day cell.

| Format | Output |
|---|---|
| `d` | `4` |
| `dd` | `04` |
| `EE d` | `Tue 4` |

### `monthFormat`

Controls month headers.

| Format | Output |
|---|---|
| `M` | `9` |
| `MM` | `09` |
| `MMM` | `Sep` |
| `MMMM` | `September` |

---

## Custom additional fields

The `customAdditional` option controls extra source fields displayed below the event title.

Default:

```text
Owner
```

If the source sheet has an `Owner` column, the owner value displays under the event title.

Multiple additional fields can be supplied as comma-separated values.

```text
Owner, Channel, Region
```

Additional fields support optional per-field tokens in the script parser.

| Token | Behavior |
|---|---|
| `true` | Show the field label |
| `false` | Hide the field label |
| `bold` | Bold the field label |
| `italic` | Italicize the field label |
| `uppercase` | Uppercase label and value |
| `lowercase` | Lowercase label and value |
| `label uppercase` | Uppercase only the label |
| `label lowercase` | Lowercase only the label |
| `value uppercase` | Uppercase only the value |
| `value lowercase` | Lowercase only the value |
| hex color | Apply color to the label |

In the current renderer, styling applies to the **additional field label and colon only**, not the value.

Example display:

```text
Campaign Launch
Owner: James
```

Only `Owner:` receives the configured additional label styling.

---

## More cells and Open Selected

If a day has more events than `maxEvents`, the calendar shows a More cell.

```text
3 More...
```

The More cell includes this note:

```text
Use Calendar Tools > Open Selected to view this date.
```

To view all events for a day:

1. Select any cell inside that day block.
2. Go to **Calendar Tools > Open Selected**.

The modal shows all events for that date, including visible events and overflow events.

---

## Refresh behavior

Calendar tabs are view-only.

Every refresh rebuilds the calendar canvas. Manual edits inside the calendar area may be overwritten.

To refresh one calendar:

1. Change controls if needed.
2. Check the refresh checkbox in `G2`.

To refresh all calendars:

```text
Calendar Tools > Refresh All Calendars
```

---

## Key Configurator

The script includes optional Key Configurator utilities.

These can apply:

- Key-based data validation
- Category-based conditional formatting

The visibility of these menu items is controlled by `showKeyConfiguratorMenuItems` in the Key tab setup section.

---

## Recreating the Key tab

The `Key` tab can be deleted and recreated.

If the `Key` tab does not exist, reload the spreadsheet and use:

```text
Calendar Tools > Create Key (and customize)
```

A new Key tab is generated from the current script defaults.

---

## Applying this script to another spreadsheet

The easiest way is to copy the whole spreadsheet:

```text
File > Make a copy
```

This preserves the bound Apps Script project.

If you only want to copy the script:

1. Open the destination spreadsheet.
2. Go to **Extensions > Apps Script**.
3. Create a new `.gs` file.
4. Paste in `Calendar_View_Builder.gs`.
5. Save.
6. Reload the spreadsheet.
7. Create a source sheet named `Events`.
8. Use **Calendar Tools > Create Key (and customize)**.
9. Use **Calendar Tools > New Calendar Sheet**.

---

## Troubleshooting

### The Calendar Tools menu does not appear

Try reloading the spreadsheet.

If it still does not appear, open **Extensions > Apps Script** and confirm the script is saved without errors.

### The calendar is blank

Check that:

- The source sheet exists.
- The source sheet has headers in row 1.
- The date column matches the `customDate` setting.
- The title column matches the `customTitle` setting.
- The selected calendar date range overlaps your event dates.

### Events do not have colors

Check that:

- The source sheet has a Category column.
- The Category values match the Category values in the Key tab.
- The Category-Color values are valid colors.

Valid color examples:

```text
#C8E6C9
#BBDEFB
red
blue
green
```

### Open Selected does not show the expected date

Refresh the calendar tab.

The selected date is calculated from the calendar grid position, so stale layout or manual edits inside the calendar canvas can cause incorrect results.

### Key tab changes do not apply

Refresh the calendar tab or reload the spreadsheet.

Some settings, such as menu visibility, require the spreadsheet to be reloaded.

---

## Important behavior

- The source sheet must be in the same spreadsheet.
- Calendar tabs are rebuilt on refresh.
- Event labels link back to the source sheet.
- Date labels stay bold.
- Unknown setup options are ignored.
- Unknown appearance options are ignored.
- Blank Key tab override values fall back to script defaults.
- The `Key` tab can be deleted and recreated from the Calendar Tools menu.
- Calendar views are intended to be edited through source data, not direct edits to the generated calendar canvas.

---

## Current version

```text
v13.5.0
```
