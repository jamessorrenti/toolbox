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
  themes/
    index.json
    berry.json
    grayscale.json
    dark-mode.json
    sunset.json
    pastels.json
    retro-console.json
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
2. Go to **Extensions > Apps Script**.
3. Create a new `.gs` file named `Calendar_View_Builder`.
4. Paste in the contents of `Calendar_View_Builder.gs`.
5. Save the Apps Script project.
6. Reload the spreadsheet.
7. Use the **Calendar Tools** menu.

### First-run authorization

The first time you run a menu action, Google prompts you to authorize the script. Click **Continue**, pick your account, and approve the requested scopes (read/write the current spreadsheet, show modal dialogs). This is normal for any Apps Script add-on — the source is right there in **Extensions > Apps Script** for you to review.

A separate authorization prompt appears the first time you run **Calendar Tools > Import Theme**. That action uses `UrlFetchApp` to read theme JSON from a public GitHub repo, which requires the additional `script.external_request` scope. Approve once and the prompt does not return.

You can either run the **Optional first-time setup** below to scaffold everything in one click, or wire up your own source tab manually.

---

## Optional first-time setup

`Calendar Tools > Initial Setup` is a one-click bootstrapper for a brand-new spreadsheet:

1. Creates an event list tab (named `Events` by default — controlled by the `defaultDataSheetName` option, see below) with these frozen header columns:

   | Title | Date | Category | Type | Status |
   |---|---|---|---|---|

2. Creates the `Key` tab with default Type icons, Category colors, Status icons, setup options, and appearance colors.
3. Runs the Key Configurator so:
   - `Type`, `Category`, and `Status` columns on the event list get dropdown validation sourced from the matching Key columns.
   - Rows in the event list get colored by `Category` based on the colors defined in the Key tab.

It is safe to re-run. Existing tabs are left in place; only missing tabs are created.

**Using a pre-existing event list?** If you already have a sheet of events (under any tab name and any column layout) you do **not** need Initial Setup. Two options:

- Point your calendar tab's `Source Data` dropdown (`G1`) at your existing tab, and make sure your date column matches `customDate` (default `Date`) and your title column matches `customTitle` (default `Title`). Both can be overridden from the `Key` tab.
- Or run **Calendar Tools > Set Key From Event List** — it walks you through picking which columns hold Date / Title / Type / Category / Status and which unique values become Key entries, then writes them into the Key tab for you. See [Set Key From Event List](#set-key-from-event-list).

You can hide or show the **Initial Setup**, **Create Event List**, **Set Key From Event List**, and Key Configurator menu items from the `Key` tab — see [Menu visibility toggles](#menu-visibility-toggles).

### Changing the default event list name

`defaultDataSheetName` is a Key-tab setup option (default `Events`). Set it to whatever name your organization uses (for example `Tactics List`) and the **Initial Setup**, **Create Event List**, and the Key Configurator will all target that name. Reload the spreadsheet after editing the value.

---

## Source sheet format

By default, the script expects a local source tab named:

```text
Events
```

You can change the default in two ways:

- From the `Key` tab — set the `defaultDataSheetName` setup option to your tab name (recommended).
- From the script itself — edit `APP_CONFIG.dataSheetName` at the top of `Calendar_View_Builder.gs`.

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

| Action | Description | Controlled by |
|---|---|---|
| Initial Setup | One-shot bootstrap: creates the event list + Key tabs and runs the Key Configurator | `showInitialMenu` |
| New Calendar Sheet | Creates a new calendar tab named `Calendar View` | — |
| Replace with Calendar | Replaces the active tab with a calendar view | — |
| Refresh All Calendars | Rebuilds all calendar tabs | — |
| Open Selected | Opens a modal for the selected calendar day | — |
| Add Q1-Q4 | Creates one tab for each quarter | — |
| Add Jan-Dec | Creates one tab for each month | — |
| Create Event List | Creates an event list tab with the default headers, if one does not exist. Auto-runs the Key Configurator when the `Key` tab is present. | `showEventListMenu` |
| Set Key From Event List | Walk-through that reads an existing event list and writes its Date / Title / Type / Category / Status values into the Key tab | `showSetKeyFromEventListMenu` |
| Import Theme | Fetches a list of themes from a public repo and writes the chosen one into the Key tab — or, optionally, into the active calendar tab's [Theme Override](#theme-override-per-calendar-tab) band | `showImportThemeMenu` (also auto-shown when any tab has an override) |
| Add Theme Override | Adds a per-tab Theme Override band to the active calendar tab, pre-populated with its currently-effective values | `showImportThemeMenu` (only when active tab is a calendar without an override) |
| Run key configurator | Applies both validation and Category colors to the event list | `showKeyConfiguratorMenuItems` |
| Set key-based validation | Applies just the validation step | `showKeyConfiguratorMenuItems` |
| Set key-based colors | Applies just the conditional formatting step | `showKeyConfiguratorMenuItems` |
| Create Key (and customize) | Creates the Key tab if it does not exist | shown only when `Key` is missing |
| Update Key (add missing features) | Appends any newly-introduced Key options to an existing Key tab without overwriting current values | shown only when the Key tab is missing options the current script version knows about |

### Menu visibility toggles

Four toggles in the `Key` tab's **Additional Setup** block control which menu sections are shown. Each toggle defaults to `TRUE` in the script (so a brand-new spreadsheet with no Key shows everything) and `FALSE` in the Key tab when the Key is created (so once you have a Key, menu items are hidden until you opt in).

| Option | Default in Key | Controls |
|---|---|---|
| `showInitialMenu` | `FALSE` | **Initial Setup** |
| `showEventListMenu` | `FALSE` | **Create Event List** |
| `showSetKeyFromEventListMenu` | `FALSE` | **Set Key From Event List** |
| `showImportThemeMenu` | `FALSE` | **Import Theme** |
| `showKeyConfiguratorMenuItems` | `FALSE` | **Run key configurator**, **Set key-based validation**, **Set key-based colors** |

Reload the spreadsheet after flipping a toggle for the menu to re-render.

---

## Set Key From Event List

`Calendar Tools > Set Key From Event List` is an interactive walk-through that builds the Key from a real event list. Useful when you have an existing list with your own Categories, Types, and Statuses.

The flow:

1. Pick the event list tab — type a number from the list, or type the tab name.
2. **Date column** — accept the detected column (default `Date`), pick a numbered column, or type a custom column name.
3. **Title column** — same as above (default `Title`).
4. **Type / Category / Status columns** — same, but each may be skipped by leaving the prompt blank.
5. **Type / Category / Status values** — for each column you kept, pick the values to add to the Key:
   - `all` includes every unique value found in that column.
   - `1,3,5` picks numbered options from the offered list.
   - `Webinar, Direct Mail, Custom Name` accepts any mix of numbers and custom names (custom names are added as-is).
6. Confirm the summary. The Key tab is updated:
   - `defaultDataSheetName` is set to the chosen event list tab so the rest of the Calendar Tools menu targets the same sheet.
   - `Type`, `Category`, and `Status` sections are replaced with the chosen values. Matching default icons / colors from the script are reused; new categories get colors rotated from the default palette.
   - `customDate` and `customTitle` in the setup section are updated to match your choices, and their dropdowns refresh to the new tab's headers.

The walk-through never touches the event list tab itself. If the Key tab does not exist, it is created first.

---

## Import Theme

`Calendar Tools > Import Theme` pulls themes from a public repo and writes the chosen palette into the Key tab — or, optionally, into the active calendar tab's [Theme Override](#theme-override-per-calendar-tab) band.

> **First-run note:** Import Theme uses `UrlFetchApp` to fetch the theme files, which triggers an extra Google authorization prompt the first time you run it (the `script.external_request` scope). Approve once and the prompt does not return. See [First-run authorization](#first-run-authorization).

The flow:

1. The script fetches a manifest (`index.json`) from the configured theme repo.
2. A numbered prompt lists every theme with its description.
3. You pick by number or by name.
4. **If you ran Import Theme from a calendar tab**, a target prompt appears:
   - **YES** — apply to the active calendar tab only (creates the Theme Override band if missing).
   - **NO** — apply to the Key tab (affects every calendar without its own override).
   - **CANCEL** — abort.
   From a non-calendar tab the prompt is skipped and the theme goes to the Key.
5. The script writes the theme's `colors` to the **Appearance** section, `setup` overrides to **Additional Setup**, and `categoryPalette` (if present) positionally to the **Category** section.
6. If targeted at a calendar tab, that calendar is refreshed automatically. If targeted at the Key, run **Refresh All Calendars** to see the change.

The default registry ships with these themes:

| Theme | Vibe |
|---|---|
| **Berry** | Deep purple title, magenta/teal/blue month rotation. The current default. |
| **Grayscale** | Pure monochrome. Title near-black, months cycle through dark, medium, light gray. |
| **Dark Mode** | Dark backgrounds, light text. Blue, green, and purple month accents. |
| **Sunset** | Warm oranges, corals, and dusky pinks. |
| **Pastels** | Soft muted pastels — lavender, blush, sage, sky. |
| **Retro Console** | Black background, neon green text, Courier New monospace. Matrix terminal vibes. |

### Theme schema

A theme file is a small JSON document:

```json
{
  "name": "Theme Display Name",
  "description": "Short tagline shown in the picker.",
  "colors": {
    "titleBackground": "#000000",
    "titleFontColor": "#FFFFFF"
  },
  "setup": {
    "fontFamily": "Courier New"
  },
  "categoryPalette": [
    "#FF0000",
    "#00FF00",
    "#0000FF"
  ]
}
```

- `colors` keys must match entries in `CALENDAR.colors` (e.g. `titleBackground`, `month1HeaderBackground`, `eventDefaultBackground`). Unknown keys are ignored.
- `setup` keys must match entries in `CALENDAR.setup` (e.g. `fontFamily`, `dateFormat`, `customAdditionalLabelsStyle`). Unknown keys are ignored.
- `categoryPalette` is positional: `palette[0]` overwrites the 1st named category in the Category section, `palette[1]` the 2nd, and so on. Categories beyond the palette length keep their existing colors.
- All three fields are optional. A theme can carry any combination.

The manifest (`index.json`) lists available themes:

```json
{
  "version": 1,
  "themes": [
    { "name": "Berry", "file": "berry.json", "description": "..." }
  ]
}
```

### Using your own theme repo

The theme URLs are read from `CALENDAR.themes` at the top of `Calendar_View_Builder.gs`:

```js
themes: {
  indexUrl: "https://raw.githubusercontent.com/jamessorrenti/toolbox/main/calendar-view-builder/themes/index.json",
  baseUrl: "https://raw.githubusercontent.com/jamessorrenti/toolbox/main/calendar-view-builder/themes/"
}
```

Fork the repo (or host your own), point these URLs at your fork, and **Import Theme** will pull from your set. The script uses `UrlFetchApp.fetch` and only requires that the URLs be HTTP-reachable from Google's network.

---

## Theme Override (per calendar tab)

By default, every calendar tab in a spreadsheet uses the same theme — the one defined by the `Key` tab's **Additional Setup** + **Appearance** + **Category** sections.

If you want one calendar to look different from the others (custom palette, different font, different Category colors, etc.) without touching the Key, you can give that tab its own **Theme Override** band. The render-path precedence is:

```
script defaults  →  Key tab overrides  →  per-tab Theme Override  →  rendered output
```

Each calendar's overrides apply only to its own render; they don't leak to other calendars during `Refresh All Calendars`.

### Adding the band

Two ways to add the Theme Override band to a calendar tab:

- **`Calendar Tools → Add Theme Override`** — visible when the active tab is a calendar without an override. Creates the band and pre-populates every value cell with the currently-effective value (script default overlaid with Key overrides), so you see "what this tab is currently using" and can edit individual cells.
- **`Calendar Tools → Import Theme` with target = "this tab only"** — same as above, but applies a chosen theme into the band.

The band lives in columns **H–M** as a single grouped column band, collapsed by default. Click the `+` above column H (or the toggle that appears when a column group is collapsed) to expand it.

### Band layout

| Col | Section | Purpose |
|---|---|---|
| H | spacer | Visual gap from the calendar grid |
| I / J | Additional Setup | Setup-option name / value. Blank value = inherit Key. |
| I / J | Category | Below a 1-row buffer, then `Category` / `Category-Color` sub-header. Blank color = inherit Key. |
| K | spacer | — |
| L / M | Appearance | Appearance-option name / hex color. Blank = inherit Key. |

All **value** cells (J and M) honor the same dropdowns and live color formatting as the Key tab.

### What can be overridden per tab

Everything in `CALENDAR.setup` *except* six spreadsheet-wide concerns: `defaultDataSheetName` and the five menu-visibility toggles. Boolean options use a 3-state dropdown (blank / TRUE / FALSE) so blank still means "inherit Key".

The Category section supports per-category color overrides. The name column is pre-filled from the Key's current categories so you don't have to type them — drop hex values into the color cells of categories you want to override on this tab.

### Import Theme menu visibility

When **any** calendar tab has a Theme Override band, `Calendar Tools → Import Theme` stays visible in the menu even if `showImportThemeMenu` is FALSE in the Key. Once you have an override, you need a way to import a theme to it.

---

## Auto-Refresh

Calendars update themselves when you switch to them, controlled by the **`autoRefresh`** checkbox in the Key tab (defaults to `TRUE`).

When enabled:

1. Every time someone edits a relevant column on the source event list (`Date`, `Title`, `Type`, `Category`, `Status`, or one of the `customAdditional` fields, plus their fallback aliases), the script stamps a `lastSourceChangeAt` timestamp.
2. When *any* user then **switches to a calendar tab**, the script checks whether that calendar's last refresh predates the most-recent source edit. If so, it silently re-renders just that one calendar.

To turn it off, uncheck `autoRefresh` in the Key tab's Additional Setup block.

### What counts as a "relevant" edit

Only the columns the renderer actually consumes invalidate calendars:

- The configured `customDate` and `customTitle` columns
- Common date / title aliases: `Date`, `Start Date`, `Event Date`, `Title`, `Name`, etc.
- `Type`, `Event Type`, `Channel`, `Tactic Type`
- `Category`, `Event Category`, `Theme`, `Product`, `Pillar`
- `Status`, `Event Status`
- `Description`, `Details` (used for hyperlink targets)
- Any columns named by `customAdditional`

Format-only changes, color, and edits to unrelated columns do not invalidate. (Direct edits to the calendar canvas, the Key tab, or the per-tab Theme Override band are unrelated to Auto-Refresh — they don't trigger source-edit invalidation, but they may need a manual refresh for very specific cases.)

### Multi-user behavior

- The `autoRefresh` toggle and the per-calendar refresh timestamps live in the spreadsheet's `Key` tab and `DocumentProperties` respectively — shared across users.
- Each user's "previous active sheet" tracker lives in their own `UserProperties`, so two editors navigating in parallel don't fight each other.

### Limitations

- The simple `onSelectionChange` trigger has a 30-second execution cap. A single calendar refresh is typically 1–3 s with batched rendering, well under the cap. If a calendar render takes longer than 30 s, the on-switch refresh might time out (manual refresh still works).
- If you're already viewing a calendar tab when someone else edits the source, your calendar won't update until you switch away and back — or click the refresh box / menu manually.

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

| Option | Script default | Key default | Description |
|---|---|---|---|
| `defaultDataSheetName` | `Events` | `Events` | Source tab the menu items target (Initial Setup, Create Event List, Key Configurator). Dropdown of available tabs; accepts custom names. |
| `showInitialMenu` | `TRUE` | `FALSE` | Show the **Initial Setup** menu item |
| `showEventListMenu` | `TRUE` | `FALSE` | Show the **Create Event List** menu item |
| `showSetKeyFromEventListMenu` | `TRUE` | `FALSE` | Show the **Set Key From Event List** menu item |
| `showImportThemeMenu` | `TRUE` | `FALSE` | Show the **Import Theme** menu item |
| `showKeyConfiguratorMenuItems` | `TRUE` | `FALSE` | Show the three Key Configurator menu items |
| `autoRefresh` | `TRUE` | `TRUE` | When `TRUE`, calendars auto-refresh on tab switch after source edits. See [Auto-Refresh](#auto-refresh). |
| `frozenWeekdayHeader` | `TRUE` | `FALSE` | Show a frozen weekday header row |
| `customDate` | `Date` | `Date` | Source column used for event dates. Dropdown of headers from `defaultDataSheetName`; accepts custom names. |
| `customTitle` | `Title` | `Title` | Source column used for event titles. Dropdown of headers. |
| `customType` | `Type` | `Type` | Source column used for event Type (icon lookup, validation source). Dropdown of headers. |
| `customCategory` | `Category` | `Category` | Source column used for event Category (color lookup, Key Configurator color rules, validation source). Dropdown of headers. |
| `customStatus` | `Status` | `Status` | Source column used for event Status (icon lookup, validation source). Dropdown of headers. |
| `maxEvents` | `4` | `4` | Maximum visible events per day before showing More |
| `customAdditional` | `Owner` | `Owner` | Additional source fields to display below event title. Dropdown of headers from `defaultDataSheetName`; accepts custom comma-separated lists. Enable **Data → Data validation → Allow multiple selections** on this cell to pick several headers from the dropdown at once. |
| `customAdditionalLabels` | `TRUE` | `TRUE` | Show labels for additional fields |
| `customAdditionalLabelsStyle` | `Bold` | `Bold` | Style applied to additional field labels |
| `q1StartMonth` | `January` | `January` | Month used as the start of Q1. Dropdown of month names. |
| `startWeekOn` | `Sunday` | `Sunday` | First day of the week. Dropdown of day names. `Monday-CompressedWeekend` starts on Monday and renders Saturday + Sunday columns at half width. |
| `dayFormat` | `EEEE` | `EEEE` | Format for weekday labels |
| `dateFormat` | `d` | `d` | Format for date labels inside calendar cells |
| `monthFormat` | `MMMM` | `MMMM` | Format for month headers |
| `fontFamily` | `Inter` | `Inter` | Default font family |

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

Each additional line is rendered with a single leading space to slightly indent it under the event title.

Example display:

```text
Campaign Launch
 Owner: James
```

Only `Owner:` receives the configured additional label styling (the leading space is included in the styled range, which is invisible).

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
v13.13.2
```
