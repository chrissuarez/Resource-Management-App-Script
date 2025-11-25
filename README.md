# Resource Management App Script

Google Apps Script that automates the paid-media resource capacity workflow for the EMEA team. The script pulls CSV reports from labeled Gmail threads, loads them into dedicated sheets, reshapes the schedule data, builds availability by country, and outputs a final capacity table for planning.

## Features
- **Email ingestion:** `importDataFromEmails()` fetches the latest CSV attachment per Gmail label (schedules, estimates vs actuals, and timecards) and overwrites the matching import tabs.
- **Schedule transformation:** `transformData()` pivots the consolidated schedule sheet into a long-form table with helper keys (`ResourceName-MM-yy`) so downstream lookups stay simple.
- **Availability matrix:** `buildAvailabilityMatrix()` combines “Active staff” records and “Country Hours” to calculate each staffer’s monthly available hours, respecting start dates, FTE, and country-specific working hours.
- **Capacity roll-up:** `buildFinalCapacity()` merges availability, schedules, and leave to calculate billable capacity, non-billable hours, TBH, and staffing metadata per resource/month.
- **One-click refresh:** `refreshAll()` chains the entire workflow, surfaced through the custom “Paid Media Resourcing” menu (`onOpen()`).
- **Region calendar automation:** `setupRegionConfigSheets()` only creates the working-pattern + holiday tables when they're empty/template, and `buildAvailabilityMatrix()` rebuilds the `Country Hours` tab from them before every run so existing formulas keep the same format.
- **Country code normalization:** Names/codes such as `South Africa`, `SA`, and `ZA` are treated as the same country so availability/capacity hours stay aligned and the report can show the preferred label.
- **Country coverage:** When new markets are added, update `COUNTRY_MAP`/`COUNTRY_DISPLAY_OVERRIDES` in `code.gs` (and the central Region Calendar/Holidays) so the pipeline recognizes their codes, hours, and holidays.

## Repository Structure
```
code.gs   // Main Apps Script file (can be pasted into Apps Script or clasp project)
```

## Setup
1. **Spreadsheet**
   - Create a Google Sheet and add the tabs referenced by the script (e.g., `IMPORT-FF Schedules`, `Est vs Act - Import`, `Actuals - Import`, `Consolidated-FF Schedules`, `Active staff`, `Country Hours`).
   - Set the spreadsheet ID in `spreadsheetId` inside `importDataFromEmails()`.
   - Populate `Region Calendar` + `Region Holidays`, then let the automation rebuild `Country Hours` (it keeps `Country | Month | Hours`, so existing lookups remain valid).
   - **Important:** When deploying to production, ensure you copy the "Config" tab from the staging sheet to the production sheet if it's missing or incomplete. This tab contains critical settings that may not be present in a new sheet.

2. **Gmail Labels**
   - Create Gmail labels that match the entries in the `emailConfigs` array (or adjust the array to your labels).
   - Ensure the reporting emails include a CSV attachment encoded as specified (defaults to `ISO-8859-1`).

3. **Apps Script project**
   - Open the Sheet → **Extensions → Apps Script**.
   - Replace the default `Code.gs` with the contents of this repository’s `code.gs`.
   - Save the project and accept the required authorizations on first run (Gmail, Sheets, etc.).

4. **Customizations (optional)**
   - Update `COUNTRY_MAP` at the bottom of `code.gs` to normalize new country names.
   - Adjust the billing-percentage heuristics inside `buildFinalCapacity()` if roles change.
   - If report threads contain multiple messages, consider updating the email parsing logic to grab the latest message in each thread.

## Usage
### Manual
- Open the spreadsheet, wait for the “Paid Media Resourcing” menu (added by `onOpen()`), then choose:
  1. `Import & Transform` – runs the full pipeline (`refreshAll()`).
  2. `Build Availability` – recalculates the availability matrix only (useful after editing staff/hour sheets).
  3. `Build Capacity` – rebuilds the final capacity table after ad-hoc schedule tweaks.
  4. `Scaffold Region Config Sheets` – fills the template tabs only if they're missing or still contain the sample data, preventing accidental overwrites of real configs.

### Scheduled
- In Apps Script, go to **Triggers** and create a time-driven trigger on `refreshAll()` (e.g., daily after the reports arrive).

## Data Requirements
- **Active staff:** Must contain headers matching the regexes the script looks for (Resource Name, Resource Country, Start Date, FTE, Hub, ResourceRole).
- **Country Hours:** Still expected in the sheet, but the script now regenerates it automatically from the Region Calendar/Holidays tables and keeps the `Country | Month | Hours` structure.
- **Consolidated schedules:** Row 2 headers for months (dates or formatted strings) and rows starting at row 3 with resource data.

## Troubleshooting
- Use **View → Logs** while running functions to see Gmail/search/import diagnostics.
- If sheets are empty, confirm the corresponding Gmail label has an unread email with a CSV attachment and that encoding is correct.
- When schema changes break header matching, adjust the regex lookups near the top of each builder function.

## Next Steps
- Add unit tests via the Apps Script Test framework or clasp to detect CSV/Sheet regressions.
- Externalize user-configurable values (labels, sheet names, billing percentages) into a single config object for easier maintenance.
