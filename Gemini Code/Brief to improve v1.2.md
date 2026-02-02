To handle "Scenario Planning" (where you intentionally want to modify/override official numbers) while still solving the "Double Counting" issue (where forecasts accidentally add to official numbers), we need a slightly more sophisticated approach than just "delete everything."

We will implement a **"Priority + Protection"** model.

### The Solution: Priority Logic & "Scenario Lock"

1.  **Priority Logic (The Fix):** We change the code from **Adding** (`Official + Override`) to **Replacing** (`Override > Official`).
    * *Previously:* Official (10) + Forecast (10) = 20 (Double Count).
    * *New Way:* If an Override exists (10), it **wins**. Final result = 10.
    * *Scenario Benefit:* If you want to move time (e.g., zero out a resource), you just type `0` in the override. The system uses the `0` instead of the Official `10`. You no longer have to calculate "negative hours" to balance it out.

2.  **Scenario Lock (The Feature):** We add a **"Lock / Scenario"** checkbox column to your Override tab.
    * **Unchecked (Default - Forecast):** The "Auto-Cleanup" script *will* clear these cells if Official data appears (keeping your sheet clean of stale forecasts).
    * **Checked (Scenario Mode):** The "Auto-Cleanup" script **skips** these rows. Your manual overrides stay forever, allowing you to permanently override the official data for scenario planning.

---

### Updated Brief for Google Anti-Gravity

**Title:** Resource Management Dashboard - "Smart" Override & Scenario Planning

**1. Context**
The user utilizes the `FF Schedule Override` tab for two distinct purposes:
1.  **Forecasting:** Adding hours for projects not yet in Salesforce (which should be removed once they become Official to avoid double counting).
2.  **Scenario Planning:** Intentionally modifying existing Official hours (e.g., moving hours from Resource A to Resource B) which must *persist* even if Official data exists.

**2. Objectives**
* **Prevent Double Counting:** Ensure Override values **replace** Official values rather than adding to them.
* **Enable "Zeroing Out":** Allow users to enter `0` in Overrides to effectively remove an Official allocation for a resource.
* **Smart Cleanup:** Automatically remove "Forecast" overrides when they become Official, but preserve "Scenario" overrides based on a user setting.

**3. Technical Requirements**

**A. Schema Change (Spreadsheet Side)**
* **Action:** In `FF Schedule Override` tab, Insert **Column A** as a Checkbox column named `Scenario Lock`.
* *(Note for Developer: Adjust column index reading in script to account for this new offset).*

**B. Modify `buildFinalSchedules` (in `Code.gs`)**
* **Allow Zeros:** Update the parsing logic (approx Line 58) to accept `0` as a valid override value. Currently, it likely skips `0` or empty cells. It must distinguish between `Empty` (no override) and `0` (override to zero).
* **Replacement Logic:** Change the calculation at **Line 73**:
    * *Current:* `var finalHours = hours + (overrideMap[overrideKey] || 0);`
    * *New:* `var finalHours = (overrideMap[overrideKey] !== undefined) ? overrideMap[overrideKey] : hours;`
    * *Result:* The Override value acts as the "Final Truth" if present.

**C. Create Function: `pruneForecasts()`**
* **Logic:**
    1.  Read the `FF Schedule Override` tab.
    2.  Iterate through rows.
    3.  **Check Column A (Lock):**
        * If **Checked (TRUE):** SKIP this row. (Preserve Scenario).
        * If **Unchecked (FALSE):** Proceed to check columns.
    4.  **Check Cells:** For each month/cell in the row:
        * If `Official Data` exists for this (Resource + Project + Month) AND `Override Cell` is not empty:
        * **Action:** `clearContent()` for that specific cell.
* **Trigger:** Add to the `refreshAll` chain (Phase 2).

**4. User Workflow Example**
* **Forecast:** User enters 10h for "New Project". *Lock is Unchecked.* -> When Official data arrives (10h), script clears the Override. Final = 10h.
* **Scenario:** User sees Official 20h for "Project A". Wants to reduce to 5h. User enters 5h in Override and **Checks the Lock**. -> Script sees Official data but *respects the Lock*. Final = 5h.