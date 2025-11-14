### **Project Context:**

You are helping me refactor a Google Apps Script file (code.gs) for a resource management spreadsheet. The goal is to make the script versatile so it can be easily adapted by other teams.

Currently, the script has team-specific business logic, sheet names, and email labels hard-coded directly into the code. We need to refactor this by moving all this configuration into the spreadsheet itself (primarily into a sheet named Config), turning code.gs into a generic "engine" that reads its instructions from the spreadsheet.

Your Task & Rules:  
I will break this refactor into several phases.

1. You will be given the code for one or more functions from code.gs.  
2. You will complete *only* the single phase I request.  
3. After you provide the modified code for that phase, **stop and wait for my approval.**  
4. Once I approve, I will give you the prompt for the next phase.  
5. For each new phase, we will create a new git branch (e.g., phase-1-role-config, phase-2-email-refactor) to preserve the previous working version.

### **Phase 1: Abstract Billable Rate Logic**

**Branch:** phase-1-role-config

Goal:  
The buildFinalCapacity function calculates a resource's billable percentage (Bill %) using hard-coded logic. We will move this logic to a new sheet named Role Config.  
**Current Hard-Coded Logic (in buildFinalCapacity):**

var bill=/VP/i.test(st.role)?0.5:/Director/i.test(st.role)?0.7:/Executive|Manager/i.test(st.role)?0.8:1;

Your Task:  
Modify the buildFinalCapacity function to read from a new sheet named Role Config instead of using the hard-coded logic above.

1. **Assume** a new sheet named Role Config exists in the spreadsheet. This sheet has two columns:  
   * Column A: Role (e.g., "VP \+", "Senior Director", "Director", "Manager", "(default)")  
   * Column B: Billable % (e.g., "50%", "70%", "80%", "100%")  
2. **Inside buildFinalCapacity**, before the main loop, read the entire Role Config sheet (from row 2 to the end).  
3. **Create a JavaScript Map** or plain object (e.g., billableMap) from this data.  
   * Store the role from Column A as the key (in lowercase).  
   * Store the percentage from Column B as the value (as a decimal, e.g., 0.5).  
   * Look for a (default) role and store its value to be used as a fallback. If no (default) is found, use 1 (100%) as the fallback.  
4. **Replace** the hard-coded var bill=... line with new logic.  
   * Get the resource's role: var resourceRole \= (st.role || '').toLowerCase();  
   * Look up this resourceRole in your billableMap.  
   * Set the bill variable to the found value, or to the default fallback value if the role isn't in the map.

Please provide the *entire modified buildFinalCapacity function* for my review.

### **Phase 2: Abstract Email Import Configuration**

**(Wait for user approval of Phase 1 before proceeding)**

**Branch:** phase-2-email-refactor

Goal:  
The importDataFromEmails function has a hard-coded array of Gmail labels and target sheet names. We will move this configuration into the Config sheet.  
**Current Hard-Coded Logic (in importDataFromEmails):**

var emailConfigs \= \[  
      {  
        label: 'dashboard-reports-earned-media-schdules',  
        sheetName: 'IMPORT-FF Schedules',  
        encoding: 'ISO-8859-1' // Original encoding  
      },  
      {  
        label: 'dashboard-reports-earned-media---est-vs-actual ',  
        sheetName: 'Est vs Act \- Import',  
        encoding: 'ISO-8859-1' // Original encoding  
      },  
      {  
        label: 'dashboard-reports-earned-media-timecards-with-projects', // Your new label  
        sheetName: 'Actuals \- Import',                        // Your new target sheet  
        encoding: 'ISO-8859-1' // Assuming UTF-8 for this new source, adjust if needed  
      }  
    \];

Your Task:  
Modify the importDataFromEmails function to build its configuration from the Config sheet.

1. **Assume** the Config sheet has a table (e.g., starting at row 10\) with the following 4 columns:  
   * Column A: Setting Type (e.g., "Email Import")  
   * Column B: Label / Name (the Gmail label)  
   * Column C: Target Sheet (the sheet name to import to)  
   * Column D: Encoding (e.g., "ISO-8859-1" or "UTF-8")  
2. **Inside importDataFromEmails**, remove the hard-coded emailConfigs array.  
3. **Read the new table** from the Config sheet (e.g., A10:D).  
4. **Dynamically build** the emailConfigs array by looping through the rows you read:  
   * Filter for rows where Column A is exactly "Email Import".  
   * For each of these rows, create an object:  
     * label: value from Column B  
     * sheetName: value from Column C  
     * encoding: value from Column D (or 'UTF-8' if blank)  
5. The rest of the function, which loops through this emailConfigs array, should remain the same.

Please provide the *entire modified importDataFromEmails function* for my review.

### **Phase 3: Create Central Config & Refactor All Functions**

**(Wait for user approval of Phase 2 before proceeding)**

**Branch:** phase-3-central-config

Goal:  
Many functions still use hard-coded sheet names (e.g., 'Consolidated-FF Schedules', 'Active staff') and values (e.g., the number 7 in transformData). We will create a central helper function to read all such settings from the Config sheet and refactor all functions to use it.  
**Your Task:**

1. **Create a new helper function** named getConfigSettings().  
   * This function should read a key-value table from the Config sheet (e.g., in range A2:C).  
   * **Assume** the Config sheet table has columns:  
     * Column A: Setting Type (e.g., "Sheet Name", "Sheet Config")  
     * Column B: Key (e.g., "Staff", "Consolidated Schedules", "Data Start Column")  
     * Column C: Value (e.g., "Active staff", "Consolidated-FF Schedules", "8")  
   * The function should process this table and return a single settings object. Example output:  
     {  
       "sheetNames": {  
         "staff": "Active staff",  
         "consolidatedSchedules": "Consolidated-FF Schedules",  
         "finalSchedules": "Final \- Schedules",  
         "availabilityMatrix": "Availability Matrix",  
         "finalCapacity": "Final \- Capacity",  
         "countryHours": "Country Hours"  
       },  
       "sheetConfig": {  
         "dataStartColumn": 8   
       }  
     }

2. **Refactor** the following functions to use this new getConfigSettings() function:  
   * transformData  
   * buildAvailabilityMatrix  
   * buildFinalCapacity  
   * refreshAll (and any other functions that call them)  
3. **In each function:**  
   * Call var config \= getConfigSettings(); at the beginning.  
   * Replace all hard-coded sheet names with the config object:  
     * ss.getSheetByName('Consolidated-FF Schedules') becomes ss.getSheetByName(config.sheetNames.consolidatedSchedules)  
     * ss.getSheetByName('Active staff') becomes ss.getSheetByName(config.sheetNames.staff)  
     * ...and so on for all sheet names.  
   * Replace the hard-coded 7 in transformData with the config value:  
     * for (var j \= 7; ...) becomes for (var j \= config.sheetConfig.dataStartColumn \- 1; ...)  
     * var nr \= row.slice(0, 7); becomes var nr \= row.slice(0, config.sheetConfig.dataStartColumn \- 1);

Please provide the **new getConfigSettings() function** AND the **fully refactored transformData, buildAvailabilityMatrix, buildFinalCapacity, and refreshAll functions** for my review.

### **Phase 4: Dynamically Scaffold Region Config**

**(Wait for user approval of Phase 3 before proceeding)**

**Branch:** phase-4-region-config

Goal:  
The setupRegionConfigSheets function creates template sheets for regions and holidays, but its examples are hard-coded. We need to read the list of in-scope regions from Config sheet cell C6 and use that to build the examples.  
Config Sheet, Cell C6 Value:  
Denmark|France|Germany|South Africa|Spain|United Kingdom|Italy|Netherlands|United Arab Emirates|Australia|Israel|India|Mexico|United States  
**Current Hard-Coded Logic (in setupRegionConfigSheets):**

var regionExamples \= \[  
    \['UK', 'UK', 1, 5, 7.5\],  
    \['EU-Central', 'DE,FR,NL', 1, 5, 8\],  
    \['GCC', 'AE,SA', 7, 4, 8\]  
  \];  
// ...  
var holidayExamples \= \[  
    \['UK', new Date('2025-01-01'), 'New Year\\'s Day'\],  
    \['DE', new Date('2025-10-03'), 'German Unity Day'\],  
    \['AE', new Date('2025-03-31'), 'Eid al-Fitr (placeholder)'\]  
  \];

Your Task:  
Modify the setupRegionConfigSheets function to dynamically build the regionExamples and holidayExamples arrays.

1. **Inside setupRegionConfigSheets**, read the value from Config sheet, cell C6.  
2. **Parse this value:** It's a string with countries separated by |. Split it into an array of country names (e.g., \['Denmark', 'France', 'Germany', ...\]).  
3. **Clear** the hard-coded regionExamples and holidayExamples arrays.  
4. **Dynamically populate** these arrays by looping through your new country array:  
   * For **regionExamples**:  
     * For each country, get its 2-letter code using the existing normalizeCountryCode\_ function (e.g., normalizeCountryCode\_('Denmark') \-\> DK).  
     * Create a row using the country code as both the "Region Code" and "Country Code". Use 1, 5, and 8 as placeholders for Start Day, End Day, and Hours.  
     * *Example row for "Denmark":* \['DK', 'DK', 1, 5, 8\]  
   * For **holidayExamples**:  
     * For each country, get its 2-letter code.  
     * Create a placeholder holiday row.  
     * *Example row for "Denmark":* \['DK', new Date('2025-01-01'), 'New Year\\'s Day (placeholder)'\]

Please provide the *entire modified setupRegionConfigSheets function* for my review.