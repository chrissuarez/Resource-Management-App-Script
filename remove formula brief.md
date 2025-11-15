### **Project Goal & Context**

You are an expert Google Apps Script developer. My project is a Google Sheet used for resource management. It's slow and unreliable because it depends on many intermediate "processing" tabs that use slow formulas like QUERY, IMPORTRANGE, and SUMIFS.

My goal is to **eliminate all formula-based processing tabs** and replace them with efficient Apps Script functions. The script will perform all data filtering, aggregation, and transformation *in memory* and write the final, clean data directly to the "Final" tabs, which are used as a Looker Studio data source.

### **Current (Bad) Workflow:**

1. **Import:** Data arrives in "Import" tabs (e.g., IMPORT-FF Schedules, Est vs Act \- Import).  
2. **Manual Filter:** I manually update an IMPORTRANGE \+ QUERY formula on my Active Staff tab to filter by a list of regions I maintain on the Config tab. This is a major bottleneck.  
3. **Formula Processing:** Other "Processing" tabs (e.g., Consolidated-FF Schedules, Est vs Act \- Aggregated) use more formulas to look up, pivot, and aggregate this data.  
4. **Final Output:** The "Final" tabs read from these slow processing tabs.

### **New (Good) Workflow:**

1. **Config:** A single getGlobalConfig() function will read all settings (URLs, sheet names, region lists, business logic) from the Config tab *once*.  
2. **Fetch & Filter:** The script will *replace* my IMPORTRANGE formula. It will fetch the raw staff data and filter it in memory using the "Regions in Scope" list from the Config tab.  
3. **Process in Memory:** The script will replace all other processing tabs, performing pivots, consolidations, and aggregations in memory.  
4. **Write Values:** The script will write the final, static values *directly* to the Final tabs.

### **Rules for Our Work**

* We will do this in phases. I will give you the prompt for one phase at a time.  
* You will provide the *complete code* for that phase.  
* **Stop and wait for my approval** before proceeding to the next phase.  
* All code must read its settings (sheet names, URLs, logic) from the config object.

### **Phase 1: The Config Helper (Foundation)**

Goal:  
Create one central helper function, getGlobalConfig(), that reads all settings from our Config sheet. This function must handle two data types:

1. Key/Value pairs from Columns B:C.  
2. The "Regions in Scope" list, which starts in cell D6.

Task:  
Provide the code for the getGlobalConfig() function. It should read all settings and return a single settings object.  
/\*\*  
 \* Reads all settings from the 'Config' sheet into a single object.  
 \* Reads 2-column key/value pairs from B:C AND specialized ranges.  
 \*/  
function getGlobalConfig() {  
  var ss \= SpreadsheetApp.getActiveSpreadsheet();  
  var configSheet \= ss.getSheetByName('Config');  
  if (\!configSheet) {  
    throw new Error("Configuration sheet named 'Config' was not found.");  
  }  
    
  // 1\. Read the Key/Value pairs from B:C  
  var configData \= configSheet.getRange('B1:C' \+ configSheet.getLastRow()).getValues();  
    
  // Helper to find a value from the 2-column array  
  function findValue(keyToFind) {  
    for (var i \= 0; i \< configData.length; i++) {  
      if (configData\[i\]\[0\] \=== keyToFind) {  
        return configData\[i\]\[1\]; // Return the value from Column C  
      }  
    }  
    return null; // Not found  
  }  
    
  // 2\. Read the "Regions in Scope" list from row 6  
  // Reads D6:Z6 to capture all potential regions  
  var regionsRow \= configSheet.getRange('D6:Z6').getValues()\[0\];   
  var regionsList \= regionsRow.filter(function(cell) {   
    return cell && cell.trim() \!== ''; // Filter out blank cells  
  });

  // 3\. Build the final settings object  
  var settings \= {  
    // \--- IMPORT SHEETS (from email) \---  
    'importSchedules': findValue('IMPORT-FF Schedules') || 'IMPORT-FF Schedules',  
    'importEstVsAct': findValue('Est vs Act \- Import') || 'Est vs Act \- Import',  
    'importActuals': findValue('Actuals \- Import') || 'Actuals \- Import',

    // \--- DATA SOURCES & PROCESSING SHEETS \---  
    'activeStaffUrl': findValue('Active Staff URL'), // \<-- For Phase 2  
    'staffSheet': findValue('Active Staff Sheet') || 'Active staff',  
    'overrideSchedules': findValue('FF Schedule Override Sheet') || 'FF Schedule Override',  
    'countryHours': findValue('Country Hours Sheet') || 'Country Hours',  
    'availabilityMatrix': findValue('Availability Matrix Sheet') || 'Availability Matrix',  
    'consolidatedSchedulesSheet': findValue('Consolidated Schedules Sheet') || 'Consolidated-FF Schedules',  
      
    // \--- FINAL OUTPUT SHEETS \---  
    'finalSchedules': findValue('Final \- Schedules Sheet') || 'Final \- Schedules',  
    'finalCapacity': findValue('Final \- Capacity Sheet') || 'Final \- Capacity',  
    'finalEstVsAct': findValue('Est vs Act \- Aggregated Sheet') || 'Est vs Act \- Aggregated',  
    'roleConfigSheet': findValue('Role Config Sheet') || 'Role Config',

    // \--- LOGIC & VALUES \---  
    'leaveProjectName': findValue('Leave Project Name') || 'JFGP All Leave',  
    'dataStartColumn': parseInt(findValue('Data Start Column') || '8'),  
    'regionCalendarId': findValue('Global Holidays') ? (findValue('Global Holidays').match(/\[-\\w\]{25,}/) || \[null\])\[0\] : null,  
      
    // \--- SPECIALIZED LISTS \---  
    'regionsInScope': regionsList // The list from D6:Z6  
  };  
    
  return settings;  
}

**Stop and wait for my approval.**

### **Phase 2: *New* Function to Import & Filter Staff**

Goal:  
Replace the manual IMPORTRANGE \+ QUERY on the Active Staff tab. This new function will fetch the raw staff data, filter it in memory using our config.regionsInScope list, and paste the clean values into our local Active Staff tab.  
Task:  
Provide the code for a new function, importAndFilterActiveStaff(config).

* It must accept the config object.  
* It must use config.activeStaffUrl to open the source sheet.  
* It must read the source data (assume "Sheet1\!A:Z" or similar).  
* It must find the "Resource Country" column (handle errors if not found).  
* It must build a Set from config.regionsInScope for fast filtering.  
* It must filter the data, keeping only rows (plus the header) where the "Resource Country" is in the Set.  
* It must write these filtered values to the local Active Staff sheet (from config.staffSheet), overwriting old data.

**Stop and wait for my approval.**

### **Phase 3: Refactor *Existing* Script Functions**

Goal:  
Update all the old functions from my original code.gs file to use our new config object, so they no longer use hard-coded sheet names or values.  
Task:  
Provide the modified versions of the following functions. They must all accept a config object (or ss, config) and use it to get sheet names, column numbers, and logic.

1. importDataFromEmails(config)  
2. transformData(config) (This will be temporary, as Phase 5 will replace it)  
3. buildAvailabilityMatrix(config)  
4. refreshCountryHoursFromRegion\_(ss, config) (Update it to use config.regionCalendarId and config.countryHours)  
5. buildFinalCapacity(config) (Update it to use config.leaveProjectName and all sheet names from config)

**Stop and wait for my approval.**

### **Phase 4: *New* Function to Replace Est vs Act Processing**

Goal:  
Create one new function that completely replaces the Est vs Act \- Processed and Est vs Act \- Aggregated formula tabs.  
Task:  
Provide the code for a new function, buildEstVsActAggregate(config).

* It must read *only* from the Est vs Act \- Import sheet (using config.importEstVsAct).  
* It will loop through the rows *in memory*, performing any cleaning or adjustments.  
* It will aggregate (SUM) the Actual Hours and Estimated Hours by Project, Resource, and Month using a JavaScript Map for high speed.  
* It will write the *final, aggregated values* directly to the Est vs Act \- Aggregated sheet (using config.finalEstVsAct), overwriting any old data.

**Stop and wait for my approval.**

### **Phase 5: *New* Function to Replace Schedule Processing**

Goal:  
Create one new function that replaces the Pivot-FF Schedules tab, the Consolidated-FF Schedules tab, AND our old transformData function.  
Task:  
Provide the code for a new function, buildFinalSchedules(config).

* It must read data from two sources:  
  1. IMPORT-FF Schedules (using config.importSchedules)  
  2. FF Schedule Override (using config.overrideSchedules)  
* It will perform the un-pivoting logic (turning date columns into rows) *in memory*.  
* It will read the FF Schedule Override data into a Map for fast lookups.  
* As it processes the un-pivoted data, it will check the Map and apply any overrides *in memory*.  
* It will write the *final, consolidated, and un-pivoted* data directly to the Final \- Schedules sheet (using config.finalSchedules), overwriting old data.

**Stop and wait for my approval.**

### **Phase 6: *New* Setup & Role Config Functions**

Goal:  
Provide the helper scripts that non-destructively set up the Config and Role Config tabs with the required data, so the main script can run.  
Task:  
Provide the code for two functions:

1. setupConfigTab(): This function reads all the keys from Phase 1 (e.g., 'Active Staff URL', 'Leave Project Name') and appends any that are *missing* to the Config tab (Columns B:C) with example values. It must not overwrite existing keys.  
2. setupRoleConfigTab(): This function checks if a sheet named Role Config (or from config.roleConfigSheet) exists. If not, it creates it and populates it with headers (Role, Billable %) and default data (e.g., "VP \+", "50%"; "(default)", "100%"). It must not overwrite an existing, populated sheet.

**Stop and wait for my approval.**

### **Phase 7: Final refreshAll() Function**

Goal:  
Tie everything together into one master "Refresh" button that runs the entire, efficient, code-based pipeline in the correct order.  
Task:  
Provide the final refreshAll() function and the onOpen() function.

* The onOpen() function should include menu items for refreshAll and for the runSetup (wrapper for Phase 6\) functions.  
* The refreshAll() function must be clean and simple, showing the correct order of operations:  
  1. var config \= getGlobalConfig(); (Get all settings first)  
  2. var ss \= SpreadsheetApp.getActiveSpreadsheet();  
  3. importDataFromEmails(config)  
  4. importAndFilterActiveStaff(config) (New staff import)  
  5. buildEstVsActAggregate(config) (New Est vs Act)  
  6. refreshCountryHoursFromRegion\_(ss, config) (Refactored)  
  7. buildAvailabilityMatrix(config) (Refactored)  
  8. buildFinalSchedules(config) (New schedule build)  
  9. buildFinalCapacity(config) (Refactored)  
* **Crucially, refreshAll must *no longer call* the old transformData function, as buildFinalSchedules has replaced it.**