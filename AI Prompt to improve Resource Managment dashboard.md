### **Project Context:**

You are helping me refactor a Google Apps Script file (code.gs) for a resource management spreadsheet. The goal is to make the script versatile so it can be easily adapted by other teams.

We have already refactored the importDataFromEmails function to read its settings from the Config sheet.

Now, we must continue refactoring the rest of the script to move all other hard-coded, team-specific logic into the Config sheet, turning code.gs into a generic "engine."

Your Task & Rules:  
I will break this refactor into several phases.

1. You will be given the code for one or more functions from code.gs.  
2. You will complete *only* the single phase I request.  
3. After you provide the modified code for that phase, **stop and wait for my approval.**  
4. Once I approve, I will give you the prompt for the next phase.  
5. For each new phase, we will create a new git branch (e.g., phase-3-role-config, phase-4-transform-refactor) to preserve the previous working version.

### **Phase 3: Abstract Billable Rate Logic**

**(This was our original Phase 1\. The code file shows it was skipped, so we must do it now.)**

**Branch:** phase-3-role-config

Goal:  
The buildFinalCapacity function calculates a resource's billable percentage (Bill %) using hard-coded logic \[source: chrissuarez/resource-management-app-script/Resource-Management-App-Script-phase-2-email-refactor/code.gs, line 316\]. We will move this logic to a new sheet named Role Config.  
**Current Hard-Coded Logic (in buildFinalCapacity):**

var bill=/VP/i.test(st.role)?0.5:/Director/i.test(st.role)?0.7:/Executive|Manager/i.test(st.role)?0.8:1;

Your Task:  
Modify the buildFinalCapacity function to read from a new sheet named Role Config instead of using the hard-coded logic above.

1. **Assume** a new sheet named Role Config exists in the spreadsheet. This sheet has two columns:  
   * Column A: Role (e.g., "VP \+", "Senior Director", "Director", "Manager", "(default)")  
   * Column B: Billable % (e.g., "50%", "70%", "80%", "100%")  
2. **Inside buildFinalCapacity**, before the main loop \[source: chrissuarez/resource-management-app-script/Resource-Management-App-Script-phase-2-email-refactor/code.gs, line 316\], add code to read the entire Role Config sheet (from row 2 to the end).  
3. **Create a JavaScript Map** or plain object (e.g., billableMap) from this data.  
   * Store the role from Column A as the key (in lowercase).  
   * Store the percentage from Column B as the value (as a decimal, e.g., 0.5).  
   * Look for a (default) role and store its value to be used as a fallback. If no (default) is found, use 1 (100%) as the fallback.  
4. **Replace** the hard-coded var bill=... line \[source: chrissuarez/resource-management-app-script/Resource-Management-App-Script-phase-2-email-refactor/code.gs, line 316\] with new logic.  
   * Get the resource's role: var resourceRole \= (st.role || '').toLowerCase();  
   * Look up this resourceRole in your billableMap.  
   * Set the bill variable to the found value, or to the default fallback value if the role isn't in the map.

Please provide the *entire modified buildFinalCapacity function* for my review.

### **Phase 4: Refactor transformData**

**(Wait for user approval of Phase 3 before proceeding)**

**Branch:** phase-4-transform-refactor

Goal:  
The transformData function \[source: chrissuarez/resource-management-app-script/Resource-Management-App-Script-phase-2-email-refactor/code.gs, line 125\] has hard-coded sheet names and column numbers. We will move these to the Config sheet.  
**Current Hard-Coded Logic:**

* ss.getSheetByName('Consolidated-FF Schedules') \[source: chrissuarez/resource-management-app-script/Resource-Management-App-Script-phase-2-email-refactor/code.gs, line 126\]  
* ss.getSheetByName('Final \- Schedules') \[source: chrissuarez/resource-management-app-script/Resource-Management-App-Script-phase-2-email-refactor/code.gs, line 171\]  
* for (var j \= 7; ...) \[source: chrissuarez/resource-management-app-script/Resource-Management-App-Script-phase-2-email-refactor/code.gs, line 134\] (meaning data starts in column 8\)  
* var nr \= row.slice(0, 7); \[source: chrissuarez/resource-management-app-script/Resource-Management-App-Script-phase-2-email-refactor/code.gs, line 137\] (meaning 7 metadata columns)

**Your Task:**

1. **Assume** we add the following rows to our Config sheet (in the key-value table, e.g., A2:C):

| Setting Type | Key | Value |
| :---- | :---- | :---- |
| Sheet Name | Consolidated Schedules | Consolidated-FF Schedules |
| Sheet Name | Final Schedules | Final \- Schedules |
| Sheet Config | Data Start Column | 8 |

2. **Modify transformData** to read these settings from the Config sheet.  
   * Read the Config sheet to find and store these 3 values.  
   * Replace the hard-coded sheet names with the values from the Config sheet.  
   * Store the Data Start Column value as a variable (e.g., dataStartCol \= 8).  
   * **Crucially**, update the hard-coded 7 to be dataStartCol \- 1\.  
     * for (var j \= 7; ...) becomes for (var j \= dataStartCol \- 1; ...)  
     * row.slice(0, 7\) becomes row.slice(0, dataStartCol \- 1\)  
     * headers.slice(0, 7\) becomes headers.slice(0, dataStartCol \- 1\)

Please provide the *entire modified transformData function* for my review.

### **Phase 5: Refactor buildAvailabilityMatrix**

**(Wait for user approval of Phase 4 before proceeding)**

**Branch:** phase-5-availability-refactor

Goal:  
The buildAvailabilityMatrix function \[source: chrissuarez/resource-management-app-script/Resource-Management-App-Script-phase-2-email-refactor/code.gs, line 190\] also has hard-coded sheet names. We will move these to the Config sheet.  
**Current Hard-Coded Logic:**

* ss.getSheetByName('Active staff') \[source: chrissuarez/resource-management-app-script/Resource-Management-App-Script-phase-2-email-refactor/code.gs, line 191\]  
* ss.getSheetByName('Country Hours') \[source: chrissuarez/resource-management-app-script/Resource-Management-App-Script-phase-2-email-refactor/code.gs, line 192\]  
* ss.getSheetByName('Availability Matrix') \[source: chrissuarez/resource-management-app-script/Resource-Management-App-Script-phase-2-email-refactor/code.gs, line 208\]

**Your Task:**

1. **Assume** we add the following rows to our Config sheet:

| Setting Type | Key | Value |
| :---- | :---- | :---- |
| Sheet Name | Staff | Active staff |
| Sheet Name | Country Hours | Country Hours |
| Sheet Name | Availability Matrix | Availability Matrix |

2. **Modify buildAvailabilityMatrix** to read these settings from the Config sheet.  
   * Read the Config sheet to find and store these 3 values.  
   * Replace all three hard-coded getSheetByName calls with the variables.

Please provide the *entire modified buildAvailabilityMatrix function* for my review.

### **Phase 6: Refactor buildFinalCapacity (Again)**

**(Wait for user approval of Phase 5 before proceeding)**

**Branch:** phase-6-capacity-refactor

Goal:  
The buildFinalCapacity function \[source: chrissuarez/resource-management-app-script/Resource-Management-App-Script-phase-2-email-refactor/code.gs, line 230\], which we already modified for billable rates, also has hard-coded sheet names and a hard-coded project name for "Leave". We must abstract these.  
**Current Hard-Coded Logic:**

* ss.getSheetByName('Availability Matrix') \[source: chrissuarez/resource-management-app-script/Resource-Management-App-Script-phase-2-email-refactor/code.gs, line 231\]  
* ss.getSheetByName('Final \- Schedules') \[source: chrissuarez/resource-management-app-script/Resource-Management-App-Script-phase-2-email-refactor/code.gs, line 232\]  
* ss.getSheetByName('Active staff') \[source: chrissuarez/resource-management-app-script/Resource-Management-App-Script-phase-2-email-refactor/code.gs, line 233\]  
* ss.getSheetByName('Final \- Capacity') \[source: chrissuarez/resource-management-app-script/Resource-Management-App-Script-phase-2-email-refactor/code.gs, line 265\]  
* if(pj==='JFGP All Leave') ... \[source: chrissuarez/resource-management-app-script/Resource-Management-App-Script-phase-2-email-refactor/code.gs, line 259\]

**Your Task:**

1. **Assume** we add the following rows to our Config sheet:

| Setting Type | Key | Value |
| :---- | :---- | :---- |
| Sheet Name | Final Capacity | Final \- Capacity |
| Project Name | Leave Project | JFGP All Leave |

2. **Modify buildFinalCapacity** to read these new settings, *in addition* to the settings it already uses (from Phase 3 & 5).  
   * Read the Config sheet to find and store all necessary values.  
   * Replace the four hard-coded sheet names with variables.  
   * Replace the hard-coded string for the leave project:  
     * if(pj==='JFGP All Leave') becomes if(pj \=== leaveProjectName)

Please provide the *entire modified buildFinalCapacity function* for my review.

### **Phase 7: Optimize with a Global Config Object**

**(Wait for user approval of Phase 6 before proceeding)**

**Branch:** phase-7-config-helper

Goal:  
In Phases 3-6, we have made our functions read from the Config sheet many times. This is repetitive and slow. We will create one helper function to read the config once and pass it to the other functions.  
**Your Task:**

1. **Create a new helper function** getGlobalConfig().  
   * This function should:  
     * Get the active spreadsheet and the Config sheet.  
     * Read the *entire* key-value table (e.g., A2:C on the Config sheet).  
     * Read the Role Config sheet and build the billableMap.  
   * It should return a single, comprehensive config object, like this:  
     {  
       "sheetNames": {  
         "staff": "Active staff",  
         "countryHours": "Country Hours",  
         "availabilityMatrix": "Availability Matrix",  
         "consolidatedSchedules": "Consolidated-FF Schedules",  
         "finalSchedules": "Final \- Schedules",  
         "finalCapacity": "Final \- Capacity"   
       },  
       "sheetConfig": {  
         "dataStartColumn": 8  
       },  
       "projectNames": {  
         "leaveProject": "JFGP All Leave"  
       },  
       "billableMap": {  
         "vp \+": 0.5,  
         "director": 0.7,  
         // ... etc.  
         "(default)": 1.0  
       }  
     }

2. **Modify refreshAll** \[source: chrissuarez/resource-management-app-script/Resource-Management-App-Script-phase-2-email-refactor/code.gs, line 282\]:  
   * Have it call var config \= getGlobalConfig(); *one time* at the beginning.  
   * Pass this config object as a parameter to the functions that need it:  
     * transformData(config)  
     * buildAvailabilityMatrix(config)  
     * buildFinalCapacity(config)  
3. **Modify transformData, buildAvailabilityMatrix, and buildFinalCapacity:**  
   * Change their signatures to accept the config object (e.g., function transformData(config)).  
   * **Remove all code** from inside these functions that reads from the Config or Role Config sheets.  
   * Use the config object passed into them directly (e.g., ss.getSheetByName(config.sheetNames.staff), var bill \= config.billableMap\[resourceRole\] || config.billableMap\['(default)'\];).

This will make the script much faster and cleaner.

Please provide the **new getGlobalConfig() function** and the **modified refreshAll, transformData, buildAvailabilityMatrix, and buildFinalCapacity functions**.