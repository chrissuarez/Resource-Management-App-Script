/**
 * Imports paid media EMEA schedules via email label and CSV, then transforms the data,
 * builds availability matrix, and final capacity tables.
 * Adds custom menu for manual refresh and convenience wrappers.
 */

// ----- 0. Import schedules and actuals from email -----
function importDataFromEmails() {
  try {
    var config = getGlobalConfig();
    var ss = SpreadsheetApp.openById(getSpreadsheetIdFromConfig_());
    var emailConfigs = [];
    var configSheet = ss.getSheetByName('Config');
    if (configSheet) {
      var startRow = 10;
      var lastRow = configSheet.getLastRow();
      if (lastRow >= startRow) {
        var configData = configSheet.getRange(startRow, 2, lastRow - startRow + 1, 4).getValues();
        configData.forEach(function(row) {
          if ((row[0] + '').trim() !== 'Email Import') return;
          var label = (row[1] + '').trim();
          var sheetName = (row[2] + '').trim();
          if (!label || !sheetName) return;
          var encoding = (row[3] + '').trim() || 'UTF-8';
          emailConfigs.push({ label: label, sheetName: sheetName, encoding: encoding });
        });
      }
    }

    if (!emailConfigs.length) {
      Logger.log('No Email Import configurations found in Config sheet.');
      return;
    }

    emailConfigs.forEach(function(config) {
      Logger.log('Processing label: ' + config.label + ' for sheet: ' + config.sheetName);
      
      // Search for emails with the label. GmailApp.search returns threads sorted newest first.
      var threads = GmailApp.search('label:' + config.label); 
      
      if (!threads.length) {
        Logger.log('No threads found for ' + config.label);
        return; // continue to next config in the forEach loop
      }
      
      // Get the newest message from the newest thread
      var message = threads[0].getMessages()[0]; 
      // If a thread could have multiple messages and you need the *very latest* message
      // across all messages in that thread, you might need to sort messages by date.
      // However, threads[0].getMessages()[0] usually gives the initial or most relevant message.
      // For single-message threads (common for reports), this is fine.
      // If you want the absolute latest message in the thread:
      // var messagesInThread = threads[0].getMessages();
      // var message = messagesInThread[messagesInThread.length - 1];


      var attachments = message.getAttachments();
      
      // Find the CSV attachment. Since you confirmed only one attachment (which is the CSV),
      // you could theoretically use attachments[0] if you are absolutely certain no other
      // inline images (like signatures) could ever be present.
      // Using .find is safer.
      var csvAttachment = attachments.find(att => att.getContentType() === 'text/csv' || att.getContentType() === 'application/csv');

      if (!csvAttachment) {
        Logger.log('No CSV attachment found in the latest email for ' + config.label + '. Email subject: ' + message.getSubject());
        return; // continue to next config
      }
      
      var charEncoding = config.encoding || 'UTF-8'; // Default to UTF-8 if not specified
      var csvDataString = csvAttachment.getDataAsString(charEncoding);
      var data = Utilities.parseCsv(csvDataString);

      if (!data || data.length === 0) {
        Logger.log('CSV data is empty for ' + config.label);
        return; // continue to next config
      }

      var sheet = ss.getSheetByName(config.sheetName);
      if (!sheet) {
        Logger.log('Sheet not found: ' + config.sheetName + '. Creating it.');
        sheet = ss.insertSheet(config.sheetName);
      } else {
        Logger.log('Clearing existing content from sheet: ' + config.sheetName);
        // This clears all content and formatting. Assumes the CSV is a full replacement.
        sheet.clearContents(); 
      }
      
      // Write data, assuming CSV includes headers starting from the first row.
      sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      Logger.log('Imported data from label "' + config.label + '" to sheet "' + config.sheetName + '". Rows: ' + data.length + ". Email subject: " + message.getSubject());
      
      // Optional: Mark email as read or archive it
      // message.markRead();
      // GmailApp.moveThreadToArchive(threads[0]);
    });
  } catch(e) {
    Logger.log('Import error: ' + e.toString() + ' Stack: ' + e.stack);
    // Consider sending an email notification for critical errors
    // MailApp.sendEmail('your-email@example.com', 'App Script Import Error', 'Error: ' + e.toString() + '\nStack: ' + e.stack);
  }
}

/**
 * Transforms the imported schedules and adds a Helper column.
 */
function buildFinalSchedules(config) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var importSheet = ss.getSheetByName(config.importSchedules || 'IMPORT-FF Schedules');
  if (!importSheet) throw new Error('Import schedules sheet not found');
  var data = importSheet.getDataRange().getValues();
  if (data.length < 2) throw new Error('Import schedules sheet has no data');
  var headers = data[0];
  var rows = data.slice(1);
  var metadataColumns = Math.max(1, parseInt(config.dataStartColumn, 10) || 8) - 1;
  var timezone = ss.getSpreadsheetTimeZone();

  function findIndex(pats) {
    return headers.findIndex(function(h){
      return pats.some(function(p){return new RegExp(p,'i').test(h);});
    });
  }
  var resourceIdx = findIndex(['Resource Name','ResourceName']);
  if (resourceIdx === -1) resourceIdx = metadataColumns - 1;
  var projectIdx = findIndex(['Project']);

  var overrideSheet = ss.getSheetByName(config.overrideSchedules || 'FF Schedule Override');
  var overrideMap = {};
  if (overrideSheet && overrideSheet.getLastRow() > 1) {
    var overrideData = overrideSheet.getDataRange().getValues();
    var oHeaders = overrideData[0];
    var oRows = overrideData.slice(1);
    function findOverrideIndex(pats) {
      return oHeaders.findIndex(function(h){
        return pats.some(function(p){return new RegExp(p,'i').test(h);});
      });
    }
    var oRes = findOverrideIndex(['Resource Name','Resource']);
    var oProj = findOverrideIndex(['Project']);
    var oDate = findOverrideIndex(['Date']);
    var oVal  = findOverrideIndex(['Hours','Value']);
    if (oRes > -1 && oDate > -1 && oVal > -1) {
      oRows.forEach(function(row){
        var res = (row[oRes] + '').trim();
        if (!res) return;
        var proj = oProj > -1 ? (row[oProj] + '').trim() : '';
        var rawDate = row[oDate];
        var dateObj = rawDate instanceof Date ? rawDate : new Date(rawDate);
        if (isNaN(dateObj)) return;
        var monthKey = Utilities.formatDate(new Date(dateObj.getFullYear(), dateObj.getMonth(), 1), timezone, 'yyyy-MM');
        var key = [res, proj, monthKey].join('|');
        overrideMap[key] = parseFloat(row[oVal]) || 0;
      });
    }
  }

  function parseHeaderDate(value) {
    if (value instanceof Date) return value;
    if (typeof value === 'string') {
      if (/^[A-Za-z]{3}-\d{2}$/.test(value)) {
        var parts = value.split('-');
        var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
        var idx = months.indexOf(parts[0]);
        if (idx > -1) return new Date(2000 + parseInt(parts[1], 10), idx, 1);
      }
      var match = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (match) {
        return new Date(parseInt(match[3],10), parseInt(match[2],10) - 1, parseInt(match[1],10));
      }
      var parsed = new Date(value);
      if (!isNaN(parsed)) return parsed;
    }
    return new Date(value);
  }

  var outRows = [];
  rows.forEach(function(row){
    for (var j = metadataColumns; j < row.length; j++) {
      var hours = row[j];
      if (hours === '' || hours === null || typeof hours === 'undefined') continue;
      var base = row.slice(0, metadataColumns);
      var dateValue = headers[j];
      var dt = parseHeaderDate(dateValue);
      if (isNaN(dt)) continue;
      var resourceName = resourceIdx > -1 ? (row[resourceIdx] + '').trim() : '';
      var projectName = projectIdx > -1 ? (row[projectIdx] + '').trim() : '';
      var monthKey = Utilities.formatDate(new Date(dt.getFullYear(), dt.getMonth(), 1), timezone, 'yyyy-MM');
      var overrideKey = [resourceName, projectName, monthKey].join('|');
      var finalHours = overrideMap.hasOwnProperty(overrideKey) ? overrideMap[overrideKey] : hours;

      var helperKey = resourceName + '-' + Utilities.formatDate(dt, timezone, 'MM-yy');
      var newRow = base.slice();
      newRow.push(dt, finalHours, helperKey);
      outRows.push(newRow);
    }
  });

  var finalSheet = ss.getSheetByName(config.finalSchedules || 'Final - Schedules') || ss.insertSheet(config.finalSchedules || 'Final - Schedules');
  finalSheet.clearContents();
  var headerRow = headers.slice(0, metadataColumns).concat(['Date','Value','Helper']);
  finalSheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
  if (outRows.length) {
    finalSheet.getRange(2, 1, outRows.length, outRows[0].length).setValues(outRows);
  }
}

// ----- 2. Build availability matrix -----
function buildAvailabilityMatrix(config) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var staffSheet = ss.getSheetByName(config.staffSheet || 'Active staff');
  var hoursSheet = ss.getSheetByName(config.countryHours || 'Country Hours');
  if (!staffSheet || !hoursSheet) throw new Error('Missing staff or hours sheet');

  var sd = staffSheet.getDataRange().getValues(), hdr = sd[0], rows = sd.slice(1);
  function find(pats){return hdr.findIndex(h=>pats.some(p=>new RegExp(p,'i').test(h)));}
  var iName = find(['ResourceName','Resource Name']), iCountry = find(['Resource Country']), iStart = find(['Start Date']), iFTE = find(['FTE']);

  var hd = hoursSheet.getDataRange().getValues().slice(1), hmap = {};
  hd.forEach(function(r){
    var ct = normalizeCountryCode_(r[0]);
    if (!ct) return;
    var dt = r[1], hrs = parseFloat(r[2]) || 0;
    var d = dt instanceof Date ? dt : parseMonthYearValue_(dt);
    var key = Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), 'yyyy-MM');
    hmap[ct+'|'+key] = hrs;
  });
  var months = Object.keys(hmap).map(k=>k.split('|')[1]).filter((v,i,a)=>a.indexOf(v)===i).sort();

  var out = ss.getSheetByName(config.availabilityMatrix || 'Availability Matrix')||ss.insertSheet(config.availabilityMatrix || 'Availability Matrix');
  out.clearContents(); out.getRange(1,1).setValue('ResourceName');
  months.forEach(function(m,i){var d=new Date(m+'-01');out.getRange(1,i+2).setValue(d).setNumberFormat('MMM-yy');});

  function cw(s,e){var c=0;for(var d=new Date(s);d<=e;d.setDate(d.getDate()+1)){if(d.getDay()>0&&d.getDay()<6)c++;}return c;}
  var outData=[];
  rows.forEach(function(r){
    var name=r[iName]+''; if(!name)return;
    var ctCode=normalizeCountryCode_(r[iCountry]);
    var st=r[iStart]?new Date(r[iStart]):null;
    var f=parseFloat(r[iFTE])||0;
    var row=[name];
    months.forEach(function(m){
      var ms=new Date(m+'-01'), me=new Date(ms.getFullYear(),ms.getMonth()+1,0);
      var key=Utilities.formatDate(ms,ss.getSpreadsheetTimeZone(),'yyyy-MM');
      var wh=ctCode? (hmap[ctCode+'|'+key]||0):0;
      var tot=cw(ms,me);
      var av=!st?tot:(st>me?0:(st<=ms?tot:cw(st,me)));
      row.push(av? (f*wh)*(av/tot): '');
    });
    outData.push(row);
  });
  if(outData.length) out.getRange(2,1,outData.length,outData[0].length).setValues(outData);
}

// ----- 3. Build final capacity table -----
function buildFinalCapacity(config) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var avail = ss.getSheetByName(config.availabilityMatrix || 'Availability Matrix');
  var sched = ss.getSheetByName(config.finalSchedules || 'Final - Schedules');
  var staff = ss.getSheetByName(config.staffSheet || 'Active staff');
  if(!avail||!sched||!staff) throw new Error('Missing sheets');

  var roleSheet = ss.getSheetByName(config.roleConfigSheet || 'Role Config') || ensureRoleConfigSheet_(ss);
  function parseBillableValue_(value) {
    if (value === null || typeof value === 'undefined') return null;
    if (typeof value === 'number') {
      if (isNaN(value)) return null;
      return value > 1 ? value / 100 : value;
    }
    var str = (value + '').trim();
    if (!str) return null;
    var cleaned = str.replace(/%/g, '');
    var num = parseFloat(cleaned);
    if (isNaN(num)) return null;
    if (str.indexOf('%') > -1 || num > 1) {
      return num / 100;
    }
    return num;
  }

  var billableMap = {};
  var defaultBillable = 1;
  if (roleSheet) {
    var lastRoleRow = roleSheet.getLastRow();
    if (lastRoleRow >= 2) {
      var roleValues = roleSheet.getRange(2, 1, lastRoleRow - 1, 2).getValues();
      roleValues.forEach(function(row) {
        var roleName = (row[0] + '').trim();
        if (!roleName) return;
        var value = parseBillableValue_(row[1]);
        if (value === null) return;
        var key = roleName.toLowerCase();
        if (key === '(default)') {
          defaultBillable = value;
        } else {
          billableMap[key] = value;
        }
      });
    }
  }

  var sd = staff.getDataRange().getValues(), sh=sd[0], sr=sd.slice(1);
  function fi(p){return sh.findIndex(h=>p.some(x=>new RegExp(x,'i').test(h)));}
  var iRes = fi(['ResourceName']), iRR=fi(['ResourceRole']), iHub=fi(['Hub']), iC=fi(['Resource Country']);
  var staffMap={};
  sr.forEach(function(r){
    var n=r[iRes]+''; if(!n) return;
    var pr=r[iRR]+''; var ps=pr.split('-');
    var countryOriginal=r[iC]+'';
    var countryCode=normalizeCountryCode_(countryOriginal);
    staffMap[n]={
      hub:r[iHub]+'',
      practice:ps[0].trim(),
      role:ps[1]?ps.slice(1).join('-').trim():'',
      countryOriginal:countryOriginal,
      countryCode:countryCode,
      country:formatCountryDisplay_(countryOriginal,countryCode)
    };
  });

  var ad=avail.getDataRange().getValues(), am=ad[0].slice(1).map(d=>Utilities.formatDate(new Date(d),ss.getSpreadsheetTimeZone(),'yyyy-MM'));
  var fullMap={}; ad.slice(1).forEach(r=>{var n=r[0]+''; r.slice(1).forEach((v,i)=>{fullMap[n+'|'+am[i]]=parseFloat(v)||0;});});

  var sd2=sched.getDataRange().getValues(), sh2=sd2[0], sr2=sd2.slice(1);
  var iProj=sh2.findIndex(h=>/Project/i.test(h)), iVal=sh2.findIndex(h=>/Value|Hours/i.test(h)), iHelp=sh2.findIndex(h=>/Helper/i.test(h));
  var leave={}, schedM={};
  sr2.forEach(r=>{var pj=r[iProj]+'', h=r[iHelp]+''; var m=h.match(/^(.+)-(\d{2})-(\d{2})$/); if(!m)return;
    var nm=m[1], mo=m[2], yr=m[3]; var key=(yr.length===2?('20'+yr):yr)+'-'+mo; var hrs=parseFloat(r[iVal])||0;
    if(pj===(config.leaveProjectName||'JFGP All Leave')) leave[nm+'|'+key]=(leave[nm+'|'+key]||0)+hrs;
    else schedM[nm+'|'+key]=(schedM[nm+'|'+key]||0)+hrs;
  });

  var fo=ss.getSheetByName(config.finalCapacity || 'Final - Capacity')||ss.insertSheet(config.finalCapacity || 'Final - Capacity'); fo.clear();
  var hdr=['Resource Name','Hub','Role','Country','Bill %','Practice','Month-Year','Full Hours','Annual Leave','NB Hours','TBH','Sched Hrs','Billable Capacity'];
  fo.getRange(1,1,1,hdr.length).setValues([hdr]); var out=[];
  Object.keys(fullMap).forEach(function(k){
    var p=k.split('|'),n=p[0],m=p[1],st=staffMap[n]||{};
    var resourceRole=(st.role||'').toLowerCase();
    var bill=billableMap.hasOwnProperty(resourceRole)?billableMap[resourceRole]:defaultBillable;
    var f=fullMap[k], al=leave[k]||0, net=f-al; var tbh=net*bill, nb=net*(1-bill), sch=schedM[k]||0, bc=tbh-sch;
    var monthDate = new Date(m+'-01');
    out.push([n,st.hub,st.role,st.country,bill,st.practice,monthDate,f,al,nb,tbh,sch,bc]);
  });
  if(out.length){
    fo.getRange(2,1,out.length,out[0].length).setValues(out);
    fo.getRange(2,5,out.length,1).setNumberFormat('0%');
    fo.getRange(2,7,out.length,1).setNumberFormat('mmm-yy');
  }
}

// ----- 4. Wrapper to run full refresh -----
function importAndFilterActiveStaff(config) {
  var sourceUrl = config.activeStaffUrl;
  if (!sourceUrl) throw new Error('Active Staff URL missing from Config sheet');
  var source = SpreadsheetApp.openByUrl(sourceUrl);
  var sheet = source.getSheets()[0];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error('Active staff source has no data');
  var headers = data[0];
  var rows = data.slice(1);
  var destSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.staffSheet || 'Active staff');
  if (!destSheet) throw new Error('Destination staff sheet not found');
  var countryIdx = headers.findIndex(function(h){return /Resource Country/i.test(h);});
  if (countryIdx === -1) throw new Error('Resource Country column missing');
  var regionSet = new Set((config.regionsInScope || []).map(function(x){return (x + '').toLowerCase();}));
  var filtered = [headers];
  rows.forEach(function(row){
    var country = (row[countryIdx] + '').toLowerCase();
    if (!regionSet.size || regionSet.has(country)) {
      filtered.push(row);
    }
  });
  destSheet.clearContents();
  destSheet.getRange(1, 1, filtered.length, filtered[0].length).setValues(filtered);
}

/** Country mapping **/
var COUNTRY_MAP = {
  'United Kingdom':'UK','Germany':'DE','Denmark':'DK','France':'FR',
  'South Africa':'ZA','Spain':'ES','Netherlands':'NL','Italy':'IT',
  'United Arab Emirates':'AE','UAE':'AE','AE':'AE',
  'Australia':'AU','AU':'AU',
  'Israel':'IL','IL':'IL',
  'India':'IN','IN':'IN',
  'Mexico':'MX','MX':'MX',
  'United States':'US','USA':'US','US':'US',
  'Japan':'JP','JP':'JP',
  'SA':'ZA','ZA':'ZA'
};
var COUNTRY_DISPLAY_OVERRIDES = {
  'ZA':'South Africa',
  'AE':'United Arab Emirates',
  'US':'United States'
};

function getSpreadsheetIdFromConfig_() {
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) return '';
  var fallback = active.getId();
  var configSheet = active.getSheetByName('Config');
  if (!configSheet) return fallback;
  var raw = (configSheet.getRange('C24').getDisplayValue() + '').trim();
  if (!raw) return fallback;
  var match = raw.match(/[-\w]{25,}/);
  return match ? match[0] : raw;
}

function normalizeCountryCode_(value) {
  if (value === null || typeof value === 'undefined') return '';
  var str = (value + '').trim();
  if (!str) return '';
  if (str.length === 2) str = str.toUpperCase();
  return COUNTRY_MAP[str] || str;
}

function formatCountryDisplay_(original, canonical) {
  var code = canonical || normalizeCountryCode_(original);
  var base = COUNTRY_DISPLAY_OVERRIDES[code] || original || code;
  if (code && COUNTRY_DISPLAY_OVERRIDES[code]) {
    return base + ' (' + code + ')';
  }
  return base || code;
}


var DEFAULT_MONTH_RANGE_MONTHS = 36;

function ensureSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/**
 * Reads configuration values from the Config sheet and returns a settings object.
 */
function getGlobalConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName('Config');
  if (!configSheet) {
    throw new Error("Configuration sheet named 'Config' was not found.");
  }

  var lastRow = configSheet.getLastRow() || 1;
  var configData = configSheet.getRange(1, 2, lastRow, 2).getDisplayValues(); // Columns B:C
  function findValue(keyToFind) {
    keyToFind = (keyToFind || '').trim();
    if (!keyToFind) return '';
    for (var i = 0; i < configData.length; i++) {
      var row = configData[i];
      if ((row[0] + '').trim() === keyToFind) {
        var value = row[1];
        return value !== null && typeof value !== 'undefined' ? (value + '').trim() : '';
      }
    }
    return '';
  }

  var regionsRow = configSheet.getRange('D6:Z6').getDisplayValues()[0] || [];
  var regionsList = regionsRow.map(function(cell){return (cell + '').trim();}).filter(function(cell){return cell;});

  var rawCalendarLink = findValue('Global Holidays');
  var calendarMatch = rawCalendarLink ? rawCalendarLink.match(/[-\w]{25,}/) : null;

  var settings = {
    importSchedules: findValue('IMPORT-FF Schedules') || 'IMPORT-FF Schedules',
    importEstVsAct: findValue('Est vs Act - Import') || 'Est vs Act - Import',
    importActuals: findValue('Actuals - Import') || 'Actuals - Import',
    activeStaffUrl: findValue('Active Staff URL'),
    staffSheet: findValue('Active Staff Sheet') || 'Active staff',
    overrideSchedules: findValue('FF Schedule Override Sheet') || 'FF Schedule Override',
    countryHours: findValue('Country Hours Sheet') || 'Country Hours',
    availabilityMatrix: findValue('Availability Matrix Sheet') || 'Availability Matrix',
    consolidatedSchedulesSheet: findValue('Consolidated Schedules Sheet') || 'Consolidated-FF Schedules',
    finalSchedules: findValue('Final - Schedules Sheet') || 'Final - Schedules',
    finalCapacity: findValue('Final - Capacity Sheet') || 'Final - Capacity',
    finalEstVsAct: findValue('Est vs Act - Aggregated Sheet') || 'Est vs Act - Aggregated',
    roleConfigSheet: findValue('Role Config Sheet') || 'Role Config',
    leaveProjectName: findValue('Leave Project Name') || 'JFGP All Leave',
    dataStartColumn: parseInt(findValue('Data Start Column') || '8', 10) || 8,
    regionCalendarId: calendarMatch ? calendarMatch[0] : (rawCalendarLink || ''),
    regionsInScope: regionsList
  };

  return settings;
}

function ensureRoleConfigSheet_(ss) {
  var sheet = ss.getSheetByName('Role Config');
  if (sheet) return sheet;
  sheet = ss.insertSheet('Role Config');
  var headers = ['Role', 'Billable %'];
  var rows = [
    ['(default)', '100%'],
    ['VP', '50%'],
    ['Director', '70%'],
    ['Executive', '80%'],
    ['Manager', '80%']
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.autoResizeColumns(1, headers.length);
  return sheet;
}

function setupConfigTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Config') || ss.insertSheet('Config');
  var lastRow = Math.max(sheet.getLastRow(), 1);
  var existingValues = sheet.getRange(1, 2, lastRow, 1).getDisplayValues();
  var existingKeys = {};
  existingValues.forEach(function(row){
    var key = (row[0] + '').trim();
    if (key) existingKeys[key] = true;
  });

  var requiredEntries = [
    { key: 'IMPORT-FF Schedules', sample: 'IMPORT-FF Schedules' },
    { key: 'Est vs Act - Import', sample: 'Est vs Act - Import' },
    { key: 'Actuals - Import', sample: 'Actuals - Import' },
    { key: 'Active Staff URL', sample: 'https://docs.google.com/spreadsheets/d/EXAMPLE/edit' },
    { key: 'Active Staff Sheet', sample: 'Active staff' },
    { key: 'FF Schedule Override Sheet', sample: 'FF Schedule Override' },
    { key: 'Country Hours Sheet', sample: 'Country Hours' },
    { key: 'Availability Matrix Sheet', sample: 'Availability Matrix' },
    { key: 'Consolidated Schedules Sheet', sample: 'Consolidated-FF Schedules' },
    { key: 'Final - Schedules Sheet', sample: 'Final - Schedules' },
    { key: 'Final - Capacity Sheet', sample: 'Final - Capacity' },
    { key: 'Est vs Act - Aggregated Sheet', sample: 'Est vs Act - Aggregated' },
    { key: 'Role Config Sheet', sample: 'Role Config' },
    { key: 'Leave Project Name', sample: 'JFGP All Leave' },
    { key: 'Data Start Column', sample: '8' },
    { key: 'Global Holidays', sample: 'https://docs.google.com/spreadsheets/d/EXAMPLE_HOLIDAYS/edit' }
  ];

  requiredEntries.forEach(function(entry){
    if (!existingKeys[entry.key]) {
      sheet.appendRow(['', entry.key, entry.sample]);
      existingKeys[entry.key] = true;
    }
  });
}

function setupRoleConfigTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = getGlobalConfig();
  var sheetName = config.roleConfigSheet || 'Role Config';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  var dataRange = sheet.getDataRange();
  var hasData = dataRange.getNumRows() > 1 && dataRange.getValues().slice(1).some(function(row){
    return row.some(function(cell){return cell !== '' && cell !== null;});
  });
  if (hasData) return;

  sheet.clear();
  var headers = ['Role','Billable %'];
  var defaults = [
    ['(default)','100%'],
    ['VP','50%'],
    ['Director','70%'],
    ['Executive','80%'],
    ['Manager','80%']
  ];
  sheet.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  sheet.getRange(2,1,defaults.length,headers.length).setValues(defaults);
  sheet.autoResizeColumns(1, headers.length);
}

function refreshCountryHoursFromRegion_(ss, config) {
  try {
    var regionId = config.regionCalendarId;
    if (!regionId) return false;
    var regionSource = SpreadsheetApp.openById(regionId);
    var regionSheet = regionSource.getSheetByName('Region Calendar');
    if (!regionSheet || regionSheet.getLastRow() < 2) return false;
    var configs = extractRegionConfigs_(regionSheet);
    if (!Object.keys(configs).length) return false;
    var timezone = ss.getSpreadsheetTimeZone();
    var holidaysSheet = regionSource.getSheetByName('Region Holidays');
    var holidays = extractHolidayMap_(holidaysSheet, timezone);
    var countryHoursSheet = ss.getSheetByName(config.countryHours || 'Country Hours') || ss.insertSheet(config.countryHours || 'Country Hours');
    var months = determineMonthsToBuild_(ss, countryHoursSheet, timezone, config);
    if (!months.length) return false;
    var countries = Object.keys(configs).sort();
    var outputRows = [];
    months.forEach(function(monthKey){
      countries.forEach(function(ct){
        var hours = calculateMonthlyHours_(configs[ct], monthKey, holidays[ct], timezone);
        outputRows.push([ct, monthKey, hours]);
      });
    });
    countryHoursSheet.clear();
    countryHoursSheet.getRange(1, 1, 1, 3).setValues([['Country','Month','Hours']]);
    if (outputRows.length) {
      countryHoursSheet.getRange(2, 1, outputRows.length, 3).setValues(outputRows);
    }
    countryHoursSheet.autoResizeColumns(1, 3);
    return true;
  } catch (err) {
    Logger.log('Country Hours rebuild error: ' + err);
    return false;
  }
}

function extractRegionConfigs_(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};
  var headers = data[0];
  var idxRegion = headers.indexOf('Region Code');
  var idxCountries = headers.indexOf('Country Codes (comma separated)');
  var idxStart = headers.indexOf('Workweek Start (Mon=1 ... Sun=7)');
  var idxEnd = headers.indexOf('Workweek End (Mon=1 ... Sun=7)');
  var idxHours = headers.indexOf('Hours Per Day');
  var configs = {};
  data.slice(1).forEach(function(row){
    var countriesCell = idxCountries > -1 ? row[idxCountries] : '';
    if (!countriesCell) return;
    var hoursPerDay = idxHours > -1 ? parseFloat(row[idxHours]) || 0 : 0;
    var startDow = idxStart > -1 ? parseInt(row[idxStart], 10) || 1 : 1;
    var endDow = idxEnd > -1 ? parseInt(row[idxEnd], 10) || 5 : 5;
    var regionCode = idxRegion > -1 ? (row[idxRegion] + '') : '';
    (countriesCell + '').split(',').forEach(function(token){
      var country = token.trim();
      if (!country) return;
      var canonical = normalizeCountryCode_(country);
      if (!canonical) return;
      configs[canonical] = {
        country: canonical,
        region: regionCode,
        hoursPerDay: hoursPerDay,
        startDow: normalizeDow_(startDow),
        endDow: normalizeDow_(endDow)
      };
    });
  });
  return configs;
}

function extractHolidayMap_(sheet, timezone) {
  var map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idxCountry = headers.indexOf('Country Code');
  var idxDate = headers.indexOf('Date');
  data.slice(1).forEach(function(row){
    var countryRaw = idxCountry > -1 ? (row[idxCountry] + '').trim() : '';
    var canonical = normalizeCountryCode_(countryRaw);
    if (!canonical) return;
    var dateVal = idxDate > -1 ? row[idxDate] : null;
    var dateObj = coerceToDate_(dateVal, timezone);
    if (!dateObj) return;
    var key = Utilities.formatDate(dateObj, timezone, 'yyyy-MM-dd');
    if (!map[canonical]) map[canonical] = {};
    map[canonical][key] = true;
  });
  return map;
}

function determineMonthsToBuild_(ss, countryHoursSheet, timezone, config) {
  var monthSet = {};
  function add(date) {
    if (!date || isNaN(date)) return;
    var key = Utilities.formatDate(new Date(date.getFullYear(), date.getMonth(), 1), timezone, 'yyyy-MM');
    monthSet[key] = true;
  }
  if (countryHoursSheet && countryHoursSheet.getLastRow() > 1) {
    var existing = countryHoursSheet.getRange(2, 2, countryHoursSheet.getLastRow() - 1, 1).getValues();
    existing.forEach(function(row){
      var dt = parseMonthYearValue_(row[0]);
      if (dt) add(dt);
    });
  }
  var scheduleSheet = ss.getSheetByName(config.consolidatedSchedulesSheet || 'Consolidated-FF Schedules');
  if (scheduleSheet && scheduleSheet.getLastRow() >= 2) {
    var lastCol = scheduleSheet.getLastColumn();
    var headers = scheduleSheet.getRange(2, 1, 1, lastCol).getValues()[0];
    for (var j = config.dataStartColumn ? config.dataStartColumn - 1 : 7; j < lastCol; j++) {
      var headerDate = coerceToDate_(headers[j], timezone);
      if (headerDate) add(headerDate);
    }
  }
  if (!Object.keys(monthSet).length) {
    var start = new Date(new Date().getFullYear(), 0, 1);
    for (var i = 0; i < DEFAULT_MONTH_RANGE_MONTHS; i++) {
      add(new Date(start.getFullYear(), start.getMonth() + i, 1));
    }
  }
  return Object.keys(monthSet).sort();
}

function calculateMonthlyHours_(configEntry, monthKey, holidayMap, timezone) {
  if (!configEntry) return 0;
  var monthDate = monthKeyToDate_(monthKey);
  if (!monthDate) return 0;
  var start = new Date(monthDate.getFullYear(), monthDate.getMonth(), 1);
  var end = new Date(monthDate.getFullYear(), monthDate.getMonth() + 1, 0);
  var workingDays = 0;
  var cursor = new Date(start);
  while (cursor <= end) {
    var dow = cursor.getDay() === 0 ? 7 : cursor.getDay();
    var dateKey = Utilities.formatDate(cursor, timezone, 'yyyy-MM-dd');
    if (isWorkingDow_(dow, configEntry.startDow, configEntry.endDow) && !(holidayMap && holidayMap[dateKey])) {
      workingDays++;
    }
    cursor.setDate(cursor.getDate() + 1);
  }
  return workingDays * (configEntry.hoursPerDay || 0);
}

function parseMonthYearValue_(value) {
  if (!value) return null;
  if (value instanceof Date) return new Date(value.getFullYear(), value.getMonth(), 1);
  var str = (value + '').trim();
  if (!str) return null;
  var iso = str.match(/^(\d{4})[-\/](\d{1,2})$/);
  if (iso) return new Date(parseInt(iso[1],10), parseInt(iso[2],10)-1, 1);
  var short = str.match(/^(\d{1,2})[-\/](\d{2,4})$/);
  if (short) {
    var mm = parseInt(short[1],10)-1;
    var yy = parseInt(short[2],10);
    if (yy < 100) yy += 2000;
    return new Date(yy, mm, 1);
  }
  var dt = new Date(str);
  return isNaN(dt) ? null : new Date(dt.getFullYear(), dt.getMonth(), 1);
}

function coerceToDate_(value, timezone) {
  if (!value) return null;
  if (value instanceof Date) return value;
  var str = (value + '').trim();
  if (!str) return null;
  if (/^[A-Za-z]{3}-\d{2}$/.test(str)) {
    var parts = str.split('-');
    var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    var idx = months.indexOf(parts[0]);
    if (idx > -1) {
      var year = 2000 + parseInt(parts[1], 10);
      return new Date(year, idx, 1);
    }
  }
  var dateMatch = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (dateMatch) {
    var d = parseInt(dateMatch[1],10);
    var m = parseInt(dateMatch[2],10)-1;
    var y = parseInt(dateMatch[3],10);
    if (y < 100) y += 2000;
    return new Date(y, m, d);
  }
  var parsed = new Date(str);
  return isNaN(parsed) ? null : parsed;
}

function isWorkingDow_(dow, startDow, endDow) {
  if (startDow <= endDow) return dow >= startDow && dow <= endDow;
  return dow >= startDow || dow <= endDow;
}

function normalizeDow_(value) {
  var v = parseInt(value, 10);
  if (isNaN(v) || v < 1 || v > 7) return 1;
  return v;
}

function monthKeyToDate_(key) {
  if (!key) return null;
  var parts = (key + '').split('-');
  if (parts.length !== 2) return null;
  var year = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10) - 1;
  if (isNaN(year) || isNaN(month)) return null;
  return new Date(year, month, 1);
}

function refreshAll() {
  var config = getGlobalConfig();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  importDataFromEmails(config);
  importAndFilterActiveStaff(config);
  buildEstVsActAggregate(config);
  refreshCountryHoursFromRegion_(ss, config);
  buildAvailabilityMatrix(config);
  buildFinalSchedules(config);
  buildFinalCapacity(config);
}

function runSetup() {
  setupConfigTab();
  setupRoleConfigTab();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Earned Media Resourcing')
    .addItem('Refresh All','refreshAll')
    .addItem('Run Setup','runSetup')
    .addToUi();
}
