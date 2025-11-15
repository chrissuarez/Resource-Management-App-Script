/**
 * Imports paid media EMEA schedules via email label and CSV, then transforms the data,
 * builds availability matrix, and final capacity tables.
 * Adds custom menu for manual refresh and convenience wrappers.
 */

// ----- 0. Import schedules and actuals from email -----
function importDataFromEmails() {
  try {
    var ss = SpreadsheetApp.openById(getSpreadsheetIdFromConfig_());
    var configSheet = ss.getSheetByName('Config');
    var emailConfigs = [];
    if (configSheet) {
      var startRow = 10;
      var startCol = 2; // table starts at column B
      var lastRow = configSheet.getLastRow();
      if (lastRow >= startRow) {
        var configData = configSheet.getRange(startRow, startCol, lastRow - startRow + 1, 4).getValues();
        configData.forEach(function(row) {
          var type = (row[0] + '').trim();
          if (type !== 'Email Import') return;
          var label = (row[1] + '').trim();
          var sheetName = (row[2] + '').trim();
          if (!label || !sheetName) return;
          var encoding = (row[3] + '').trim() || 'UTF-8';
          emailConfigs.push({
            label: label,
            sheetName: sheetName,
            encoding: encoding
          });
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
function transformData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName('Consolidated-FF Schedules');
  if (!sourceSheet) throw new Error('Source sheet not found');

  // Read headers and data
  var lastCol = sourceSheet.getLastColumn();
  var lastRow = sourceSheet.getLastRow();
  var headers = sourceSheet.getRange(2, 1, 1, lastCol).getValues()[0];
  var data    = sourceSheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
  var outRows = [];

  // Iterate each row & pivot month columns into rows
  data.forEach(function(row) {
    for (var j = 7; j < row.length; j++) {
      if (row[j] !== '') {
        // Base columns A–G
        var nr = row.slice(0, 7);
        // Add Date & Value
        var dateValue  = headers[j];
        var hoursValue = row[j];
        nr.push(dateValue, hoursValue);

        // Build Helper key: ResourceName-MM-yy
        var resName = nr[6] + '';
        var dt;

        // Case 1: header is a Date object
        if (dateValue instanceof Date) {
          dt = dateValue;

        // Case 2: header is text 'MMM-yy'
        } else if (typeof dateValue === 'string' && /^[A-Za-z]{3}-\d{2}$/.test(dateValue)) {
          var parts = dateValue.split('-');
          var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
          var mIndex = months.indexOf(parts[0]);
          var yFull  = 2000 + parseInt(parts[1], 10);
          dt = new Date(yFull, mIndex, 1);

        // Case 3: header is text 'dd/MM/yyyy'
        } else if (typeof dateValue === 'string') {
          var m = dateValue.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
          if (m) {
            dt = new Date(parseInt(m[3],10), parseInt(m[2],10) - 1, parseInt(m[1],10));
          } else {
            dt = new Date(dateValue);
          }

        // Fallback
        } else {
          dt = new Date(dateValue);
        }

        var helperKey = resName + '-' +
          Utilities.formatDate(dt, ss.getSpreadsheetTimeZone(), 'MM-yy');
        nr.push(helperKey);

        outRows.push(nr);
      }
    }
  });

  // Write to 'Final - Schedules' sheet
  var ts = ss.getSheetByName('Final - Schedules') || ss.insertSheet('Final - Schedules');
  if (ts.getLastRow() > 1) {
    ts.getRange(2, 1, ts.getLastRow() - 1, ts.getLastColumn()).clearContent();
  }

  // Set headers including Helper
  var newHdr = headers.slice(0, 7).concat(['Date', 'Value', 'Helper']);
  ts.getRange(1, 1, 1, newHdr.length).setValues([newHdr]);

  // Paste transformed rows
  if (outRows.length) {
    ts.getRange(2, 1, outRows.length, outRows[0].length).setValues(outRows);
  }
}

// ----- 2. Build availability matrix -----
function buildAvailabilityMatrix() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var staffSheet = ss.getSheetByName('Active staff');
  var hoursSheet = ss.getSheetByName('Country Hours');
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

  var out = ss.getSheetByName('Availability Matrix')||ss.insertSheet('Availability Matrix');
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
function buildFinalCapacity() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var avail = ss.getSheetByName('Availability Matrix');
  var sched = ss.getSheetByName('Final - Schedules');
  var staff = ss.getSheetByName('Active staff');
  if(!avail||!sched||!staff) throw new Error('Missing sheets');

  var roleSheet = ensureRoleConfigSheet_(ss);
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
    if(pj==='JFGP All Leave') leave[nm+'|'+key]=(leave[nm+'|'+key]||0)+hrs;
    else schedM[nm+'|'+key]=(schedM[nm+'|'+key]||0)+hrs;
  });

  var fo=ss.getSheetByName('Final - Capacity')||ss.insertSheet('Final - Capacity'); fo.clear();
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
function refreshAll() {
  importDataFromEmails();
  transformData();
  buildAvailabilityMatrix();
  buildFinalCapacity();
}

// ----- 5. Add custom menu -----
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Earned Media Resourcing')
    .addItem('Import & Transform','refreshAll')
    .addItem('Build Availability (after adjusting team/hours)','buildAvailabilityMatrix')
    .addItem('Build Capacity (after adhoc schedule updates)','buildFinalCapacity')
    .addToUi();
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
