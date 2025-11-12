/**
 * Imports paid media EMEA schedules via email label and CSV, then transforms the data,
 * builds availability matrix, and final capacity tables.
 * Adds custom menu for manual refresh and convenience wrappers.
 */

// ----- 0. Import schedules and actuals from email -----
function importDataFromEmails() {
  try {
    var emailConfigs = [
      {
        label: 'dashboard-reports-paid-media---emea-schedules',
        sheetName: 'IMPORT-FF Schedules',
        encoding: 'ISO-8859-1' // Original encoding
      },
      {
        label: 'dashboard-reports-paid-media-est-vs-actuals-emea',
        sheetName: 'Est vs Act - Import',
        encoding: 'ISO-8859-1' // Original encoding
      },
      {
        label: 'dashboard-reports-paid-media-timecards-emea', // Your new label
        sheetName: 'Actuals - Import',                        // Your new target sheet
        encoding: 'ISO-8859-1' // Assuming UTF-8 for this new source, adjust if needed
      }
    ];
    // IMPORTANT: Use the ID of your spreadsheet
    var spreadsheetId = '1WuMWJaB5iBYQx0pQntuaPJIZrDYFKMUaRPcF4t_lru8'; 
    var ss = SpreadsheetApp.openById(spreadsheetId);

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
  refreshCountryHoursFromRegion_(ss);
  var staffSheet = ss.getSheetByName('Active staff');
  var hoursSheet = ss.getSheetByName('Country Hours');
  if (!staffSheet || !hoursSheet) throw new Error('Missing staff or hours sheet');

  var sd = staffSheet.getDataRange().getValues(), hdr = sd[0], rows = sd.slice(1);
  function find(pats){return hdr.findIndex(h=>pats.some(p=>new RegExp(p,'i').test(h)));}
  var iName = find(['ResourceName','Resource Name']), iCountry = find(['Resource Country']), iStart = find(['Start Date']), iFTE = find(['FTE']);

  var hd = hoursSheet.getDataRange().getValues().slice(1), hmap = {};
  hd.forEach(function(r){
    var ct=r[0]+'', dt=r[1], hrs=parseFloat(r[2])||0;
    var d = dt instanceof Date ? dt : parseMonthYearValue_(dt);
    var key = Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), 'yyyy-MM');
    var c = COUNTRY_MAP[ct]||ct;
    hmap[c+'|'+key] = hrs;
  });
  var months = Object.keys(hmap).map(k=>k.split('|')[1]).filter((v,i,a)=>a.indexOf(v)===i).sort();

  var out = ss.getSheetByName('Availability Matrix')||ss.insertSheet('Availability Matrix');
  out.clearContents(); out.getRange(1,1).setValue('ResourceName');
  months.forEach(function(m,i){var d=new Date(m+'-01');out.getRange(1,i+2).setValue(d).setNumberFormat('MMM-yy');});

  function cw(s,e){var c=0;for(var d=new Date(s);d<=e;d.setDate(d.getDate()+1)){if(d.getDay()>0&&d.getDay()<6)c++;}return c;}
  var outData=[];
  rows.forEach(function(r){
    var name=r[iName]+''; if(!name)return;
    var ct=r[iCountry]+''; var st=r[iStart]?new Date(r[iStart]):null;
    var f=parseFloat(r[iFTE])||0;
    var row=[name];
    months.forEach(function(m){
      var ms=new Date(m+'-01'), me=new Date(ms.getFullYear(),ms.getMonth()+1,0);
      var key=Utilities.formatDate(ms,ss.getSpreadsheetTimeZone(),'yyyy-MM');
      var wh=hmap[(COUNTRY_MAP[ct]||ct)+'|'+key]||0;
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

  var sd = staff.getDataRange().getValues(), sh=sd[0], sr=sd.slice(1);
  function fi(p){return sh.findIndex(h=>p.some(x=>new RegExp(x,'i').test(h)));}
  var iRes = fi(['ResourceName']), iRR=fi(['ResourceRole']), iHub=fi(['Hub']), iC=fi(['Resource Country']);
  var staffMap={};
  sr.forEach(function(r){var n=r[iRes]+''; if(n){var pr=r[iRR]+''; var ps=pr.split('-');
    staffMap[n]={hub:r[iHub]+'',practice:ps[0].trim(),role:ps[1]?ps.slice(1).join('-').trim():'',country:r[iC]+''};
  }});

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
  Object.keys(fullMap).forEach(k=>{var p=k.split('|'),n=p[0],m=p[1],st=staffMap[n]||{}; var bill=/VP/i.test(st.role)?0.5:/Director/i.test(st.role)?0.7:/Executive|Manager/i.test(st.role)?0.8:1;
    var f=fullMap[k], al=leave[k]||0, net=f-al; var tbh=net*bill, nb=net*(1-bill), sch=schedM[k]||0, bc=tbh-sch;
    var mstr=Utilities.formatDate(new Date(m+'-01'),ss.getSpreadsheetTimeZone(),'MMM-yy');
    out.push([n,st.hub,st.role,st.country,bill,st.practice,mstr,f,al,nb,tbh,sch,bc]);
  });
  if(out.length){fo.getRange(2,1,out.length,out[0].length).setValues(out); fo.getRange(2,5,out.length,1).setNumberFormat('0%');}
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
    .createMenu('Paid Media Resourcing')
    .addItem('Import & Transform','refreshAll')
    .addItem('Build Availability (after adjusting team/hours)','buildAvailabilityMatrix')
    .addItem('Build Capacity (after adhoc schedule updates)','buildFinalCapacity')
    .addItem('Scaffold Region Config Sheets','setupRegionConfigSheets')
    .addToUi();
}

/** Country mapping **/
var COUNTRY_MAP = {
  'United Kingdom':'UK','Germany':'DE','Denmark':'DK','France':'FR',
  'South Africa':'SA','Spain':'ES','Netherlands':'NL','Italy':'IT'
};

var DEFAULT_MONTH_RANGE_MONTHS = 36;

/**
 * Creates template sheets for working patterns and bank holidays so
 * regional hours can be maintained without code changes.
 */
function setupRegionConfigSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var regionHeaders = [
    'Region Code',
    'Country Codes (comma separated)',
    'Workweek Start (Mon=1 ... Sun=7)',
    'Workweek End (Mon=1 ... Sun=7)',
    'Hours Per Day'
  ];
  var regionExamples = [
    ['UK', 'UK', 1, 5, 7.5],
    ['EU-Central', 'DE,FR,NL', 1, 5, 8],
    ['GCC', 'AE,SA', 7, 4, 8]
  ];
  var regionSheet = ss.getSheetByName('Region Calendar');
  if (!regionSheet) {
    regionSheet = ss.insertSheet('Region Calendar');
  }
  if (sheetHasOnlyTemplateData_(regionSheet, regionHeaders, regionExamples)) {
    writeTemplate_(regionSheet, regionHeaders, regionExamples);
  } else {
    Logger.log('Skipping Region Calendar scaffold: existing data detected.');
  }

  var holidayHeaders = ['Country Code', 'Date', 'Holiday Name'];
  var holidayExamples = [
    ['UK', new Date('2025-01-01'), 'New Year\'s Day'],
    ['DE', new Date('2025-10-03'), 'German Unity Day'],
    ['AE', new Date('2025-03-31'), 'Eid al-Fitr (placeholder)']
  ];
  var holidaysSheet = ss.getSheetByName('Region Holidays');
  if (!holidaysSheet) {
    holidaysSheet = ss.insertSheet('Region Holidays');
  }
  if (sheetHasOnlyTemplateData_(holidaysSheet, holidayHeaders, holidayExamples)) {
    writeTemplate_(holidaysSheet, holidayHeaders, holidayExamples, [[2, 'yyyy-mm-dd']]);
  } else {
    Logger.log('Skipping Region Holidays scaffold: existing data detected.');
  }
}

function ensureSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function writeTemplate_(sheet, headers, rows, dateColumns) {
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  if (rows && rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  if (dateColumns && dateColumns.length) {
    dateColumns.forEach(function(entry) {
      var col = Array.isArray(entry) ? entry[0] : entry;
      var fmt = Array.isArray(entry) && entry[1] ? entry[1] : 'yyyy-mm-dd';
      sheet.getRange(2, col, Math.max(rows ? rows.length : 1, 1), 1).setNumberFormat(fmt);
    });
  }
  sheet.autoResizeColumns(1, headers.length);
}

function sheetHasOnlyTemplateData_(sheet, headers, templateRows) {
  var colCount = headers.length;
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return true;
  var values = sheet.getRange(2, 1, lastRow - 1, colCount).getValues();
  var dataRows = values.filter(function(row){
    return row.some(function(cell){ return cell !== '' && cell !== null && !(typeof cell === 'number' && isNaN(cell)); });
  });
  if (!dataRows.length) return true;
  if (!templateRows || !templateRows.length) return false;
  if (dataRows.length !== templateRows.length) return false;
  var templateStrings = templateRows.map(function(r){return normalizeRowValues_(r, colCount);}).sort();
  var sheetStrings = dataRows.map(function(r){return normalizeRowValues_(r, colCount);}).sort();
  for (var i = 0; i < sheetStrings.length; i++) {
    if (sheetStrings[i] !== templateStrings[i]) return false;
  }
  return true;
}

function normalizeRowValues_(row, colCount) {
  var parts = [];
  for (var i = 0; i < colCount; i++) {
    var value = row[i];
    if (value instanceof Date) {
      parts.push(value.getTime());
    } else if (value === null || typeof value === 'undefined') {
      parts.push('');
    } else {
      parts.push((value + '').trim());
    }
  }
  return parts.join('|');
}

function refreshCountryHoursFromRegion_(ss) {
  try {
    var regionSheet = ss.getSheetByName('Region Calendar');
    if (!regionSheet || regionSheet.getLastRow() < 2) return false;
    var configs = extractRegionConfigs_(regionSheet);
    if (!Object.keys(configs).length) return false;
    var timezone = ss.getSpreadsheetTimeZone();
    var holidays = extractHolidayMap_(ss.getSheetByName('Region Holidays'), timezone);
    var countryHoursSheet = ensureSheet_(ss, 'Country Hours');
    var months = determineMonthsToBuild_(ss, countryHoursSheet, timezone);
    if (!months.length) return false;
    var countries = Object.keys(configs).sort();
    var rows = [];
    months.forEach(function(monthKey){
      countries.forEach(function(ct){
        var hours = calculateMonthlyHours_(configs[ct], monthKey, holidays[ct], timezone);
        rows.push([ct, monthKey, hours]);
      });
    });
    countryHoursSheet.clear();
    countryHoursSheet.getRange(1, 1, 1, 3).setValues([['Country','Month','Hours']]);
    if (rows.length) {
      countryHoursSheet.getRange(2, 1, rows.length, 3).setValues(rows);
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
      configs[country] = {
        country: country,
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
    var country = idxCountry > -1 ? (row[idxCountry] + '').trim() : '';
    if (!country) return;
    var dateVal = idxDate > -1 ? row[idxDate] : null;
    var dateObj = coerceToDate_(dateVal, timezone);
    if (!dateObj) return;
    var key = Utilities.formatDate(dateObj, timezone, 'yyyy-MM-dd');
    if (!map[country]) map[country] = {};
    map[country][key] = true;
  });
  return map;
}

function determineMonthsToBuild_(ss, countryHoursSheet, timezone) {
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
  var scheduleSheet = ss.getSheetByName('Consolidated-FF Schedules');
  if (scheduleSheet && scheduleSheet.getLastRow() >= 2) {
    var lastCol = scheduleSheet.getLastColumn();
    var headers = scheduleSheet.getRange(2, 1, 1, lastCol).getValues()[0];
    for (var j = 7; j < lastCol; j++) {
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

function calculateMonthlyHours_(config, monthKey, holidayMap, timezone) {
  if (!config) return 0;
  var monthDate = monthKeyToDate_(monthKey);
  if (!monthDate) return 0;
  var start = new Date(monthDate.getFullYear(), monthDate.getMonth(), 1);
  var end = new Date(monthDate.getFullYear(), monthDate.getMonth() + 1, 0);
  var workingDays = 0;
  var cursor = new Date(start);
  while (cursor <= end) {
    var dow = cursor.getDay() === 0 ? 7 : cursor.getDay();
    var dateKey = Utilities.formatDate(cursor, timezone, 'yyyy-MM-dd');
    if (isWorkingDow_(dow, config.startDow, config.endDow) && !(holidayMap && holidayMap[dateKey])) {
      workingDays++;
    }
    cursor.setDate(cursor.getDate() + 1);
  }
  return workingDays * (config.hoursPerDay || 0);
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
