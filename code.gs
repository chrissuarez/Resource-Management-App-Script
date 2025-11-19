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
      Logger.log('Debug: "Config" sheet found. Last row is: ' + lastRow + '. Required start row is: ' + startRow);

      if (lastRow >= startRow) {
        var configData = configSheet.getRange(startRow, 2, lastRow - startRow + 1, 4).getValues();
        configData.forEach(function(row, index) {
          Logger.log('Debug: Processing row ' + (startRow + index) + '. Value in column B is: "' + row[0] + '"');
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
      
      var message = threads[0].getMessages()[0]; 
      var attachments = message.getAttachments();
      var csvAttachment = attachments.find(att => att.getContentType() === 'text/csv' || att.getContentType() === 'application/csv');

      if (!csvAttachment) {
        Logger.log('No CSV attachment found in the latest email for ' + config.label + '. Email subject: ' + message.getSubject());
        return; // continue to next config
      }
      
      var charEncoding = config.encoding || 'UTF-8';
      var csvDataString = csvAttachment.getDataAsString(charEncoding);
      var data = Utilities.parseCsv(csvDataString);

      if (!data || data.length === 0) {
        Logger.log('CSV data is empty for ' + config.label);
        return;
      }

      var sheet = ss.getSheetByName(config.sheetName);
      if (!sheet) {
        Logger.log('Sheet not found: ' + config.sheetName + '. Creating it.');
        sheet = ss.insertSheet(config.sheetName);
      } else {
        Logger.log('Clearing existing content from sheet: ' + config.sheetName);
        clearSheetContents_(sheet);
      }
      
      sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      Logger.log('Imported data from label "' + config.label + '" to sheet "' + config.sheetName + '". Rows: ' + data.length + '. Email subject: ' + message.getSubject());
    });
  } catch(e) {
    Logger.log('Import error: ' + e.toString() + ' Stack: ' + e.stack);
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
  var requestedMetadataCols = Math.max(1, parseInt(config.dataStartColumn, 10) || 8) - 1;
  var metadataColumns = Math.min(headers.length - 1, requestedMetadataCols);
  var timezone = ss.getSpreadsheetTimeZone();
  var dataStartIdx = metadataColumns;
  for (var probe = 0; probe < headers.length; probe++) {
    var idx = probe;
    if (idx < requestedMetadataCols - 1) continue;
    var maybeDate = coerceToDate_(headers[idx], timezone);
    if (maybeDate) { dataStartIdx = idx; break; }
  }

  function normalizeHeaderValue(h) {
    return (h === null || typeof h === 'undefined') ? '' : (h + '').replace(/\u00a0/g,' ').replace(/\s+/g,' ').trim().toLowerCase();
  }
  var normalizedHeaders = headers.map(normalizeHeaderValue);
  function findIndex(pats) {
    return normalizedHeaders.findIndex(function(h){
      return pats.some(function(p){return new RegExp(p,'i').test(h);});
    });
  }
  var resourceIdx = findIndex(['Resource Name','ResourceName','Resource:\s*Resource Name','^Resource$']);
  if (resourceIdx === -1) resourceIdx = Math.max(0, metadataColumns - 1);
  var projectIdx = findIndex(['Project']);
  if (projectIdx === -1) throw new Error('Project column not found in ' + (config.importSchedules || 'IMPORT-FF Schedules'));
  var accountIdx = findIndex(['Account','Client']);
  var hoursIdx = findIndex(['Estimated Hours','Est Hours','Hours','Value','Estimated-Hours']);
  var dateIdx = findIndex(['End Date','Date','End-Date']);
  if (hoursIdx === -1 && headers.length === 4) hoursIdx = 2;
  if (dateIdx === -1 && headers.length === 4) dateIdx = 3;
  var isLongForm = hoursIdx > -1 && dateIdx > -1 && rows.length;
  var metaStop = Math.min(metadataColumns, headers.length);
  if (isLongForm) {
    var avoidHours = hoursIdx > -1 ? hoursIdx : metaStop;
    var avoidDate = dateIdx > -1 ? dateIdx : metaStop;
    metaStop = Math.min(metaStop, avoidHours, avoidDate);
  }

  var staffSheet = ss.getSheetByName(config.staffSheet || 'Active staff');
  var staffMap = {};
  if (staffSheet && staffSheet.getLastRow() > 1) {
    var staffData = staffSheet.getDataRange().getValues();
    var sh = staffData[0], sr = staffData.slice(1);
    function sFind(pats){return sh.findIndex(function(h){return pats.some(function(p){return new RegExp(p,'i').test(h);});});}
    var sName = sFind(['ResourceName','Resource Name']);
    var sRole = sFind(['ResourceRole','Resource Role']);
    var sCountry = sFind(['Resource Country','Resource Location','Location']);
    var sHub = sFind(['Hub','Resource Hub']);
    sr.forEach(function(r){
      var name = sName > -1 ? (r[sName] + '').trim() : '';
      if (!name) return;
      var roleRaw = sRole > -1 ? (r[sRole] + '').trim() : '';
      var parts = roleRaw.split('-');
      var practice = parts[0] ? parts[0].trim() : '';
      var roleRest = parts.length > 1 ? parts.slice(1).join('-').trim() : '';
      staffMap[name] = {
        role: roleRaw || roleRest,
        practice: practice,
        location: sCountry > -1 ? (r[sCountry] + '').trim() : '',
        hub: sHub > -1 ? (r[sHub] + '').trim() : ''
      };
    });
  }

  var lookupSheet = ss.getSheetByName('Lookups');
  var accountLookup = {};
  var hubLookup = {};
  if (lookupSheet && lookupSheet.getLastRow() > 1) {
    var lookupData = lookupSheet.getDataRange().getValues();
    lookupData.slice(1).forEach(function(row){
      var projectVal = (row[0] + '').trim();
      var accountVal = (row[1] + '').trim();
      var resVal = (row[6] + '').trim();
      var hubVal = (row[7] + '').trim();
      if (projectVal && accountVal) accountLookup[projectVal] = accountVal;
      if (resVal && hubVal) hubLookup[resVal] = hubVal;
    });
  }

  var overrideSheet = ss.getSheetByName(config.overrideSchedules || 'FF Schedule Override');
  var overrideMap = {};
  var overrideAccountMap = {};
  if (overrideSheet && overrideSheet.getLastRow() > 1) {
    var overrideData = overrideSheet.getDataRange().getValues();
    var oHeaders = overrideData[0];
    var oRows = overrideData.slice(1);
    var oResIdx = oHeaders.findIndex(function(h){return /Resource/i.test(h);});
    var oProjIdx = oHeaders.findIndex(function(h){return /Project/i.test(h);});
    var oAccIdx = oHeaders.findIndex(function(h){return /Account/i.test(h);});
    var startMonthCol = 3;
    for (var iRow = 0; i < oRows.length; iRow++) {
      var row = oRows[iRow];
      var oRes = oResIdx > -1 ? (row[oResIdx] + '').trim() : '';
      var oProj = oProjIdx > -1 ? (row[oProjIdx] + '').trim() : '';
      if (!oRes || !oProj) continue;
      var oAcc = oAccIdx > -1 ? (row[oAccIdx] + '').trim() : '';
      for (var col = startMonthCol; col < row.length; col++) {
        var headerVal = oHeaders[col];
        var dtMonth = parseMonthYearValue_(headerVal);
        if (!dtMonth) continue;
        var monthKey = Utilities.formatDate(dtMonth, timezone, 'yyyy-MM');
        var key = [oRes, oProj, monthKey].join('|');
        var hoursVal = parseFloat(row[col]);
        if (isNaN(hoursVal) || hoursVal === 0) continue;
        overrideMap[key] = (overrideMap[key] || 0) + hoursVal;
        if (oAcc) overrideAccountMap[key] = oAcc;
      }
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
  var officialKeys = {};
  if (isLongForm) {
    rows.forEach(function(row){
      var hours = row[hoursIdx];
      if (hours === '' || hours === null || typeof hours === 'undefined') return;
      var dt = coerceToDate_(row[dateIdx], timezone);
      if (!dt || isNaN(dt)) return;
      var resourceName = resourceIdx > -1 ? (row[resourceIdx] + '').trim() : '';
      var projectName = projectIdx > -1 ? (row[projectIdx] + '').trim() : '';
      if (!projectName) return;
      if (Object.keys(staffMap).length && !staffMap[resourceName]) return;
      var monthKey = Utilities.formatDate(new Date(dt.getFullYear(), dt.getMonth(), 1), timezone, 'yyyy-MM');
      var overrideKey = [resourceName, projectName, monthKey].join('|');
      var finalHours = hours + (overrideMap[overrideKey] || 0);
      var helperKey = resourceName + '-' + Utilities.formatDate(dt, timezone, 'MM-yy');
      var staff = staffMap[resourceName] || {};
      var accountVal = accountLookup[projectName] || overrideAccountMap[overrideKey] || (accountIdx > -1 ? row[accountIdx] : '');
      var hubVal = hubLookup[resourceName] || staff.hub || '';
      var newRow = [];
      if (Object.keys(staffMap).length) {
        newRow.push(staff.role || '', staff.practice || '', staff.location || '', hubVal);
      }
      newRow.push(accountVal, projectName, resourceName, dt, finalHours, helperKey);
      outRows.push(newRow);
      officialKeys[overrideKey] = true;
    });
  } else {
    rows.forEach(function(row){
      for (var j = dataStartIdx; j < row.length; j++) {
        var hours = row[j];
        if (hours === '' || hours === null || typeof hours === 'undefined') continue;
        var dateValue = headers[j];
        var dt = parseHeaderDate(dateValue);
        if (isNaN(dt)) continue;
        var resourceName = resourceIdx > -1 ? (row[resourceIdx] + '').trim() : '';
        var projectName = projectIdx > -1 ? (row[projectIdx] + '').trim() : '';
        if (!projectName) continue;
        if (Object.keys(staffMap).length && !staffMap[resourceName]) return;
        var monthKey = Utilities.formatDate(new Date(dt.getFullYear(), dt.getMonth(), 1), timezone, 'yyyy-MM');
        var overrideKey = [resourceName, projectName, monthKey].join('|');
        var finalHours = hours + (overrideMap[overrideKey] || 0);

        var helperKey = resourceName + '-' + Utilities.formatDate(dt, timezone, 'MM-yy');
        var staff = staffMap[resourceName] || {};
        var accountVal = accountLookup[projectName] || overrideAccountMap[overrideKey] || (accountIdx > -1 ? row[accountIdx] : '');
        var hubVal = hubLookup[resourceName] || staff.hub || '';
        var newRow = [];
        if (Object.keys(staffMap).length) {
          newRow.push(staff.role || '', staff.practice || '', staff.location || '', hubVal);
        }
        newRow.push(accountVal, projectName, resourceName, dt, finalHours, helperKey);
        outRows.push(newRow);
        officialKeys[overrideKey] = true;
      }
    });
  }

  Object.keys(overrideMap).forEach(function(key){
    if (officialKeys[key]) return;
    var parts = key.split('|');
    var resourceName = parts[0], projectName = parts[1], monthKey = parts[2];
    if (Object.keys(staffMap).length && !staffMap[resourceName]) return;
    var dt = monthKeyToDate_(monthKey);
    if (!dt || isNaN(dt)) return;
    var staff = staffMap[resourceName] || {};
    var accountVal = accountLookup[projectName] || overrideAccountMap[key] || '';
    var hubVal = hubLookup[resourceName] || staff.hub || '';
    var helperKey = resourceName + '-' + Utilities.formatDate(dt, timezone, 'MM-yy');
    var newRow = [];
    if (Object.keys(staffMap).length) {
      newRow.push(staff.role || '', staff.practice || '', staff.location || '', hubVal);
    }
    newRow.push(accountVal, projectName, resourceName, dt, overrideMap[key], helperKey);
    outRows.push(newRow);
  });

  var finalSheet = ss.getSheetByName(config.finalSchedules || 'Final - Schedules') || ss.insertSheet(config.finalSchedules || 'Final - Schedules');
  clearSheetContents_(finalSheet);
  var staffHeaders = Object.keys(staffMap).length ? ['Resource Role','Practice','Resource Location','Resource Hub'] : [];
  var accountHeader = accountIdx > -1 ? headers[accountIdx] : 'Account';
  var headerRow = staffHeaders.concat([accountHeader, headers[projectIdx] || 'Project', headers[resourceIdx] || 'Resource Name','Date','Value','Helper']);
  finalSheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
  if (outRows.length) {
    finalSheet.getRange(2, 1, outRows.length, outRows[0].length).setValues(outRows);
  } else {
    Logger.log('buildFinalSchedules produced 0 rows. isLongForm=' + isLongForm + ', hoursIdx=' + hoursIdx + ', dateIdx=' + dateIdx + ', dataStartIdx=' + dataStartIdx + ', metadataColumns=' + metadataColumns + ', headers=' + headers.join(' | '));
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
  clearSheetContents_(out); out.getRange(1,1).setValue('ResourceName');
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
  var sched = ss.getSheetByName(config.finalSchedules || 'Final - Schedules');
  var staff = ss.getSheetByName(config.staffSheet || 'Active staff');
  var chSheet = ss.getSheetByName(config.countryHours || 'Country Hours');
  if(!sched||!staff||!chSheet) throw new Error('Missing capacity prerequisite sheets (Final - Schedules, Active staff, or Country Hours)');

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

  var lookupSheet = ss.getSheetByName('Lookups');
  var hubLookup = {};
  if (lookupSheet && lookupSheet.getLastRow() > 1) {
    var lookupData = lookupSheet.getDataRange().getValues();
    var lHeaders = lookupData[0];
    var lRows = lookupData.slice(1);
    var lResIdx = lHeaders.findIndex(function(h){return /Resource/i.test(h);});
    var lHubIdx = lHeaders.findIndex(function(h){return /Hub/i.test(h);});
    lRows.forEach(function(row){
      var res = lResIdx > -1 ? (row[lResIdx] + '').trim() : '';
      var hub = lHubIdx > -1 ? (row[lHubIdx] + '').trim() : '';
      if (res && hub) hubLookup[res] = hub;
    });
  }

  var sd = staff.getDataRange().getValues(), sh=sd[0], sr=sd.slice(1);
  function fi(p){return sh.findIndex(h=>p.some(x=>new RegExp(x,'i').test(h)));}
  var iRes = fi(['ResourceName']), iRR=fi(['ResourceRole']), iHub=fi(['Hub']), iC=fi(['Resource Country']), iStart=fi(['Start Date']), iFte=fi(['FTE']);
  var staffMap={};
  sr.forEach(function(r){
    var n=r[iRes]+''; if(!n) return;
    var pr=r[iRR]+''; var ps=pr.split('-');
    var countryOriginal=r[iC]+'';
    var countryCode=normalizeCountryCode_(countryOriginal);
    var startDate = iStart > -1 && r[iStart] ? new Date(r[iStart]) : null;
    var fte = iFte > -1 ? parseFloat(r[iFte]) || 0 : 0;
    var hubValue = iHub > -1 ? (r[iHub] + '').trim() : '';
    if (!hubValue && hubLookup[n]) hubValue = hubLookup[n];
    staffMap[n]={
      hub:hubValue,
      practice:ps[0].trim(),
      role:ps[1]?ps.slice(1).join('-').trim():'',
      countryOriginal:countryOriginal,
      countryCode:countryCode,
      country:formatCountryDisplay_(countryOriginal,countryCode),
      startDate:startDate,
      fte:fte
    };
  });

  var countryHoursMap = {};
  var monthSet = {};
  if (chSheet && chSheet.getLastRow() > 1) {
    var chData = chSheet.getDataRange().getValues().slice(1);
    chData.forEach(function(r){
      var cCode = normalizeCountryCode_(r[0]);
      var mVal = r[1];
      var hrs = parseFloat(r[2]) || 0;
      var mDate = mVal instanceof Date ? mVal : coerceToDate_(mVal, ss.getSpreadsheetTimeZone());
      if (!cCode || !mDate) return;
      var mKey = Utilities.formatDate(new Date(mDate.getFullYear(), mDate.getMonth(), 1), ss.getSpreadsheetTimeZone(), 'yyyy-MM');
      countryHoursMap[cCode+'|'+mKey] = hrs;
      monthSet[mKey] = true;
    });
  }

  function countWorkingDaysInclusive_(startDate, endDate) {
    var count = 0;
    var cursor = new Date(startDate);
    while (cursor <= endDate) {
      var dow = cursor.getDay();
      if (dow !== 0 && dow !== 6) count++;
      cursor.setDate(cursor.getDate() + 1);
    }
    return count;
  }

  function availableHoursForMonth_(staffEntry, monthKey) {
    var base = countryHoursMap[(staffEntry.countryCode || '') + '|' + monthKey];
    if (!base) return 0;
    var fte = staffEntry.fte || 0;
    if (!fte) return 0;
    var monthStart = new Date(monthKey + '-01');
    var monthEnd = new Date(monthStart.getFullYear(), monthStart.getMonth() + 1, 0);
    if (!staffEntry.startDate) {
      return base * fte;
    }
    var start = staffEntry.startDate;
    if (start > monthEnd) return 0;
    var effectiveStart = start > monthStart ? start : monthStart;
    var workingTotal = countWorkingDaysInclusive_(monthStart, monthEnd);
    if (!workingTotal) return 0;
    var workingRemaining = countWorkingDaysInclusive_(effectiveStart, monthEnd);
    var factor = workingRemaining / workingTotal;
    return base * fte * factor;
  }

  var fullMap = {};
  Object.keys(staffMap).forEach(function(name){
    var staffEntry = staffMap[name];
    Object.keys(monthSet).forEach(function(monthKey){
      var hours = availableHoursForMonth_(staffEntry, monthKey);
      if (hours) {
        fullMap[name + '|' + monthKey] = hours;
      }
    });
  });

  var sd2=sched.getDataRange().getValues(), sh2=sd2[0], sr2=sd2.slice(1);
  var iProj=sh2.findIndex(h=>/Project/i.test(h)), iVal=sh2.findIndex(h=>/Value|Hours/i.test(h)), iHelp=sh2.findIndex(h=>/Helper/i.test(h));
  var leave={}, schedM={};
  sr2.forEach(r=>{var pj=r[iProj]+'', h=r[iHelp]+''; var m=h.match(/^(.+)-(\d{2})-(\d{2})$/); if(!m)return;
    var nm=m[1], mo=m[2], yr=m[3]; var key=(yr.length===2?('20'+yr):yr)+'-'+mo; var hrs=parseFloat(r[iVal])||0;
    if(pj===(config.leaveProjectName||'JFGP All Leave')) leave[nm+'|'+key]=(leave[nm+'|'+key]||0)+hrs;
    else schedM[nm+'|'+key]=(schedM[nm+'|'+key]||0)+hrs;
  });

  function ensureFullMapKey(resourceName, monthKey) {
    if (!staffMap[resourceName]) return;
    var composite = resourceName + '|' + monthKey;
    if (fullMap.hasOwnProperty(composite)) return;
    var hours = availableHoursForMonth_(staffMap[resourceName], monthKey);
    fullMap[composite] = hours || 0;
  }
  Object.keys(leave).forEach(function(key){
    var parts = key.split('|');
    ensureFullMapKey(parts[0], parts[1]);
  });
  Object.keys(schedM).forEach(function(key){
    var parts = key.split('|');
    ensureFullMapKey(parts[0], parts[1]);
  });

  var fo=ss.getSheetByName(config.finalCapacity || 'Final - Capacity')||ss.insertSheet(config.finalCapacity || 'Final - Capacity');
  clearSheetContents_(fo);
  var hdr=['Resource Name','Hub','Role','Country','Bill %','Practice','Month-Year','Full Hours','Annual Leave','NB Hours','TBH','SchedHrs','Billable Capacity'];
  fo.getRange(1,1,1,hdr.length).setValues([hdr]); var out=[];
  Object.keys(fullMap).forEach(function(k){
    var p=k.split('|'),n=p[0],m=p[1],st=staffMap[n]||{};
    var resourceRole=(st.role||'').toLowerCase();
    var bill=billableMap.hasOwnProperty(resourceRole)?billableMap[resourceRole]:defaultBillable;
    var monthDate = new Date(m+'-01');
    var chKey = (st.countryCode||'')+'|'+m;
    var fullHours = countryHoursMap[chKey] || fullMap[k] || 0;
    var al=leave[k]||0, net=fullHours-al; var tbh=net*bill, nb=net*(1-bill), sch=schedM[k]||0, bc=tbh-sch;
    out.push([n,st.hub,st.role,st.country,bill,st.practice,monthDate,fullHours,al,nb,tbh,sch,bc]);
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
  var source;
  try {
    source = openSpreadsheetByUrlOrId_(sourceUrl);
  } catch (err) {
    throw new Error('Failed to load Active Staff source from Config ("Active Staff URL"). ' + err.message);
  }
  var sourceSheetName = (config.activeStaffSourceSheet || '').trim();
  var sheet = sourceSheetName ? source.getSheetByName(sourceSheetName) : null;
  if (!sheet) {
    var sheets = source.getSheets();
    if (sourceSheetName) {
      throw new Error('Active Staff source sheet "' + sourceSheetName + '" was not found in the provided spreadsheet.');
    }
    if (!sheets.length) throw new Error('Active Staff source spreadsheet has no sheets.');
    sheet = sheets[0];
  }
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error('Active staff source has no data');
  var headers = data[0];
  var rows = data.slice(1);
  var destSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.staffSheet || 'Active staff');
  if (!destSheet) throw new Error('Destination staff sheet not found');
  var practiceIdx = headers.findIndex(function(h){return /^Parent Practice$/i.test(h);});
  var countryIdx = headers.findIndex(function(h){return /Resource Country/i.test(h);});
  var titleIdx = headers.findIndex(function(h){return /Resource Title/i.test(h);});
  if (practiceIdx === -1 || countryIdx === -1 || titleIdx === -1) {
    throw new Error('Required columns missing in Active Staff source (need Parent Practice, Resource Country, Resource Title).');
  }

  var filterCfg = getStaffFilterConfig_();
  var practiceFilter = (filterCfg.practice || '').toLowerCase();
  var regionPattern = filterCfg.regionPattern ? new RegExp(filterCfg.regionPattern, 'i') : null;

  var filtered = [headers];
  rows.forEach(function(row){
    var practice = (row[practiceIdx] + '').toLowerCase();
    var country = (row[countryIdx] + '').trim();
    var title = (row[titleIdx] + '').toLowerCase();

    if (practiceFilter && practice !== practiceFilter) return;
    if (regionPattern && !regionPattern.test(country)) return;
    if (/contractor/i.test(title)) return;

    filtered.push(row);
  });

  clearSheetContents_(destSheet);
  destSheet.getRange(1, 1, filtered.length, filtered[0].length).setValues(filtered);
}

/**
 * Aggregates Est vs Act data by Resource, Project, Month.
 */
function buildEstVsActAggregate(config) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var importSheet = ss.getSheetByName(config.importEstVsAct || 'Est vs Act - Import');
  if (!importSheet) throw new Error('Est vs Act import sheet not found');
  var destName = config.finalEstVsAct || 'Est vs Act - Aggregated';
  var destSheet = ss.getSheetByName(destName) || ss.insertSheet(destName);
  var data = importSheet.getDataRange().getValues();
  var headers = data.length ? data[0] : [];
  var rows = data.length > 1 ? data.slice(1) : [];
  var timezone = ss.getSpreadsheetTimeZone();

  function findIndex(patterns) {
    if (!headers || !headers.length) return -1;
    return headers.findIndex(function(h) {
      return patterns.some(function(p) {
        return new RegExp(p, 'i').test(h);
      });
    });
  }

  var idxResource = findIndex(['Resource Name', 'ResourceName']);
  var idxProject = findIndex(['Project']);
  var idxDate = findIndex(['Month', 'Month-Year', 'Month Year', 'Date', 'Week']);
  var idxEst = findIndex(['Estimated Hours', 'Estimate', 'Est Hours', 'Est. Hours']);
  var idxActual = findIndex(['Actual Hours', 'Actuals', 'Act Hours', 'Actual']);
  var idxSingleHours = -1, idxType = -1;
  if (idxEst === -1 && idxActual === -1) {
    idxSingleHours = findIndex(['Hours', 'Value']);
    idxType = findIndex(['Type', 'Category', 'Hours Type']);
  }

  if (idxResource === -1 || idxProject === -1 || idxDate === -1 || (idxEst === -1 && idxActual === -1 && idxSingleHours === -1)) {
    throw new Error('Est vs Act import sheet is missing required columns');
  }

  var aggregate = {};
  function addHours(key, estVal, actVal) {
    if (estVal === null && actVal === null) return;
    if (!aggregate[key]) aggregate[key] = { est: 0, act: 0 };
    if (estVal !== null) aggregate[key].est += estVal;
    if (actVal !== null) aggregate[key].act += actVal;
  }

  function parseHours(value) {
    var num = parseFloat(value);
    return isNaN(num) ? null : num;
  }

  rows.forEach(function(row) {
    var resource = idxResource > -1 ? (row[idxResource] + '').trim() : '';
    var project = idxProject > -1 ? (row[idxProject] + '').trim() : '';
    var dateValue = idxDate > -1 ? coerceToDate_(row[idxDate], timezone) : null;
    if (!resource || !project || !dateValue) return;
    var monthKey = Utilities.formatDate(new Date(dateValue.getFullYear(), dateValue.getMonth(), 1), timezone, 'yyyy-MM');
    var key = [resource, project, monthKey].join('|');
    var estVal = idxEst > -1 ? parseHours(row[idxEst]) : null;
    var actVal = idxActual > -1 ? parseHours(row[idxActual]) : null;
    if (idxSingleHours > -1) {
      var singleVal = parseHours(row[idxSingleHours]);
      if (singleVal !== null) {
        var typeVal = idxType > -1 ? (row[idxType] + '').toLowerCase() : '';
        if (/est/.test(typeVal)) {
          estVal = (estVal || 0) + singleVal;
        } else if (/act/.test(typeVal)) {
          actVal = (actVal || 0) + singleVal;
        } else {
          actVal = (actVal || 0) + singleVal;
        }
      }
    }
    addHours(key, estVal, actVal);
  });

  var output = [];
  Object.keys(aggregate).sort(function(a, b) {
    var pa = a.split('|'), pb = b.split('|');
    if (pa[0] !== pb[0]) return pa[0].localeCompare(pb[0]);
    if (pa[1] !== pb[1]) return pa[1].localeCompare(pb[1]);
    return pa[2].localeCompare(pb[2]);
  }).forEach(function(key) {
    var parts = key.split('|');
    var monthDate = monthKeyToDate_(parts[2]) || new Date(parts[2] + '-01');
    var est = aggregate[key].est || 0;
    var act = aggregate[key].act || 0;
    output.push([parts[0], parts[1], monthDate, est, act, act - est]);
  });

  clearSheetContents_(destSheet);
  var header = ['Resource Name', 'Project', 'Month', 'Estimated Hours', 'Actual Hours', 'Variance'];
  destSheet.getRange(1, 1, 1, header.length).setValues([header]);
  if (output.length) {
    destSheet.getRange(2, 1, output.length, output[0].length).setValues(output);
    destSheet.getRange(2, 3, output.length, 1).setNumberFormat('mmm-yy');
  }
}

/**
 * Rebuilds the "All Rows Needed Data Source" tab (or configured Variance Source Sheet) 
 * without relying on Sheet QUERY, then reapplies the downstream array formulas in E:L. 
 * This keeps the existing header if present; otherwise a default 12-column header is used.
 */
function rebuildVarianceSourceSheet_(config) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var destName = config.varianceSourceSheet || 'All Rows Needed Data Source';
  var destSheet = ss.getSheetByName(destName) || ss.insertSheet(destName);

  var defaultHeader = [
    'Resource',
    'Project',
    'Start Date',
    'Date',
    'Account',
    'Date Adj',
    'Capability Partner',
    'Parent Practice',
    'Practice',
    'ResourceRole',
    'Relative Month',
    'Year - Month'
  ];

  var header = defaultHeader;
  if (destSheet.getLastRow() >= 1 && destSheet.getLastColumn() >= 1) {
    var existingHeader = destSheet.getRange(1, 1, 1, Math.max(destSheet.getLastColumn(), defaultHeader.length)).getValues()[0];
    if (existingHeader.some(function(c) { return c !== null && c !== ''; })) {
      header = existingHeader;
    }
  }

  var estSheet = ss.getSheetByName(config.importEstVsAct || 'Est vs Act - Import');
  var actSheet = ss.getSheetByName(config.importActuals || 'Actuals - Import');
  var lookupsSheet = ss.getSheetByName('Lookups');
  var activeStaffSheet = ss.getSheetByName(config.staffSheet || 'Active staff');
  var timezone = ss.getSpreadsheetTimeZone();

  var accountMap = {};
  if (lookupsSheet && lookupsSheet.getLastRow() > 1) {
    var lData = lookupsSheet.getDataRange().getValues();
    lData.slice(1).forEach(function(r) {
      var proj = (r[0] + '').trim();
      var acc = (r[1] + '').trim();
      if (proj && acc) accountMap[proj] = acc;
    });
  }

  var resourceMap = {};
  if (activeStaffSheet && activeStaffSheet.getLastRow() > 1) {
    var asData = activeStaffSheet.getDataRange().getValues();
    var IDX_NAME = 7;
    var IDX_CAP = 8;
    var IDX_PARENT = 0;
    var IDX_PRACTICE = 1;
    var IDX_ROLE = 3;
    asData.slice(1).forEach(function(r){
      var resName = r.length > IDX_NAME ? (r[IDX_NAME] + '').trim() : '';
      if (!resName) return;
      var cap = r.length > IDX_CAP ? (r[IDX_CAP] + '').trim() : '';
      var parent = r.length > IDX_PARENT ? (r[IDX_PARENT] + '').trim() : '';
      var practice = r.length > IDX_PRACTICE ? (r[IDX_PRACTICE] + '').trim() : '';
      var roleRaw = r.length > IDX_ROLE ? (r[IDX_ROLE] + '').trim() : '';
      var role = roleRaw;
      if (practice && roleRaw && roleRaw.toLowerCase().indexOf((practice + ' -').toLowerCase()) === 0) {
        role = roleRaw.substring((practice + ' - ').length);
      }
      resourceMap[resName] = {
        cap: cap || 'Not in lookup',
        parent: parent || 'Not in lookup',
        practice: practice || 'Not in lookup',
        role: role || 'Not in lookup'
      };
    });
  }

  var combined = [];

  if (estSheet && estSheet.getLastRow() > 1) {
    var estValues = estSheet.getDataRange().getValues().slice(1);
    estValues.forEach(function(r) {
      var project = (r[0] + '').trim();
      var resource = (r[1] + '').trim();
      if (!project || !resource) return;
      combined.push([resource, project, '', r[4]]);
    });
  }

  if (actSheet && actSheet.getLastRow() > 1) {
    var actValues = actSheet.getDataRange().getValues().slice(1);
    actValues.forEach(function(r) {
      var project = (r[1] + '').trim();
      if (!project) return;
      combined.push([r[0], r[1], r[2], r[3]]);
    });
  }

  combined.sort(function(a, b) {
    return (a[1] || '').localeCompare(b[1] || '');
  });

  clearSheetContents_(destSheet);
  destSheet.getRange(1, 1, 1, header.length).setValues([header]);

  if (combined.length) {
    var configSheet = ss.getSheetByName('Config');
    var anchorCell = configSheet ? configSheet.getRange('C7').getValue() : null;
    var anchorDate = coerceToDate_(anchorCell, timezone);
    var output = combined.map(function(row){
      var res = row[0], proj = row[1], startDate = row[2], dateVal = row[3];
      var acct = accountMap[proj] || '#N/A';
      var dateAdj = dateVal ? coerceToDate_(dateVal, timezone) : null;
      var resInfo = resourceMap[res] || { cap:'Not in lookup', parent:'Not in lookup', practice:'Not in lookup', role:'Not in lookup' };
      var relMonth = (res && dateAdj && anchorDate) ? ((dateAdj.getFullYear() - anchorDate.getFullYear()) * 12 + (dateAdj.getMonth() - anchorDate.getMonth())) : '';
      var yearMonth = dateAdj ? Utilities.formatDate(new Date(dateAdj.getFullYear(), dateAdj.getMonth() + 1, 0), timezone, 'yyyy - MM') : '';
      return [
        res,
        proj,
        startDate || '',
        dateAdj || '',
        acct,
        dateAdj || '',
        resInfo.cap,
        resInfo.parent,
        resInfo.practice,
        resInfo.role,
        relMonth,
        yearMonth
      ];
    });
    destSheet.getRange(2, 1, output.length, header.length).setValues(output);
  }

  destSheet.autoResizeColumns(1, Math.max(header.length, 12));
}

/**
 * Rebuilds the Variance tab (replacing previous sheet queries) by grouping
 * the "All Rows Needed Data Source" sheet with the same logic as:
 * SELECT L, E, B, A, G, H, I, J, sum(C) WHERE F IS NOT NULL AND K <= 0 AND K > -4 GROUP BY L, E, B, A, G, H, I, J
 */
function buildVarianceTab(config) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceName = config.varianceSourceSheet || 'All Rows Needed Data Source';
  var destName = config.varianceSheet || 'Variance';
  var source = ss.getSheetByName(sourceName);
  if (!source) throw new Error('Variance source sheet not found: ' + sourceName);
  var data = source.getDataRange().getValues();
  if (!data.length) return;
  var headers = data[0];
  if (headers.length < 12) {
    throw new Error('Variance source sheet "' + sourceName + '" is missing expected columns (needs at least 12).');
  }

  var agg = {};
  data.slice(1).forEach(function(row){
    if (!row.length) return;
    var fVal = row[5];
    if (fVal === null || typeof fVal === 'undefined' || (fVal + '').trim() === '') return;
    var kVal = parseFloat(row[10]);
    if (isNaN(kVal) || kVal > 0 || kVal <= -4) return;
    var cVal = parseFloat(row[2]);
    if (isNaN(cVal)) cVal = 0;
    var keyParts = [row[11], row[4], row[1], row[0], row[6], row[7], row[8], row[9]];
    var key = keyParts.join('|');
    if (!agg[key]) {
      agg[key] = { cols: keyParts, sum: 0 };
    }
    agg[key].sum += cVal;
  });

  var estActSheet = ss.getSheetByName(config.finalEstVsAct || 'Est vs Act - Aggregated');
  var estActMap = {};
  if (estActSheet && estActSheet.getLastRow() > 1) {
    var eaData = estActSheet.getDataRange().getValues().slice(1);
    eaData.forEach(function(r){
      var res = (r[0] + '').trim();
      var proj = (r[1] + '').trim();
      var mVal = r[2];
      if (!res || !proj || !mVal) return;
      var monthDate = mVal instanceof Date ? mVal : coerceToDate_(mVal, ss.getSpreadsheetTimeZone());
      if (!monthDate || isNaN(monthDate)) return;
      var monthKeyPretty = Utilities.formatDate(monthDate, ss.getSpreadsheetTimeZone(), 'yyyy - MM');
      var key = monthKeyPretty + proj + res;
      estActMap[key] = { est: parseFloat(r[3]) || 0, act: parseFloat(r[4]) || 0 };
    });
  }

  var staffInfo = {};
  var staffSheet = ss.getSheetByName(config.staffSheet || 'Active staff');
  if (staffSheet && staffSheet.getLastRow() > 1) {
    var sData = staffSheet.getDataRange().getValues();
    var sHeaders = sData[0];
    function sFind(pats){return sHeaders.findIndex(function(h){return pats.some(function(p){return new RegExp(p,'i').test(h);});});}
    var iName = sFind(['ResourceName','Resource Name']);
    var iRegion = sFind(['Region']);
    var iCountry = sFind(['Resource Country','Country']);
    var iHub = sFind(['Hub']);
    sData.slice(1).forEach(function(r){
      var name = iName > -1 ? (r[iName] + '').trim() : '';
      if (!name) return;
      staffInfo[name] = {
        region: iRegion > -1 ? (r[iRegion] + '').trim() : '',
        country: iCountry > -1 ? (r[iCountry] + '').trim() : '',
        hub: iHub > -1 ? (r[iHub] + '').trim() : ''
      };
    });
  }

  var hubLookup = {};
  var lookupsAll = [];
  var lookupSheet = ss.getSheetByName('Lookups');
  if (lookupSheet && lookupSheet.getLastRow() > 1) {
    var lData = lookupSheet.getDataRange().getValues();
    var lHead = lData[0];
    var lResIdx = lHead.findIndex(function(h){return /Resource/i.test(h);});
    var lHubIdx = lHead.findIndex(function(h){return /Hub/i.test(h);});
    lookupsAll = lData.slice(1);
    lookupsAll.forEach(function(r){
      var res = lResIdx > -1 ? (r[lResIdx] + '').trim() : '';
      var hub = lHubIdx > -1 ? (r[lHubIdx] + '').trim() : '';
      if (res && hub) hubLookup[res] = hub;
    });
  }

  var outHeader = [
    headers[11] || 'Year - Month',
    headers[4] || 'Account',
    headers[1] || 'Project',
    headers[0] || 'Resource',
    headers[6] || 'Capability Partner',
    headers[7] || 'Parent Practice',
    headers[8] || 'Practice',
    headers[9] || 'ResourceRole',
    'Act TC',
    'Billable',
    'Sched.',
    'Est Act Hrs',
    'Act.',
    'Var',
    'Region',
    'Country',
    'Resource Hub',
    'Relative Year',
    'Relative Month',
    'Project - Adj'
  ];

  function deriveBillable_(projectName) {
    var p = (projectName || '').toString();
    if (!p) return '';
    var lower = p.toLowerCase();
    if (/opp|client admin/.test(lower)) return 'Growth';
    if (/JFGP|JFTR/i.test(p)) return 'Internal';
    return 'Billable';
  }

  function adjustAccount_(projectName, accountVal) {
    if (/^JFGP/i.test(projectName || '')) return 'Jellyfish Internal';
    if (/^JFTR/i.test(projectName || '')) return 'Jellyfish Training';
    return accountVal;
  }

  var today = new Date();
  var todayYear = today.getFullYear();
  var todayMonth = today.getMonth() + 1;

  function parseMonthString_(str) {
    if (!str || typeof str !== 'string') return null;
    var m = str.match(/^(\d{4})\s*-\s*(\d{2})$/);
    if (!m) return null;
    var year = parseInt(m[1], 10);
    var month = parseInt(m[2], 10);
    if (isNaN(year) || isNaN(month)) return null;
    return { year: year, month: month };
  }

  function adjustedProject_(projectName, resourceName, ymString) {
    if ((projectName || '') !== 'JFGP All Leave') return projectName;
    var parsed = parseMonthString_(ymString);
    if (!parsed) return projectName;
    var endOfMonth = new Date(parsed.year, parsed.month, 0);
    var match = lookupsAll.find(function(row){ return (row[0] + '').trim() === resourceName || (row[6] + '').trim() === resourceName;});
    var colD = match ? match[3] : null;
    var dt = colD instanceof Date ? colD : null;
    if (dt && dt > endOfMonth) return 'Maternity Leave';
    return 'All Leave';
  }

  var rows = Object.keys(agg).filter(function(key){
    var val = (agg[key].cols[4] || '').trim();
    return !/^not in lookup$/i.test(val);
  }).sort(function(a, b){
    return a.localeCompare(b);
  }).map(function(key){
    var entry = agg[key];
    var cols = entry.cols.slice();
    cols[1] = adjustAccount_(cols[2], cols[1]);
    var billable = deriveBillable_(entry.cols[2]);
    var lookupKey = (entry.cols[0] || '') + (entry.cols[2] || '') + (entry.cols[3] || '');
    var estAct = estActMap[lookupKey] || { est: 0, act: 0 };
    var actValue = estAct.act && estAct.act !== 0 ? estAct.act : entry.sum;
    var variance = actValue - (estAct.est || 0);
    var staff = staffInfo[entry.cols[3]] || {};
    var hubVal = staff.hub || hubLookup[entry.cols[3]] || '';
    var ymParsed = parseMonthString_(entry.cols[0]);
    var relYear = ymParsed ? (ymParsed.year - todayYear) : '';
    var relMonth = ymParsed ? ((todayYear - ymParsed.year) * 12 + (todayMonth - ymParsed.month)) : '';
    var projectAdj = adjustedProject_(entry.cols[2], entry.cols[3], entry.cols[0]);
    return cols.concat([entry.sum, billable, estAct.est, estAct.act, actValue, variance, staff.region || '', staff.country || '', hubVal, relYear, relMonth, projectAdj]);
  });

  var dest = ss.getSheetByName(destName) || ss.insertSheet(destName);
  clearSheetContents_(dest);
  dest.getRange(1, 1, 1, outHeader.length).setValues([outHeader]);
  if (rows.length) {
    dest.getRange(2, 1, rows.length, outHeader.length).setValues(rows);
  }
}

/**
 * Copies the Lookups data from the spreadsheet URL/ID specified in Config!C2
 * (range "Import!A1:E") into the local Lookups sheet, replacing the IMPORTRANGE.
 */
function refreshLookupsFromConfig_(config) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceRef = config && config.lookupsImportUrl ? (config.lookupsImportUrl + '').trim() : '';
  if (!sourceRef) {
    Logger.log('Lookups import skipped: Config!C2 is empty.');
    return false;
  }
  try {
    var source = openSpreadsheetByUrlOrId_(sourceRef);
    var sourceSheet = source.getSheetByName('Import');
    if (!sourceSheet) {
      Logger.log('Lookups import skipped: sheet "Import" not found in source.');
      return false;
    }
    var values = sourceSheet.getDataRange().getValues();
    if (!values || !values.length) {
      Logger.log('Lookups import skipped: source has no data.');
      return false;
    }
    var dest = ss.getSheetByName('Lookups') || ss.insertSheet('Lookups');
    var clearRows = Math.max(dest.getLastRow(), values.length);
    if (clearRows) {
      dest.getRange(1, 1, clearRows, 5).clearContent();
    }
    dest.getRange(1, 1, values.length, values[0].length).setValues(values);
    dest.getRange('L1').setValue('Accounts');
    dest.getRange('L2').setFormula('=SORT(UNIQUE(FILTER(B2:B, E2:E <> "Project Close Out")),1,1)');
    dest.getRange('N1').setValue('Practices for Actuals Report');
    dest.getRange('N2').setFormula('=JOIN(", ",N3:N20)');
    dest.getRange('N3').setFormula('=UNIQUE(\'Final - Capacity\'!F3:F)');
    dest.autoResizeColumns(1, Math.max(values[0].length, 14));
    return true;
  } catch (err) {
    Logger.log('Lookups import failed: ' + err);
    return false;
  }
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
  var raw = (configSheet.getRange('C23').getDisplayValue() + '').trim(); // Changed from C24 to C23
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

function clearSheetContents_(sheet) {
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow && lastCol) {
    sheet.getRange(1, 1, lastRow, lastCol).clearContent();
  }
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
  var lookupsImportUrl = (configSheet.getRange('C2').getDisplayValue() + '').trim();

  var settings = {
    importSchedules: findValue('IMPORT-FF Schedules') || 'IMPORT-FF Schedules',
    importEstVsAct: findValue('Est vs Act - Import') || 'Est vs Act - Import',
    importActuals: findValue('Actuals - Import') || 'Actuals - Import',
    activeStaffUrl: findValue('Active Staff URL'),
    activeStaffSourceSheet: findValue('Active Staff Source Sheet'),
    activeStaffPracticeFilter: findValue('Active Staff Practice Filter'),
    activeStaffCountryRegex: findValue('Active Staff Country Regex'),
    staffSheet: findValue('Active Staff Sheet') || 'Active staff',
    overrideSchedules: findValue('FF Schedule Override Sheet') || 'FF Schedule Override',
    countryHours: findValue('Country Hours Sheet') || 'Country Hours',
    availabilityMatrix: findValue('Availability Matrix Sheet') || 'Availability Matrix',
    consolidatedSchedulesSheet: findValue('Consolidated Schedules Sheet') || 'Consolidated-FF Schedules',
    finalSchedules: findValue('Final - Schedules Sheet') || 'Final - Schedules',
    finalCapacity: findValue('Final - Capacity Sheet') || 'Final - Capacity',
    finalEstVsAct: findValue('Est vs Act - Aggregated Sheet') || 'Est vs Act - Aggregated',
    varianceSourceSheet: findValue('Variance Source Sheet') || 'All Rows Needed Data Source',
    varianceSheet: findValue('Variance Sheet') || 'Variance',
    roleConfigSheet: findValue('Role Config Sheet') || 'Role Config',
    leaveProjectName: findValue('Leave Project Name') || 'JFGP All Leave',
    dataStartColumn: parseInt(findValue('Data Start Column') || '8', 10) || 8,
    regionCalendarId: calendarMatch ? calendarMatch[0] : (rawCalendarLink || ''),
    regionsInScope: regionsList,
    lookupsImportUrl: lookupsImportUrl
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
    { key: 'Active Staff Source Sheet', sample: 'Staff Export' },
    { key: 'Active Staff Practice Filter', sample: 'Earned Media (H)' },
    { key: 'Active Staff Country Regex', sample: 'UK|United Kingdom|United States|US' },
    { key: 'Active Staff Sheet', sample: 'Active staff' },
    { key: 'FF Schedule Override Sheet', sample: 'FF Schedule Override' },
    { key: 'Country Hours Sheet', sample: 'Country Hours' },
    { key: 'Availability Matrix Sheet', sample: 'Availability Matrix' },
    { key: 'Consolidated Schedules Sheet', sample: 'Consolidated-FF Schedules' },
    { key: 'Final - Schedules Sheet', sample: 'Final - Schedules' },
    { key: 'Final - Capacity Sheet', sample: 'Final - Capacity' },
    { key: 'Est vs Act - Aggregated Sheet', sample: 'Est vs Act - Aggregated' },
    { key: 'Variance Source Sheet', sample: 'All Rows Needed Data Source' },
    { key: 'Variance Sheet', sample: 'Variance' },
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

  clearSheetContents_(sheet);
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
    clearSheetContents_(countryHoursSheet);
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
  // Phase 1, Step 1: fast email import. 
  importDataFromEmails(config);
  SpreadsheetApp.flush();
  // Schedule the next step of Phase 1.
  scheduleNextStep_('phase1_step2_importActiveStaff');
}

function scheduleNextStep_(functionName) {
  // Clean up any pending triggers to avoid duplicates.
  ScriptApp.getProjectTriggers().forEach(function(trig){
    if (trig.getHandlerFunction && (
        trig.getHandlerFunction() === functionName || 
        trig.getHandlerFunction().startsWith('phase1_') ||
        trig.getHandlerFunction().startsWith('phase2_') ||
        trig.getHandlerFunction() === 'refreshAllPhase2'
      )) {
      ScriptApp.deleteTrigger(trig);
    }
  });
  if (functionName) {
    ScriptApp.newTrigger(functionName)
      .timeBased()
      .after(20 * 1000)
      .create();
  }
}

function refreshAllPhase2() {
  // This function is now deprecated and will be broken into smaller, chained functions. 
  // It is kept here for reference during the refactor but should not be called directly.
  var config = getGlobalConfig();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  buildEstVsActAggregate(config);
  rebuildVarianceSourceSheet_(config);
  refreshCountryHoursFromRegion_(ss, config);
  buildFinalSchedules(config);
  buildFinalCapacity(config);
  buildVarianceTab(config);
}

function runSetup() {
  setupConfigTab();
  setupRoleConfigTab();
}

function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Earned Media Resourcing')
      .addItem('Refresh All','refreshAll')
      .addItem('Run Setup','runSetup')
      .addToUi();
  } catch (err) {
    Logger.log('onOpen skipped (no UI available): ' + err);
  }
}

function openSpreadsheetByUrlOrId_(input) {
  if (!input) throw new Error('Spreadsheet reference is empty');
  var trimmed = (input + '').trim();
  if (!trimmed) throw new Error('Spreadsheet reference is empty');
  var attempts = [];
  var idCandidates = [];
  var pathMatch = trimmed.match(/\/d\/([-\w]+)/i);
  if (pathMatch && pathMatch[1]) {
    idCandidates.push(pathMatch[1]);
  }
  var matches = trimmed.match(/[-\w]{20,}/g);
  if (matches && matches.length) {
    matches.forEach(function(token) {
      if (idCandidates.indexOf(token) === -1) {
        idCandidates.push(token);
      }
    });
  }
  for (var i = 0; i < idCandidates.length; i++) {
    var candidate = idCandidates[i];
    try {
      return SpreadsheetApp.openById(candidate);
    } catch (err) {
      attempts.push('openById(' + candidate + '): ' + err);
    }
  }
  if (/^https?:\/\//i.test(trimmed)) {
    try {
      return SpreadsheetApp.openByUrl(trimmed);
    } catch (err2) {
      attempts.push('openByUrl: ' + err2);
    }
  }
  var message = 'Unable to open spreadsheet from reference "' + trimmed + '".';
  if (attempts.length) {
    message += ' Attempts: ' + attempts.join(' | ');
  } else {
    message += ' Provide a full Google Sheets URL or ID.';
  }
  throw new Error(message);
}

function getStaffFilterConfig_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName('Config');
  var practice = '';
  var regionPattern = '';
  if (configSheet) {
    try {
      var cfg = getGlobalConfig();
      practice = (cfg.activeStaffPracticeFilter || '').trim();
      regionPattern = (cfg.activeStaffCountryRegex || '').trim();
    } catch (err) {
      // fall back to direct cell reads below
    }
    if (!practice) {
      practice = (configSheet.getRange('C4').getDisplayValue() + '').trim();
    }
    if (!regionPattern) {
      regionPattern = (configSheet.getRange('C6').getDisplayValue() + '').trim();
    }
  }
  return { practice: practice, regionPattern: regionPattern };
}

function testVariance() {
  buildVarianceTab(getGlobalConfig());
}

// ----- Chained Execution for Phase 1 ----- 

function phase1_step2_importActiveStaff() {
  try {
    var config = getGlobalConfig();
    Logger.log('Starting Phase 1, Step 2: importAndFilterActiveStaff');
    importAndFilterActiveStaff(config);
    Logger.log('Completed Step 2. Scheduling Step 3.');
    scheduleNextStep_('phase1_step3_refreshLookups');
  } catch (e) {
    Logger.log('Error in step 2 (importAndFilterActiveStaff): ' + e.toString());
  }
}

function phase1_step3_refreshLookups() {
  try {
    var config = getGlobalConfig();
    Logger.log('Starting Phase 1, Step 3: refreshLookupsFromConfig_');
    refreshLookupsFromConfig_(config);
    SpreadsheetApp.flush();
    Logger.log('Completed Step 3. Scheduling Phase 2.');
    scheduleNextStep_('phase2_step1_buildEstVsAct');
  } catch (e) {
    Logger.log('Error in step 3 (refreshLookupsFromConfig_): ' + e.toString());
  }
}

// ----- Chained Execution for Phase 2 ----- 

function phase2_step1_buildEstVsAct() {
  try {
    var config = getGlobalConfig();
    Logger.log('Starting Phase 2, Step 1: buildEstVsActAggregate');
    buildEstVsActAggregate(config);
    Logger.log('Completed Step 1. Scheduling Step 2.');
    scheduleNextStep_('phase2_step2_rebuildVarianceSource');
  } catch (e) {
    Logger.log('Error in step 1 (buildEstVsActAggregate): ' + e.toString());
  }
}

function phase2_step2_rebuildVarianceSource() {
  try {
    var config = getGlobalConfig();
    Logger.log('Starting Phase 2, Step 2: rebuildVarianceSourceSheet_');
    rebuildVarianceSourceSheet_(config);
    Logger.log('Completed Step 2. Scheduling Step 3.');
    scheduleNextStep_('phase2_step3_refreshCountryHours');
  } catch (e) {
    Logger.log('Error in step 2 (rebuildVarianceSourceSheet_): ' + e.toString());
  }
}

function phase2_step3_refreshCountryHours() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var config = getGlobalConfig();
    Logger.log('Starting Phase 2, Step 3: refreshCountryHoursFromRegion_');
    refreshCountryHoursFromRegion_(ss, config);
    Logger.log('Completed Step 3. Scheduling Step 4.');
    scheduleNextStep_('phase2_step4_buildFinalSchedules');
  } catch (e) {
    Logger.log('Error in step 3 (refreshCountryHoursFromRegion_): ' + e.toString());
  }
}

function phase2_step4_buildFinalSchedules() {
  try {
    var config = getGlobalConfig();
    Logger.log('Starting Phase 2, Step 4: buildFinalSchedules');
    buildFinalSchedules(config);
    Logger.log('Completed Step 4. Scheduling Step 5.');
    scheduleNextStep_('phase2_step5_buildFinalCapacity');
  } catch (e) {
    Logger.log('Error in step 4 (buildFinalSchedules): ' + e.toString());
  }
}

function phase2_step5_buildFinalCapacity() {
  try {
    var config = getGlobalConfig();
    Logger.log('Starting Phase 2, Step 5: buildFinalCapacity');
    buildFinalCapacity(config);
    Logger.log('Completed Step 5. Scheduling Step 6.');
    scheduleNextStep_('phase2_step6_buildVarianceTab');
  } catch (e) {
    Logger.log('Error in step 5 (buildFinalCapacity): ' + e.toString());
  }
}

function phase2_step6_buildVarianceTab() {
  try {
    var config = getGlobalConfig();
    Logger.log('Starting Phase 2, Step 6: buildVarianceTab');
    buildVarianceTab(config);
    Logger.log('Phase 2 complete. All steps finished.');
    // This is the last step, so we delete any remaining triggers for this chain. 
    scheduleNextStep_(null); 
  } catch (e) {
    Logger.log('Error in step 6 (buildVarianceTab): ' + e.toString());
  }
}
