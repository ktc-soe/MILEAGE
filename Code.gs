/**
 * Mileage Record Sheet — bound Google Apps Script backend
 * Paste this file into Code.gs in Extensions > Apps Script for the target sheet.
 * Existing monthly sheets and their header layouts are detected automatically.
 */
const MILEAGE_APP = Object.freeze({
  fuelKmPerLitre: 60,
  preferredIdColumn: 13, // M, normally blank in the supplied workbook
  headerScanRows: 15,
  newMonthHeaderRow: 4,
  maxTextLength: 500,
  spreadsheetProperty: 'MILEAGE_SPREADSHEET_ID'
});

/** Run once from the Apps Script editor while it is open from the target Sheet. */
function setupMileageApp() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open Apps Script from the target Google Sheet using Extensions > Apps Script, then run setupMileageApp.');
  PropertiesService.getScriptProperties().setProperty(MILEAGE_APP.spreadsheetProperty, ss.getId());
  return 'Ready for ' + ss.getName() + '. The web app is bound to this spreadsheet.';
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Mileage Record Sheet')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Return every recognized trip row across all monthly tabs. */
function getMileageData() {
  return withMileageLock_(function () {
    return readMileageData_();
  });
}

/** Add or update a trip. The record is written to the tab matching its date. */
function saveMileageTrip(payload) {
  return withMileageLock_(function () {
    var ss = getBoundSpreadsheet_();
    var trip = validateTrip_(payload);
    var source = trip.id ? findRecordById_(ss, trip.id, true) : null;
    if (trip.id && !source) {
      throw new Error('That trip could not be found. Refresh the mileage sheet and try again.');
    }

    var targetName = monthNameForDate_(trip.date);
    var targetSheet = findMonthSheet_(ss, targetName);
    if (!targetSheet) {
      targetSheet = createMonthSheet_(ss, targetName, trip.vehicle || findDefaultVehicle_(ss));
    }
    var targetSchema = detectMileageSchema_(targetSheet);
    if (!targetSchema) {
      throw new Error('The ' + targetName + ' tab exists but its mileage headers were not recognized. The existing sheet was not changed.');
    }
    targetSchema.idCol = ensureIdColumn_(targetSheet, targetSchema);

    var sameSheet = source && source.sheet.getSheetId() === targetSheet.getSheetId();
    var targetRow = sameSheet ? source.row : nextDataRow_(targetSheet, targetSchema);
    if (!sameSheet) copyPreviousRowFormat_(targetSheet, targetSchema, targetRow);

    var recordId = trip.id || Utilities.getUuid();
    writeTripToRow_(targetSheet, targetSchema, targetRow, trip, recordId);
    SpreadsheetApp.flush();

    if (source && !sameSheet) clearTripRow_(source.sheet, source.schema, source.row);
    SpreadsheetApp.flush();
    return readMileageData_();
  });
}

/** Clear a trip's cells without deleting/shifting the user's worksheet rows. */
function deleteMileageTrip(recordId) {
  return withMileageLock_(function () {
    var ss = getBoundSpreadsheet_();
    var found = findRecordById_(ss, String(recordId || ''), true);
    if (!found) throw new Error('That trip could not be found. Refresh the mileage sheet and try again.');
    clearTripRow_(found.sheet, found.schema, found.row);
    SpreadsheetApp.flush();
    return readMileageData_();
  });
}

function readMileageData_() {
  var ss = getBoundSpreadsheet_();
  var tz = getSpreadsheetTimeZoneString_(ss);
  var records = [];
  var months = [];
  var defaultVehicle = findDefaultVehicle_(ss);
  var sheetOrder = 0;

  ss.getSheets().forEach(function (sheet) {
    var schema = detectMileageSchema_(sheet);
    if (!schema) return;

    schema.idCol = ensureIdColumn_(sheet, schema);
    var lastRow = lastDataRow_(sheet, schema);
    months.push({
      name: sheet.getName(),
      sheetId: sheet.getSheetId(),
      hasVehicle: !!schema.cols.vehicle,
      hasService: !!schema.cols.service,
      hasRefill: !!schema.cols.refill
    });

    if (lastRow <= schema.headerRow) {
      sheetOrder += 1;
      return;
    }

    var firstRow = schema.headerRow + 1;
    var rowCount = lastRow - schema.headerRow;
    var values = sheet.getRange(firstRow, 1, rowCount, schema.idCol).getValues();
    var displays = sheet.getRange(firstRow, 1, rowCount, schema.idCol).getDisplayValues();
    var idRange = sheet.getRange(firstRow, schema.idCol, rowCount, 1);
    var ids = idRange.getValues();
    var changedIds = false;

    for (var offset = 0; offset < rowCount; offset += 1) {
      var rowNumber = firstRow + offset;
      var raw = values[offset];
      var display = displays[offset];
      var dateValue = raw[schema.cols.date - 1];
      if (isBlank_(dateValue)) continue;

      var id = String(ids[offset][0] || '').trim();
      if (!id) {
        id = Utilities.getUuid();
        ids[offset][0] = id;
        changedIds = true;
      }
      var start = numberValue_(raw[schema.cols.start - 1]);
      var end = numberValue_(raw[schema.cols.end - 1]);
      var km = start === null || end === null ? null : roundNumber_(end - start, 2);
      var vehicle = schema.cols.vehicle ? displayValue_(display, schema.cols.vehicle) : defaultVehicle;
      records.push({
        id: id,
        date: dateToIso_(dateValue, tz),
        driver: displayValue_(display, schema.cols.driver),
        vehicle: vehicle || defaultVehicle || '',
        startOdometer: start,
        endOdometer: end,
        km: km,
        purpose: displayValue_(display, schema.cols.purpose),
        fuelUsed: km === null ? null : roundNumber_(km / MILEAGE_APP.fuelKmPerLitre, 2),
        whoPays: displayValue_(display, schema.cols.whoPays),
        refill: displayValue_(display, schema.cols.refill),
        service: displayValue_(display, schema.cols.service),
        sheetName: sheet.getName(),
        sheetId: sheet.getSheetId(),
        rowNumber: rowNumber,
        _sheetOrder: sheetOrder
      });
    }
    if (changedIds) idRange.setValues(ids);
    sheetOrder += 1;
  });

  records.sort(function (a, b) {
    var byDate = String(b.date || '').localeCompare(String(a.date || ''));
    if (byDate) return byDate;
    if (a._sheetOrder !== b._sheetOrder) return b._sheetOrder - a._sheetOrder;
    return b.rowNumber - a.rowNumber;
  });

  var vehicleSet = {};
  records.forEach(function (record) { if (record.vehicle) vehicleSet[record.vehicle] = true; });
  if (defaultVehicle) vehicleSet[defaultVehicle] = true;
  var vehicleOptions = Object.keys(vehicleSet).sort();
  var latest = records.length ? records[0] : null;
  records.forEach(function (record) { delete record._sheetOrder; });

  return {
    spreadsheetName: ss.getName(),
    records: records,
    months: months,
    defaultVehicle: defaultVehicle || (vehicleOptions[0] || ''),
    vehicleOptions: vehicleOptions,
    nextStartOdometer: latest && latest.endOdometer !== null ? latest.endOdometer : '',
    fuelKmPerLitre: MILEAGE_APP.fuelKmPerLitre
  };
}

function getBoundSpreadsheet_() {
  var id = PropertiesService.getScriptProperties().getProperty(MILEAGE_APP.spreadsheetProperty);
  if (id) return SpreadsheetApp.openById(id);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  throw new Error('Run setupMileageApp once from the Apps Script editor opened from the target Google Sheet.');
}

function withMileageLock_(callback) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function detectMileageSchema_(sheet) {
  var lastColumn = Math.max(1, sheet.getLastColumn());
  var scanColumns = Math.min(Math.max(lastColumn, MILEAGE_APP.preferredIdColumn), 30);
  var scanRows = Math.min(Math.max(sheet.getLastRow(), 1), MILEAGE_APP.headerScanRows);
  var cells = sheet.getRange(1, 1, scanRows, scanColumns).getDisplayValues();

  for (var r = 0; r < cells.length; r += 1) {
    var normalized = cells[r].map(normalizeHeader_);
    var cols = {
      date: findHeader_(normalized, function (h) { return h === 'date'; }),
      driver: findHeader_(normalized, function (h) { return h === 'driver'; }),
      vehicle: findHeader_(normalized, function (h) { return h === 'vehicle'; }),
      start: findHeader_(normalized, function (h) { return h === 'start' || h.indexOf('startodometer') === 0; }),
      end: findHeader_(normalized, function (h) { return h === 'end' || h.indexOf('endodometer') === 0; }),
      km: findHeader_(normalized, function (h) { return h.indexOf('km') === 0 || h.indexOf('distance') === 0; }),
      purpose: findHeader_(normalized, function (h) { return h.indexOf('purpose') >= 0 || h.indexOf('destination') >= 0; }),
      fuel: findHeader_(normalized, function (h) { return h.indexOf('fuelused') === 0; }),
      whoPays: findHeader_(normalized, function (h) { return h.indexOf('whopays') === 0 || h === 'payer'; }),
      refill: findHeader_(normalized, function (h) { return h.indexOf('refill') === 0; }),
      service: findHeader_(normalized, function (h) { return h.indexOf('service') === 0; })
    };
    if (cols.date && cols.driver && cols.start && cols.end && cols.purpose) {
      return { headerRow: r + 1, cols: cols, idCol: null };
    }
  }
  return null;
}

function findHeader_(headers, predicate) {
  for (var i = 0; i < headers.length; i += 1) {
    if (predicate(headers[i])) return i + 1;
  }
  return null;
}

function normalizeHeader_(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function ensureIdColumn_(sheet, schema) {
  var lastContentColumn = sheet.getLastColumn();
  var lastRow = Math.max(sheet.getLastRow(), schema.headerRow);
  var idCol = null;
  var scanEnd = Math.max(lastContentColumn, MILEAGE_APP.preferredIdColumn);
  for (var col = MILEAGE_APP.preferredIdColumn; col <= scanEnd; col += 1) {
    var label = normalizeHeader_(sheet.getRange(schema.headerRow, col).getDisplayValue());
    if (label === 'recordid') { idCol = col; break; }
    var columnValues = sheet.getRange(1, col, lastRow, 1).getDisplayValues();
    var used = columnValues.some(function (row) { return String(row[0] || '').trim() !== ''; });
    if (!used) { idCol = col; break; }
  }
  if (!idCol) idCol = scanEnd + 1;
  if (idCol > sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(), idCol - sheet.getMaxColumns());
  if (normalizeHeader_(sheet.getRange(schema.headerRow, idCol).getDisplayValue()) !== 'recordid') {
    sheet.getRange(schema.headerRow, idCol).setValue('Record ID');
  }
  try { sheet.hideColumns(idCol); } catch (e) { /* already hidden */ }
  return idCol;
}

function lastDataRow_(sheet, schema) {
  var last = sheet.getLastRow();
  if (last <= schema.headerRow) return schema.headerRow;
  var dates = sheet.getRange(schema.headerRow + 1, schema.cols.date, last - schema.headerRow, 1).getValues();
  for (var i = dates.length - 1; i >= 0; i -= 1) {
    if (!isBlank_(dates[i][0])) return schema.headerRow + 1 + i;
  }
  return schema.headerRow;
}

function nextDataRow_(sheet, schema) {
  return Math.max(schema.headerRow + 1, lastDataRow_(sheet, schema) + 1);
}

function findRecordById_(ss, recordId, ensureIds) {
  if (!recordId) return null;
  var found = null;
  ss.getSheets().some(function (sheet) {
    var schema = detectMileageSchema_(sheet);
    if (!schema) return false;
    schema.idCol = ensureIds ? ensureIdColumn_(sheet, schema) : findIdColumn_(sheet, schema);
    if (!schema.idCol) return false;
    var last = lastDataRow_(sheet, schema);
    if (last <= schema.headerRow) return false;
    var ids = sheet.getRange(schema.headerRow + 1, schema.idCol, last - schema.headerRow, 1).getDisplayValues();
    for (var i = 0; i < ids.length; i += 1) {
      if (String(ids[i][0]) === recordId) {
        found = { sheet: sheet, schema: schema, row: schema.headerRow + 1 + i };
        return true;
      }
    }
    return false;
  });
  return found;
}

function findIdColumn_(sheet, schema) {
  var lastColumn = Math.max(sheet.getLastColumn(), MILEAGE_APP.preferredIdColumn);
  for (var col = MILEAGE_APP.preferredIdColumn; col <= lastColumn; col += 1) {
    if (normalizeHeader_(sheet.getRange(schema.headerRow, col).getDisplayValue()) === 'recordid') return col;
  }
  return null;
}

function findMonthSheet_(ss, monthName) {
  var wanted = String(monthName).toUpperCase();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i += 1) {
    if (sheets[i].getName().toUpperCase() === wanted) return sheets[i];
  }
  return null;
}

function createMonthSheet_(ss, monthName, vehicle) {
  var sheet = ss.insertSheet(monthName);
  var headerRow = MILEAGE_APP.newMonthHeaderRow;
  sheet.getRange('A1').setValue('Mileage Record Sheet' + (vehicle ? ' - ' + vehicle : ''));
  sheet.getRange('C2').setValue('FUEL RATE (L/km):');
  sheet.getRange('E2').setValue('1L = 60KM');
  var headers = ['Date', 'Driver', 'Vehicle', 'Start', 'End', "KM's", 'Purpose / Destination', 'Fuel used (L)', 'Who pays?', 'Refill (litres)', 'SERVICE'];
  sheet.getRange(headerRow, 2, 1, headers.length).setValues([headers]);
  sheet.getRange(headerRow, 2, 1, headers.length)
    .setBackground('#167957').setFontColor('#ffffff').setFontWeight('bold').setWrap(true);
  sheet.setFrozenRows(headerRow);
  sheet.setColumnWidth(2, 112);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 125);
  sheet.setColumnWidths(5, 2, 110);
  sheet.setColumnWidth(7, 105);
  sheet.setColumnWidth(8, 240);
  sheet.setColumnWidth(9, 110);
  sheet.setColumnWidth(10, 130);
  sheet.setColumnWidth(11, 120);
  sheet.setColumnWidth(12, 150);
  var schema = detectMileageSchema_(sheet);
  if (schema) {
    schema.idCol = ensureIdColumn_(sheet, schema);
    return sheet;
  }
  throw new Error('Could not prepare a new monthly sheet.');
}

function findDefaultVehicle_(ss) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i += 1) {
    var sheet = sheets[i];
    var schema = detectMileageSchema_(sheet);
    if (!schema) continue;
    if (schema.cols.vehicle) {
      var last = lastDataRow_(sheet, schema);
      if (last > schema.headerRow) {
        var vehicles = sheet.getRange(schema.headerRow + 1, schema.cols.vehicle, last - schema.headerRow, 1).getDisplayValues();
        for (var j = 0; j < vehicles.length; j += 1) {
          if (String(vehicles[j][0]).trim()) return String(vehicles[j][0]).trim();
        }
      }
    }
  }
  for (var k = 0; k < sheets.length; k += 1) {
    var title = String(sheets[k].getRange(1, 1, 2, Math.min(3, sheets[k].getMaxColumns())).getDisplayValues().map(function (r) { return r.join(' '); }).join(' '));
    var match = title.match(/Mileage\s+Record\s+Sheet\s*[-–—]\s*(.+)/i);
    if (match && match[1]) return match[1].trim();
  }
  return '';
}

function monthNameForDate_(isoDate) {
  var months = ['JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE','JULY','AUGUST','SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER'];
  var month = Number(String(isoDate).split('-')[1]);
  if (!month || month < 1 || month > 12) throw new Error('Enter a valid trip date.');
  return months[month - 1];
}

function validateTrip_(payload) {
  if (!payload || typeof payload !== 'object') throw new Error('Trip details are missing.');
  var date = String(payload.date || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Enter a valid trip date.');
  var parts = date.split('-').map(Number);
  var check = new Date(parts[0], parts[1] - 1, parts[2], 12, 0, 0);
  if (check.getFullYear() !== parts[0] || check.getMonth() !== parts[1] - 1 || check.getDate() !== parts[2]) throw new Error('Enter a valid trip date.');

  var start = numberValue_(payload.startOdometer);
  var end = numberValue_(payload.endOdometer);
  if (start === null || start < 0) throw new Error('Enter a valid start odometer reading.');
  if (end === null || end < start) throw new Error('End odometer must be at least the start reading.');
  var driver = limitedText_(payload.driver, 150, 'Driver');
  var purpose = limitedText_(payload.purpose, MILEAGE_APP.maxTextLength, 'Purpose / Destination');
  if (!driver) throw new Error('Enter a driver name.');
  if (!purpose) throw new Error('Enter a purpose or destination.');

  return {
    id: String(payload.id || '').trim(),
    date: date,
    driver: driver,
    vehicle: limitedText_(payload.vehicle, 150, 'Vehicle'),
    startOdometer: start,
    endOdometer: end,
    purpose: purpose,
    whoPays: limitedText_(payload.whoPays, 150, 'Who pays?'),
    refill: limitedText_(payload.refill, 100, 'Refill (litres)'),
    service: limitedText_(payload.service, 200, 'Service')
  };
}

function limitedText_(value, maxLength, label) {
  var text = String(value === null || value === undefined ? '' : value).trim();
  if (text.length > maxLength) throw new Error(label + ' is too long (maximum ' + maxLength + ' characters).');
  return text;
}

function writeTripToRow_(sheet, schema, row, trip, recordId) {
  var c = schema.cols;
  sheet.getRange(row, c.date).setValue(new Date(Number(trip.date.slice(0, 4)), Number(trip.date.slice(5, 7)) - 1, Number(trip.date.slice(8, 10)), 12, 0, 0));
  sheet.getRange(row, c.driver).setNumberFormat('@').setValue(safeTextCell_(trip.driver));
  if (c.vehicle) sheet.getRange(row, c.vehicle).setNumberFormat('@').setValue(safeTextCell_(trip.vehicle));
  sheet.getRange(row, c.start).setValue(trip.startOdometer);
  sheet.getRange(row, c.end).setValue(trip.endOdometer);
  if (c.purpose) sheet.getRange(row, c.purpose).setNumberFormat('@').setValue(safeTextCell_(trip.purpose));
  if (c.whoPays) sheet.getRange(row, c.whoPays).setNumberFormat('@').setValue(safeTextCell_(trip.whoPays));
  if (c.refill) sheet.getRange(row, c.refill).setNumberFormat('@').setValue(safeTextCell_(trip.refill));
  if (c.service) sheet.getRange(row, c.service).setNumberFormat('@').setValue(safeTextCell_(trip.service));

  if (c.km) {
    sheet.getRange(row, c.km).setFormula('=' + columnLetter_(c.end) + row + '-' + columnLetter_(c.start) + row);
  }
  if (c.fuel) {
    var fuelFormula = c.km
      ? '=' + columnLetter_(c.km) + row + '/' + MILEAGE_APP.fuelKmPerLitre
      : '=(' + columnLetter_(c.end) + row + '-' + columnLetter_(c.start) + row + ')/' + MILEAGE_APP.fuelKmPerLitre;
    sheet.getRange(row, c.fuel).setFormula(fuelFormula);
  }
  sheet.getRange(row, schema.idCol).setValue(recordId);
}

function clearTripRow_(sheet, schema, row) {
  sheet.getRange(row, 1, 1, schema.idCol).clearContent();
}

function copyPreviousRowFormat_(sheet, schema, targetRow) {
  var previous = targetRow - 1;
  if (previous <= schema.headerRow) return;
  try {
    sheet.getRange(previous, 1, 1, schema.idCol).copyTo(
      sheet.getRange(targetRow, 1, 1, schema.idCol),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false
    );
  } catch (e) {
    // Formatting is optional; the row's data is still written safely.
  }
}

function columnLetter_(column) {
  var result = '';
  var n = column;
  while (n > 0) {
    var rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

function displayValue_(row, column) {
  return column ? String(row[column - 1] || '').trim() : '';
}

function numberValue_(value) {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'number') return isFinite(value) ? value : null;
  var parsed = Number(String(value).replace(/,/g, '').trim());
  return isFinite(parsed) ? parsed : null;
}

function dateToIso_(value, timezone) {
  var safeTimezone = normalizeTimeZone_(timezone);
  if (value instanceof Date && !isNaN(value.getTime())) {
    return formatDateInTimeZone_(value, safeTimezone, 'yyyy-MM-dd');
  }
  var text = String(value || '').trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(text)) return text.slice(0, 10);
  var parsed = new Date(text);
  if (!isNaN(parsed.getTime())) return formatDateInTimeZone_(parsed, safeTimezone, 'yyyy-MM-dd');
  return text;
}

function getSpreadsheetTimeZoneString_(spreadsheet) {
  var timezone = null;
  try {
    timezone = spreadsheet.getSpreadsheetTimeZone();
  } catch (e) {
    timezone = null;
  }
  if (typeof timezone === 'string' && timezone.trim()) return timezone.trim();

  try {
    timezone = Session.getScriptTimeZone();
  } catch (e2) {
    timezone = null;
  }
  return normalizeTimeZone_(timezone);
}

function normalizeTimeZone_(timezone) {
  if (typeof timezone === 'string' && timezone.trim()) return timezone.trim();
  try {
    var scriptTimezone = Session.getScriptTimeZone();
    if (typeof scriptTimezone === 'string' && scriptTimezone.trim()) return scriptTimezone.trim();
  } catch (e) {
    // Use the universal timezone below if the script timezone is unavailable.
  }
  return 'Etc/GMT';
}

function formatDateInTimeZone_(date, timezone, format) {
  var safeTimezone = normalizeTimeZone_(timezone);
  try {
    return Utilities.formatDate(date, safeTimezone, format);
  } catch (error) {
    var fallbackTimezone = normalizeTimeZone_(null);
    if (safeTimezone !== fallbackTimezone) {
      return Utilities.formatDate(date, fallbackTimezone, format);
    }
    throw new Error('Unable to format a mileage date using a valid timezone. ' + error.message);
  }
}

function isBlank_(value) {
  return value === null || value === undefined || String(value).trim() === '';
}

function roundNumber_(value, decimals) {
  var factor = Math.pow(10, decimals);
  return Math.round((Number(value) + Number.EPSILON) * factor) / factor;
}

function safeTextCell_(value) {
  var text = String(value || '');
  return /^[=+@-]/.test(text) ? "'" + text : text;
}
