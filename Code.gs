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

/** Return every recognized trip and vehicle-event row across monthly tabs. */
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
    if (source && recordTypeForRow_(source.sheet, source.schema, source.row) !== 'TRIP') {
      throw new Error('Use the separate Refill or Service action to edit this event.');
    }
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
    targetSchema.eventTypeCol = ensureEventTypeColumn_(targetSheet, targetSchema);

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

/** Add or update a refill/service event as its own dated row in the matching month tab. */
function saveMileageEvent(payload) {
  return withMileageLock_(function () {
    var ss = getBoundSpreadsheet_();
    var event = validateMileageEvent_(payload);
    var source = event.id ? findRecordById_(ss, event.id, true) : null;
    if (event.id && !source) throw new Error('That event could not be found. Refresh the sheet and try again.');
    if (source && recordTypeForRow_(source.sheet, source.schema, source.row) !== event.type) {
      throw new Error('This row is not a ' + event.type.toLowerCase() + ' event.');
    }

    var targetName = monthNameForDate_(event.date);
    var targetSheet = findMonthSheet_(ss, targetName);
    if (!targetSheet) targetSheet = createMonthSheet_(ss, targetName, event.vehicle || findDefaultVehicle_(ss));
    var targetSchema = detectMileageSchema_(targetSheet);
    if (!targetSchema) throw new Error('The ' + targetName + ' tab exists but its mileage headers were not recognized.');
    ensureEventValueColumn_(targetSheet, targetSchema, event.type);
    targetSchema.idCol = ensureIdColumn_(targetSheet, targetSchema);
    targetSchema.eventTypeCol = ensureEventTypeColumn_(targetSheet, targetSchema);

    var sameSheet = source && source.sheet.getSheetId() === targetSheet.getSheetId();
    var targetRow = sameSheet ? source.row : nextDataRow_(targetSheet, targetSchema);
    if (!sameSheet) copyPreviousRowFormat_(targetSheet, targetSchema, targetRow);
    clearTripRow_(targetSheet, targetSchema, targetRow);
    var recordId = event.id || Utilities.getUuid();
    writeMileageEventToRow_(targetSheet, targetSchema, targetRow, event, recordId);
    SpreadsheetApp.flush();

    if (source && !sameSheet) clearTripRow_(source.sheet, source.schema, source.row);
    SpreadsheetApp.flush();
    return readMileageData_();
  });
}

/** Clear an entry's cells without deleting/shifting the user's worksheet rows. */
function deleteMileageTrip(recordId) {
  return withMileageLock_(function () {
    var ss = getBoundSpreadsheet_();
    var found = findRecordById_(ss, String(recordId || ''), true);
    if (!found) throw new Error('That entry could not be found. Refresh the sheet and try again.');
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
    schema.eventTypeCol = ensureEventTypeColumn_(sheet, schema);
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
    var readColumns = managedRowWidth_(schema);
    var values = sheet.getRange(firstRow, 1, rowCount, readColumns).getValues();
    var displays = sheet.getRange(firstRow, 1, rowCount, readColumns).getDisplayValues();
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
      var recordType = normalizeRecordType_(displayValue_(display, schema.eventTypeCol));
      var isTrip = recordType === 'TRIP';
      var start = isTrip ? numberValue_(raw[schema.cols.start - 1]) : null;
      var end = isTrip ? numberValue_(raw[schema.cols.end - 1]) : null;
      var km = !isTrip || start === null || end === null ? null : roundNumber_(end - start, 2);
      var vehicle = schema.cols.vehicle ? displayValue_(display, schema.cols.vehicle) : defaultVehicle;
      records.push({
        id: id,
        recordType: recordType,
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
  var trips = records.filter(function (record) { return record.recordType === 'TRIP'; });
  var latest = trips.length ? trips[0] : null;
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

function ensureEventTypeColumn_(sheet, schema) {
  var eventTypeCol = findEventTypeColumn_(sheet, schema);
  var lastRow = Math.max(sheet.getLastRow(), schema.headerRow);
  var scanEnd = Math.max(sheet.getLastColumn(), schema.idCol + 1);
  if (!eventTypeCol) {
    for (var col = schema.idCol + 1; col <= scanEnd; col += 1) {
      var values = sheet.getRange(1, col, lastRow, 1).getDisplayValues();
      var used = values.some(function (row) { return String(row[0] || '').trim() !== ''; });
      if (!used) { eventTypeCol = col; break; }
    }
  }
  if (!eventTypeCol) eventTypeCol = scanEnd + 1;
  if (eventTypeCol > sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(), eventTypeCol - sheet.getMaxColumns());
  if (normalizeHeader_(sheet.getRange(schema.headerRow, eventTypeCol).getDisplayValue()) !== 'eventtype') {
    sheet.getRange(schema.headerRow, eventTypeCol).setValue('Event Type');
  }
  try { sheet.hideColumns(eventTypeCol); } catch (e) { /* already hidden */ }
  return eventTypeCol;
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
    schema.eventTypeCol = ensureIds ? ensureEventTypeColumn_(sheet, schema) : findEventTypeColumn_(sheet, schema);
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
    schema.eventTypeCol = ensureEventTypeColumn_(sheet, schema);
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
    whoPays: limitedText_(payload.whoPays, 150, 'Who pays?')
  };
}

function validateMileageEvent_(payload) {
  if (!payload || typeof payload !== 'object') throw new Error('Event details are missing.');
  var type = String(payload.type || '').trim().toUpperCase();
  if (type !== 'REFILL' && type !== 'SERVICE') throw new Error('Choose either a refill or service event.');
  var date = String(payload.date || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Enter a valid event date.');
  var parts = date.split('-').map(Number);
  var check = new Date(parts[0], parts[1] - 1, parts[2], 12, 0, 0);
  if (check.getFullYear() !== parts[0] || check.getMonth() !== parts[1] - 1 || check.getDate() !== parts[2]) throw new Error('Enter a valid event date.');

  var refillLitres = null;
  var service = '';
  if (type === 'REFILL') {
    refillLitres = numberValue_(payload.refillLitres);
    if (refillLitres === null || refillLitres <= 0) throw new Error('Enter a refill amount greater than zero.');
  } else {
    service = limitedText_(payload.service, MILEAGE_APP.maxTextLength, 'Service description');
    if (!service) throw new Error('Enter a service description.');
  }
  return {
    id: String(payload.id || '').trim(),
    type: type,
    date: date,
    vehicle: limitedText_(payload.vehicle, 150, 'Vehicle'),
    whoPays: limitedText_(payload.whoPays, 150, 'Who pays?'),
    refillLitres: refillLitres,
    service: service
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
  sheet.getRange(row, schema.eventTypeCol).setValue('TRIP');
}

function clearTripRow_(sheet, schema, row) {
  sheet.getRange(row, 1, 1, managedRowWidth_(schema)).clearContent();
}

function writeMileageEventToRow_(sheet, schema, row, event, recordId) {
  var c = schema.cols;
  sheet.getRange(row, c.date).setValue(new Date(Number(event.date.slice(0, 4)), Number(event.date.slice(5, 7)) - 1, Number(event.date.slice(8, 10)), 12, 0, 0));
  if (c.driver) sheet.getRange(row, c.driver).clearContent();
  if (c.vehicle) sheet.getRange(row, c.vehicle).setNumberFormat('@').setValue(safeTextCell_(event.vehicle || findDefaultVehicle_(sheet.getParent())));
  if (c.start) sheet.getRange(row, c.start).clearContent();
  if (c.end) sheet.getRange(row, c.end).clearContent();
  if (c.km) sheet.getRange(row, c.km).clearContent();
  if (c.fuel) sheet.getRange(row, c.fuel).clearContent();
  if (c.purpose) sheet.getRange(row, c.purpose).clearContent();
  if (c.whoPays) sheet.getRange(row, c.whoPays).setNumberFormat('@').setValue(safeTextCell_(event.whoPays));
  if (c.refill) sheet.getRange(row, c.refill).clearContent();
  if (c.service) sheet.getRange(row, c.service).clearContent();
  if (event.type === 'REFILL') sheet.getRange(row, c.refill).setValue(event.refillLitres);
  if (event.type === 'SERVICE') sheet.getRange(row, c.service).setNumberFormat('@').setValue(safeTextCell_(event.service));
  sheet.getRange(row, schema.idCol).setValue(recordId);
  sheet.getRange(row, schema.eventTypeCol).setValue(event.type);
}

function ensureEventValueColumn_(sheet, schema, type) {
  var key = type === 'REFILL' ? 'refill' : 'service';
  if (schema.cols[key]) return schema.cols[key];
  var column = Math.max(1, sheet.getLastColumn() + 1);
  if (column > sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(), column - sheet.getMaxColumns());
  var header = sheet.getRange(schema.headerRow, column);
  if (column > 1) {
    try { sheet.getRange(schema.headerRow, column - 1).copyTo(header, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false); } catch (e) { /* header styling is optional */ }
  }
  header.setValue(type === 'REFILL' ? 'Refill (litres)' : 'SERVICE').setWrap(true);
  try { sheet.setColumnWidth(column, 145); } catch (e2) { /* column sizing is optional */ }
  schema.cols[key] = column;
  return column;
}

function managedRowWidth_(schema) {
  var width = Math.max(schema.idCol || 0, schema.eventTypeCol || 0);
  Object.keys(schema.cols || {}).forEach(function (key) { width = Math.max(width, schema.cols[key] || 0); });
  return Math.max(width, 1);
}

function findEventTypeColumn_(sheet, schema) {
  var idCol = schema.idCol || findIdColumn_(sheet, schema);
  if (!idCol) return null;
  var lastColumn = Math.max(sheet.getLastColumn(), idCol + 1);
  for (var col = idCol + 1; col <= lastColumn; col += 1) {
    if (normalizeHeader_(sheet.getRange(schema.headerRow, col).getDisplayValue()) === 'eventtype') return col;
  }
  return null;
}

function normalizeRecordType_(value) {
  var type = String(value || '').trim().toUpperCase();
  return type === 'REFILL' || type === 'SERVICE' ? type : 'TRIP';
}

function recordTypeForRow_(sheet, schema, row) {
  var column = schema.eventTypeCol || findEventTypeColumn_(sheet, schema);
  return normalizeRecordType_(column ? sheet.getRange(row, column).getDisplayValue() : '');
}

function copyPreviousRowFormat_(sheet, schema, targetRow) {
  var previous = targetRow - 1;
  if (previous <= schema.headerRow) return;
  try {
    var rowWidth = managedRowWidth_(schema);
    sheet.getRange(previous, 1, 1, rowWidth).copyTo(
      sheet.getRange(targetRow, 1, 1, rowWidth),
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
