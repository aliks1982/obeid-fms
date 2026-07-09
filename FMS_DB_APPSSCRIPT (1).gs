
/***********************************************************************
 * FMS COMMAND — PLATFORM DATA APPS SCRIPT
 * Spreadsheet: FMS Command — Platform Data
 * Sheet ID: 13zPJDPlYhbqqZ2rXq8Ev7fm0Oj4FHI54dP0kSaR8rnw
 *
 * SETUP INSTRUCTIONS:
 * 1. Open the FMS Command — Platform Data spreadsheet
 * 2. Extensions → Apps Script
 * 3. Delete any existing code, paste this entire file
 * 4. Save → Deploy → New deployment
 *    Type: Web App | Execute as: Me | Access: Anyone
 * 5. Copy the Web App URL → paste in FMS Platform Settings → DB URL
 ***********************************************************************/

const DB_SHEET_ID = '13zPJDPlYhbqqZ2rXq8Ev7fm0Oj4FHI54dP0kSaR8rnw';

// All record types → their Google Sheet tab names
const SHEET_MAP = {
  eoc:          'ESR_Rounds',
  training:     'Training',
  incident:     'Incidents',
  secInc:       'Security_Incidents',
  secDrill:     'Security_Drills',
  chemical:     'Hazmat_Chemicals',
  wasteAudit:   'Waste_Audits',
  emDrill:      'Emergency_Drills',
  fireAlarm:    'Fire_Alarm_Tests',
  fireSup:      'Fire_Suppression',
  fireExit:     'Fire_Exit_Checks',
  fireDrill:    'Fire_Drills',
  ppm:          'Equipment_PPM',
  genTest:      'Generator_Tests',
  waterTest:    'Water_Tests',
  pmgCheck:     'PMG_Checks',
  airPurity:    'Air_Purity',
  tldReview:    'Radiation_TLD',
  leadApron:    'Lead_Aprons',
  shieldCert:   'Shielding_Certs',
  rad_tld:      'Radiation_TLD',
  rad_apron:    'Lead_Aprons',
  rad_shield:   'Shielding_Certs',
  doc:          'Documents',
  workorders:   'Work_Orders',
  eq_added:     'Equipment_Added',
  eq_relocations: 'Equipment_Relocations',
  contracts:    'Contracts',
  evaluations:  'Contract_Evaluations',
  risk_register:'Risk_Register',
  mg_lox:       'MedGas_LOX',
  lox_gauge:    'LOX_Gauge_Readings',  // Physical gauge inch H2O readings — FMS 32.8
  mg_o2backup:  'MedGas_O2Backup',
  mg_n2o:       'MedGas_N2O',
  mg_vacuum:    'MedGas_Vacuum',
  mg_aircomp:   'MedGas_AirComp',
  mg_airbackup: 'MedGas_AirBackup',
  mg_ambulance: 'MedGas_Ambulance',
  mg_cylinder:  'MedGas_Cylinder',
  mg_o2refill:  'MedGas_O2Refill',
  mg_airrefill: 'MedGas_AirRefill',
  app_settings:      'App_Settings',
  platform_settings: 'Platform_Settings',  // users, PINs, extra accounts — syncs across all devices
  kpi2_entries:      'KPI2_Entries',        // FMS KPI data entries
  commTest:          'Communication_Tests', // FMS 39.2 — PA/PABX/nurse call system tests
  pmg_annex:         'PMG_Annexes',         // FMS 32.2 — Annexes A,B,C,D,F,G,I records
  pmg_ppm:           'PMG_PPM',             // FMS 32.1 — PMG preventive maintenance
  cyl_test:          'Cylinder_Tests',      // FMS 32.9 — cylinder hydrostatic tests
  ach_readings:      'ACH_Readings',        // MOH IPC — room pressure/ACH/temp/humidity (Google Form)
};

// ── HELPERS ──────────────────────────────────────────────────────
function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    Logger.log('Created sheet: ' + name);
  }
  return sheet;
}

function flattenRecord(rec) {
  // Flatten a JSON record into ordered columns
  // Always put _id, _ts, _by, _rl first, then all other keys
  var priority = ['_id','_ts','_by','_rl','_type','_label'];
  var keys = Object.keys(rec);
  var ordered = priority.filter(k => keys.includes(k));
  keys.forEach(function(k) { if (!ordered.includes(k)) ordered.push(k); });
  return { keys: ordered, values: ordered.map(function(k) {
    var v = rec[k];
    if (v === null || v === undefined) return '';
    if (typeof v === 'object') return JSON.stringify(v);
    return String(v);
  })};
}

function ensureHeaders(sheet, keys) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(keys);
    sheet.getRange(1, 1, 1, keys.length).setFontWeight('bold')
         .setBackground('#1a3a5c').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    return;
  }
  // Check if headers match — update if new keys added
  var existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var existingSet = existingHeaders.map(String);
  var needsUpdate = false;
  keys.forEach(function(k) { if (!existingSet.includes(k)) needsUpdate = true; });
  if (needsUpdate) {
    // Append missing headers
    keys.forEach(function(k, i) {
      if (!existingSet.includes(k)) {
        var col = existingSet.length + 1;
        existingSet.push(k);
        sheet.getRange(1, col).setValue(k).setFontWeight('bold')
             .setBackground('#1a3a5c').setFontColor('#ffffff');
      }
    });
  }
}

// ── doGet — fetch all records for a type ─────────────────────────
function doGet(e) {
  try {
    var action = e.parameter.action || 'getAll';
    var type   = e.parameter.type   || '';
    var ss     = SpreadsheetApp.openById(DB_SHEET_ID);

    // ── GET-based upsert (CORS-safe alternative to POST) ──
    if (action === 'upsert' && type) {
      var record = {};
      try { record = JSON.parse(decodeURIComponent(e.parameter.record || '{}')); } catch(ex) {}
      var sheetName = SHEET_MAP[type];
      if (!sheetName) return jsonErrGet('Unknown type: ' + type);
      var sheet = getOrCreateSheet(ss, sheetName);
      var flat  = flattenRecord(record);
      ensureHeaders(sheet, flat.keys);
      var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(String);
      var idCol   = headers.indexOf('_id');
      if (idCol >= 0 && sheet.getLastRow() > 1) {
        var idVals = sheet.getRange(2, idCol+1, sheet.getLastRow()-1, 1).getValues();
        for (var i = 0; i < idVals.length; i++) {
          if (String(idVals[i][0]) === String(record._id)) {
            var rowData = headers.map(function(h) {
              var v = record[h];
              if (v === undefined || v === null) return '';
              if (typeof v === 'object') return JSON.stringify(v);
              return String(v);
            });
            sheet.getRange(i+2, 1, 1, rowData.length).setValues([rowData]);
            return ContentService.createTextOutput(JSON.stringify({ ok: true, action: 'updated' }))
              .setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
      sheet.appendRow(flat.keys.map(function(k) {
        var v = record[k];
        if (v === undefined || v === null) return '';
        if (typeof v === 'object') return JSON.stringify(v);
        return String(v);
      }));
      return ContentService.createTextOutput(JSON.stringify({ ok: true, action: 'inserted' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'getAll' && !type) {
      // Return all data from all sheets
      var result = {};
      Object.keys(SHEET_MAP).forEach(function(key) {
        var sheetName = SHEET_MAP[key];
        var sheet     = ss.getSheetByName(sheetName);
        if (!sheet || sheet.getLastRow() < 2) {
          result[key] = [];
          return;
        }
        var data    = sheet.getDataRange().getValues();
        var headers = data[0].map(String);
        var rows    = data.slice(1).map(function(row) {
          var obj = {};
          headers.forEach(function(h, i) {
            obj[h] = row[i] !== undefined && row[i] !== null ? String(row[i]) : '';
          });
          return obj;
        }).filter(function(r) { return r._id && r._id !== ''; });
        result[key] = rows;
      });
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, data: result, fetchedAt: new Date().toISOString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'getType' && type) {
      var sheetName = SHEET_MAP[type];
      if (!sheetName) return jsonErr('Unknown type: ' + type);
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() < 2) {
        return ContentService.createTextOutput(JSON.stringify({ ok:true, type:type, rows:[] }))
               .setMimeType(ContentService.MimeType.JSON);
      }
      var data    = sheet.getDataRange().getValues();
      var headers = data[0].map(String);
      var rows    = data.slice(1).map(function(row) {
        var obj = {};
        headers.forEach(function(h, i) { obj[h] = String(row[i] || ''); });
        return obj;
      }).filter(function(r) { return r._id; });
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, type: type, rows: rows }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return jsonErr('Unknown action');
  } catch(err) {
    Logger.log('doGet error: ' + err);
    return jsonErr(err.toString());
  }
}

// ── doPost — save a record ────────────────────────────────────────
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action  = payload.action;
    var type    = payload.type;
    var record  = payload.record;
    var ss      = SpreadsheetApp.openById(DB_SHEET_ID);

    Logger.log('doPost: action=' + action + ' type=' + type);

    if (action === 'bulkUpsert' && payload.table && Array.isArray(payload.records)) {
      var bSheetName = SHEET_MAP[payload.table] || payload.table;
      var bSheet = getOrCreateSheet(ss, bSheetName);
      var count = 0;
      for (var bi = 0; bi < payload.records.length; bi++) {
        var brec = payload.records[bi];
        if (!brec || !brec._id) continue;
        var bflat = flattenRecord(brec);
        ensureHeaders(bSheet, bflat.keys);
        var bHeaders = bSheet.getRange(1,1,1,bSheet.getLastColumn()).getValues()[0].map(String);
        var bIdCol = bHeaders.indexOf('_id');
        var bRow = bHeaders.map(function(h){ return brec[h] !== undefined ? String(brec[h]) : ''; });
        var found = false;
        if (bIdCol >= 0 && bSheet.getLastRow() > 1) {
          var bIdVals = bSheet.getRange(2, bIdCol+1, bSheet.getLastRow()-1, 1).getValues();
          for (var ri = 0; ri < bIdVals.length; ri++) {
            if (String(bIdVals[ri][0]) === String(brec._id)) {
              bSheet.getRange(ri+2, 1, 1, bRow.length).setValues([bRow]);
              found = true; break;
            }
          }
        }
        if (!found) bSheet.appendRow(bRow);
        count++;
      }
      return jsonOkPost({ bulkUpserted: count });
    }

    if (action === 'upsert' && type && record) {
      var sheetName = SHEET_MAP[type];
      if (!sheetName) return jsonErrPost('Unknown type: ' + type);
      var sheet = getOrCreateSheet(ss, sheetName);
      var flat  = flattenRecord(record);
      ensureHeaders(sheet, flat.keys);

      // Check if record already exists (by _id) → update row
      var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(String);
      var idCol   = headers.indexOf('_id');
      if (idCol >= 0 && sheet.getLastRow() > 1) {
        var idVals = sheet.getRange(2, idCol+1, sheet.getLastRow()-1, 1).getValues();
        for (var i = 0; i < idVals.length; i++) {
          if (String(idVals[i][0]) === String(record._id)) {
            // Update existing row
            var rowData = headers.map(function(h) {
              var v = record[h];
              if (v === undefined || v === null) return '';
              if (typeof v === 'object') return JSON.stringify(v);
              return String(v);
            });
            sheet.getRange(i+2, 1, 1, rowData.length).setValues([rowData]);
            return ContentService.createTextOutput(JSON.stringify({ok:true,action:'updated',id:record._id}))
                   .setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
      // Append new row
      var rowData = flat.keys.map(function(k) {
        var existH = headers.indexOf(k);
        return existH >= 0 ? flat.values[flat.keys.indexOf(k)] : '';
      });
      sheet.appendRow(flat.values);
      return ContentService.createTextOutput(JSON.stringify({ok:true,action:'inserted',id:record._id}))
             .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'delete' && type && payload.id) {
      var sheetName = SHEET_MAP[type];
      if (!sheetName) return jsonErrPost('Unknown type');
      var sheet   = ss.getSheetByName(sheetName); if (!sheet) return jsonErrPost('Sheet not found');
      var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(String);
      var idCol   = headers.indexOf('_id');
      if (idCol < 0) return jsonErrPost('No _id column');
      var idVals  = sheet.getRange(2,idCol+1,Math.max(1,sheet.getLastRow()-1),1).getValues();
      for (var i = 0; i < idVals.length; i++) {
        if (String(idVals[i][0]) === String(payload.id)) {
          sheet.deleteRow(i+2);
          return ContentService.createTextOutput(JSON.stringify({ok:true,action:'deleted'}))
                 .setMimeType(ContentService.MimeType.JSON);
        }
      }
      return jsonErrPost('Record not found');
    }

    return jsonErrPost('Unknown action: ' + action);
  } catch(err) {
    Logger.log('doPost error: ' + err);
    return jsonErrPost(err.toString());
  }
}

function jsonErr(msg) {
  return ContentService.createTextOutput(JSON.stringify({ok:false,error:msg}))
         .setMimeType(ContentService.MimeType.JSON);
}
function jsonErrGet(msg) {
  return ContentService.createTextOutput(JSON.stringify({ok:false,error:msg})).setMimeType(ContentService.MimeType.JSON);
}
function jsonErrPost(msg) { return jsonErr(msg); }
