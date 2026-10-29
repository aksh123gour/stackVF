/**
 * STACK FITNESS PWA - GOOGLE APPS SCRIPT CLOUD BACKUP & TWO-WAY SYNC
 * 
 * Instructions:
 * 1. Open Google Sheets (https://sheets.new) and name it "STACK Fitness Backup".
 * 2. Click "Extensions" -> "Apps Script".
 * 3. Delete any default code in Code.gs, paste this entire file, and click "Save" (disk icon).
 * 4. Click "Deploy" -> "New deployment".
 * 5. Click the gear icon next to "Select type" and select "Web app".
 * 6. Set Description: "STACK Sync API"
 *    Set "Execute as": "Me"
 *    Set "Who has access": "Anyone" (allows your mobile PWA to communicate securely)
 * 7. Click "Deploy", authorize access with your Google account.
 * 8. Copy the "Web app URL" (ends in /exec).
 * 9. Paste this URL into STACK -> Settings -> "Google Cloud Sync URL".
 */

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'get';
    
    if (action === 'ping') {
      return jsonResponse({ status: 'ok', timestamp: new Date().toISOString() });
    }
    
    // Fetch latest backup state
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('STACK_BACKUP');
    if (!sheet) {
      return jsonResponse({ status: 'empty', message: 'No backup found yet' });
    }
    
    const rawData = sheet.getRange('A1').getValue();
    if (!rawData) {
      return jsonResponse({ status: 'empty', message: 'Backup is empty' });
    }
    
    const state = JSON.parse(rawData);
    return jsonResponse({ status: 'success', state: state, updatedAt: sheet.getRange('B1').getValue() });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ status: 'error', message: 'No payload provided' });
    }
    
    const payload = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Store Full JSON State in dedicated backup sheet
    let backupSheet = ss.getSheetByName('STACK_BACKUP');
    if (!backupSheet) {
      backupSheet = ss.insertSheet('STACK_BACKUP');
    }
    
    const state = payload.state || payload;
    const jsonStr = JSON.stringify(state);
    const nowStr = new Date().toISOString();
    
    backupSheet.getRange('A1').setValue(jsonStr);
    backupSheet.getRange('B1').setValue(nowStr);
    
    // 2. Also populate a human-readable "Workouts Log" sheet
    if (state.sessions && Array.isArray(state.sessions)) {
      updateWorkoutsSheet(ss, state.sessions);
    }
    
    // 3. Also populate a human-readable "Bodyweight Log" sheet
    if (state.bodyweight && Array.isArray(state.bodyweight)) {
      updateBodyweightSheet(ss, state.bodyweight);
    }
    
    return jsonResponse({ status: 'success', updatedAt: nowStr });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function updateWorkoutsSheet(ss, sessions) {
  let sheet = ss.getSheetByName('Workouts');
  if (!sheet) {
    sheet = ss.insertSheet('Workouts');
    sheet.appendRow(['Date', 'Workout Name', 'Duration', 'Volume (kg)', 'Exercises Count', 'Exercise Details']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  
  // Clear old rows except header and write fresh list
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 6).clearContent();
  }
  
  if (!sessions.length) return;
  
  const rows = sessions.map(function(s) {
    const elapsedMins = s.elapsed ? Math.round(s.elapsed / 60) + ' min' : '—';
    const exNames = (s.exercises || []).map(function(e) {
      const setsCount = (e.sets || []).length;
      return e.name + ' (' + setsCount + ' sets)';
    }).join('; ');
    
    return [
      s.date || '',
      s.dayName || 'Workout',
      elapsedMins,
      s.vol || 0,
      (s.exercises || []).length,
      exNames
    ];
  });
  
  sheet.getRange(2, 1, rows.length, 6).setValues(rows);
}

function updateBodyweightSheet(ss, bwList) {
  let sheet = ss.getSheetByName('Bodyweight');
  if (!sheet) {
    sheet = ss.insertSheet('Bodyweight');
    sheet.appendRow(['Date', 'Weight (kg)']);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 2).clearContent();
  }
  
  if (!bwList.length) return;
  
  const rows = bwList.map(function(b) {
    return [b.date || '', b.kg || ''];
  });
  
  sheet.getRange(2, 1, rows.length, 2).setValues(rows);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
