/**
 * ============================================================================
 * STACK FITNESS PWA — MASTER GOOGLE SHEETS CLOUD DATA HUB (VERSION 2.0)
 * ============================================================================
 * 
 * This script transforms any standard Google Sheet into an automated, 
 * professional fitness telemetry and analytics hub for STACK.
 * 
 * Features Generated Automatically on Every Sync:
 *   1. "Executive Hub": KPI Cards, metric tiles, and automated visual charts.
 *   2. "PR & Strength Progress": All-time personal records, 1RM, and progression.
 *   3. "Workouts Log": Complete chronological history with volume heatmaps.
 *   4. "Sets & Reps Detail": Granular telemetry of every logged set.
 *   5. "Consistency Heatmap": Daily adherence calendar with color-coded matrix.
 *   6. "Nutrition & Macros": Calorie, macro, and hydration tracking logs.
 *   7. "Bodyweight & Measurements": Weight tracking with 7-day moving averages.
 *   8. "STACK_BACKUP": Protected full JSON store for 1-click two-way cloud restore.
 * 
 * ----------------------------------------------------------------------------
 * QUICK SETUP INSTRUCTIONS:
 * ----------------------------------------------------------------------------
 * 1. Open Google Sheets (https://sheets.new) and name it: "STACK Fitness Hub".
 * 2. In the menu, click: Extensions -> Apps Script.
 * 3. Delete any code in Code.gs, paste this entire file, and click "Save" (Ctrl+S).
 * 4. Click: Deploy -> New deployment.
 * 5. Click the gear icon (Select type) -> choose "Web app".
 * 6. Set Description: "STACK Master Hub API"
 *    Set "Execute as": "Me"
 *    Set "Who has access": "Anyone" (Required so your PWA can securely POST sync data)
 * 7. Click "Deploy", approve permissions with your Google account.
 * 8. Copy the "Web app URL" (ends in /exec).
 * 9. In STACK PWA -> Settings -> Paste into "Google Apps Script Web App URL".
 * 10. Tap "Sync Now" — your Google Sheet will instantly format and populate!
 * ============================================================================
 */

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'get';
    
    if (action === 'ping') {
      return jsonResponse({ status: 'ok', timestamp: new Date().toISOString() });
    }
    
    // Fetch latest backup state for PWA Restore
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
    return jsonResponse({
      status: 'success',
      state: state,
      updatedAt: sheet.getRange('B1').getValue() || new Date().toISOString()
    });
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
    const state = payload.state || payload;
    const nowStr = new Date().toISOString();
    
    // 1. Raw System Backup (for two-way mobile restore)
    updateRawBackupSheet(ss, state, nowStr);
    
    // 2. PR & Strength Progression Hub
    updatePRsSheet(ss, state);
    
    // 3. Workouts Log & Volume Heatmap
    updateWorkoutsLogSheet(ss, state);
    
    // 4. Granular Sets & Reps Telemetry
    updateSetsDetailSheet(ss, state);
    
    // 5. Consistency & Habit Heatmap
    updateConsistencySheet(ss, state);
    
    // 6. Nutrition & Macro Log
    updateNutritionSheet(ss, state);
    
    // 7. Bodyweight & Measurements Progress
    updateBodyweightSheet(ss, state);
    
    // 8. Executive Overview Dashboard & Visual Charts
    updateExecutiveDashboard(ss, state);
    
    return jsonResponse({
      status: 'success',
      updatedAt: nowStr,
      message: 'STACK Data Hub successfully updated with all analytics & charts'
    });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ============================================================================
// 1. RAW BACKUP SHEET
// ============================================================================
function updateRawBackupSheet(ss, state, nowStr) {
  let sheet = ss.getSheetByName('STACK_BACKUP');
  if (!sheet) {
    sheet = ss.insertSheet('STACK_BACKUP');
    sheet.setTabColor('#4B5563');
  }
  sheet.getRange('A1').setValue(JSON.stringify(state));
  sheet.getRange('B1').setValue(nowStr);
}

// ============================================================================
// 2. PR & STRENGTH PROGRESSION SHEET
// ============================================================================
function updatePRsSheet(ss, state) {
  let sheet = ss.getSheetByName('PR & Strength');
  if (!sheet) {
    sheet = ss.insertSheet('PR & Strength');
    sheet.setTabColor('#10B981');
  }
  
  sheet.clear();
  
  // Title Banner
  sheet.getRange('A1:H1').merge()
    .setValue('🏆 PERSONAL RECORDS & STRENGTH PROGRESSION')
    .setFontWeight('bold')
    .setFontSize(13)
    .setBackground('#0F172A')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('left');
    
  sheet.getRange('A2:H2').merge()
    .setValue('Automated strength telemetry · Estimated 1RM calculated via Brzycki Formula: Weight × (36 / (37 - Reps))')
    .setFontSize(9)
    .setFontColor('#64748B')
    .setBackground('#0F172A')
    .setHorizontalAlignment('left');
    
  const headers = [
    'Exercise Name',
    'Muscle Group',
    'All-Time PR (kg)',
    'Reps @ PR',
    'Est. 1RM (kg)',
    'Date Achieved',
    'Lifetime Sets Logged',
    'All-Time Volume (kg)'
  ];
  
  sheet.getRange(3, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1E293B')
    .setFontColor('#F8FAFC')
    .setHorizontalAlignment('center');
    
  sheet.setFrozenRows(3);
  
  // Aggregate PRs and Exercise stats from sessions & prs
  const prMap = state.prs || {};
  const sessions = state.sessions || [];
  const exStats = {};
  
  // Scan all past sessions for lifetime sets & volume
  sessions.forEach(function(sess) {
    (sess.exercises || []).forEach(function(ex) {
      const name = ex.name || 'Exercise';
      if (!exStats[name]) {
        exStats[name] = { sets: 0, vol: 0, muscle: ex.muscle || 'Full Body', bestWeight: 0, bestReps: 0, e1rm: 0, date: '' };
      }
      (ex.sets || []).forEach(function(s) {
        const w = Number(s.weight) || 0;
        const r = Number(s.reps) || 0;
        exStats[name].sets += 1;
        exStats[name].vol += (w * r);
        const est = r > 0 ? (r === 1 ? w : Math.round(w * (36 / (37 - r)))) : w;
        if (est > exStats[name].e1rm) {
          exStats[name].e1rm = est;
          exStats[name].bestWeight = w;
          exStats[name].bestReps = r;
          exStats[name].date = sess.date || '';
        }
      });
    });
  });
  
  // Merge explicit state.prs
  Object.keys(prMap).forEach(function(key) {
    const item = prMap[key];
    const name = item.exName || key;
    if (!exStats[name]) {
      exStats[name] = { sets: 0, vol: item.vol || 0, muscle: item.muscle || 'Full Body', bestWeight: item.weight || 0, bestReps: item.reps || 0, e1rm: item.e1rm || 0, date: item.date || '' };
    } else {
      if ((item.weight || 0) > exStats[name].bestWeight) exStats[name].bestWeight = item.weight;
      if ((item.e1rm || 0) > exStats[name].e1rm) exStats[name].e1rm = item.e1rm;
      if (item.date) exStats[name].date = item.date;
    }
  });
  
  const rows = Object.keys(exStats).map(function(name) {
    const s = exStats[name];
    return [
      name,
      formatMuscle(s.muscle),
      s.bestWeight || '—',
      s.bestReps || '—',
      s.e1rm || '—',
      s.date || '—',
      s.sets,
      Math.round(s.vol)
    ];
  });
  
  // Sort by Estimated 1RM descending
  rows.sort(function(a, b) {
    const aVal = typeof a[4] === 'number' ? a[4] : 0;
    const bVal = typeof b[4] === 'number' ? b[4] : 0;
    return bVal - aVal;
  });
  
  if (rows.length > 0) {
    sheet.getRange(4, 1, rows.length, headers.length).setValues(rows);
    
    // Style numbers
    sheet.getRange(4, 3, rows.length, 1).setNumberFormat('#,##0.0 "kg"');
    sheet.getRange(4, 5, rows.length, 1).setNumberFormat('#,##0.0 "kg"');
    sheet.getRange(4, 8, rows.length, 1).setNumberFormat('#,##0 "kg"');
    
    // Soft alternating rows
    for (let i = 0; i < rows.length; i++) {
      const bg = (i % 2 === 0) ? '#FFFFFF' : '#F8FAFC';
      sheet.getRange(4 + i, 1, 1, headers.length).setBackground(bg);
    }
    
    // Green highlight on PR column
    sheet.getRange(4, 3, rows.length, 1).setFontColor('#059669').setFontWeight('bold');
    sheet.getRange(4, 5, rows.length, 1).setFontColor('#0284C7').setFontWeight('bold');
  }
  
  sheet.autoResizeColumns(1, headers.length);
}

// ============================================================================
// 3. WORKOUTS LOG SHEET WITH VOLUME HEATMAP
// ============================================================================
function updateWorkoutsLogSheet(ss, state) {
  let sheet = ss.getSheetByName('Workouts Log');
  if (!sheet) {
    sheet = ss.insertSheet('Workouts Log');
    sheet.setTabColor('#EF4444');
  }
  
  sheet.clear();
  
  sheet.getRange('A1:I1').merge()
    .setValue('🏋️ WORKOUT SESSIONS LOG')
    .setFontWeight('bold')
    .setFontSize(13)
    .setBackground('#0F172A')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('left');
    
  const headers = [
    'Date',
    'Day of Week',
    'Workout Name',
    'Duration (min)',
    'Volume (kg)',
    'Exercises',
    'Total Sets',
    'Average Intensity (kg/set)',
    'Exercise Breakdown'
  ];
  
  sheet.getRange(2, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1E293B')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');
    
  sheet.setFrozenRows(2);
  
  const sessions = state.sessions || [];
  if (!sessions.length) {
    sheet.autoResizeColumns(1, headers.length);
    return;
  }
  
  const rows = sessions.map(function(s) {
    const d = s.date ? new Date(s.date + 'T12:00:00Z') : new Date();
    const dayName = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][d.getUTCDay()];
    const durMins = s.elapsed ? Math.round(s.elapsed / 60) : 0;
    
    let totalSets = 0;
    const exSummaries = (s.exercises || []).map(function(ex) {
      const setsCount = (ex.sets || []).length;
      totalSets += setsCount;
      const topSet = (ex.sets || []).reduce(function(max, cur) {
        return (cur.weight > (max.weight || 0)) ? cur : max;
      }, {});
      const topStr = topSet.weight ? ' [Top: ' + topSet.weight + 'kg×' + topSet.reps + ']' : '';
      return ex.name + ' (' + setsCount + 's)' + topStr;
    }).join(' • ');
    
    const vol = Number(s.vol) || 0;
    const avgIntensity = totalSets > 0 ? Math.round(vol / totalSets) : 0;
    
    return [
      s.date || '',
      dayName,
      (s.dayName || 'Workout').toUpperCase(),
      durMins,
      vol,
      (s.exercises || []).length,
      totalSets,
      avgIntensity,
      exSummaries
    ];
  });
  
  sheet.getRange(3, 1, rows.length, headers.length).setValues(rows);
  
  // Format numbers
  sheet.getRange(3, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(3, 4, rows.length, 1).setNumberFormat('#,##0 "min"');
  sheet.getRange(3, 5, rows.length, 1).setNumberFormat('#,##0 "kg"').setFontWeight('bold');
  sheet.getRange(3, 8, rows.length, 1).setNumberFormat('#,##0 "kg/s"');
  
  // Volume Gradient Heatmap conditional formatting
  const volRange = sheet.getRange(3, 5, rows.length, 1);
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint('#FFFFFF')
    .setGradientMidpointWithValue('#FED7AA', SpreadsheetApp.InterpolationType.PERCENTILE, '50')
    .setGradientMaxpointWithValue('#EF4444', SpreadsheetApp.InterpolationType.PERCENTILE, '90')
    .setRanges([volRange])
    .build();
  sheet.setConditionalFormatRules([rule]);
  
  sheet.autoResizeColumns(1, headers.length);
}

// ============================================================================
// 4. GRANULAR SETS & REPS DETAIL SHEET
// ============================================================================
function updateSetsDetailSheet(ss, state) {
  let sheet = ss.getSheetByName('Sets & Reps Detail');
  if (!sheet) {
    sheet = ss.insertSheet('Sets & Reps Detail');
    sheet.setTabColor('#8B5CF6');
  }
  
  sheet.clear();
  
  sheet.getRange('A1:I1').merge()
    .setValue('📊 SET-BY-SET TELEMETRY & LOG')
    .setFontWeight('bold')
    .setFontSize(13)
    .setBackground('#0F172A')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('left');
    
  const headers = [
    'Date',
    'Workout Routine',
    'Exercise Name',
    'Muscle Group',
    'Set #',
    'Weight (kg)',
    'Reps',
    'Set Volume (kg)',
    'Record Flag'
  ];
  
  sheet.getRange(2, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1E293B')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');
    
  sheet.setFrozenRows(2);
  
  const sessions = state.sessions || [];
  const rows = [];
  
  sessions.forEach(function(sess) {
    const sDate = sess.date || '';
    const rName = sess.dayName || 'Workout';
    
    (sess.exercises || []).forEach(function(ex) {
      const eName = ex.name || 'Exercise';
      const muscle = formatMuscle(ex.muscle);
      
      (ex.sets || []).forEach(function(set, idx) {
        const w = Number(set.weight) || 0;
        const r = Number(set.reps) || 0;
        const v = Math.round(w * r);
        const isPr = set.pr ? '⭐ PR' : '';
        
        rows.push([
          sDate,
          rName,
          eName,
          muscle,
          idx + 1,
          w,
          r,
          v,
          isPr
        ]);
      });
    });
  });
  
  if (rows.length > 0) {
    sheet.getRange(3, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(3, 6, rows.length, 1).setNumberFormat('#,##0.0 "kg"');
    sheet.getRange(3, 8, rows.length, 1).setNumberFormat('#,##0 "kg"');
    
    // Highlight PR tags
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][8]) {
        sheet.getRange(3 + i, 1, 1, headers.length).setBackground('#F0FDF4');
        sheet.getRange(3 + i, 9).setFontWeight('bold').setFontColor('#15803D');
      }
    }
  }
  
  sheet.autoResizeColumns(1, headers.length);
}

// ============================================================================
// 5. CONSISTENCY & HABIT HEATMAP SHEET
// ============================================================================
function updateConsistencySheet(ss, state) {
  let sheet = ss.getSheetByName('Consistency Heatmap');
  if (!sheet) {
    sheet = ss.insertSheet('Consistency Heatmap');
    sheet.setTabColor('#F59E0B');
  }
  
  sheet.clear();
  
  sheet.getRange('A1:F1').merge()
    .setValue('📅 HABIT & CONSISTENCY MATRIX')
    .setFontWeight('bold')
    .setFontSize(13)
    .setBackground('#0F172A')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('left');
    
  const headers = [
    'Date',
    'Day of Week',
    'Training Status',
    'Nutrition Status',
    'Daily Water (L)',
    'Adherence Score'
  ];
  
  sheet.getRange(2, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1E293B')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');
    
  sheet.setFrozenRows(2);
  
  const cons = state.cons || {};
  const waterMap = state.water || {};
  const restDays = state.restDays || [];
  
  // Collect all unique dates from cons, water, and last 30 days
  const dateSet = new Set(Object.keys(cons));
  Object.keys(waterMap).forEach(d => dateSet.add(d));
  
  // Ensure at least past 14 days are represented
  const now = new Date();
  for (let i = 0; i < 14; i++) {
    const d = new Date(now.getTime() - i * 86400000);
    dateSet.add(d.toISOString().split('T')[0]);
  }
  
  const sortedDates = Array.from(dateSet).sort().reverse();
  const rows = sortedDates.map(function(ds) {
    const c = cons[ds] || {};
    const dObj = new Date(ds + 'T12:00:00Z');
    const dayName = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][dObj.getUTCDay()];
    
    let trainStatus = 'REST';
    if (c.training) {
      trainStatus = '✅ TRAINED';
    } else if (restDays.includes(ds)) {
      trainStatus = '💤 PLANNED REST';
    }
    
    const nutStatus = c.nutrition ? '🥗 LOGGED' : '—';
    const waterL = waterMap[ds] ? Number(waterMap[ds]).toFixed(1) + ' L' : '—';
    
    let score = 0;
    if (c.training) score += 50;
    else if (restDays.includes(ds)) score += 50;
    if (c.nutrition) score += 50;
    
    return [
      ds,
      dayName,
      trainStatus,
      nutStatus,
      waterL,
      score + '%'
    ];
  });
  
  if (rows.length > 0) {
    sheet.getRange(3, 1, rows.length, headers.length).setValues(rows);
    
    // Colorize cells
    for (let i = 0; i < rows.length; i++) {
      const trainCell = sheet.getRange(3 + i, 3);
      if (rows[i][2].includes('TRAINED')) {
        trainCell.setBackground('#DCFCE7').setFontColor('#166534').setFontWeight('bold');
      } else if (rows[i][2].includes('PLANNED REST')) {
        trainCell.setBackground('#E0E7FF').setFontColor('#3730A3');
      }
      
      const nutCell = sheet.getRange(3 + i, 4);
      if (rows[i][3].includes('LOGGED')) {
        nutCell.setBackground('#FEF3C7').setFontColor('#92400E').setFontWeight('bold');
      }
    }
  }
  
  sheet.autoResizeColumns(1, headers.length);
}

// ============================================================================
// 6. NUTRITION & MACRO SHEET
// ============================================================================
function updateNutritionSheet(ss, state) {
  let sheet = ss.getSheetByName('Nutrition Log');
  if (!sheet) {
    sheet = ss.insertSheet('Nutrition Log');
    sheet.setTabColor('#10B981');
  }
  
  sheet.clear();
  
  sheet.getRange('A1:I1').merge()
    .setValue('🥗 DAILY NUTRITION & MACROS LOG')
    .setFontWeight('bold')
    .setFontSize(13)
    .setBackground('#0F172A')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('left');
    
  const headers = [
    'Date',
    'Total Calories',
    'Calorie Goal',
    'Calorie Delta',
    'Protein (g)',
    'Carbs (g)',
    'Fats (g)',
    'Water (L)',
    'Meals Count'
  ];
  
  sheet.getRange(2, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1E293B')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');
    
  sheet.setFrozenRows(2);
  
  const mealsMap = state.meals || {};
  const waterMap = state.water || {};
  const st = state.settings || {};
  const calGoal = st.calGoal || 2000;
  
  const dates = Object.keys(mealsMap).sort().reverse();
  if (!dates.length) {
    sheet.autoResizeColumns(1, headers.length);
    return;
  }
  
  const rows = dates.map(function(ds) {
    const mealList = mealsMap[ds] || [];
    let totCal = 0, totProt = 0, totCarb = 0, totFat = 0;
    
    mealList.forEach(function(m) {
      totCal += Number(m.cals) || 0;
      totProt += Number(m.prot) || 0;
      totCarb += Number(m.carb) || 0;
      totFat += Number(m.fat) || 0;
    });
    
    const delta = totCal - calGoal;
    const waterL = waterMap[ds] ? Number(waterMap[ds]).toFixed(2) : 0;
    
    return [
      ds,
      totCal,
      calGoal,
      delta,
      Math.round(totProt),
      Math.round(totCarb),
      Math.round(totFat),
      waterL,
      mealList.length
    ];
  });
  
  sheet.getRange(3, 1, rows.length, headers.length).setValues(rows);
  sheet.getRange(3, 2, rows.length, 3).setNumberFormat('#,##0 "kcal"');
  sheet.getRange(3, 5, rows.length, 3).setNumberFormat('#,##0 "g"');
  sheet.getRange(3, 8, rows.length, 1).setNumberFormat('0.00 "L"');
  
  sheet.autoResizeColumns(1, headers.length);
}

// ============================================================================
// 7. BODYWEIGHT & MEASUREMENTS SHEET
// ============================================================================
function updateBodyweightSheet(ss, state) {
  let sheet = ss.getSheetByName('Bodyweight');
  if (!sheet) {
    sheet = ss.insertSheet('Bodyweight');
    sheet.setTabColor('#3B82F6');
  }
  
  sheet.clear();
  
  sheet.getRange('A1:D1').merge()
    .setValue('⚖️ BODYWEIGHT & COMPOSITION TRACKER')
    .setFontWeight('bold')
    .setFontSize(13)
    .setBackground('#0F172A')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('left');
    
  const headers = [
    'Date',
    'Bodyweight (kg)',
    'Change from Baseline (kg)',
    '7-Day Moving Avg (kg)'
  ];
  
  sheet.getRange(2, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1E293B')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');
    
  sheet.setFrozenRows(2);
  
  const bwList = (state.bodyweight || []).slice().sort((a,b) => (a.date || '').localeCompare(b.date || ''));
  if (!bwList.length) {
    sheet.autoResizeColumns(1, headers.length);
    return;
  }
  
  const baseline = bwList[0].kg || 0;
  const rows = bwList.map(function(b, idx) {
    const curKg = Number(b.kg) || 0;
    const diff = Number((curKg - baseline).toFixed(2));
    
    // 7-day rolling window avg
    const startIdx = Math.max(0, idx - 6);
    let sum = 0;
    for (let k = startIdx; k <= idx; k++) {
      sum += Number(bwList[k].kg) || 0;
    }
    const avg = Number((sum / (idx - startIdx + 1)).toFixed(2));
    
    return [
      b.date || '',
      curKg,
      diff,
      avg
    ];
  });
  
  sheet.getRange(3, 1, rows.length, headers.length).setValues(rows);
  sheet.getRange(3, 2, rows.length, 3).setNumberFormat('0.00 "kg"');
  
  sheet.autoResizeColumns(1, headers.length);
}

// ============================================================================
// 8. EXECUTIVE DASHBOARD & EMBEDDED CHARTS
// ============================================================================
function updateExecutiveDashboard(ss, state) {
  let sheet = ss.getSheetByName('Executive Hub');
  if (!sheet) {
    sheet = ss.insertSheet('Executive Hub', 0);
    sheet.setTabColor('#FF3B30');
  }
  
  sheet.clear();
  
  // Remove prior charts
  const existingCharts = sheet.getCharts();
  existingCharts.forEach(c => sheet.removeChart(c));
  
  // Header Banner
  sheet.getRange('A1:J1').merge()
    .setValue('⚡ STACK FITNESS — EXECUTIVE PERFORMANCE HUB')
    .setFontWeight('bold')
    .setFontSize(16)
    .setBackground('#09090B')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('left');
    
  const athleteName = (state.settings && state.settings.name) ? state.settings.name.toUpperCase() : 'ATHLETE';
  sheet.getRange('A2:J2').merge()
    .setValue('Athlete: ' + athleteName + '  |  Last Synced: ' + new Date().toLocaleString() + '  |  Cloud Engine: Active')
    .setFontSize(9)
    .setFontColor('#A1A1AA')
    .setBackground('#09090B')
    .setHorizontalAlignment('left');
    
  // KPI CALCULATIONS
  const sessions = state.sessions || [];
  const totWorkouts = sessions.length;
  let totVolume = 0;
  let totDurSec = 0;
  sessions.forEach(function(s) {
    totVolume += (Number(s.vol) || 0);
    totDurSec += (Number(s.elapsed) || 0);
  });
  const avgDurMins = totWorkouts > 0 ? Math.round((totDurSec / totWorkouts) / 60) : 0;
  
  const streakDays = (state.streak && state.streak.cur) ? state.streak.cur : 0;
  const bestStreak = (state.streak && state.streak.best) ? state.streak.best : 0;
  const prsCount = Object.keys(state.prs || {}).length;
  
  const bwList = state.bodyweight || [];
  const latestBw = bwList.length > 0 ? bwList[bwList.length - 1].kg : '—';
  
  // KPI BLOCKS (Row 4-6)
  const kpiData = [
    ['TOTAL SESSIONS', 'TOTAL VOLUME', 'ACTIVE STREAK', 'ALL-TIME PRs', 'CURRENT WEIGHT', 'AVG DURATION'],
    [
      totWorkouts + ' workouts',
      Math.round(totVolume).toLocaleString() + ' kg',
      streakDays + ' days (Best ' + bestStreak + ')',
      prsCount + ' records',
      latestBw + ' kg',
      avgDurMins + ' mins'
    ]
  ];
  
  // Format KPI cards
  const cols = [
    { start: 1, end: 2 },
    { start: 3, end: 4 },
    { start: 5, end: 6 },
    { start: 7, end: 7 },
    { start: 8, end: 8 },
    { start: 9, end: 10 }
  ];
  
  for (let i = 0; i < cols.length; i++) {
    const colStart = cols[i].start;
    const colEnd = cols[i].end;
    
    // Label cell
    sheet.getRange(4, colStart, 1, colEnd - colStart + 1).merge()
      .setValue(kpiData[0][i])
      .setFontWeight('bold')
      .setFontSize(9)
      .setBackground('#18181B')
      .setFontColor('#A1A1AA')
      .setHorizontalAlignment('center');
      
    // Value cell
    sheet.getRange(5, colStart, 2, colEnd - colStart + 1).merge()
      .setValue(kpiData[1][i])
      .setFontWeight('bold')
      .setFontSize(14)
      .setBackground('#27272A')
      .setFontColor('#FFFFFF')
      .setHorizontalAlignment('center');
  }
  
  // Red accent line under KPI blocks
  sheet.getRange('A7:J7').setBackground('#FF3B30');
  
  // CREATE EMBEDDED CHARTS
  // 1. Volume Progression Chart
  const wLogSheet = ss.getSheetByName('Workouts Log');
  if (wLogSheet && wLogSheet.getLastRow() >= 3) {
    const lastRow = Math.min(wLogSheet.getLastRow(), 22);
    const dateRange = wLogSheet.getRange(3, 1, lastRow - 2, 1);
    const volRange = wLogSheet.getRange(3, 5, lastRow - 2, 1);
    
    const volChart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dateRange)
      .addRange(volRange)
      .setPosition(9, 1, 0, 0)
      .setOption('title', 'Session Volume Progression (kg)')
      .setOption('legend', { position: 'none' })
      .setOption('colors', ['#FF3B30'])
      .setOption('hAxis', { title: 'Workout Date' })
      .setOption('vAxis', { title: 'Volume (kg)' })
      .setOption('width', 580)
      .setOption('height', 320)
      .build();
      
    sheet.insertChart(volChart);
  }
  
  // 2. Bodyweight Trend Chart
  const bwSheet = ss.getSheetByName('Bodyweight');
  if (bwSheet && bwSheet.getLastRow() >= 3) {
    const lastRow = bwSheet.getLastRow();
    const dateRange = bwSheet.getRange(3, 1, lastRow - 2, 1);
    const weightRange = bwSheet.getRange(3, 2, lastRow - 2, 1);
    
    const bwChart = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dateRange)
      .addRange(weightRange)
      .setPosition(9, 6, 0, 0)
      .setOption('title', 'Bodyweight Trend (kg)')
      .setOption('legend', { position: 'none' })
      .setOption('colors', ['#3B82F6'])
      .setOption('curveType', 'function')
      .setOption('pointSize', 5)
      .setOption('hAxis', { title: 'Date' })
      .setOption('vAxis', { title: 'Weight (kg)' })
      .setOption('width', 540)
      .setOption('height', 320)
      .build();
      
    sheet.insertChart(bwChart);
  }
  
  sheet.autoResizeColumns(1, 10);
}

// ============================================================================
// UTILITIES
// ============================================================================
function formatMuscle(m) {
  if (!m) return 'Full Body';
  const dict = {
    chest: 'Chest', upper_chest: 'Upper Chest', lower_chest: 'Lower Chest', full_chest: 'Full Chest',
    lats: 'Lats', upper_traps: 'Upper Traps', lower_traps: 'Lower Traps', lower_back: 'Lower Back',
    front_delts: 'Front Delts', side_delts: 'Side Delts', rear_delts: 'Rear Delts',
    biceps: 'Biceps', triceps: 'Triceps', forearms: 'Forearms',
    full_quads: 'Quads', hamstrings: 'Hamstrings', glutes: 'Glutes', calves: 'Calves',
    upper_abs: 'Upper Abs', lower_abs: 'Lower Abs', obliques: 'Obliques'
  };
  return dict[m] || m.replace(/_/g, ' ').toUpperCase();
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
