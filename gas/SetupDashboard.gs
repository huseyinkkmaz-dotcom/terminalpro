/**
 * TERMINAL PRO — V23.1 SETUP SCRIPT (TIME-AWARE BATCHING)
 *
 * KEY CHANGE FROM V23.0:
 *   - All individual setFormula() loops replaced with batch setFormulas() calls
 *   - setupAllBatched() uses PropertiesService to save/resume progress
 *   - createAutoTrigger() installs a 10-minute trigger to run unattended
 *   - You can close your laptop — Google's servers finish the job in the background
 *
 * HOW TO USE:
 *   1. Run createAutoTrigger() once — installs the 10-minute auto-runner
 *   2. Run setupAllBatched() once  — it will process as much as it can
 *   3. Close your laptop. The trigger picks up where it left off every 10 min.
 *   4. Check Execution Log to see progress. When you see "ALL PHASES COMPLETE",
 *      run clearSetupState() to clean up, then remove the auto-trigger.
 *
 * PHASES:
 *   0 = Intra Levels (headers + stats + historical formulas)
 *   1 = Intra Live sheet
 *   2 = Supporting sheets (OpenTrades, ClosedTrades, AlertsLog, ZScoreAge)
 *   3 = Generate credit pairs
 *   4 = Credit Levels (headers + stats + historical formulas)
 *   5 = Credit Live sheet
 *   6 = Install hourly/daily triggers
 *   7 = DONE
 *
 * ORIGINAL FUNCTIONS (setupDashboard, generateCreditPairs, etc.) still work
 * for manual runs on small datasets — they now also use batch formula writes.
 */

// ============================================================
// TIME-AWARE BATCHED SETUP — the main entry point
// ============================================================

/** Max execution time before saving state (4 min 10 sec of 6 min limit) */
var MAX_RUNTIME_MS = 250000;

/**
 * Run this function. It picks up from the last saved phase.
 * If it runs out of time, it saves progress and exits cleanly.
 * The 10-minute trigger (from createAutoTrigger) will resume it.
 */
function setupAllBatched() {
  var startTime = new Date().getTime();
  var props = PropertiesService.getScriptProperties();
  var stateJson = props.getProperty('SETUP_STATE');
  var state = stateJson ? JSON.parse(stateJson) : { phase: 0 };

  Logger.log('=== setupAllBatched START === Phase: ' + state.phase);

  var ss = SpreadsheetApp.getActive();

  try {
    // --- PHASE 0: Intra-company Levels sheet ---
    if (state.phase === 0) {
      Logger.log('Phase 0: Building Levels sheet...');
      var pairsSheet = ss.getSheetByName('Pairs');
      if (!pairsSheet) { Logger.log('ERROR: Pairs sheet not found. Create it first.'); return; }
      var pairs = pairsSheet.getDataRange().getValues();
      var numPairs = pairs.length - 1;
      if (numPairs < 1) { Logger.log('ERROR: Pairs sheet is empty.'); return; }

      var levels = getOrCreateSheet_(ss, 'Levels');
      levels.getRange(1, 1, 1, 7).setValues([['PairID', 'Mean', 'StDev', 'Lower_5', 'Upper_95', 'HistCount', 'Historical Spread Data →']]);
      levels.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');

      // Stats formulas — batch write
      var statsF = [];
      for (var i = 0; i < numPairs; i++) {
        var r = i + 2;
        statsF.push([
          '=Pairs!A' + (i + 2),
          '=IFERROR(AVERAGE(G' + r + ':' + r + '), 0)',
          '=IFERROR(STDEV(G' + r + ':' + r + '), 0.001)',
          '=IFERROR(PERCENTILE(G' + r + ':' + r + ', 0.05), 0)',
          '=IFERROR(PERCENTILE(G' + r + ':' + r + ', 0.95), 0)',
          '=COUNTA(G' + r + ':' + r + ')'
        ]);
      }
      levels.getRange(2, 1, statsF.length, 6).setFormulas(statsF);

      // Historical spread formulas — BATCH (was individual setFormula per row!)
      var histF = [];
      for (var i = 0; i < numPairs; i++) {
        var tA = String(pairs[i + 1][1]).trim();
        var tB = String(pairs[i + 1][2]).trim();
        if (!tA || !tB) { histF.push(['']); continue; }
        histF.push([
          '=IFERROR(TRANSPOSE(ARRAYFORMULA(QUERY(GOOGLEFINANCE("' + tA + '","price",TODAY()-90,TODAY()),"select Col2 offset 1",0)-QUERY(GOOGLEFINANCE("' + tB + '","price",TODAY()-90,TODAY()),"select Col2 offset 1",0))),)'
        ]);
      }
      levels.getRange(2, 7, histF.length, 1).setFormulas(histF);
      levels.setFrozenRows(1);
      Logger.log('Phase 0 complete: ' + numPairs + ' intra Levels rows written (batched).');

      state.phase = 1;
      if (isTimeUp_(startTime)) { saveState_(props, state, 'Phase 0 done. Pausing before Live sheet.'); return; }
    }

    // --- PHASE 1: Intra-company Live sheet ---
    if (state.phase === 1) {
      Logger.log('Phase 1: Building Live sheet...');
      var pairsSheet = ss.getSheetByName('Pairs');
      var pairs = pairsSheet.getDataRange().getValues();
      var numPairs = pairs.length - 1;
      buildLiveSheet_(ss, 'Live', 'Pairs', 'Levels', numPairs, pairs);
      Logger.log('Phase 1 complete: Live sheet built with ' + numPairs + ' rows.');

      state.phase = 2;
      if (isTimeUp_(startTime)) { saveState_(props, state, 'Phase 1 done. Pausing before supporting sheets.'); return; }
    }

    // --- PHASE 2: Supporting sheets ---
    if (state.phase === 2) {
      Logger.log('Phase 2: Ensuring supporting sheets...');
      ensureSheet_(ss, 'OpenTrades', ['PairID', 'EntryZ', 'CostA', 'CostB', 'SizeA', 'SizeB', 'Timestamp']);
      ensureSheet_(ss, 'ClosedTrades', ['PairID', 'EntryZ', 'CostA', 'CostB', 'SizeA', 'SizeB', 'OpenDate', 'CloseDate', 'PnL']);
      ensureSheet_(ss, 'AlertsLog', ['Timestamp', 'PairID', 'Z-Score', 'Spread']);
      var ageSheet = ss.getSheetByName('ZScoreAge');
      if (!ageSheet) {
        ageSheet = ss.insertSheet('ZScoreAge');
        ageSheet.getRange(1, 1, 1, 3).setValues([['PairID', 'FirstCrossTimestamp', 'Source']]);
        ageSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
      }
      Logger.log('Phase 2 complete: Supporting sheets ready.');

      state.phase = 3;
      if (isTimeUp_(startTime)) { saveState_(props, state, 'Phase 2 done. Pausing before credit pair generation.'); return; }
    }

    // --- PHASE 3: Generate credit pairs ---
    if (state.phase === 3) {
      Logger.log('Phase 3: Generating credit pairs...');
      var master = ss.getSheetByName('Master');
      if (!master) {
        Logger.log('WARNING: No Master sheet. Skipping credit pairs.');
        state.phase = 6; // Skip to triggers
      } else {
        var creditPairs = generateCreditPairsBatched_(ss);
        Logger.log('Phase 3 complete: ' + creditPairs + ' credit pairs generated.');
        state.phase = 4;
      }
      if (isTimeUp_(startTime)) { saveState_(props, state, 'Phase 3 done. Pausing before CreditLevels.'); return; }
    }

    // --- PHASE 4: Credit Levels sheet ---
    if (state.phase === 4) {
      Logger.log('Phase 4: Building CreditLevels sheet...');
      var cpSheet = ss.getSheetByName('CreditPairs');
      if (!cpSheet || cpSheet.getLastRow() < 2) {
        Logger.log('WARNING: CreditPairs empty. Skipping.');
        state.phase = 6;
      } else {
        var pairs = cpSheet.getDataRange().getValues();
        var numPairs = pairs.length - 1;

        var cLevels = getOrCreateSheet_(ss, 'CreditLevels');
        cLevels.getRange(1, 1, 1, 7).setValues([['PairID', 'Mean', 'StDev', 'Lower_5', 'Upper_95', 'HistCount', 'Historical Spread Data →']]);
        cLevels.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#2d1a2e').setFontColor('#ffffff');

        // Stats formulas — batch
        var statsF = [];
        for (var i = 0; i < numPairs; i++) {
          var r = i + 2;
          statsF.push([
            '=CreditPairs!A' + (i + 2),
            '=IFERROR(AVERAGE(G' + r + ':' + r + '), 0)',
            '=IFERROR(STDEV(G' + r + ':' + r + '), 0.001)',
            '=IFERROR(PERCENTILE(G' + r + ':' + r + ', 0.05), 0)',
            '=IFERROR(PERCENTILE(G' + r + ':' + r + ', 0.95), 0)',
            '=COUNTA(G' + r + ':' + r + ')'
          ]);
        }
        cLevels.getRange(2, 1, statsF.length, 6).setFormulas(statsF);

        // Historical spread formulas — BATCH
        var histF = [];
        for (var i = 0; i < numPairs; i++) {
          var tA = String(pairs[i + 1][1]).trim();
          var tB = String(pairs[i + 1][2]).trim();
          if (!tA || !tB) { histF.push(['']); continue; }
          histF.push([
            '=IFERROR(TRANSPOSE(ARRAYFORMULA(QUERY(GOOGLEFINANCE("' + tA + '","price",TODAY()-90,TODAY()),"select Col2 offset 1",0)-QUERY(GOOGLEFINANCE("' + tB + '","price",TODAY()-90,TODAY()),"select Col2 offset 1",0))),)'
          ]);
        }
        cLevels.getRange(2, 7, histF.length, 1).setFormulas(histF);
        cLevels.setFrozenRows(1);
        Logger.log('Phase 4 complete: ' + numPairs + ' credit Levels rows written (batched).');
        state.phase = 5;
      }
      if (isTimeUp_(startTime)) { saveState_(props, state, 'Phase 4 done. Pausing before CreditLive.'); return; }
    }

    // --- PHASE 5: Credit Live sheet ---
    if (state.phase === 5) {
      Logger.log('Phase 5: Building CreditLive sheet...');
      var cpSheet = ss.getSheetByName('CreditPairs');
      var pairs = cpSheet.getDataRange().getValues();
      var numPairs = pairs.length - 1;
      buildLiveSheet_(ss, 'CreditLive', 'CreditPairs', 'CreditLevels', numPairs, pairs);
      Logger.log('Phase 5 complete: CreditLive built with ' + numPairs + ' rows.');

      state.phase = 6;
      if (isTimeUp_(startTime)) { saveState_(props, state, 'Phase 5 done. Pausing before triggers.'); return; }
    }

    // --- PHASE 6: Install triggers ---
    if (state.phase === 6) {
      Logger.log('Phase 6: Installing triggers...');
      // Remove existing
      var triggers = ScriptApp.getProjectTriggers();
      for (var i = 0; i < triggers.length; i++) {
        var fn = triggers[i].getHandlerFunction();
        if (fn === 'snapshotZScores' || fn === 'dailyCreditRefresh') {
          ScriptApp.deleteTrigger(triggers[i]);
        }
      }
      ScriptApp.newTrigger('snapshotZScores').timeBased().everyHours(1).create();
      ScriptApp.newTrigger('dailyCreditRefresh').timeBased().atHour(5).everyDays(1).create();
      Logger.log('Phase 6 complete: Triggers installed.');

      state.phase = 7;
    }

    // --- PHASE 7: DONE ---
    if (state.phase >= 7) {
      props.deleteProperty('SETUP_STATE');
      // Remove the auto-setup trigger since we're done
      removeAutoSetupTrigger_();
      SpreadsheetApp.flush();
      Logger.log('=== ALL PHASES COMPLETE === Setup finished successfully!');
      showMsg_('Setup complete! All sheets built. GOOGLEFINANCE data will populate over the next few minutes.');
      return;
    }

  } catch (e) {
    Logger.log('ERROR in phase ' + state.phase + ': ' + e.toString());
    // Save state so the trigger can retry from this phase
    saveState_(props, state, 'Error occurred, will retry: ' + e.toString());
  }
}

/** Check if we're approaching the time limit */
function isTimeUp_(startTime) {
  return (new Date().getTime() - startTime) > MAX_RUNTIME_MS;
}

/** Save current phase to PropertiesService and log */
function saveState_(props, state, msg) {
  props.setProperty('SETUP_STATE', JSON.stringify(state));
  Logger.log('⏸ ' + msg + ' Saved at phase ' + state.phase + '. Will resume on next trigger run.');
}

/**
 * Creates a 10-minute trigger that runs setupAllBatched().
 * Run this ONCE, then run setupAllBatched() to kick it off.
 * The trigger auto-removes itself when setup is complete.
 */
function createAutoTrigger() {
  // Remove any existing auto-setup trigger first
  removeAutoSetupTrigger_();

  ScriptApp.newTrigger('setupAllBatched')
    .timeBased()
    .everyMinutes(10)
    .create();

  Logger.log('Auto-trigger installed: setupAllBatched will run every 10 minutes.');
  showMsg_(
    'Auto-trigger created!\n\n' +
    'Now run setupAllBatched() once to start.\n' +
    'It will continue automatically every 10 minutes until done.\n' +
    'You can close your laptop — check Execution Log for progress.'
  );
}

/** Remove the auto-setup trigger (called when setup is complete) */
function removeAutoSetupTrigger_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'setupAllBatched') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/** Reset saved progress — use if setup gets stuck */
function clearSetupState() {
  PropertiesService.getScriptProperties().deleteProperty('SETUP_STATE');
  removeAutoSetupTrigger_();
  Logger.log('Setup state cleared. Run createAutoTrigger() + setupAllBatched() to start fresh.');
  showMsg_('Setup state cleared. You can start fresh.');
}

/** Show current setup progress */
function checkSetupProgress() {
  var stateJson = PropertiesService.getScriptProperties().getProperty('SETUP_STATE');
  var phaseNames = [
    '0: Intra Levels', '1: Intra Live', '2: Supporting Sheets',
    '3: Credit Pair Generation', '4: Credit Levels', '5: Credit Live',
    '6: Triggers', '7: DONE'
  ];
  if (!stateJson) {
    showMsg_('No setup in progress. Run createAutoTrigger() + setupAllBatched() to start.');
  } else {
    var state = JSON.parse(stateJson);
    showMsg_('Setup in progress.\n\nCurrent phase: ' + (phaseNames[state.phase] || state.phase) +
             '\n\nPhases remaining: ' + (7 - state.phase));
  }
}

// ============================================================
// CREDIT PAIR GENERATION (headless, used by setupAllBatched)
// ============================================================
/** Returns the number of pairs generated */
function generateCreditPairsBatched_(ss) {
  var master = ss.getSheetByName('Master');
  if (!master) return 0;
  var data = master.getDataRange().getValues();
  if (data.length < 2) return 0;

  // Build company map from Pairs sheet (union-find)
  var companyOf = {};
  var intraPairs = {};
  var pairsSheet = ss.getSheetByName('Pairs');
  if (pairsSheet) {
    var pData = pairsSheet.getDataRange().getValues();
    for (var i = 1; i < pData.length; i++) {
      var ptA = String(pData[i][1]).trim().toUpperCase();
      var ptB = String(pData[i][2]).trim().toUpperCase();
      if (!ptA || !ptB) continue;
      intraPairs[cleanId_(ptA) + "_" + cleanId_(ptB)] = true;
      intraPairs[cleanId_(ptB) + "_" + cleanId_(ptA)] = true;
      var labelA = companyOf[ptA], labelB = companyOf[ptB];
      if (labelA && labelB) {
        if (labelA !== labelB) { for (var t in companyOf) { if (companyOf[t] === labelB) companyOf[t] = labelA; } }
      } else if (labelA) { companyOf[ptB] = labelA; }
      else if (labelB) { companyOf[ptA] = labelB; }
      else { var lbl = ptA.replace(/-.*/, ''); companyOf[ptA] = lbl; companyOf[ptB] = lbl; }
    }
  }

  // Group by credit rating
  var groups = {};
  for (var i = 1; i < data.length; i++) {
    var ticker = String(data[i][0]).trim();
    var coupon = data[i][2];
    var curYield = data[i][3];
    var rating = String(data[i][4]).trim();
    if (!ticker || !rating) continue;
    if (coupon === "" || coupon === null || coupon === undefined) continue;
    if (curYield === "" || curYield === null || curYield === undefined) continue;
    var upper = ticker.toUpperCase();
    if (!companyOf[upper]) companyOf[upper] = upper.replace(/-.*/, '');
    if (!groups[rating]) groups[rating] = [];
    groups[rating].push(ticker);
  }

  // Generate combinations — no cap
  var allPairs = [];
  for (var rating in groups) {
    var tickers = groups[rating];
    if (tickers.length < 2) continue;
    tickers.sort();
    for (var a = 0; a < tickers.length; a++) {
      for (var b = a + 1; b < tickers.length; b++) {
        var tA = tickers[a], tB = tickers[b];
        if (intraPairs[cleanId_(tA) + "_" + cleanId_(tB)]) continue;
        var coA = companyOf[tA.toUpperCase()] || tA.toUpperCase().replace(/-.*/, '');
        var coB = companyOf[tB.toUpperCase()] || tB.toUpperCase().replace(/-.*/, '');
        if (coA === coB) continue;
        allPairs.push([tA + "|" + tB, tA, tB, "CreditArb:" + rating]);
      }
    }
  }

  // Write to CreditPairs sheet
  var cpSheet = getOrCreateSheet_(ss, 'CreditPairs');
  cpSheet.getRange(1, 1, 1, 4).setValues([['PairID', 'TickerA', 'TickerB', 'Sector']]);
  cpSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#2d1a2e').setFontColor('#ffffff');
  if (allPairs.length > 0) {
    cpSheet.getRange(2, 1, allPairs.length, 4).setValues(allPairs);
  }
  return allPairs.length;
}

// ============================================================
// ORIGINAL FUNCTIONS (still work for manual runs, now with batch writes)
// ============================================================

function setupAll() {
  setupDashboard();
  generateCreditPairs();
  setupCreditSheets();
  setupTriggers();
}

// 1. INTRA-COMPANY SETUP
function setupDashboard() {
  var ss = SpreadsheetApp.getActive();
  var pairsSheet = ss.getSheetByName('Pairs');
  if (!pairsSheet) { showMsg_('ERROR: "Pairs" sheet not found.\n\nRequired: A: PairID | B: TickerA | C: TickerB | D: Sector'); return; }
  var masterSheet = ss.getSheetByName('Master');
  if (!masterSheet) { showMsg_('WARNING: "Master" sheet not found. Yields and coupon filtering will not work. Continuing...'); }
  var pairs = pairsSheet.getDataRange().getValues();
  var numPairs = pairs.length - 1;
  if (numPairs < 1) { showMsg_('ERROR: Pairs sheet is empty.'); return; }
  showMsg_('Setting up ' + numPairs + ' intra-company pairs.\nClick OK to proceed.');

  // --- LEVELS ---
  var levels = getOrCreateSheet_(ss, 'Levels');
  levels.getRange(1, 1, 1, 7).setValues([['PairID', 'Mean', 'StDev', 'Lower_5', 'Upper_95', 'HistCount', 'Historical Spread Data →']]);
  levels.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
  var statsF = [];
  for (var i = 0; i < numPairs; i++) {
    var r = i + 2;
    statsF.push([
      '=Pairs!A' + (i + 2),
      '=IFERROR(AVERAGE(G' + r + ':' + r + '), 0)',
      '=IFERROR(STDEV(G' + r + ':' + r + '), 0.001)',
      '=IFERROR(PERCENTILE(G' + r + ':' + r + ', 0.05), 0)',
      '=IFERROR(PERCENTILE(G' + r + ':' + r + ', 0.95), 0)',
      '=COUNTA(G' + r + ':' + r + ')'
    ]);
  }
  levels.getRange(2, 1, statsF.length, 6).setFormulas(statsF);

  // Historical formulas — BATCH (was individual loop!)
  var histF = [];
  for (var i = 0; i < numPairs; i++) {
    var tA = String(pairs[i + 1][1]).trim();
    var tB = String(pairs[i + 1][2]).trim();
    if (!tA || !tB) { histF.push(['']); continue; }
    histF.push([
      '=IFERROR(TRANSPOSE(ARRAYFORMULA(QUERY(GOOGLEFINANCE("' + tA + '","price",TODAY()-90,TODAY()),"select Col2 offset 1",0)-QUERY(GOOGLEFINANCE("' + tB + '","price",TODAY()-90,TODAY()),"select Col2 offset 1",0))),)'
    ]);
  }
  levels.getRange(2, 7, histF.length, 1).setFormulas(histF);
  levels.setFrozenRows(1);

  // --- LIVE ---
  buildLiveSheet_(ss, 'Live', 'Pairs', 'Levels', numPairs, pairs);

  // --- SUPPORTING SHEETS ---
  ensureSheet_(ss, 'OpenTrades', ['PairID', 'EntryZ', 'CostA', 'CostB', 'SizeA', 'SizeB', 'Timestamp']);
  ensureSheet_(ss, 'ClosedTrades', ['PairID', 'EntryZ', 'CostA', 'CostB', 'SizeA', 'SizeB', 'OpenDate', 'CloseDate', 'PnL']);
  ensureSheet_(ss, 'AlertsLog', ['Timestamp', 'PairID', 'Z-Score', 'Spread']);

  var ageSheet = ss.getSheetByName('ZScoreAge');
  if (!ageSheet) {
    ageSheet = ss.insertSheet('ZScoreAge');
    ageSheet.getRange(1, 1, 1, 3).setValues([['PairID', 'FirstCrossTimestamp', 'Source']]);
    ageSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  }

  SpreadsheetApp.flush();
  showMsg_('Intra-company setup complete for ' + numPairs + ' pairs.\n\nNext: run generateCreditPairs() then setupCreditSheets().');
}

// 2. CREDIT PAIR GENERATION (interactive version)
function generateCreditPairs() {
  var ss = SpreadsheetApp.getActive();
  var master = ss.getSheetByName('Master');
  if (!master) { showMsg_('ERROR: "Master" sheet not found.'); return; }
  var data = master.getDataRange().getValues();
  if (data.length < 2) { showMsg_('ERROR: Master sheet is empty.'); return; }

  var companyOf = {};
  var intraPairs = {};
  var pairsSheet = ss.getSheetByName('Pairs');
  if (pairsSheet) {
    var pData = pairsSheet.getDataRange().getValues();
    for (var i = 1; i < pData.length; i++) {
      var ptA = String(pData[i][1]).trim().toUpperCase();
      var ptB = String(pData[i][2]).trim().toUpperCase();
      if (!ptA || !ptB) continue;
      intraPairs[cleanId_(ptA) + "_" + cleanId_(ptB)] = true;
      intraPairs[cleanId_(ptB) + "_" + cleanId_(ptA)] = true;
      var labelA = companyOf[ptA], labelB = companyOf[ptB];
      if (labelA && labelB) {
        if (labelA !== labelB) { for (var t in companyOf) { if (companyOf[t] === labelB) companyOf[t] = labelA; } }
      } else if (labelA) { companyOf[ptB] = labelA; }
      else if (labelB) { companyOf[ptA] = labelB; }
      else { var lbl = ptA.replace(/-.*/, ''); companyOf[ptA] = lbl; companyOf[ptB] = lbl; }
    }
  }

  var groups = {};
  for (var i = 1; i < data.length; i++) {
    var ticker = String(data[i][0]).trim();
    var coupon = data[i][2]; var curYield = data[i][3]; var rating = String(data[i][4]).trim();
    if (!ticker || !rating) continue;
    if (coupon === "" || coupon === null || coupon === undefined) continue;
    if (curYield === "" || curYield === null || curYield === undefined) continue;
    var upper = ticker.toUpperCase();
    if (!companyOf[upper]) companyOf[upper] = upper.replace(/-.*/, '');
    if (!groups[rating]) groups[rating] = [];
    groups[rating].push(ticker);
  }

  var allPairs = [];
  for (var rating in groups) {
    var tickers = groups[rating];
    if (tickers.length < 2) continue;
    tickers.sort();
    for (var a = 0; a < tickers.length; a++) {
      for (var b = a + 1; b < tickers.length; b++) {
        var tA = tickers[a], tB = tickers[b];
        if (intraPairs[cleanId_(tA) + "_" + cleanId_(tB)]) continue;
        var coA = companyOf[tA.toUpperCase()] || tA.toUpperCase().replace(/-.*/, '');
        var coB = companyOf[tB.toUpperCase()] || tB.toUpperCase().replace(/-.*/, '');
        if (coA === coB) continue;
        allPairs.push([tA + "|" + tB, tA, tB, "CreditArb:" + rating]);
      }
    }
  }

  var cpSheet = getOrCreateSheet_(ss, 'CreditPairs');
  cpSheet.getRange(1, 1, 1, 4).setValues([['PairID', 'TickerA', 'TickerB', 'Sector']]);
  cpSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#2d1a2e').setFontColor('#ffffff');
  if (allPairs.length > 0) { cpSheet.getRange(2, 1, allPairs.length, 4).setValues(allPairs); }

  var summary = [];
  var ratingCounts = {};
  allPairs.forEach(function(p) { var r = p[3].replace('CreditArb:', ''); ratingCounts[r] = (ratingCounts[r] || 0) + 1; });
  for (var r in ratingCounts) summary.push(r + ': ' + ratingCounts[r]);
  showMsg_('Generated ' + allPairs.length + ' credit arb pairs.\n\nBreakdown:\n' + summary.join('\n') + '\n\nNext: run setupCreditSheets().');
}

// 3. CREDIT SHEETS SETUP
function setupCreditSheets() {
  var ss = SpreadsheetApp.getActive();
  var cpSheet = ss.getSheetByName('CreditPairs');
  if (!cpSheet) { showMsg_('ERROR: "CreditPairs" sheet not found. Run generateCreditPairs() first.'); return; }
  var pairs = cpSheet.getDataRange().getValues();
  var numPairs = pairs.length - 1;
  if (numPairs < 1) { showMsg_('ERROR: CreditPairs is empty.'); return; }

  var cLevels = getOrCreateSheet_(ss, 'CreditLevels');
  cLevels.getRange(1, 1, 1, 7).setValues([['PairID', 'Mean', 'StDev', 'Lower_5', 'Upper_95', 'HistCount', 'Historical Spread Data →']]);
  cLevels.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#2d1a2e').setFontColor('#ffffff');
  var statsF = [];
  for (var i = 0; i < numPairs; i++) {
    var r = i + 2;
    statsF.push([
      '=CreditPairs!A' + (i + 2),
      '=IFERROR(AVERAGE(G' + r + ':' + r + '), 0)',
      '=IFERROR(STDEV(G' + r + ':' + r + '), 0.001)',
      '=IFERROR(PERCENTILE(G' + r + ':' + r + ', 0.05), 0)',
      '=IFERROR(PERCENTILE(G' + r + ':' + r + ', 0.95), 0)',
      '=COUNTA(G' + r + ':' + r + ')'
    ]);
  }
  cLevels.getRange(2, 1, statsF.length, 6).setFormulas(statsF);

  // Historical formulas — BATCH
  var histF = [];
  for (var i = 0; i < numPairs; i++) {
    var tA = String(pairs[i + 1][1]).trim();
    var tB = String(pairs[i + 1][2]).trim();
    if (!tA || !tB) { histF.push(['']); continue; }
    histF.push([
      '=IFERROR(TRANSPOSE(ARRAYFORMULA(QUERY(GOOGLEFINANCE("' + tA + '","price",TODAY()-90,TODAY()),"select Col2 offset 1",0)-QUERY(GOOGLEFINANCE("' + tB + '","price",TODAY()-90,TODAY()),"select Col2 offset 1",0))),)'
    ]);
  }
  cLevels.getRange(2, 7, histF.length, 1).setFormulas(histF);
  cLevels.setFrozenRows(1);

  buildLiveSheet_(ss, 'CreditLive', 'CreditPairs', 'CreditLevels', numPairs, pairs);
  SpreadsheetApp.flush();
  showMsg_('Credit sheets built!\n\n• CreditLevels: ' + numPairs + ' pairs\n• CreditLive: ' + numPairs + ' pairs\n\nGOOGLEFINANCE may take 30-60s to populate.');
}

// 4. TRIGGERS
function setupTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === 'snapshotZScores' || fn === 'dailyCreditRefresh') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('snapshotZScores').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('dailyCreditRefresh').timeBased().atHour(5).everyDays(1).create();
  showMsg_('Triggers installed!\n\n• snapshotZScores(): Every hour\n• dailyCreditRefresh(): Daily at 5 AM');
}

// ============================================================
// 5. SNAPSHOT + AGE TRACKING (Hourly Trigger)
// ============================================================
function snapshotZScores() {
  try {
    var ss = SpreadsheetApp.getActive();
    var now = new Date();
    var sheets = [
      { name: 'Live', source: 'intra' },
      { name: 'CreditLive', source: 'credit' }
    ];
    var logSheet = ss.getSheetByName('AlertsLog');
    var ageSheet = ss.getSheetByName('ZScoreAge');
    if (!logSheet || !ageSheet) return;

    var ageData = ageSheet.getDataRange().getValues();
    var ageMap = {};
    for (var a = 1; a < ageData.length; a++) {
      var aid = cleanId_(String(ageData[a][0]));
      if (aid) { ageMap[aid] = { row: a + 1, timestamp: ageData[a][1], source: ageData[a][2] || '' }; }
    }

    var logRows = [];
    var ageUpdates = [];
    for (var s = 0; s < sheets.length; s++) {
      var liveSheet = ss.getSheetByName(sheets[s].name);
      if (!liveSheet) continue;
      var data = liveSheet.getDataRange().getValues();
      if (data.length <= 1) continue;
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var pairId = row[0];
        if (!pairId) continue;
        var priceA = parseFloat(row[3]) || 0;
        var priceB = parseFloat(row[4]) || 0;
        if (priceA <= 0 || priceB <= 0) continue;
        var histCount = parseFloat(row[16]) || 0;
        if (histCount < 60) continue;
        var couponA = row[8], couponB = row[9];
        if (couponA === "" || couponA === null || couponB === "" || couponB === null) continue;
        var zScore = parseFloat(row[12]) || 0;
        var spread = parseFloat(row[5]) || 0;
        var cid = cleanId_(String(pairId));
        logRows.push([now, String(pairId), zScore, spread]);
        var isActive = Math.abs(zScore) >= 1.5;
        var existing = ageMap[cid];
        if (isActive && !existing) {
          ageUpdates.push({ action: 'add', pairId: String(pairId), cleanId: cid, source: sheets[s].source });
        } else if (!isActive && existing) {
          ageUpdates.push({ action: 'remove', cleanId: cid, row: existing.row });
        }
      }
    }

    if (logRows.length > 0) { logSheet.getRange(logSheet.getLastRow() + 1, 1, logRows.length, 4).setValues(logRows); }

    var rowsToDelete = [];
    for (var u = 0; u < ageUpdates.length; u++) { if (ageUpdates[u].action === 'remove') rowsToDelete.push(ageUpdates[u].row); }
    rowsToDelete.sort(function(a, b) { return b - a; });
    for (var d = 0; d < rowsToDelete.length; d++) { ageSheet.deleteRow(rowsToDelete[d]); }

    var newAgeRows = [];
    for (var u = 0; u < ageUpdates.length; u++) { if (ageUpdates[u].action === 'add') newAgeRows.push([ageUpdates[u].pairId, now, ageUpdates[u].source]); }
    if (newAgeRows.length > 0) { ageSheet.getRange(ageSheet.getLastRow() + 1, 1, newAgeRows.length, 3).setValues(newAgeRows); }

    var totalRows = logSheet.getLastRow();
    if (totalRows > 5500) { logSheet.deleteRows(2, totalRows - 5000); }
  } catch (e) {
    console.error("snapshotZScores error: " + e);
  }
}

// ============================================================
// 6. DAILY CREDIT REFRESH (Daily Trigger) — now with batch writes
// ============================================================
function dailyCreditRefresh() {
  try {
    var ss = SpreadsheetApp.getActive();
    var master = ss.getSheetByName('Master');
    if (!master) return;
    var data = master.getDataRange().getValues();
    if (data.length < 2) return;

    var companyOf = {};
    var intraPairs = {};
    var pairsSheet = ss.getSheetByName('Pairs');
    if (pairsSheet) {
      var pData = pairsSheet.getDataRange().getValues();
      for (var i = 1; i < pData.length; i++) {
        var ptA = String(pData[i][1]).trim().toUpperCase();
        var ptB = String(pData[i][2]).trim().toUpperCase();
        if (!ptA || !ptB) continue;
        intraPairs[cleanId_(ptA) + "_" + cleanId_(ptB)] = true;
        intraPairs[cleanId_(ptB) + "_" + cleanId_(ptA)] = true;
        var labelA = companyOf[ptA], labelB = companyOf[ptB];
        if (labelA && labelB) {
          if (labelA !== labelB) { for (var t in companyOf) { if (companyOf[t] === labelB) companyOf[t] = labelA; } }
        } else if (labelA) { companyOf[ptB] = labelA; }
        else if (labelB) { companyOf[ptA] = labelB; }
        else { var lbl = ptA.replace(/-.*/, ''); companyOf[ptA] = lbl; companyOf[ptB] = lbl; }
      }
    }

    var groups = {};
    for (var i = 1; i < data.length; i++) {
      var ticker = String(data[i][0]).trim();
      var coupon = data[i][2]; var curYield = data[i][3]; var rating = String(data[i][4]).trim();
      if (!ticker || !rating) continue;
      if (coupon === "" || coupon === null || coupon === undefined) continue;
      if (curYield === "" || curYield === null || curYield === undefined) continue;
      var upper = ticker.toUpperCase();
      if (!companyOf[upper]) companyOf[upper] = upper.replace(/-.*/, '');
      if (!groups[rating]) groups[rating] = [];
      groups[rating].push(ticker);
    }

    var allPairs = [];
    for (var rating in groups) {
      var tickers = groups[rating];
      if (tickers.length < 2) continue;
      tickers.sort();
      for (var a = 0; a < tickers.length; a++) {
        for (var b = a + 1; b < tickers.length; b++) {
          var tA = tickers[a], tB = tickers[b];
          if (intraPairs[cleanId_(tA) + "_" + cleanId_(tB)]) continue;
          var coA = companyOf[tA.toUpperCase()] || tA.toUpperCase().replace(/-.*/, '');
          var coB = companyOf[tB.toUpperCase()] || tB.toUpperCase().replace(/-.*/, '');
          if (coA === coB) continue;
          allPairs.push([tA + "|" + tB, tA, tB, "CreditArb:" + rating]);
        }
      }
    }

    var cpSheet = getOrCreateSheet_(ss, 'CreditPairs');
    cpSheet.getRange(1, 1, 1, 4).setValues([['PairID', 'TickerA', 'TickerB', 'Sector']]);
    if (allPairs.length > 0) { cpSheet.getRange(2, 1, allPairs.length, 4).setValues(allPairs); }

    var numPairs = allPairs.length;
    if (numPairs < 1) return;

    // CreditLevels — batch writes
    var cLevels = getOrCreateSheet_(ss, 'CreditLevels');
    cLevels.getRange(1, 1, 1, 7).setValues([['PairID', 'Mean', 'StDev', 'Lower_5', 'Upper_95', 'HistCount', 'Historical →']]);
    var sf = [];
    for (var i = 0; i < numPairs; i++) {
      var r = i + 2;
      sf.push([
        '=CreditPairs!A' + (i + 2),
        '=IFERROR(AVERAGE(G' + r + ':' + r + '),0)',
        '=IFERROR(STDEV(G' + r + ':' + r + '),0.001)',
        '=IFERROR(PERCENTILE(G' + r + ':' + r + ',0.05),0)',
        '=IFERROR(PERCENTILE(G' + r + ':' + r + ',0.95),0)',
        '=COUNTA(G' + r + ':' + r + ')'
      ]);
    }
    cLevels.getRange(2, 1, sf.length, 6).setFormulas(sf);

    // Historical formulas — BATCH
    var histF = [];
    for (var i = 0; i < numPairs; i++) {
      var tA = String(allPairs[i][1]).trim();
      var tB = String(allPairs[i][2]).trim();
      if (!tA || !tB) { histF.push(['']); continue; }
      histF.push([
        '=IFERROR(TRANSPOSE(ARRAYFORMULA(QUERY(GOOGLEFINANCE("' + tA + '","price",TODAY()-90,TODAY()),"select Col2 offset 1",0)-QUERY(GOOGLEFINANCE("' + tB + '","price",TODAY()-90,TODAY()),"select Col2 offset 1",0))),)'
      ]);
    }
    cLevels.getRange(2, 7, histF.length, 1).setFormulas(histF);

    buildLiveSheet_(ss, 'CreditLive', 'CreditPairs', 'CreditLevels', numPairs, [['h']].concat(allPairs));
    SpreadsheetApp.flush();
  } catch (e) {
    console.error("dailyCreditRefresh error: " + e);
  }
}

// ============================================================
// SHARED HELPERS
// ============================================================
function buildLiveSheet_(ss, sheetName, pairsRef, levelsRef, numPairs, pairsData) {
  var live = getOrCreateSheet_(ss, sheetName);
  var headers = [
    'PairID', 'TickerA', 'TickerB', 'PriceA', 'PriceB', 'Spread',
    'YieldA', 'YieldB', 'CouponA', 'CouponB',
    'Mean', 'StDev', 'Z-Score', 'Lower', 'Upper',
    'Sector', 'HistCount',
    'AvgLiqA', 'AvgLiqB', 'AvgLiq', 'CurVolA', 'CurVolB', 'CurVol', 'VolSpike'
  ];
  live.getRange(1, 1, 1, headers.length).setValues([headers]);
  live.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
  var formulas = [];
  for (var i = 0; i < numPairs; i++) {
    var r = i + 2;
    var p = i + 2;
    formulas.push([
      '=' + pairsRef + '!A' + p,
      '=' + pairsRef + '!B' + p,
      '=' + pairsRef + '!C' + p,
      '=IFERROR(GOOGLEFINANCE(B' + r + '),0)',
      '=IFERROR(GOOGLEFINANCE(C' + r + '),0)',
      '=D' + r + '-E' + r,
      '=IFERROR(INDEX(Master!D:D,MATCH(B' + r + ',Master!A:A,0)),"")',
      '=IFERROR(INDEX(Master!D:D,MATCH(C' + r + ',Master!A:A,0)),"")',
      '=IFERROR(INDEX(Master!C:C,MATCH(B' + r + ',Master!A:A,0)),"")',
      '=IFERROR(INDEX(Master!C:C,MATCH(C' + r + ',Master!A:A,0)),"")',
      '=IFERROR(INDEX(' + levelsRef + '!B:B,MATCH(A' + r + ',' + levelsRef + '!A:A,0)),0)',
      '=IFERROR(INDEX(' + levelsRef + '!C:C,MATCH(A' + r + ',' + levelsRef + '!A:A,0)),0.001)',
      '=IF(L' + r + '<=0.001,0,(F' + r + '-K' + r + ')/L' + r + ')',
      '=IFERROR(INDEX(' + levelsRef + '!D:D,MATCH(A' + r + ',' + levelsRef + '!A:A,0)),0)',
      '=IFERROR(INDEX(' + levelsRef + '!E:E,MATCH(A' + r + ',' + levelsRef + '!A:A,0)),0)',
      '=' + pairsRef + '!D' + p,
      '=IFERROR(INDEX(' + levelsRef + '!F:F,MATCH(A' + r + ',' + levelsRef + '!A:A,0)),0)',
      '=IFERROR(GOOGLEFINANCE(B' + r + ',"volumeavg"),0)',
      '=IFERROR(GOOGLEFINANCE(C' + r + ',"volumeavg"),0)',
      '=(R' + r + '+S' + r + ')/2',
      '=IFERROR(GOOGLEFINANCE(B' + r + ',"volume"),0)',
      '=IFERROR(GOOGLEFINANCE(C' + r + ',"volume"),0)',
      '=(U' + r + '+V' + r + ')/2',
      '=IF(AND(T' + r + '>0,W' + r + '>1.5*T' + r + '),TRUE,FALSE)'
    ]);
  }
  live.getRange(2, 1, formulas.length, 24).setFormulas(formulas);
  live.setFrozenRows(1);
}
function getOrCreateSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (sheet) { sheet.clear(); return sheet; }
  return ss.insertSheet(name);
}
function ensureSheet_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
}
function cleanId_(id) {
  return id ? String(id).toUpperCase().replace(/[^A-Z0-9]/g, '') : "";
}
function showMsg_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}
