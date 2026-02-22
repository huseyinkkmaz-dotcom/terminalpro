/**
 * TERMINAL PRO — V23.1 SETUP SCRIPT (TIME-AWARE BATCHING)
 *
 * KEY CHANGE FROM V23.0:
 *   - All individual setFormula() loops replaced with batch setFormulas() calls
 *   - setupAllBatched() uses PropertiesService to save/resume progress
 *   - createAutoTrigger() installs ALL recurring triggers
 *   - updateLivePrices() snapshots GOOGLEFINANCE values to WebCache every 10 min
 *   - setupAllBatched() is a ONE-TIME heavy build (not recurring)
 *   - Frontend reads from WebCache — fast, no formula recalculation
 *
 * HOW TO USE:
 *   1. Run createAutoTrigger() once — installs all triggers (10-min, hourly, daily)
 *   2. Run setupAllBatched() once  — builds all sheets (resumes if it times out)
 *   3. Close your laptop. Check Execution Log for progress.
 *   4. Once "ALL PHASES COMPLETE" appears, WebCache is auto-populated.
 *   5. To force a full rebuild later: clearSetupState() then setupAllBatched().
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

  // ZOMBIE GUARD: If no saved state and Live sheet already has data,
  // setup was already completed. Don't restart from phase 0 (which wipes sheets).
  if (!stateJson) {
    var existingLive = ss.getSheetByName('Live');
    if (existingLive && existingLive.getLastRow() > 1) {
      Logger.log('setupAllBatched: Live sheet already built (' + (existingLive.getLastRow() - 1) + ' rows). Skipping — setup already complete.');
      Logger.log('To force a full rebuild: run clearSetupState() first, then setupAllBatched().');
      return;
    }
  }

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
      SpreadsheetApp.flush();

      // Historical spread formulas — BATCH (was individual setFormula per row!)
      var histF = [];
      for (var i = 0; i < numPairs; i++) {
        var tA = String(pairs[i + 1][1]).trim();
        var tB = String(pairs[i + 1][2]).trim();
        if (!tA || !tB) { histF.push(['']); continue; }
        histF.push([
          '=IFERROR(TRANSPOSE(ARRAYFORMULA(QUERY(GOOGLEFINANCE("' + tA + '","price",TODAY()-120,TODAY()),"select Col2 offset 1",0)-QUERY(GOOGLEFINANCE("' + tB + '","price",TODAY()-120,TODAY()),"select Col2 offset 1",0))),)'
        ]);
      }
      levels.getRange(2, 7, histF.length, 1).setFormulas(histF);
      SpreadsheetApp.flush();
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
      ensureSheet_(ss, 'DivDates', ['Ticker', 'ExDivDate', 'LastFetched']);
      ensureSheet_(ss, 'Watchlist', ['PairID', 'AddedDate', 'AddedZ', 'Mode']);
      ensureSheet_(ss, 'TreasuryHist', ['Date', 'US2Y', 'US5Y', 'US7Y', 'US10Y', 'US30Y']);
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
        SpreadsheetApp.flush();

        // Historical spread formulas — BATCH
        var histF = [];
        for (var i = 0; i < numPairs; i++) {
          var tA = String(pairs[i + 1][1]).trim();
          var tB = String(pairs[i + 1][2]).trim();
          if (!tA || !tB) { histF.push(['']); continue; }
          histF.push([
            '=IFERROR(TRANSPOSE(ARRAYFORMULA(QUERY(GOOGLEFINANCE("' + tA + '","price",TODAY()-120,TODAY()),"select Col2 offset 1",0)-QUERY(GOOGLEFINANCE("' + tB + '","price",TODAY()-120,TODAY()),"select Col2 offset 1",0))),)'
          ]);
        }
        cLevels.getRange(2, 7, histF.length, 1).setFormulas(histF);
        SpreadsheetApp.flush();
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

    // --- PHASE 6: Finalize + populate WebCache ---
    if (state.phase === 6) {
      Logger.log('Phase 6: Finalizing...');
      // Install triggers only if they don't already exist (createAutoTrigger may have set them)
      var existingTriggers = {};
      var triggers = ScriptApp.getProjectTriggers();
      for (var i = 0; i < triggers.length; i++) {
        existingTriggers[triggers[i].getHandlerFunction()] = true;
      }
      if (!existingTriggers['snapshotZScores']) {
        ScriptApp.newTrigger('snapshotZScores').timeBased().everyHours(1).create();
        Logger.log('Installed snapshotZScores trigger.');
      }
      if (!existingTriggers['dailyCreditRefresh']) {
        ScriptApp.newTrigger('dailyCreditRefresh').timeBased().atHour(5).everyDays(1).create();
        Logger.log('Installed dailyCreditRefresh trigger.');
      }
      if (!existingTriggers['updateLivePrices']) {
        ScriptApp.newTrigger('updateLivePrices').timeBased().everyMinutes(10).create();
        Logger.log('Installed updateLivePrices trigger.');
      }
      // Populate WebCache immediately so frontend has data right away
      SpreadsheetApp.flush();
      Utilities.sleep(3000); // Brief pause for GOOGLEFINANCE formula recalculation
      try {
        updateLivePrices();
        Logger.log('Phase 6 complete: WebCache populated.');
      } catch (cacheErr) {
        Logger.log('Phase 6: WebCache population failed (GOOGLEFINANCE may still be loading): ' + cacheErr.toString());
        Logger.log('updateLivePrices will retry on the next 10-minute trigger cycle.');
      }

      state.phase = 7;
    }

    // --- PHASE 7: DONE ---
    if (state.phase >= 7) {
      props.deleteProperty('SETUP_STATE');
      SpreadsheetApp.flush();
      Logger.log('=== ALL PHASES COMPLETE === Setup finished successfully!');
      Logger.log('Triggers active: updateLivePrices (10 min), snapshotZScores (1 hr), dailyCreditRefresh (5 AM).');
      showMsg_('Setup complete! All sheets built.\n\nWebCache is populated — your dashboard should show data now.\nTriggers are active for automatic updates.');
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
 * Installs ALL recurring triggers. Run this ONCE.
 *
 * Schedule:
 *   - updateLivePrices:    every 10 min  (light — snapshot GOOGLEFINANCE → WebCache)
 *   - snapshotZScores:     every 1 hour  (trend ribbons + age tracking)
 *   - dailyCreditRefresh:  daily at 5 AM (regenerate credit pairs)
 *
 * setupAllBatched is NOT on a recurring trigger — it's a one-time build.
 * To rebuild: run clearSetupState() then setupAllBatched() manually.
 */
function createAutoTrigger() {
  // Remove ALL existing triggers managed by this system
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === 'setupAllBatched' || fn === 'updateLivePrices' ||
        fn === 'snapshotZScores' || fn === 'dailyCreditRefresh' || fn === 'fetchDividendDates' ||
        fn === 'dailyMacroRefresh' || fn === 'computeMacroValuationsBatch') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // LIGHT: updateLivePrices every 10 minutes (snapshots GOOGLEFINANCE → WebCache)
  ScriptApp.newTrigger('updateLivePrices')
    .timeBased()
    .everyMinutes(10)
    .create();

  // HOURLY: snapshotZScores (trend ribbons + Z-score age tracking)
  ScriptApp.newTrigger('snapshotZScores')
    .timeBased()
    .everyHours(1)
    .create();

  // DAILY: dailyCreditRefresh at 5 AM (regenerate credit pairs from Master)
  ScriptApp.newTrigger('dailyCreditRefresh')
    .timeBased()
    .atHour(5)
    .everyDays(1)
    .create();

  // DAILY: fetchDividendDates at 6 AM (scrape Yahoo Finance for ex-div dates)
  ScriptApp.newTrigger('fetchDividendDates')
    .timeBased()
    .atHour(6)
    .everyDays(1)
    .create();

  // DAILY: dailyMacroRefresh at 7 AM (macro valuation: preferreds vs treasuries)
  ScriptApp.newTrigger('dailyMacroRefresh')
    .timeBased()
    .atHour(7)
    .everyDays(1)
    .create();

  Logger.log('All triggers installed:\n• updateLivePrices: every 10 min\n• snapshotZScores: every hour\n• dailyCreditRefresh: daily 5 AM\n• fetchDividendDates: daily 6 AM\n• dailyMacroRefresh: daily 7 AM');
  showMsg_(
    'Triggers installed!\n\n' +
    '• updateLivePrices: every 10 min (snapshots prices to WebCache)\n' +
    '• snapshotZScores: every hour (trend tracking)\n' +
    '• dailyCreditRefresh: daily 5 AM (credit pair regeneration)\n' +
    '• fetchDividendDates: daily 6 AM (Yahoo Finance ex-div dates)\n' +
    '• dailyMacroRefresh: daily 7 AM (macro valuation vs treasuries)\n\n' +
    'Now run setupAllBatched() to build the sheets.\n' +
    'Then run updateLivePrices() to populate WebCache immediately.'
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

/** Reset saved progress — use if setup gets stuck or you want to force a full rebuild */
function clearSetupState() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty('SETUP_STATE');
  props.deleteProperty('WEBCACHE_UPDATED');
  removeAutoSetupTrigger_();
  Logger.log('Setup state cleared. Run createAutoTrigger() + setupAllBatched() to start fresh.');
  showMsg_('Setup state cleared. You can start fresh.\n\nNext: run createAutoTrigger() then setupAllBatched().');
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
// ONE-CLICK: Update historical lookback from 90 → 120 days
// ============================================================

/**
 * Updates GOOGLEFINANCE historical formulas in Levels and CreditLevels
 * from 90-day to 120-day lookback, IN-PLACE (no full rebuild needed).
 *
 * Run this ONCE after deploying the updated SetupDashboard.gs.
 * It rewrites column G formulas, then GOOGLEFINANCE recalculates (~30-60s),
 * then HistCount will increase from ~59 to ~85.
 *
 * After running: wait 60s, then run updateLivePrices() to refresh WebCache.
 */
function updateHistoricalLookback() {
  var ss = SpreadsheetApp.getActive();
  var updated = 0;

  var targets = [
    { levelsName: 'Levels', pairsName: 'Pairs' },
    { levelsName: 'CreditLevels', pairsName: 'CreditPairs' }
  ];

  for (var t = 0; t < targets.length; t++) {
    var levelsSheet = ss.getSheetByName(targets[t].levelsName);
    var pairsSheet = ss.getSheetByName(targets[t].pairsName);
    if (!levelsSheet || !pairsSheet) {
      Logger.log('updateHistoricalLookback: ' + targets[t].levelsName + ' or ' + targets[t].pairsName + ' not found, skipping.');
      continue;
    }

    var pairs = pairsSheet.getDataRange().getValues();
    var numPairs = pairs.length - 1;
    if (numPairs < 1) continue;

    // Rebuild column G formulas with 120-day lookback
    var histF = [];
    for (var i = 0; i < numPairs; i++) {
      var tA = String(pairs[i + 1][1]).trim();
      var tB = String(pairs[i + 1][2]).trim();
      if (!tA || !tB) { histF.push(['']); continue; }
      histF.push([
        '=IFERROR(TRANSPOSE(ARRAYFORMULA(QUERY(GOOGLEFINANCE("' + tA + '","price",TODAY()-120,TODAY()),"select Col2 offset 1",0)-QUERY(GOOGLEFINANCE("' + tB + '","price",TODAY()-120,TODAY()),"select Col2 offset 1",0))),)'
      ]);
    }

    levelsSheet.getRange(2, 7, histF.length, 1).setFormulas(histF);
    SpreadsheetApp.flush();
    updated += numPairs;
    Logger.log('updateHistoricalLookback: Updated ' + numPairs + ' formulas in ' + targets[t].levelsName + ' (90d → 120d).');
  }

  if (updated > 0) {
    Logger.log('updateHistoricalLookback: Done. ' + updated + ' total formulas updated.');
    Logger.log('GOOGLEFINANCE will recalculate in ~30-60 seconds.');
    Logger.log('After that, run updateLivePrices() to refresh WebCache.');
    showMsg_(
      'Historical lookback updated to 120 days!\n\n' +
      updated + ' formulas rewritten.\n\n' +
      'Wait 60 seconds for GOOGLEFINANCE to recalculate,\n' +
      'then run updateLivePrices() to refresh the cache.'
    );
  } else {
    showMsg_('No Levels/CreditLevels sheets found to update.');
  }
}

// ============================================================
// DIVIDEND DATE FETCHING — Yahoo Finance scraping
// ============================================================

/**
 * Convert our ticker format to Yahoo Finance format.
 * BAC-B → BAC-PB (insert P before series letter)
 * AGNCP → AGNCP (non-hyphenated: as-is)
 */
function toYahooTicker_(ticker) {
  var t = String(ticker).trim().toUpperCase();
  var idx = t.indexOf('-');
  if (idx > 0) {
    return t.substring(0, idx) + '-P' + t.substring(idx + 1);
  }
  return t;
}

/**
 * Gets Yahoo Finance API authentication (cookies + crumb).
 * Required for v7/finance/quote and v10/finance/quoteSummary endpoints.
 */
function getYahooCrumb_() {
  // Step 1: Get cookies from Yahoo
  var resp = UrlFetchApp.fetch('https://fc.yahoo.com/', {
    muteHttpExceptions: true,
    followRedirects: true
  });
  var headers = resp.getAllHeaders();
  var cookies = '';
  var setCookies = headers['Set-Cookie'];
  if (setCookies) {
    if (!Array.isArray(setCookies)) setCookies = [setCookies];
    cookies = setCookies.map(function(c) { return c.split(';')[0]; }).join('; ');
  }

  // Step 2: Get crumb using cookies
  var crumbResp = UrlFetchApp.fetch('https://query2.finance.yahoo.com/v1/test/getcrumb', {
    headers: {
      'Cookie': cookies,
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
    },
    muteHttpExceptions: true
  });

  var crumb = crumbResp.getContentText().trim();
  Logger.log('getYahooCrumb_: cookies=' + (cookies ? 'yes' : 'no') + ', crumb=' + (crumb ? crumb.substring(0, 6) + '...' : 'EMPTY'));
  return { cookies: cookies, crumb: crumb };
}

/**
 * Fetches dividend dates for all tickers in Master sheet.
 * Writes results to DivDates sheet. Run daily via trigger at 6 AM.
 *
 * Two-phase approach:
 *   Phase 1: v7/finance/quote batch (50/request, fast) — uses dividendDate field
 *   Phase 2: v8/finance/chart fallback (no auth needed) — for tickers Phase 1 missed
 *
 * If the date is in the past, adds 91 days until it's >= today (quarterly projection).
 * Skips tickers fetched within last 20 hours.
 */
function fetchDividendDates() {
  var ss = SpreadsheetApp.getActive();
  var master = ss.getSheetByName('Master');
  if (!master) { Logger.log('fetchDividendDates: No Master sheet'); return; }

  var data = master.getDataRange().getValues();
  var tickers = [];
  for (var i = 1; i < data.length; i++) {
    var t = String(data[i][0]).trim();
    if (t) tickers.push(t);
  }
  if (tickers.length === 0) return;

  // Get or create DivDates sheet
  var divSheet = ss.getSheetByName('DivDates');
  if (!divSheet) {
    divSheet = ss.insertSheet('DivDates');
    divSheet.getRange(1, 1, 1, 3).setValues([['Ticker', 'NextDivDate', 'LastFetched']]);
    divSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  }

  // Load existing data — skip recently fetched tickers
  var dataMap = {};
  if (divSheet.getLastRow() > 1) {
    var existData = divSheet.getRange(2, 1, divSheet.getLastRow() - 1, 3).getValues();
    for (var i = 0; i < existData.length; i++) {
      var key = String(existData[i][0]).toUpperCase().trim();
      dataMap[key] = existData[i];
    }
  }

  var now = new Date();
  var toFetch = [];
  for (var i = 0; i < tickers.length; i++) {
    var key = tickers[i].toUpperCase();
    var existing = dataMap[key];
    if (existing && existing[2] instanceof Date) {
      var hoursSince = (now.getTime() - existing[2].getTime()) / (1000 * 60 * 60);
      if (hoursSince < 20) continue;
    }
    toFetch.push(tickers[i]);
  }

  Logger.log('fetchDividendDates: ' + toFetch.length + ' to fetch, ' + (tickers.length - toFetch.length) + ' cached.');
  if (toFetch.length === 0) return;

  // ── Phase 1: v7/finance/quote batch (50 tickers per request) ──
  var auth = getYahooCrumb_();
  if (!auth.crumb) {
    Logger.log('fetchDividendDates: Failed to get Yahoo crumb. Skipping Phase 1.');
  }

  var tickerMap = {};  // Yahoo symbol → our ticker
  for (var i = 0; i < toFetch.length; i++) {
    tickerMap[toYahooTicker_(toFetch[i]).toUpperCase()] = toFetch[i];
  }

  var successCount = 0;
  var BATCH = 50;

  if (auth.crumb) {
    for (var b = 0; b < toFetch.length; b += BATCH) {
      var batch = toFetch.slice(b, Math.min(b + BATCH, toFetch.length));
      var symbols = batch.map(function(t) { return toYahooTicker_(t); }).join(',');
      var url = 'https://query2.finance.yahoo.com/v7/finance/quote?symbols=' +
                encodeURIComponent(symbols) + '&crumb=' + encodeURIComponent(auth.crumb);

      try {
        var resp = UrlFetchApp.fetch(url, {
          headers: {
            'Cookie': auth.cookies,
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
          },
          muteHttpExceptions: true
        });

        if (resp.getResponseCode() !== 200) {
          Logger.log('Phase 1: Batch at offset ' + b + ' HTTP ' + resp.getResponseCode());
          continue;
        }

        var json = JSON.parse(resp.getContentText());
        var results = (json.quoteResponse && json.quoteResponse.result) || [];

        if (b === 0) {
          Logger.log('Phase 1: ' + results.length + '/' + batch.length + ' results.');
          if (results.length > 0) {
            var s = results[0];
            Logger.log('Sample: ' + s.symbol +
              ' exDividendDate=' + (s.exDividendDate || 'N/A') +
              ' (payDate=' + (s.dividendDate || 'N/A') + ', ignored)');
          }
        }

        for (var i = 0; i < results.length; i++) {
          var q = results[i];
          var yahooSym = String(q.symbol).toUpperCase();
          var ourTicker = tickerMap[yahooSym];
          if (!ourTicker) continue;
          var key = ourTicker.toUpperCase();

          // ONLY use exDividendDate — dividendDate is the PAY date, not ex-date
          var rawTs = (q.exDividendDate && q.exDividendDate > 0) ? q.exDividendDate : 0;

          if (rawTs > 0) {
            var d = projectNextDivDate_(new Date(rawTs * 1000), now);
            if (d) {
              dataMap[key] = [ourTicker, d, now];
              successCount++;
              continue;
            }
          }
          // Mark as checked but no date found (so Phase 2 can try)
          if (!dataMap[key] || !dataMap[key][1]) {
            dataMap[key] = [ourTicker, '', now];
          }
        }
      } catch (e) {
        Logger.log('Phase 1: Batch error at offset ' + b + ': ' + e.toString());
      }

      if (b + BATCH < toFetch.length) Utilities.sleep(500);
    }
  }

  Logger.log('Phase 1 done: ' + successCount + '/' + toFetch.length + ' got dates.');

  // ── Phase 2: v8/finance/chart fallback for misses (NO auth needed) ──
  var misses = [];
  for (var i = 0; i < toFetch.length; i++) {
    var key = toFetch[i].toUpperCase();
    if (!dataMap[key] || !(dataMap[key][1] instanceof Date)) {
      misses.push(toFetch[i]);
    }
  }

  if (misses.length > 0) {
    Logger.log('Phase 2: Fetching ' + misses.length + ' tickers via chart history.');
    var CHART_BATCH = 20;
    var nowSec = Math.floor(now.getTime() / 1000);
    var oneYearAgo = nowSec - 365 * 86400;
    var phase2Count = 0;

    for (var b = 0; b < misses.length; b += CHART_BATCH) {
      var batch = misses.slice(b, Math.min(b + CHART_BATCH, misses.length));
      var requests = [];
      for (var i = 0; i < batch.length; i++) {
        var yahooSym = toYahooTicker_(batch[i]);
        requests.push({
          url: 'https://query1.finance.yahoo.com/v8/finance/chart/' +
               encodeURIComponent(yahooSym) +
               '?period1=' + oneYearAgo + '&period2=' + nowSec +
               '&interval=1d&events=div',
          muteHttpExceptions: true,
          headers: {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
          }
        });
      }

      try {
        var responses = UrlFetchApp.fetchAll(requests);
        for (var i = 0; i < responses.length; i++) {
          var ticker = batch[i];
          var key = ticker.toUpperCase();
          try {
            if (responses[i].getResponseCode() !== 200) {
              dataMap[key] = [ticker, '', now];
              continue;
            }
            var json = JSON.parse(responses[i].getContentText());
            var chartResult = json.chart && json.chart.result;
            if (!chartResult || !chartResult[0]) {
              dataMap[key] = [ticker, '', now];
              continue;
            }
            var events = chartResult[0].events;
            if (!events || !events.dividends) {
              dataMap[key] = [ticker, '', now];
              continue;
            }

            // Get the most recent dividend date (keys are Unix timestamps)
            var divKeys = Object.keys(events.dividends);
            var latest = 0;
            for (var k = 0; k < divKeys.length; k++) {
              var ts = Number(divKeys[k]);
              if (ts > latest) latest = ts;
            }

            if (latest > 0) {
              var d = projectNextDivDate_(new Date(latest * 1000), now);
              if (d) {
                dataMap[key] = [ticker, d, now];
                phase2Count++;
                if (phase2Count === 1) {
                  Logger.log('Phase 2 first success: ' + ticker + ' → ' + d.toISOString().split('T')[0]);
                }
                continue;
              }
            }
            dataMap[key] = [ticker, '', now];
          } catch (e) {
            dataMap[key] = [ticker, '', now];
          }
        }
      } catch (e) {
        Logger.log('Phase 2: Batch error at offset ' + b + ': ' + e.toString());
      }

      if (b + CHART_BATCH < misses.length) Utilities.sleep(1000);
    }

    successCount += phase2Count;
    Logger.log('Phase 2 done: ' + phase2Count + '/' + misses.length + ' recovered.');
  }

  // ── Write results to sheet ──
  var allRows = [];
  for (var key in dataMap) {
    allRows.push(dataMap[key]);
  }
  allRows.sort(function(a, b) { return String(a[0]).localeCompare(String(b[0])); });

  divSheet.clear();
  divSheet.getRange(1, 1, 1, 3).setValues([['Ticker', 'NextDivDate', 'LastFetched']]);
  divSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  if (allRows.length > 0) {
    divSheet.getRange(2, 1, allRows.length, 3).setValues(allRows);
  }

  Logger.log('fetchDividendDates: Done. ' + successCount + '/' + toFetch.length + ' got dates. ' + allRows.length + ' total tickers stored.');
}

/**
 * Projects the next dividend date from a known dividend date.
 * If the date is in the past, adds 91 days (quarterly) until it's >= today.
 * Returns null if the result is unreasonable (> 1 year out).
 */
function projectNextDivDate_(divDate, today) {
  var d = new Date(divDate.getTime());
  var maxFuture = new Date(today.getTime() + 400 * 86400000); // ~13 months max

  // If already in the future, return as-is
  if (d >= today) return d;

  // Add 91 days (quarterly) until it's in the future
  while (d < today) {
    d = new Date(d.getTime() + 91 * 86400000);
  }

  // Sanity check: don't return dates too far out
  if (d > maxFuture) return null;
  return d;
}

// ============================================================
// LIGHT UPDATE — snapshots GOOGLEFINANCE values to WebCache (every 10 min)
// ============================================================

/**
 * Light update function. Runs every 10 minutes via trigger.
 * Reads computed GOOGLEFINANCE values from Live/CreditLive sheets
 * and writes them as static values to WebCache/WebCacheCredit.
 * This decouples the API from GOOGLEFINANCE recalculation timing.
 * Target execution: <30 seconds.
 */
function updateLivePrices() {
  try {
    var ss = SpreadsheetApp.getActive();
    var updated = 0;

    var sources = [
      { from: 'Live', to: 'WebCache' },
      { from: 'CreditLive', to: 'WebCacheCredit' }
    ];

    for (var i = 0; i < sources.length; i++) {
      var src = ss.getSheetByName(sources[i].from);
      if (!src || src.getLastRow() <= 1) {
        Logger.log('updateLivePrices: ' + sources[i].from + ' not found or empty, skipping.');
        continue;
      }

      // Read all computed values (GOOGLEFINANCE formulas → resolved values)
      var data = src.getDataRange().getValues();
      if (data.length <= 1) continue;

      // Sanity check: at least some rows should have non-zero prices (col D = index 3)
      // If GOOGLEFINANCE hasn't populated yet, don't overwrite existing WebCache
      var validPrices = 0;
      for (var r = 1; r < data.length && r < 20; r++) {
        if (parseFloat(data[r][3]) > 0) validPrices++;
      }
      if (validPrices === 0) {
        Logger.log('updateLivePrices: ' + sources[i].from + ' has no valid prices (GOOGLEFINANCE still loading?). Keeping existing cache.');
        continue;
      }

      // Get or create target cache sheet
      var cache = ss.getSheetByName(sources[i].to);
      if (!cache) {
        cache = ss.insertSheet(sources[i].to);
      } else {
        cache.clear();
      }

      // Write headers
      cache.getRange(1, 1, 1, data[0].length).setValues([data[0]]);
      cache.getRange(1, 1, 1, data[0].length).setFontWeight('bold');
      SpreadsheetApp.flush();

      // Write data rows in chunks of 500
      var rows = data.slice(1);
      if (rows.length === 0) continue;

      var CHUNK = 500;
      for (var c = 0; c < rows.length; c += CHUNK) {
        var chunk = rows.slice(c, Math.min(c + CHUNK, rows.length));
        try {
          cache.getRange(c + 2, 1, chunk.length, data[0].length).setValues(chunk);
          SpreadsheetApp.flush();
        } catch (writeErr) {
          Logger.log('updateLivePrices: write error at row ' + (c + 2) + ' in ' + sources[i].to + ': ' + writeErr.toString());
        }
      }

      updated += rows.length;
      Logger.log('updateLivePrices: cached ' + rows.length + ' rows from ' + sources[i].from + ' → ' + sources[i].to);
    }

    // Store timestamp for diagnostics
    PropertiesService.getScriptProperties().setProperty('WEBCACHE_UPDATED', new Date().toISOString());
    Logger.log('updateLivePrices: complete. ' + updated + ' total rows cached.');

  } catch (e) {
    Logger.log('updateLivePrices ERROR: ' + e.toString());
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

  // Group by credit rating (collect nonRated for cross-rating pairs)
  var groups = {};
  var nonRated = [];
  for (var i = 1; i < data.length; i++) {
    var ticker = String(data[i][0]).trim();
    var coupon = data[i][2];
    var curYield = data[i][3];
    var rating = String(data[i][4]).trim();
    if (!ticker) continue;
    if (coupon === "" || coupon === null || coupon === undefined) continue;
    if (curYield === "" || curYield === null || curYield === undefined) continue;
    var upper = ticker.toUpperCase();
    if (!companyOf[upper]) companyOf[upper] = upper.replace(/-.*/, '');
    if (rating) {
      if (!groups[rating]) groups[rating] = [];
      groups[rating].push(ticker);
    } else {
      nonRated.push(ticker);
    }
  }

  // Same-rating pairs (with dedup map)
  var allPairs = [];
  var seen = {};
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
        var key = tA + "|" + tB;
        if (seen[key]) continue;
        seen[key] = true;
        allPairs.push([key, tA, tB, "CreditArb:" + rating]);
      }
    }
  }

  // Cross-rating pairs (OTHER sector) — matches generateCreditPairs() and dailyCreditRefresh()
  var crossPairs = generateCrossRatingPairs_(groups, nonRated, companyOf, intraPairs, seen);
  allPairs = allPairs.concat(crossPairs);

  // Cap total credit pairs to stay within GOOGLEFINANCE limits (~6 calls per pair in Live)
  var MAX_TOTAL_CREDIT_PAIRS = 2000;
  if (allPairs.length > MAX_TOTAL_CREDIT_PAIRS) {
    Logger.log('CreditPairs capped from ' + allPairs.length + ' to ' + MAX_TOTAL_CREDIT_PAIRS);
    allPairs = allPairs.slice(0, MAX_TOTAL_CREDIT_PAIRS);
  }

  // Write to CreditPairs sheet
  var cpSheet = getOrCreateSheet_(ss, 'CreditPairs');
  cpSheet.getRange(1, 1, 1, 4).setValues([['PairID', 'TickerA', 'TickerB', 'Sector']]);
  cpSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#2d1a2e').setFontColor('#ffffff');
  if (allPairs.length > 0) {
    cpSheet.getRange(2, 1, allPairs.length, 4).setValues(allPairs);
    SpreadsheetApp.flush();
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
      '=IFERROR(TRANSPOSE(ARRAYFORMULA(QUERY(GOOGLEFINANCE("' + tA + '","price",TODAY()-120,TODAY()),"select Col2 offset 1",0)-QUERY(GOOGLEFINANCE("' + tB + '","price",TODAY()-120,TODAY()),"select Col2 offset 1",0))),)'
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

// ── Cross-Credit-Rating Pair Generator (hardcoded tier rules) ──
function generateCrossRatingPairs_(groups, nonRated, companyOf, intraPairs, seen) {
  var toSet = function(arr) {
    var s = {};
    for (var i = 0; i < arr.length; i++) s[arr[i].toUpperCase()] = true;
    return s;
  };
  var exBBB = toSet(['BNH', 'BNJ', 'KTH', 'KTN']);
  var exBBBm = toSet(['BEPI', 'BEPJ', 'BEP-A', 'BHFAL', 'BHFAN', 'BHFAO', 'BHFAP',
                       'BIPI', 'BIPJ', 'BIP-A', 'BML-G', 'BML-J', 'BML-L',
                       'ETI-', 'IPB', 'SR-A']);
  var incBBp = toSet(['FITBO', 'HBANZ', 'CFG-E', 'CFG-H', 'DLR-J', 'DLR-K', 'DLR-L',
                       'HBANL', 'HBANM', 'HBANP', 'T-A', 'T-C', 'UNMA',
                       'WFC-C', 'WFC-Y', 'WFC-D', 'WFC-Z', 'WFC-A', 'SIGIP', 'MTB-K']);
  var exBB = toSet(['ANG-D', 'GJH']);
  var incBBm = toSet(['TCBIO', 'ASB-E', 'ASB-F', 'SYF-A']);

  var filt = function(arr, ex) {
    return (arr || []).filter(function(t) { return !ex[t.toUpperCase()]; });
  };

  // Build set of all valid tickers across all tiers + non-rated
  var allValid = {};
  for (var r in groups) {
    for (var i = 0; i < groups[r].length; i++) allValid[groups[r][i].toUpperCase()] = groups[r][i];
  }
  for (var i = 0; i < nonRated.length; i++) allValid[nonRated[i].toUpperCase()] = nonRated[i];

  var bbbPlus = groups['BBB+'] || [];
  var bbb = filt(groups['BBB'], exBBB);
  var bbbMinus = filt(groups['BBB-'], exBBBm);
  var bb = filt(groups['BB'], exBB);
  // Include-only pools: match against all valid Master tickers
  var bbPlusPool = [];
  for (var k in incBBp) { if (allValid[k]) bbPlusPool.push(allValid[k]); }
  var bbMinusPool = [];
  for (var k in incBBm) { if (allValid[k]) bbMinusPool.push(allValid[k]); }
  var nr = nonRated || [];

  var MAX_CROSS_PAIRS = 300;
  var pairs = [];
  var capHit = false;
  var addPairs = function(poolA, poolB) {
    if (capHit) return;
    for (var a = 0; a < poolA.length; a++) {
      for (var b = 0; b < poolB.length; b++) {
        if (pairs.length >= MAX_CROSS_PAIRS) { capHit = true; return; }
        var tA = poolA[a], tB = poolB[b];
        if (tA.toUpperCase() === tB.toUpperCase()) continue;
        if (intraPairs[cleanId_(tA) + '_' + cleanId_(tB)]) continue;
        var coA = companyOf[tA.toUpperCase()] || tA.toUpperCase().replace(/-.*/, '');
        var coB = companyOf[tB.toUpperCase()] || tB.toUpperCase().replace(/-.*/, '');
        if (coA === coB) continue;
        var sorted = [tA, tB].sort();
        var key = sorted[0] + '|' + sorted[1];
        if (seen[key]) continue;
        seen[key] = true;
        pairs.push([key, sorted[0], sorted[1], 'OTHER']);
      }
    }
  };

  addPairs(bbbPlus, bbb);             // 1. BBB+ vs BBB
  addPairs(bbbPlus, bbbMinus);        // 2. BBB+ vs BBB-
  addPairs(bbb, bbbMinus);            // 3. BBB vs BBB-
  addPairs(bbbPlus, bbPlusPool);      // 4. BBB+ vs BB+ (include-only)
  addPairs(bbb, bbPlusPool);          // 5. BBB vs BB+ (include-only)
  addPairs(bbbMinus, bbPlusPool);     // 6. BBB- vs BB+ (include-only)
  addPairs(bbbPlus, bb);              // 7. BBB+ vs BB
  addPairs(bbb, bb);                  // 8. BBB vs BB
  addPairs(bbbMinus, bb);             // 9. BBB- vs BB
  addPairs(bbPlusPool, bb);           // 10. BB+ (include-only) vs BB
  // 11. BB- (include-only) vs everything except BBB+
  var rule11Pool = bbb.concat(bbbMinus).concat(bbPlusPool).concat(bb).concat(nr);
  addPairs(bbMinusPool, rule11Pool);
  addPairs(nr, nr);                   // 12. Non-Rated vs Non-Rated

  Logger.log('generateCrossRatingPairs_: ' + pairs.length + ' cross-rating pairs');
  return pairs;
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
  var nonRated = [];
  for (var i = 1; i < data.length; i++) {
    var ticker = String(data[i][0]).trim();
    var coupon = data[i][2]; var curYield = data[i][3]; var rating = String(data[i][4]).trim();
    if (!ticker) continue;
    if (coupon === "" || coupon === null || coupon === undefined) continue;
    if (curYield === "" || curYield === null || curYield === undefined) continue;
    var upper = ticker.toUpperCase();
    if (!companyOf[upper]) companyOf[upper] = upper.replace(/-.*/, '');
    if (rating) {
      if (!groups[rating]) groups[rating] = [];
      groups[rating].push(ticker);
    } else {
      nonRated.push(ticker);
    }
  }

  // Same-rating pairs
  var allPairs = [];
  var seen = {};
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
        var key = tA + "|" + tB;
        seen[key] = true;
        allPairs.push([key, tA, tB, "CreditArb:" + rating]);
      }
    }
  }

  // Cross-rating pairs (OTHER sector)
  var crossPairs = generateCrossRatingPairs_(groups, nonRated, companyOf, intraPairs, seen);
  allPairs = allPairs.concat(crossPairs);

  // Cap total credit pairs to stay within GOOGLEFINANCE limits
  var MAX_TOTAL_CREDIT_PAIRS = 2000;
  if (allPairs.length > MAX_TOTAL_CREDIT_PAIRS) {
    Logger.log('CreditPairs capped from ' + allPairs.length + ' to ' + MAX_TOTAL_CREDIT_PAIRS);
    allPairs = allPairs.slice(0, MAX_TOTAL_CREDIT_PAIRS);
  }

  var cpSheet = getOrCreateSheet_(ss, 'CreditPairs');
  cpSheet.getRange(1, 1, 1, 4).setValues([['PairID', 'TickerA', 'TickerB', 'Sector']]);
  cpSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#2d1a2e').setFontColor('#ffffff');
  if (allPairs.length > 0) { cpSheet.getRange(2, 1, allPairs.length, 4).setValues(allPairs); }

  var summary = [];
  var ratingCounts = {};
  allPairs.forEach(function(p) { var r = p[3]; ratingCounts[r] = (ratingCounts[r] || 0) + 1; });
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
      '=IFERROR(TRANSPOSE(ARRAYFORMULA(QUERY(GOOGLEFINANCE("' + tA + '","price",TODAY()-120,TODAY()),"select Col2 offset 1",0)-QUERY(GOOGLEFINANCE("' + tB + '","price",TODAY()-120,TODAY()),"select Col2 offset 1",0))),)'
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
        // Skip blacklisted tickers
        var tA = String(row[1]).toUpperCase().trim();
        var tB = String(row[2]).toUpperCase().trim();
        if (isBlacklisted_(tA) || isBlacklisted_(tB)) continue;
        // Skip intra-only tickers in credit sheets
        if (sheets[s].source === 'credit' && (isIntraOnly_(tA) || isIntraOnly_(tB))) continue;
        var priceA = parseFloat(row[3]) || 0;
        var priceB = parseFloat(row[4]) || 0;
        if (priceA <= 0 || priceB <= 0) continue;
        var stdev = parseFloat(row[11]) || 0;
        var histCount = parseFloat(row[16]) || 0;
        if (stdev <= 0.001 || histCount < 40) continue;  // No valid or insufficient history
        var couponA = row[8], couponB = row[9];
        if (couponA === "" || couponA === null || couponB === "" || couponB === null) continue;
        var zScore = parseFloat(row[12]) || 0;
        var spread = parseFloat(row[5]) || 0;
        var cid = cleanId_(String(pairId));
        // Only log pairs with |Z| >= 1.0 to keep AlertsLog volume manageable
        // (trend ribbons only display for pairs at |Z| >= 1.8 anyway)
        if (Math.abs(zScore) >= 1.0) {
          logRows.push([now, String(pairId), zScore, spread]);
        }
        var isActive = Math.abs(zScore) >= 1.8;
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
    if (totalRows > 25000) { logSheet.deleteRows(2, totalRows - 20000); }
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
    var nonRated = [];
    for (var i = 1; i < data.length; i++) {
      var ticker = String(data[i][0]).trim();
      var coupon = data[i][2]; var curYield = data[i][3]; var rating = String(data[i][4]).trim();
      if (!ticker) continue;
      if (coupon === "" || coupon === null || coupon === undefined) continue;
      if (curYield === "" || curYield === null || curYield === undefined) continue;
      var upper = ticker.toUpperCase();
      if (!companyOf[upper]) companyOf[upper] = upper.replace(/-.*/, '');
      if (rating) {
        if (!groups[rating]) groups[rating] = [];
        groups[rating].push(ticker);
      } else {
        nonRated.push(ticker);
      }
    }

    // Same-rating pairs
    var allPairs = [];
    var seen = {};
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
          var key = tA + "|" + tB;
          seen[key] = true;
          allPairs.push([key, tA, tB, "CreditArb:" + rating]);
        }
      }
    }

    // Cross-rating pairs (OTHER sector)
    var crossPairs = generateCrossRatingPairs_(groups, nonRated, companyOf, intraPairs, seen);
    allPairs = allPairs.concat(crossPairs);

    // Cap total credit pairs to stay within GOOGLEFINANCE limits
    var MAX_TOTAL_CREDIT_PAIRS = 2000;
    if (allPairs.length > MAX_TOTAL_CREDIT_PAIRS) {
      Logger.log('dailyCreditRefresh: capped from ' + allPairs.length + ' to ' + MAX_TOTAL_CREDIT_PAIRS);
      allPairs = allPairs.slice(0, MAX_TOTAL_CREDIT_PAIRS);
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
        '=IFERROR(TRANSPOSE(ARRAYFORMULA(QUERY(GOOGLEFINANCE("' + tA + '","price",TODAY()-120,TODAY()),"select Col2 offset 1",0)-QUERY(GOOGLEFINANCE("' + tB + '","price",TODAY()-120,TODAY()),"select Col2 offset 1",0))),)'
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
  // Write in chunks of 500 rows to avoid timeout on large datasets (500+ credit pairs)
  var CHUNK_SIZE = 500;
  for (var c = 0; c < formulas.length; c += CHUNK_SIZE) {
    var chunk = formulas.slice(c, Math.min(c + CHUNK_SIZE, formulas.length));
    var startRow = c + 2; // row 1 is headers, data starts at row 2
    live.getRange(startRow, 1, chunk.length, 24).setFormulas(chunk);
    SpreadsheetApp.flush();
    Logger.log(sheetName + ': wrote rows ' + startRow + '-' + (startRow + chunk.length - 1) + ' of ' + (formulas.length + 1));
    if (c + CHUNK_SIZE < formulas.length) {
      Utilities.sleep(200);
    }
  }
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
/** Blacklist check — mirrors BLACKLIST in Code.gs */
var BLACKLIST_ = {
  'GJH':1,'GJP':1,'GJO':1,'GJR':1,'GJS':1,'GJT':1,
  'EPR-E':1,'EPR-G':1,'EPR-C':1,
  'KTH':1,'KTN':1,
  'ONBPO':1,'ONBPP':1,
  'BEPI':1,'BIP-A':1,'BIPJ':1,'BIPI':1,'BIPH':1,'BEPH':1,'BEP-A':1,'BEPJ':1,
  'IPB':1,
  'SR-A':1,'RIV-A':1,'ACP-A':1,'OPP-A':1,
  'GAB-H':1,'GAB-K':1,
  'GGT-E':1,'GGT-G':1,
  'OPP-B':1,'GAM-B':1,
  'GDV-H':1,'GNT-A':1,'GUT-C':1,'GAB-G':1
};
var INTRA_ONLY_ = {
  'BHFAN':1,'BHFAO':1,'BHFAM':1,'BHFAP':1,'BHFAL':1,
  'HFRO-A':1,'HFRO-B':1
};
function isBlacklisted_(ticker) {
  return !!BLACKLIST_[String(ticker).toUpperCase().trim()];
}
function isIntraOnly_(ticker) {
  var t = String(ticker).toUpperCase().trim();
  if (INTRA_ONLY_[t]) return true;
  if (t.indexOf('SCE-') === 0) return true;
  return false;
}
function showMsg_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

// ============================================================
// MACRO VALUATION — Preferreds vs. US Treasuries
// ============================================================
// Benchmarks: fetched from FRED public CSV (no API key needed)
var MACRO_BENCHMARKS = [
  { key: 'US2Y',  fred: 'DGS2',  label: 'US 2Y' },
  { key: 'US5Y',  fred: 'DGS5',  label: 'US 5Y' },
  { key: 'US7Y',  fred: 'DGS7',  label: 'US 7Y' },
  { key: 'US10Y', fred: 'DGS10', label: 'US 10Y' },
  { key: 'US30Y', fred: 'DGS30', label: 'US 30Y' }
];

// MacroCalc column headers (45 columns)
// [0-5] Ticker, Coupon, CurPrice, CurYield, Sector, Credit
// [6-12] US2Y: Spr, Mean, Z, Pct, Hi, Lo, Dir  (×5 benchmarks = 35 cols)
// [41-44] AvgZ, Signal, BestBM, LastComputed
var MACRO_HEADERS = (function() {
  var h = ['Ticker', 'Coupon', 'CurPrice', 'CurYield', 'Sector', 'Credit'];
  for (var b = 0; b < MACRO_BENCHMARKS.length; b++) {
    var k = MACRO_BENCHMARKS[b].key;
    h.push(k + '_Spr', k + '_Mean', k + '_Z', k + '_Pct', k + '_Hi', k + '_Lo', k + '_Dir');
  }
  h.push('AvgZ', 'Signal', 'BestBM', 'LastComputed');
  return h;
})();

/**
 * Fetches 120 days of treasury yield history from FRED (public CSV, no API key).
 * Writes to TreasuryHist sheet: Date | US2Y | US5Y | US7Y | US10Y | US30Y
 */
function fetchTreasuryHistory() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('TreasuryHist');
  if (!sheet) {
    sheet = ss.insertSheet('TreasuryHist');
  } else {
    sheet.clear();
  }

  var endDate = Utilities.formatDate(new Date(), 'GMT', 'yyyy-MM-dd');
  var startDate = Utilities.formatDate(new Date(new Date().getTime() - 150 * 86400000), 'GMT', 'yyyy-MM-dd');

  // Fetch all 5 benchmarks in parallel
  var requests = [];
  for (var i = 0; i < MACRO_BENCHMARKS.length; i++) {
    requests.push({
      url: 'https://fred.stlouisfed.org/graph/fredgraph.csv?id=' + MACRO_BENCHMARKS[i].fred +
           '&cosd=' + startDate + '&coed=' + endDate,
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' }
    });
  }

  var responses = UrlFetchApp.fetchAll(requests);

  // Parse each CSV into dateStr → yield
  var yieldMaps = [];
  for (var i = 0; i < responses.length; i++) {
    var map = {};
    if (responses[i].getResponseCode() === 200) {
      var lines = responses[i].getContentText().split('\n');
      for (var j = 1; j < lines.length; j++) {
        var parts = lines[j].split(',');
        if (parts.length >= 2 && parts[1].trim() !== '.' && parts[1].trim() !== '') {
          var val = parseFloat(parts[1]);
          if (!isNaN(val)) map[parts[0].trim()] = val;
        }
      }
    }
    yieldMaps.push(map);
    Logger.log('fetchTreasuryHistory: ' + MACRO_BENCHMARKS[i].key + ' — ' + Object.keys(map).length + ' data points');
  }

  // Build unified sorted date list
  var allDates = {};
  for (var i = 0; i < yieldMaps.length; i++) {
    for (var d in yieldMaps[i]) allDates[d] = true;
  }
  var dates = Object.keys(allDates).sort();

  // Write to sheet
  var headers = ['Date'];
  for (var i = 0; i < MACRO_BENCHMARKS.length; i++) headers.push(MACRO_BENCHMARKS[i].key);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  if (dates.length > 0) {
    var rows = [];
    for (var d = 0; d < dates.length; d++) {
      var row = [dates[d]];
      for (var i = 0; i < yieldMaps.length; i++) {
        row.push(yieldMaps[i][dates[d]] || '');
      }
      rows.push(row);
    }
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    // Force date column to plain text so Sheets doesn't auto-convert to Date objects
    sheet.getRange(2, 1, rows.length, 1).setNumberFormat('@');
  }

  Logger.log('fetchTreasuryHistory: Done. ' + dates.length + ' dates written.');
}

/**
 * Daily orchestrator — called by trigger at 7 AM.
 * Step 1: Refresh treasury yields from FRED.
 * Step 2: Schedule batched computation as a SEPARATE execution (avoids timeout).
 */
function dailyMacroRefresh() {
  try {
    Logger.log('dailyMacroRefresh: Starting treasury fetch...');
    fetchTreasuryHistory();
    Logger.log('dailyMacroRefresh: Treasury fetch complete. Scheduling computation in 1 min...');
    // Schedule computation as a separate execution to avoid sharing the 6-min limit
    cleanupMacroBatchTriggers_();
    ScriptApp.newTrigger('computeMacroValuationsBatch')
      .timeBased()
      .after(60000)
      .create();
  } catch (e) {
    Logger.log('dailyMacroRefresh ERROR: ' + e.toString());
  }
}

/**
 * Batched macro valuation computation. Processes 20 tickers per iteration.
 * Uses PropertiesService to save/resume progress across GAS time limits.
 * When more tickers remain + time runs out, schedules a 1-min continuation trigger.
 */
function computeMacroValuationsBatch() {
  var startTime = new Date().getTime();
  var props = PropertiesService.getScriptProperties();
  var stateJson = props.getProperty('MACRO_BATCH');
  var state = stateJson ? JSON.parse(stateJson) : { idx: 0 };
  var ss = SpreadsheetApp.getActive();

  Logger.log('computeMacroValuationsBatch: Starting from index ' + state.idx);

  // Read Master sheet for ticker list
  var master = ss.getSheetByName('Master');
  if (!master) { Logger.log('computeMacroValuationsBatch: No Master sheet'); return; }
  var masterData = master.getDataRange().getValues();

  // Build ticker list (only fixed-rate preferreds with coupon)
  var tickers = [];
  for (var i = 1; i < masterData.length; i++) {
    var ticker = String(masterData[i][0]).trim();
    var coupon = masterData[i][2]; // Column C: Coupon Yield
    var curYield = masterData[i][3]; // Column D: Current Yield
    var credit = String(masterData[i][4] || '').trim();
    if (!ticker) continue;
    if (coupon === '' || coupon === null || coupon === undefined) continue;
    if (isBlacklisted_(ticker)) continue;
    // Master sheet may store coupon as decimal (0.06) or percentage (6.00).
    // All preferred coupons are > 1%, so if value < 1 it's a decimal → multiply by 100.
    var couponPct = parseFloat(coupon) || 0;
    if (couponPct > 0 && couponPct < 1) couponPct = couponPct * 100;
    var yieldPct = parseFloat(curYield) || 0;
    if (yieldPct > 0 && yieldPct < 1) yieldPct = yieldPct * 100;
    tickers.push({
      ticker: ticker,
      coupon: couponPct,
      curYield: yieldPct,
      credit: credit,
      sector: ''
    });
  }

  // Get sector from Pairs sheet
  var pairsSheet = ss.getSheetByName('Pairs');
  if (pairsSheet) {
    var pData = pairsSheet.getDataRange().getValues();
    var sectorMap = {};
    for (var i = 1; i < pData.length; i++) {
      var sec = String(pData[i][3] || '').trim();
      if (sec) {
        var tA = String(pData[i][1]).trim().toUpperCase();
        var tB = String(pData[i][2]).trim().toUpperCase();
        if (tA) sectorMap[tA] = sec;
        if (tB) sectorMap[tB] = sec;
      }
    }
    for (var i = 0; i < tickers.length; i++) {
      tickers[i].sector = sectorMap[tickers[i].ticker.toUpperCase()] || '';
    }
  }

  Logger.log('computeMacroValuationsBatch: ' + tickers.length + ' total tickers');

  if (state.idx >= tickers.length) {
    props.deleteProperty('MACRO_BATCH');
    snapshotMacroCache_();
    cleanupMacroBatchTriggers_();
    Logger.log('computeMacroValuationsBatch: All done! (was already complete)');
    return;
  }

  // Read TreasuryHist
  var treasurySheet = ss.getSheetByName('TreasuryHist');
  if (!treasurySheet || treasurySheet.getLastRow() <= 1) {
    Logger.log('computeMacroValuationsBatch: No TreasuryHist data. Run fetchTreasuryHistory() first.');
    return;
  }
  var treasuryData = treasurySheet.getDataRange().getValues();

  // Build treasury date → yields lookup
  var treasuryByDate = {};
  var treasuryDates = [];
  for (var i = 1; i < treasuryData.length; i++) {
    var rawDate = treasuryData[i][0];
    var dateStr;
    // Google Sheets auto-converts date strings to Date objects — handle both
    if (rawDate instanceof Date) {
      dateStr = Utilities.formatDate(rawDate, 'GMT', 'yyyy-MM-dd');
    } else {
      dateStr = String(rawDate).trim();
    }
    if (!dateStr) continue;
    var yields = [];
    for (var b = 0; b < MACRO_BENCHMARKS.length; b++) {
      yields.push(parseFloat(treasuryData[i][b + 1]) || 0);
    }
    treasuryByDate[dateStr] = yields;
    treasuryDates.push(dateStr);
  }
  treasuryDates.sort();
  if (treasuryDates.length > 100) treasuryDates = treasuryDates.slice(-100);

  // Get or create MacroCalc sheet
  var calcSheet = ss.getSheetByName('MacroCalc');
  if (!calcSheet) {
    calcSheet = ss.insertSheet('MacroCalc');
    calcSheet.getRange(1, 1, 1, MACRO_HEADERS.length).setValues([MACRO_HEADERS]);
    calcSheet.getRange(1, 1, 1, MACRO_HEADERS.length).setFontWeight('bold');
  } else if (state.idx === 0) {
    // First batch — clear existing data rows (keep headers)
    if (calcSheet.getLastRow() > 1) {
      calcSheet.getRange(2, 1, calcSheet.getLastRow() - 1, MACRO_HEADERS.length).clearContent();
    }
  }

  var BATCH_SIZE = 20;
  var now = new Date();
  var nowSec = Math.floor(now.getTime() / 1000);
  var histStartSec = nowSec - 150 * 86400;

  // Process batches in a loop until time runs out
  while (state.idx < tickers.length) {
    // Time check
    if ((new Date().getTime() - startTime) > MAX_RUNTIME_MS) {
      props.setProperty('MACRO_BATCH', JSON.stringify(state));
      scheduleMacroContinuation_();
      Logger.log('computeMacroValuationsBatch: Paused at ' + state.idx + '/' + tickers.length + '. Continuation scheduled.');
      return;
    }

    var endIdx = Math.min(state.idx + BATCH_SIZE, tickers.length);
    var batch = tickers.slice(state.idx, endIdx);

    // Fetch 90-day price history for batch via Yahoo Finance v8/chart
    var requests = [];
    for (var i = 0; i < batch.length; i++) {
      var yahooSym = toYahooTicker_(batch[i].ticker);
      requests.push({
        url: 'https://query1.finance.yahoo.com/v8/finance/chart/' +
             encodeURIComponent(yahooSym) +
             '?period1=' + histStartSec + '&period2=' + nowSec + '&interval=1d',
        muteHttpExceptions: true,
        headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' }
      });
    }

    var responses = UrlFetchApp.fetchAll(requests);

    // Process each ticker in the batch
    var resultRows = [];
    for (var i = 0; i < batch.length; i++) {
      var t = batch[i];
      var row = [t.ticker, t.coupon, 0, 0, t.sector, t.credit];

      // Parse Yahoo response for price history
      var priceByDate = {};
      var curPrice = 0;
      try {
        if (responses[i].getResponseCode() === 200) {
          var json = JSON.parse(responses[i].getContentText());
          var chartResult = json.chart && json.chart.result && json.chart.result[0];
          if (chartResult) {
            var timestamps = chartResult.timestamp || [];
            var closes = (chartResult.indicators && chartResult.indicators.quote &&
                          chartResult.indicators.quote[0] && chartResult.indicators.quote[0].close) || [];
            for (var j = 0; j < timestamps.length; j++) {
              if (closes[j] !== null && closes[j] !== undefined) {
                var d = new Date(timestamps[j] * 1000);
                var ds = Utilities.formatDate(d, 'GMT', 'yyyy-MM-dd');
                priceByDate[ds] = closes[j];
                curPrice = closes[j];
              }
            }
          }
        }
      } catch (e) {
        // Skip this ticker on parse error
      }

      row[2] = curPrice;
      // Current Yield = Coupon% × $25 par / Price  (standard preferred stock formula)
      row[3] = curPrice > 0 ? parseFloat((t.coupon * 25 / curPrice).toFixed(4)) : 0;

      // Compute spread stats for each benchmark
      var allZ = [];
      var bestZ = 0, bestBM = '';

      for (var b = 0; b < MACRO_BENCHMARKS.length; b++) {
        var spreads = [];
        for (var d = 0; d < treasuryDates.length; d++) {
          var dateStr = treasuryDates[d];
          var price = priceByDate[dateStr];
          var tYield = treasuryByDate[dateStr] ? treasuryByDate[dateStr][b] : 0;
          if (price && price > 0 && tYield > 0) {
            var prefYield = t.coupon * 25 / price;  // Coupon% × $25 par / Price
            var spread = (prefYield - tYield) * 100; // bps
            spreads.push(spread);
          }
        }

        var curSpread = 0, mean = 0, z = 0, pct = 50, hi = 0, lo = 0, dir = '→';

        if (spreads.length >= 20) {
          var calcSpreads = spreads.slice(-90);
          // Mean
          var sum = 0;
          for (var s = 0; s < calcSpreads.length; s++) sum += calcSpreads[s];
          mean = sum / calcSpreads.length;
          // StDev
          var sqSum = 0;
          for (var s = 0; s < calcSpreads.length; s++) sqSum += Math.pow(calcSpreads[s] - mean, 2);
          var stdev = Math.sqrt(sqSum / calcSpreads.length);
          // Current spread
          curSpread = calcSpreads[calcSpreads.length - 1];
          // Z-score
          z = stdev > 0 ? (curSpread - mean) / stdev : 0;
          // Percentile
          var below = 0;
          for (var s = 0; s < calcSpreads.length; s++) {
            if (calcSpreads[s] <= curSpread) below++;
          }
          pct = Math.round((below / calcSpreads.length) * 100);
          // Hi/Lo
          hi = calcSpreads[0]; lo = calcSpreads[0];
          for (var s = 1; s < calcSpreads.length; s++) {
            if (calcSpreads[s] > hi) hi = calcSpreads[s];
            if (calcSpreads[s] < lo) lo = calcSpreads[s];
          }
          // Direction (7-day slope)
          if (calcSpreads.length >= 7) {
            var recent = calcSpreads.slice(-7);
            var slope = recent[recent.length - 1] - recent[0];
            if (slope > 5) dir = '↗';
            else if (slope < -5) dir = '↘';
          }

          allZ.push(z);
          if (Math.abs(z) > Math.abs(bestZ)) {
            bestZ = z;
            bestBM = MACRO_BENCHMARKS[b].key;
          }
        }

        row.push(
          Math.round(curSpread), Math.round(mean),
          parseFloat(z.toFixed(2)), pct,
          Math.round(hi), Math.round(lo), dir
        );
      }

      // Aggregates
      var avgZ = 0;
      if (allZ.length > 0) {
        var zSum = 0;
        for (var z2 = 0; z2 < allZ.length; z2++) zSum += allZ[z2];
        avgZ = zSum / allZ.length;
      }
      var signal = 'FAIR';
      if (avgZ >= 2.5) signal = 'EXTREMELY CHEAP';
      else if (avgZ >= 1.5) signal = 'CHEAP';
      else if (avgZ <= -2.5) signal = 'EXTREMELY EXPENSIVE';
      else if (avgZ <= -1.5) signal = 'EXPENSIVE';

      row.push(parseFloat(avgZ.toFixed(2)), signal, bestBM || 'N/A', now.toISOString());
      resultRows.push(row);
    }

    // Write batch results to MacroCalc
    if (resultRows.length > 0) {
      calcSheet.getRange(state.idx + 2, 1, resultRows.length, MACRO_HEADERS.length).setValues(resultRows);
      SpreadsheetApp.flush();
    }

    Logger.log('computeMacroValuationsBatch: Processed ' + state.idx + '-' + endIdx + ' of ' + tickers.length);
    state.idx = endIdx;

    // Brief pause between batches to avoid rate limits
    if (state.idx < tickers.length) Utilities.sleep(1000);
  }

  // All done
  props.deleteProperty('MACRO_BATCH');
  snapshotMacroCache_();
  cleanupMacroBatchTriggers_();
  Logger.log('computeMacroValuationsBatch: COMPLETE. All ' + tickers.length + ' tickers processed.');
}

/** Snapshot MacroCalc → MacroCache (same pattern as updateLivePrices) */
function snapshotMacroCache_() {
  var ss = SpreadsheetApp.getActive();
  var src = ss.getSheetByName('MacroCalc');
  if (!src || src.getLastRow() <= 1) return;
  var data = src.getDataRange().getValues();
  var cache = ss.getSheetByName('MacroCache');
  if (!cache) {
    cache = ss.insertSheet('MacroCache');
  } else {
    cache.clear();
  }
  cache.getRange(1, 1, data.length, data[0].length).setValues(data);
  cache.getRange(1, 1, 1, data[0].length).setFontWeight('bold');
  PropertiesService.getScriptProperties().setProperty('MACRO_CACHE_UPDATED', new Date().toISOString());
  Logger.log('snapshotMacroCache_: Cached ' + (data.length - 1) + ' rows.');
}

/** Schedule a 1-minute continuation trigger for batch processing */
function scheduleMacroContinuation_() {
  cleanupMacroBatchTriggers_();
  ScriptApp.newTrigger('computeMacroValuationsBatch')
    .timeBased()
    .after(60000)
    .create();
}

/** Remove any existing macro batch continuation triggers */
function cleanupMacroBatchTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'computeMacroValuationsBatch') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/** Manual reset if macro batch gets stuck */
function clearMacroBatchState() {
  PropertiesService.getScriptProperties().deleteProperty('MACRO_BATCH');
  cleanupMacroBatchTriggers_();
  Logger.log('Macro batch state cleared.');
  showMsg_('Macro batch state cleared. Run dailyMacroRefresh() to start fresh.');
}
