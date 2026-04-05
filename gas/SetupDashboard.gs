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
 *   4 = TickerData + TickerHistory (GOOGLEFINANCE per ticker, not per pair)
 *   5 = Initial credit cache computation (computeCreditCache)
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
          '=IFERROR(STDEV(G' + r + ':' + r + '), 0)',
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
      ensureSheet_(ss, 'OpenTrades', ['PairID', 'EntryZ', 'CostA', 'CostB', 'SizeA', 'SizeB', 'Timestamp', 'PaidDiv', 'ReceivedDiv', 'TargetExitZ', 'ProfitCapturePct', 'TargetPnL', 'PartialAtPct', 'SourcePortfolio', 'MaxHoldDays']);
      ensureSheet_(ss, 'ClosedTrades', ['PairID', 'EntryZ', 'CostA', 'CostB', 'SizeA', 'SizeB', 'OpenDate', 'CloseDate', 'PnL', 'ExitPriceA', 'ExitPriceB', 'ExitZ', 'CloseType', 'PaidDiv', 'ReceivedDiv', 'Notes', 'TargetExitZ', 'TargetPnL', 'SourcePortfolio']);
      ensureSheet_(ss, 'AlertsLog', ['Timestamp', 'PairID', 'Z-Score', 'Spread']);
      var ageSheet = ss.getSheetByName('ZScoreAge');
      if (!ageSheet) {
        ageSheet = ss.insertSheet('ZScoreAge');
        ageSheet.getRange(1, 1, 1, 3).setValues([['PairID', 'FirstCrossTimestamp', 'Source']]);
        ageSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
      }
      ensureSheet_(ss, 'DivDates', ['Ticker', 'ExDivDate', 'LastFetched']);
      ensureSheet_(ss, 'Watchlist', ['PairID', 'AddedDate', 'AddedZ', 'Mode', 'AddedExpProfit', 'AddedSpread']);
      ensureSheet_(ss, 'BasketCache', ['Key', 'Value', 'UpdatedAt']);
      ensureSheet_(ss, 'ScreenerCache', ['PairID','TickerA','TickerB','Mode','Z','ExpProfit','WR30','WR60','WR90','AvgMAE','P75MAE','WidenProb','Triggers','EP30','EP60','UpdatedAt','WR15','EP15','EP90']);
      ensureSheet_(ss, 'ModelPortfolioCache', ['Rank','Pairs','BasketZ','BlendedWR','ExpProfit','WidenProb','SectorMix','AvgCorrelation','Metrics','UpdatedAt']);
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

    // --- PHASE 4: Ticker-level GOOGLEFINANCE sheets ---
    // Instead of CreditLevels (1 GOOGLEFINANCE per pair = 30K+ calls), we build:
    //   TickerData: ~200 tickers × 3 formulas = ~600 calls (live price, volumeavg, volume)
    //   TickerHistory: ~200 tickers × 1 formula = ~200 calls (1-year historical prices)
    // Then computeCreditCache() computes all pair stats from these ticker sheets.
    if (state.phase === 4) {
      Logger.log('Phase 4: Building TickerData + TickerHistory sheets...');
      var tickers = buildTickerDataSheet_(ss);
      buildTickerHistorySheet_(ss, tickers);
      SpreadsheetApp.flush();
      Logger.log('Phase 4 complete: ' + tickers.length + ' tickers in TickerData + TickerHistory.');

      state.phase = 5;
      if (isTimeUp_(startTime)) { saveState_(props, state, 'Phase 4 done. Pausing before credit cache.'); return; }
    }

    // --- PHASE 5: Initial credit cache computation ---
    // TickerData/TickerHistory GOOGLEFINANCE formulas may not have resolved yet.
    // computeCreditCache() will skip if no valid prices found.
    // The 10-minute updateLivePrices trigger will retry automatically.
    if (state.phase === 5) {
      Logger.log('Phase 5: Computing initial credit cache from ticker data...');
      Utilities.sleep(5000); // Brief pause for GOOGLEFINANCE formulas to start loading
      SpreadsheetApp.flush();
      try {
        computeCreditCache();
      } catch (ccErr) {
        Logger.log('Phase 5: Initial credit cache failed (GOOGLEFINANCE still loading): ' + ccErr.toString());
        Logger.log('updateLivePrices trigger will compute it on the next 10-minute cycle.');
      }
      Logger.log('Phase 5 complete.');

      state.phase = 6;
      if (isTimeUp_(startTime)) { saveState_(props, state, 'Phase 5 done. Pausing before triggers.'); return; }
    }

    // --- PHASE 6: Finalize + populate WebCache ---
    if (state.phase === 6) {
      Logger.log('Phase 6: Finalizing...');
      // Triggers are installed by createAutoTrigger() — not duplicated here.
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
        fn === 'dailyMacroRefresh' || fn === 'computeMacroValuationsBatch' || fn === 'updateBasketAnalytics' ||
        fn === 'runNightlyScreener' || fn === 'runExitAlertCheck' || fn === 'runModelPortfolioGenerator' ||
        fn === 'updateSweepCache') {
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

  // DAILY: updateBasketAnalytics at 8 AM (portfolio basket Z-scores)
  ScriptApp.newTrigger('updateBasketAnalytics')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();

  // DAILY: runNightlyScreener at 9 AM (pre-compute probability analysis for top alerts)
  ScriptApp.newTrigger('runNightlyScreener')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();

  // DAILY: updateSweepCache at 9:30 AM (pre-compute optimal entry/exit sweeps for alert pairs)
  ScriptApp.newTrigger('updateSweepCache')
    .timeBased()
    .atHour(9)
    .nearMinute(30)
    .everyDays(1)
    .create();

  // DAILY: runModelPortfolioGenerator at 10 AM (build optimal model portfolios from sweep + alert data)
  ScriptApp.newTrigger('runModelPortfolioGenerator')
    .timeBased()
    .atHour(10)
    .everyDays(1)
    .create();

  // HOURLY: runExitAlertCheck (proactive exit signals for open trades)
  ScriptApp.newTrigger('runExitAlertCheck')
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log('All triggers installed:\n• updateLivePrices: every 10 min\n• snapshotZScores: every hour\n• runExitAlertCheck: every hour\n• dailyCreditRefresh: daily 5 AM\n• fetchDividendDates: daily 6 AM\n• dailyMacroRefresh: daily 7 AM\n• updateBasketAnalytics: daily 8 AM\n• runNightlyScreener: daily 9 AM\n• updateSweepCache: daily 9:30 AM\n• runModelPortfolioGenerator: daily 10 AM');
  showMsg_(
    'Triggers installed!\n\n' +
    '• updateLivePrices: every 10 min (snapshots prices to WebCache)\n' +
    '• snapshotZScores: every hour (trend tracking)\n' +
    '• runExitAlertCheck: every hour (proactive exit signals)\n' +
    '• dailyCreditRefresh: daily 5 AM (credit pair regeneration)\n' +
    '• fetchDividendDates: daily 6 AM (Yahoo Finance ex-div dates)\n' +
    '• dailyMacroRefresh: daily 7 AM (macro valuation vs treasuries)\n' +
    '• updateBasketAnalytics: daily 8 AM (portfolio basket Z-scores)\n' +
    '• runNightlyScreener: daily 9 AM (pre-compute top alert probabilities)\n' +
    '• updateSweepCache: daily 9:30 AM (pre-compute optimal entry/exit sweeps)\n' +
    '• runModelPortfolioGenerator: daily 10 AM (build optimal model portfolios)\n\n' +
    'Now run setupAllBatched() to build the sheets.\n' +
    'Then run updateLivePrices() to populate WebCache immediately.'
  );
}

// ============================================================
// BASKET ANALYTICS — Daily trigger
// ============================================================
/**
 * Pre-computes basket analytics and caches results to BasketCache sheet.
 * Called daily at 8 AM by trigger. Results read by getBasketAnalytics() in Code.gs.
 */
function updateBasketAnalytics() {
  try {
    var result = getBasketAnalytics();
    if (!result || !result.trades) {
      Logger.log('updateBasketAnalytics: No open trades, skipping.');
      return;
    }
    var ss = SpreadsheetApp.getActive();
    ensureSheet_(ss, 'BasketCache', ['Key', 'Value', 'UpdatedAt']);
    var sheet = ss.getSheetByName('BasketCache');
    // Clear old data and write fresh cache
    if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
    var now = new Date().toISOString();
    var rows = [
      ['trades', result.trades, now],
      ['weightedAvgZ', result.weightedAvgZ, now],
      ['totalNotional', result.totalNotional, now]
    ];
    if (result.rollingZ) {
      var windows = ['30d', '60d', '90d'];
      for (var i = 0; i < windows.length; i++) {
        var w = result.rollingZ[windows[i]];
        if (w) {
          rows.push(['rollingZ_' + windows[i] + '_z', w.z, now]);
          rows.push(['rollingZ_' + windows[i] + '_mean', w.mean, now]);
          rows.push(['rollingZ_' + windows[i] + '_std', w.std, now]);
          rows.push(['rollingZ_' + windows[i] + '_dataPoints', w.dataPoints, now]);
        }
      }
    }
    if (result.sectorPct) {
      for (var sec in result.sectorPct) {
        rows.push(['sector_' + sec, result.sectorPct[sec], now]);
      }
    }
    if (result.strategyMix) {
      rows.push(['strategy_intra', result.strategyMix.intra, now]);
      rows.push(['strategy_credit', result.strategyMix.credit, now]);
    }
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
    Logger.log('updateBasketAnalytics: Cached ' + rows.length + ' entries for ' + result.trades + ' trades.');
  } catch(e) {
    Logger.log('updateBasketAnalytics ERROR: ' + e.message);
  }
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
    '3: Credit Pair Generation', '4: TickerData + TickerHistory',
    '5: Credit Cache Computation', '6: Triggers', '7: DONE'
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
 * Fetches current prices from Yahoo Finance for a list of tickers.
 * Uses v7/finance/quote (batch, with crumb auth) as primary,
 * falls back to v8/finance/chart (individual, no auth) for misses.
 * Returns {TICKER_UPPER: price} map (only tickers with valid prices).
 */
function fetchYahooPrices_(tickers) {
  var prices = {};
  if (!tickers || tickers.length === 0) return prices;

  // Deduplicate
  var seen = {};
  var unique = [];
  for (var i = 0; i < tickers.length; i++) {
    var key = String(tickers[i]).trim().toUpperCase();
    if (key && !seen[key]) {
      seen[key] = true;
      unique.push(tickers[i]);
    }
  }

  // Build Yahoo symbol → our ticker map
  var tickerMap = {};
  for (var i = 0; i < unique.length; i++) {
    tickerMap[toYahooTicker_(unique[i]).toUpperCase()] = unique[i].toUpperCase();
  }

  // Phase 1: v7/finance/quote batch (50 tickers per request)
  var auth = getYahooCrumb_();
  var BATCH = 50;

  if (auth.crumb) {
    for (var b = 0; b < unique.length; b += BATCH) {
      var batch = unique.slice(b, Math.min(b + BATCH, unique.length));
      var symbols = batch.map(function(t) { return toYahooTicker_(t); }).join(',');
      var url = 'https://query2.finance.yahoo.com/v7/finance/quote?symbols=' +
                encodeURIComponent(symbols) + '&fields=regularMarketPrice&crumb=' + encodeURIComponent(auth.crumb);
      try {
        var resp = UrlFetchApp.fetch(url, {
          headers: {
            'Cookie': auth.cookies,
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
          },
          muteHttpExceptions: true
        });
        if (resp.getResponseCode() === 200) {
          var json = JSON.parse(resp.getContentText());
          var results = (json.quoteResponse && json.quoteResponse.result) || [];
          for (var i = 0; i < results.length; i++) {
            var q = results[i];
            var yahooSym = String(q.symbol).toUpperCase();
            var ourTicker = tickerMap[yahooSym];
            if (ourTicker && q.regularMarketPrice > 0) {
              prices[ourTicker] = q.regularMarketPrice;
            }
          }
        }
      } catch (e) {
        Logger.log('fetchYahooPrices_ Phase 1 batch error at offset ' + b + ': ' + e.toString());
      }
      if (b + BATCH < unique.length) Utilities.sleep(300);
    }
  }

  // Phase 2: v8/finance/chart fallback for misses (no auth needed)
  var misses = [];
  for (var i = 0; i < unique.length; i++) {
    var key = unique[i].toUpperCase();
    if (!prices[key]) misses.push(unique[i]);
  }

  if (misses.length > 0) {
    var CHART_BATCH = 20;
    for (var b = 0; b < misses.length; b += CHART_BATCH) {
      var batch = misses.slice(b, Math.min(b + CHART_BATCH, misses.length));
      var requests = [];
      for (var i = 0; i < batch.length; i++) {
        requests.push({
          url: 'https://query1.finance.yahoo.com/v8/finance/chart/' +
               encodeURIComponent(toYahooTicker_(batch[i])) + '?interval=1d&range=1d',
          muteHttpExceptions: true,
          headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' }
        });
      }
      try {
        var responses = UrlFetchApp.fetchAll(requests);
        for (var i = 0; i < responses.length; i++) {
          try {
            if (responses[i].getResponseCode() !== 200) continue;
            var json = JSON.parse(responses[i].getContentText());
            var result = json.chart && json.chart.result && json.chart.result[0];
            if (!result || !result.meta || !result.meta.regularMarketPrice) continue;
            var price = result.meta.regularMarketPrice;
            if (price > 0) {
              prices[batch[i].toUpperCase()] = price;
            }
          } catch (e) { /* skip individual ticker errors */ }
        }
      } catch (e) {
        Logger.log('fetchYahooPrices_ Phase 2 batch error at offset ' + b + ': ' + e.toString());
      }
      if (b + CHART_BATCH < misses.length) Utilities.sleep(500);
    }
  }

  Logger.log('fetchYahooPrices_: got prices for ' + Object.keys(prices).length + '/' + unique.length + ' tickers');
  return prices;
}

/**
 * Compares GOOGLEFINANCE prices against Yahoo Finance prices.
 * Returns a map of {TICKER_UPPER: yahooPrice} for tickers where the prices
 * diverge by more than the given threshold (default 2%).
 * Logs all corrections for audit trail.
 */
function detectPriceDiscrepancies_(gfPrices, yahooPrices, threshold) {
  if (!threshold) threshold = 0.02; // 2% default
  var corrections = {};
  var count = 0;

  for (var ticker in gfPrices) {
    var gfPrice = gfPrices[ticker];
    var yPrice = yahooPrices[ticker];
    if (!yPrice || yPrice <= 0 || !gfPrice || gfPrice <= 0) continue;

    var pctDiff = Math.abs(gfPrice - yPrice) / yPrice;
    if (pctDiff > threshold) {
      corrections[ticker] = yPrice;
      count++;
      Logger.log('PRICE CORRECTION: ' + ticker + ' GF=$' + gfPrice.toFixed(2) +
                 ' → Yahoo=$' + yPrice.toFixed(2) + ' (diff=' + (pctDiff * 100).toFixed(1) + '%)');
    }
  }

  if (count > 0) {
    Logger.log('detectPriceDiscrepancies_: ' + count + ' corrections needed out of ' + Object.keys(gfPrices).length + ' tickers');
  } else {
    Logger.log('detectPriceDiscrepancies_: all prices within ' + (threshold * 100) + '% tolerance');
  }
  return corrections;
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
    divSheet.getRange(1, 1, 1, 4).setValues([['Ticker', 'NextDivDate', 'LastFetched', 'IsEstimated']]);
    divSheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  }
  // Ensure 4th column header exists (upgrade existing sheets)
  if (divSheet.getLastColumn() < 4) {
    divSheet.getRange(1, 4).setValue('IsEstimated').setFontWeight('bold');
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
            var result = projectNextDivDate_(new Date(rawTs * 1000), now);
            if (result) {
              dataMap[key] = [ourTicker, result.date, now, result.isEstimated ? 'ESTIMATED' : 'CONFIRMED'];
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
              var result = projectNextDivDate_(new Date(latest * 1000), now);
              if (result) {
                dataMap[key] = [ticker, result.date, now, result.isEstimated ? 'ESTIMATED' : 'CONFIRMED'];
                phase2Count++;
                if (phase2Count === 1) {
                  Logger.log('Phase 2 first success: ' + ticker + ' → ' + result.date.toISOString().split('T')[0] + (result.isEstimated ? ' (estimated)' : ' (confirmed)'));
                }
                continue;
              }
            }
            dataMap[key] = [ticker, '', now, ''];
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
  divSheet.getRange(1, 1, 1, 4).setValues([['Ticker', 'NextDivDate', 'LastFetched', 'IsEstimated']]);
  divSheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  if (allRows.length > 0) {
    // Ensure all rows have 4 columns
    for (var ri = 0; ri < allRows.length; ri++) {
      while (allRows[ri].length < 4) allRows[ri].push('');
    }
    divSheet.getRange(2, 1, allRows.length, 4).setValues(allRows);
  }

  Logger.log('fetchDividendDates: Done. ' + successCount + '/' + toFetch.length + ' got dates. ' + allRows.length + ' total tickers stored.');
}

/**
 * Projects the next dividend date from a known dividend date.
 * If the date is in the past, adds 91 days (quarterly) until it's >= today.
 * Returns null if the result is unreasonable (> 1 year out).
 */
/**
 * Projects the next dividend date from a known dividend date.
 * Returns {date, isEstimated} — isEstimated=true if the date was projected
 * (not directly from Yahoo Finance). Projected dates add 91 days (quarterly).
 */
function projectNextDivDate_(divDate, today) {
  var d = new Date(divDate.getTime());
  var maxFuture = new Date(today.getTime() + 400 * 86400000); // ~13 months max

  // If already in the future, return as confirmed (directly from source)
  if (d >= today) return { date: d, isEstimated: false };

  // Add 91 days (quarterly) until it's in the future — this is an ESTIMATE
  while (d < today) {
    d = new Date(d.getTime() + 91 * 86400000);
  }

  // Sanity check: don't return dates too far out
  if (d > maxFuture) return null;
  return { date: d, isEstimated: true };
}

// ============================================================
// LIGHT UPDATE — snapshots GOOGLEFINANCE values to WebCache (every 10 min)
// ============================================================

/**
 * Light update function. Runs every 10 minutes via trigger.
 * Reads computed GOOGLEFINANCE values from Live sheet (intra)
 * and writes them as static values to WebCache/WebCacheCredit.
 * This decouples the API from GOOGLEFINANCE recalculation timing.
 * Target execution: <30 seconds.
 */
function updateLivePrices() {
  try {
    var ss = SpreadsheetApp.getActive();
    var updated = 0;

    // --- INTRA-COMPANY: snapshot Live → WebCache (unchanged, small dataset) ---
    var src = ss.getSheetByName('Live');
    if (src && src.getLastRow() > 1) {
      var data = src.getDataRange().getValues();
      if (data.length > 1) {
        // Sanity check: at least some rows should have non-zero prices
        var validPrices = 0;
        for (var r = 1; r < data.length && r < 20; r++) {
          if (parseFloat(data[r][3]) > 0) validPrices++;
        }
        if (validPrices > 0) {
          // --- Yahoo Finance price validation failsafe ---
          // Extract unique tickers and their GOOGLEFINANCE prices
          var gfPrices = {};
          var allTickers = [];
          for (var r = 1; r < data.length; r++) {
            var tA = String(data[r][1]).trim().toUpperCase();
            var tB = String(data[r][2]).trim().toUpperCase();
            var pA = parseFloat(data[r][3]) || 0;
            var pB = parseFloat(data[r][4]) || 0;
            if (tA && pA > 0) { gfPrices[tA] = pA; allTickers.push(data[r][1]); }
            if (tB && pB > 0) { gfPrices[tB] = pB; allTickers.push(data[r][2]); }
          }

          var yahooPrices = fetchYahooPrices_(allTickers);
          var corrections = detectPriceDiscrepancies_(gfPrices, yahooPrices, 0.02);

          // Apply corrections to snapshot data before writing to WebCache
          if (Object.keys(corrections).length > 0) {
            for (var r = 1; r < data.length; r++) {
              var tA = String(data[r][1]).trim().toUpperCase();
              var tB = String(data[r][2]).trim().toUpperCase();
              var changed = false;
              if (corrections[tA]) { data[r][3] = corrections[tA]; changed = true; }
              if (corrections[tB]) { data[r][4] = corrections[tB]; changed = true; }
              if (changed) {
                // Recalculate spread (col 5) and Z-score (col 12)
                data[r][5] = data[r][3] - data[r][4];
                var stdev = parseFloat(data[r][11]) || 0.001;
                if (stdev > 0.001) {
                  data[r][12] = (data[r][5] - (parseFloat(data[r][10]) || 0)) / stdev;
                }
              }
            }
          }

          var cache = ss.getSheetByName('WebCache');
          if (!cache) { cache = ss.insertSheet('WebCache'); } else { cache.clear(); }
          cache.getRange(1, 1, 1, data[0].length).setValues([data[0]]);
          cache.getRange(1, 1, 1, data[0].length).setFontWeight('bold');
          var rows = data.slice(1);
          var CHUNK = 500;
          for (var c = 0; c < rows.length; c += CHUNK) {
            var chunk = rows.slice(c, Math.min(c + CHUNK, rows.length));
            cache.getRange(c + 2, 1, chunk.length, data[0].length).setValues(chunk);
          }
          SpreadsheetApp.flush();
          updated += rows.length;
          Logger.log('updateLivePrices: cached ' + rows.length + ' intra rows (Live → WebCache)');
        } else {
          Logger.log('updateLivePrices: Live has no valid prices (GOOGLEFINANCE still loading?). Keeping existing WebCache.');
        }
      }
    }

    // --- CREDIT: compute all pairs from ticker-level data → WebCacheCredit ---
    // This replaces the old CreditLive → WebCacheCredit snapshot.
    // computeCreditCache reads TickerData + TickerHistory (GOOGLEFINANCE per ticker)
    // and computes ALL credit pair stats in GAS memory. No per-pair GOOGLEFINANCE needed.
    computeCreditCache();

    // Store timestamp for diagnostics
    PropertiesService.getScriptProperties().setProperty('WEBCACHE_UPDATED', new Date().toISOString());
    Logger.log('updateLivePrices: complete.');

  } catch (e) {
    Logger.log('updateLivePrices ERROR: ' + e.toString());
  }
}

// ============================================================
// CREDIT PAIR GENERATION (headless, used by setupAllBatched)
// ============================================================
/** Returns the number of pairs generated */
function generateCreditPairsBatched_(ss) {
  var result = generateCreditPairData_(ss);
  if (result.allPairs.length === 0) return 0;
  writeCreditPairsToSheet_(ss, result.allPairs);
  SpreadsheetApp.flush();
  return result.allPairs.length;
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
      '=IFERROR(STDEV(G' + r + ':' + r + '), 0)',
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
  ensureSheet_(ss, 'OpenTrades', ['PairID', 'EntryZ', 'CostA', 'CostB', 'SizeA', 'SizeB', 'Timestamp', 'PaidDiv', 'ReceivedDiv', 'TargetExitZ', 'ProfitCapturePct', 'TargetPnL', 'PartialAtPct', 'SourcePortfolio', 'MaxHoldDays']);
  ensureSheet_(ss, 'ClosedTrades', ['PairID', 'EntryZ', 'CostA', 'CostB', 'SizeA', 'SizeB', 'OpenDate', 'CloseDate', 'PnL', 'ExitPriceA', 'ExitPriceB', 'ExitZ', 'CloseType', 'PaidDiv', 'ReceivedDiv', 'Notes', 'TargetExitZ', 'TargetPnL', 'SourcePortfolio']);
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
function generateCrossRatingPairs_(groups, companyOf, intraPairs, seen) {
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

  // Build set of all valid tickers across all rated groups
  var allValid = {};
  for (var r in groups) {
    for (var i = 0; i < groups[r].length; i++) allValid[groups[r][i].toUpperCase()] = groups[r][i];
  }

  var bbbPlus = groups['BBB+'] || [];
  var bbb = filt(groups['BBB'], exBBB);
  var bbbMinus = filt(groups['BBB-'], exBBBm);
  var bb = filt(groups['BB'], exBB);
  // Include-only pools: match against all valid Master tickers
  var bbPlusPool = [];
  for (var k in incBBp) { if (allValid[k]) bbPlusPool.push(allValid[k]); }
  var bbMinusPool = [];
  for (var k in incBBm) { if (allValid[k]) bbMinusPool.push(allValid[k]); }

  var pairs = [];
  var addPairs = function(poolA, poolB) {
    for (var a = 0; a < poolA.length; a++) {
      for (var b = 0; b < poolB.length; b++) {
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
  // 11. BB- (include-only) vs BBB, BBB-, BB+ (include-only), BB — NO Non-Rated
  var rule11Pool = bbb.concat(bbbMinus).concat(bbPlusPool).concat(bb);
  addPairs(bbMinusPool, rule11Pool);
  // Rule 12 (Non-Rated vs Non-Rated) DELETED — too much noise

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

  var result = generateCreditPairData_(ss);
  var allPairs = result.allPairs;
  writeCreditPairsToSheet_(ss, allPairs);

  var summary = [];
  var ratingCounts = {};
  allPairs.forEach(function(p) { var r = p[3]; ratingCounts[r] = (ratingCounts[r] || 0) + 1; });
  for (var r in ratingCounts) summary.push(r + ': ' + ratingCounts[r]);
  showMsg_('Generated ' + allPairs.length + ' credit arb pairs.\n\nBreakdown:\n' + summary.join('\n') + '\n\nNext: run setupCreditSheets().');
}

// 3. CREDIT SHEETS SETUP (now uses ticker-level approach)
function setupCreditSheets() {
  var ss = SpreadsheetApp.getActive();
  var cpSheet = ss.getSheetByName('CreditPairs');
  if (!cpSheet) { showMsg_('ERROR: "CreditPairs" sheet not found. Run generateCreditPairs() first.'); return; }
  var numPairs = cpSheet.getLastRow() - 1;
  if (numPairs < 1) { showMsg_('ERROR: CreditPairs is empty.'); return; }

  // Build ticker-level sheets (GOOGLEFINANCE per unique ticker, not per pair)
  var tickers = buildTickerDataSheet_(ss);
  buildTickerHistorySheet_(ss, tickers);
  SpreadsheetApp.flush();

  // Wait for GOOGLEFINANCE to start loading, then compute credit cache
  Utilities.sleep(5000);
  SpreadsheetApp.flush();
  computeCreditCache();

  showMsg_('Credit sheets built!\n\n• CreditPairs: ' + numPairs + ' pairs\n• TickerData: ' + tickers.length + ' tickers\n• TickerHistory: ' + tickers.length + ' tickers\n• WebCacheCredit: computed from ticker data\n\nAll pair stats computed in GAS — no per-pair GOOGLEFINANCE formulas needed.');
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
      { name: 'WebCache', source: 'intra', fallback: 'Live' },
      { name: 'WebCacheCredit', source: 'credit', fallback: null }
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
      if ((!liveSheet || liveSheet.getLastRow() <= 1) && sheets[s].fallback) {
        liveSheet = ss.getSheetByName(sheets[s].fallback);
      }
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
    Logger.log("snapshotZScores error: " + e);
  }
}

// ============================================================
// 6. DAILY CREDIT REFRESH (Daily Trigger) — now with batch writes
// ============================================================
function dailyCreditRefresh() {
  try {
    var ss = SpreadsheetApp.getActive();
    var result = generateCreditPairData_(ss);
    var allPairs = result.allPairs;
    writeCreditPairsToSheet_(ss, allPairs);
    Logger.log('dailyCreditRefresh: wrote ' + allPairs.length + ' pairs to CreditPairs.');

    // Refresh TickerData + TickerHistory (picks up any new tickers added to Master)
    var tickers = buildTickerDataSheet_(ss);
    buildTickerHistorySheet_(ss, tickers);
    SpreadsheetApp.flush();

    // Compute credit cache immediately (pair stats from ticker-level data)
    Utilities.sleep(3000);
    SpreadsheetApp.flush();
    computeCreditCache();

    Logger.log('dailyCreditRefresh: complete. ' + allPairs.length + ' credit pairs, ' + tickers.length + ' tickers.');
  } catch (e) {
    Logger.log("dailyCreditRefresh error: " + e);
  }
}

// ============================================================
// CREDIT PAIR GENERATION (shared logic)
// ============================================================

/**
 * Generates credit arb pairs from Master + Pairs sheets.
 * Returns { allPairs: [[key, tA, tB, sector], ...], companyOf: {}, intraPairs: {} }
 */
function generateCreditPairData_(ss) {
  var master = ss.getSheetByName('Master');
  if (!master) return { allPairs: [], companyOf: {}, intraPairs: {} };
  var data = master.getDataRange().getValues();
  if (data.length < 2) return { allPairs: [], companyOf: {}, intraPairs: {} };

  // Build union-find company map from Pairs sheet
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

  // Group tickers by credit rating
  var groups = {};
  for (var i = 1; i < data.length; i++) {
    var ticker = String(data[i][0]).trim();
    var coupon = data[i][2]; var curYield = data[i][3];
    var rawRating = data[i][4];
    var rating = (rawRating != null && rawRating !== '') ? String(rawRating).trim() : '';
    if (rating === 'undefined' || rating === 'null') rating = '';
    if (!ticker) continue;
    if (coupon === "" || coupon === null || coupon === undefined) continue;
    if (curYield === "" || curYield === null || curYield === undefined) continue;
    var upper = ticker.toUpperCase();
    if (!companyOf[upper]) companyOf[upper] = upper.replace(/-.*/, '');
    if (!rating || rating === 'Non-Rated' || rating === 'NR') continue;
    if (!groups[rating]) groups[rating] = [];
    groups[rating].push(ticker);
  }

  // Same-rating pairs with dedup
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

  // Cross-rating pairs
  var crossPairs = generateCrossRatingPairs_(groups, companyOf, intraPairs, seen);
  allPairs = allPairs.concat(crossPairs);

  return { allPairs: allPairs, companyOf: companyOf, intraPairs: intraPairs };
}

/**
 * Writes credit pairs to CreditPairs sheet.
 */
function writeCreditPairsToSheet_(ss, allPairs) {
  var cpSheet = getOrCreateSheet_(ss, 'CreditPairs');
  cpSheet.getRange(1, 1, 1, 4).setValues([['PairID', 'TickerA', 'TickerB', 'Sector']]);
  cpSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#2d1a2e').setFontColor('#ffffff');
  if (allPairs.length > 0) {
    cpSheet.getRange(2, 1, allPairs.length, 4).setValues(allPairs);
  }
  return allPairs.length;
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
// ============================================================
// TICKER-LEVEL SHEETS — GOOGLEFINANCE per ticker, not per pair
// ============================================================

/**
 * Builds TickerData sheet: one row per unique ticker with live GOOGLEFINANCE formulas.
 * ~200 tickers × 3 formulas = ~600 GOOGLEFINANCE calls (vs 29K × 6 = 174K in CreditLive).
 * Returns array of ticker strings for use by buildTickerHistorySheet_.
 */
function buildTickerDataSheet_(ss) {
  var master = ss.getSheetByName('Master');
  if (!master) return [];
  var data = master.getDataRange().getValues();
  var tickers = [];
  for (var i = 1; i < data.length; i++) {
    var t = String(data[i][0]).trim();
    var coupon = data[i][2];
    var curYield = data[i][3];
    if (!t) continue;
    if (coupon === '' || coupon === null || coupon === undefined) continue;
    if (curYield === '' || curYield === null || curYield === undefined) continue;
    tickers.push(t);
  }
  if (tickers.length === 0) return [];

  var sheet = getOrCreateSheet_(ss, 'TickerData');
  var headers = ['Ticker', 'Price', 'VolAvg', 'Volume'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');

  // Column A: static ticker names
  var tickerValues = tickers.map(function(t) { return [t]; });
  sheet.getRange(2, 1, tickers.length, 1).setValues(tickerValues);

  // Columns B-D: GOOGLEFINANCE formulas
  var formulas = [];
  for (var i = 0; i < tickers.length; i++) {
    var r = i + 2;
    formulas.push([
      '=IFERROR(GOOGLEFINANCE(A' + r + '),0)',
      '=IFERROR(GOOGLEFINANCE(A' + r + ',"volumeavg"),0)',
      '=IFERROR(GOOGLEFINANCE(A' + r + ',"volume"),0)'
    ]);
  }
  sheet.getRange(2, 2, tickers.length, 3).setFormulas(formulas);
  sheet.setFrozenRows(1);
  Logger.log('buildTickerDataSheet_: ' + tickers.length + ' tickers (' + (tickers.length * 3) + ' GOOGLEFINANCE calls)');
  return tickers;
}

/**
 * Builds TickerHistory sheet: one row per unique ticker with 1-year GOOGLEFINANCE historical prices.
 * ~200 tickers × 1 formula = ~200 GOOGLEFINANCE calls.
 * Each formula expands horizontally into ~252 price columns.
 */
function buildTickerHistorySheet_(ss, tickers) {
  if (!tickers || tickers.length === 0) return;

  var sheet = getOrCreateSheet_(ss, 'TickerHistory');
  sheet.getRange(1, 1, 1, 2).setValues([['Ticker', 'Historical Prices →']]);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');

  // Column A: static ticker names
  var tickerValues = tickers.map(function(t) { return [t]; });
  sheet.getRange(2, 1, tickers.length, 1).setValues(tickerValues);

  // Column B: GOOGLEFINANCE historical (expands horizontally)
  var histFormulas = tickers.map(function(t, i) {
    var r = i + 2;
    return ['=IFERROR(TRANSPOSE(QUERY(GOOGLEFINANCE(A' + r + ',"price",TODAY()-370,TODAY()),"select Col2 offset 1",0)),)'];
  });
  sheet.getRange(2, 2, tickers.length, 1).setFormulas(histFormulas);
  sheet.setFrozenRows(1);
  Logger.log('buildTickerHistorySheet_: ' + tickers.length + ' tickers (' + tickers.length + ' GOOGLEFINANCE calls)');
}

/**
 * Computes ALL credit pair statistics from ticker-level data and writes to WebCacheCredit.
 * This replaces the CreditLevels + CreditLive formula approach entirely.
 *
 * Flow:
 *   1. Read TickerData (live prices/volumes per ticker) — ~200 rows
 *   2. Read TickerHistory (1-year prices per ticker) — ~200 rows × ~252 cols
 *   3. Read Master (yields, coupons per ticker) — ~200 rows
 *   4. Read CreditPairs (all pair definitions) — up to 30K rows
 *   5. For each pair: compute spread, mean, stdev, Z-score, percentiles, vol spike
 *   6. Write results to WebCacheCredit (same 24-column layout)
 *
 * Total GOOGLEFINANCE calls: ~800 (in TickerData + TickerHistory sheets)
 * vs old approach: 29K × 8 = 232K (in CreditLevels + CreditLive sheets)
 */
function computeCreditCache() {
  try {
    var ss = SpreadsheetApp.getActive();

    // 1. Read TickerData (live prices, volumes)
    var tdSheet = ss.getSheetByName('TickerData');
    if (!tdSheet || tdSheet.getLastRow() <= 1) {
      Logger.log('computeCreditCache: TickerData not found or empty. Run setup first.');
      return;
    }
    var tdData = tdSheet.getDataRange().getValues();
    var tickerMap = {}; // {TICKER: {price, volAvg, volume}}
    for (var i = 1; i < tdData.length; i++) {
      var t = String(tdData[i][0]).trim().toUpperCase();
      if (!t) continue;
      tickerMap[t] = {
        price: Number(tdData[i][1]) || 0,
        volAvg: Number(tdData[i][2]) || 0,
        volume: Number(tdData[i][3]) || 0
      };
    }

    // 1b. Yahoo Finance price validation failsafe
    var gfPrices = {};
    var allCreditTickers = [];
    for (var t in tickerMap) {
      if (tickerMap[t].price > 0) {
        gfPrices[t] = tickerMap[t].price;
        allCreditTickers.push(t);
      }
    }
    var yahooPrices = fetchYahooPrices_(allCreditTickers);
    var corrections = detectPriceDiscrepancies_(gfPrices, yahooPrices, 0.02);
    for (var t in corrections) {
      if (tickerMap[t]) {
        tickerMap[t].price = corrections[t];
      }
    }

    // 2. Read TickerHistory (1-year prices per ticker)
    var thSheet = ss.getSheetByName('TickerHistory');
    if (!thSheet || thSheet.getLastRow() <= 1) {
      Logger.log('computeCreditCache: TickerHistory not found or empty. Run setup first.');
      return;
    }
    var thData = thSheet.getDataRange().getValues();
    var histMap = {}; // {TICKER: [price1, price2, ...]}
    for (var i = 1; i < thData.length; i++) {
      var t = String(thData[i][0]).trim().toUpperCase();
      if (!t) continue;
      var prices = [];
      for (var j = 1; j < thData[i].length; j++) {
        var v = thData[i][j];
        if (v !== '' && v !== null && v !== undefined && !isNaN(v) && Number(v) !== 0) {
          prices.push(Number(v));
        }
      }
      histMap[t] = prices;
    }

    // 3. Read Master (yields, coupons)
    var masterSheet = ss.getSheetByName('Master');
    if (!masterSheet) {
      Logger.log('computeCreditCache: Master sheet not found.');
      return;
    }
    var masterData = masterSheet.getDataRange().getValues();
    var masterMap = {}; // {TICKER: {coupon, curYield}}
    for (var i = 1; i < masterData.length; i++) {
      var t = String(masterData[i][0]).trim().toUpperCase();
      if (!t) continue;
      masterMap[t] = {
        coupon: masterData[i][2],
        curYield: masterData[i][3]
      };
    }

    // 4. Read CreditPairs (all pair definitions)
    var cpSheet = ss.getSheetByName('CreditPairs');
    if (!cpSheet || cpSheet.getLastRow() <= 1) {
      Logger.log('computeCreditCache: CreditPairs not found or empty.');
      return;
    }
    var cpData = cpSheet.getDataRange().getValues();
    Logger.log('computeCreditCache: processing ' + (cpData.length - 1) + ' pairs...');

    // 5. Compute all pairs
    var results = [];
    var validPrices = 0;
    for (var i = 1; i < cpData.length; i++) {
      var pairId = cpData[i][0];
      var tickerA = String(cpData[i][1]).trim();
      var tickerB = String(cpData[i][2]).trim();
      var sector = cpData[i][3] || '';
      if (!tickerA || !tickerB) continue;

      var tA = tickerA.toUpperCase();
      var tB = tickerB.toUpperCase();
      var dA = tickerMap[tA] || { price: 0, volAvg: 0, volume: 0 };
      var dB = tickerMap[tB] || { price: 0, volAvg: 0, volume: 0 };
      var mA = masterMap[tA] || { coupon: '', curYield: '' };
      var mB = masterMap[tB] || { coupon: '', curYield: '' };

      var priceA = dA.price;
      var priceB = dB.price;
      if (priceA > 0 && priceB > 0) validPrices++;
      var spread = priceA - priceB;

      // Historical spread computation
      // IMPORTANT: Align from the END of each array, not the start.
      // TickerHistory arrays are oldest-first, but different tickers may have
      // different lengths (different listing dates, data gaps). Aligning from
      // the end ensures the most recent entries (same trading days) are paired.
      // Adaptive lookback: use 3× half-life if computable, else default 90 days.
      var hA = histMap[tA] || [];
      var hB = histMap[tB] || [];
      var DEFAULT_LOOKBACK = 90;

      // First pass: compute nominal spread series with all available history for half-life
      var maxAvail = Math.min(hA.length, hB.length);
      var spreadForHL = [];
      for (var k = 0; k < maxAvail; k++) {
        spreadForHL.push(hA[hA.length - maxAvail + k] - hB[hB.length - maxAvail + k]);
      }

      // Compute half-life to determine adaptive lookback
      var LOOKBACK = DEFAULT_LOOKBACK;
      if (spreadForHL.length >= 30) {
        // Simple OU half-life: Δy = α + β*y_{t-1}, HL = -ln(2)/β
        var hlSumX = 0, hlSumY = 0, hlSumXX = 0, hlSumXY = 0;
        var hlT = spreadForHL.length - 1;
        for (var k = 0; k < hlT; k++) {
          var x = spreadForHL[k];
          var y = spreadForHL[k + 1] - spreadForHL[k];
          hlSumX += x; hlSumY += y; hlSumXX += x * x; hlSumXY += x * y;
        }
        var hlDenom = hlT * hlSumXX - hlSumX * hlSumX;
        if (Math.abs(hlDenom) > 1e-14) {
          var hlBeta = (hlT * hlSumXY - hlSumX * hlSumY) / hlDenom;
          if (hlBeta < 0) {
            var hl = -Math.log(2) / hlBeta;
            if (hl > 0 && hl < 300) {
              // Adaptive: lookback = 3× half-life, clamped to [30, 180]
              LOOKBACK = Math.max(30, Math.min(180, Math.round(hl * 3)));
            }
          }
        }
      }

      var minLen = Math.min(hA.length, hB.length, LOOKBACK);
      var spreads = [];
      for (var k = 0; k < minLen; k++) {
        spreads.push(hA[hA.length - minLen + k] - hB[hB.length - minLen + k]);
      }

      var histCount = spreads.length;
      var mean = 0, stdev = 0.001, lower = 0, upper = 0;
      if (histCount > 0) {
        // Mean
        var sum = 0;
        for (var k = 0; k < histCount; k++) sum += spreads[k];
        mean = sum / histCount;

        // StDev (sample)
        var sumSq = 0;
        for (var k = 0; k < histCount; k++) sumSq += (spreads[k] - mean) * (spreads[k] - mean);
        stdev = histCount > 1 ? Math.sqrt(sumSq / (histCount - 1)) : 0.001;
        if (stdev < 0.001) stdev = 0.001;

        // Percentiles (5th and 95th)
        var sorted = spreads.slice().sort(function(a, b) { return a - b; });
        var idx5 = Math.floor(histCount * 0.05);
        var idx95 = Math.floor(histCount * 0.95);
        lower = sorted[Math.max(0, idx5)];
        upper = sorted[Math.min(histCount - 1, idx95)];
      }

      var zScore = (spread - mean) / stdev;

      var volAvgA = dA.volAvg;
      var volAvgB = dB.volAvg;
      var avgLiq = (volAvgA + volAvgB) / 2;
      var curVolA = dA.volume;
      var curVolB = dB.volume;
      var curVol = (curVolA + curVolB) / 2;
      var volSpike = (avgLiq > 0 && curVol > 1.5 * avgLiq);

      results.push([
        pairId, tickerA, tickerB, priceA, priceB, spread,
        mA.curYield !== '' ? mA.curYield : '', mB.curYield !== '' ? mB.curYield : '',
        mA.coupon !== '' ? mA.coupon : '', mB.coupon !== '' ? mB.coupon : '',
        mean, stdev, zScore, lower, upper,
        sector, histCount,
        volAvgA, volAvgB, avgLiq, curVolA, curVolB, curVol, volSpike
      ]);
    }

    // Sanity check: if TickerData hasn't loaded yet, don't overwrite
    if (validPrices === 0) {
      Logger.log('computeCreditCache: no valid prices found (TickerData still loading?). Skipping write.');
      return;
    }

    // 6. Write to WebCacheCredit (overwrite, don't clear first to minimize downtime)
    var wcSheet = ss.getSheetByName('WebCacheCredit');
    if (!wcSheet) {
      wcSheet = ss.insertSheet('WebCacheCredit');
    }
    var headers = [
      'PairID', 'TickerA', 'TickerB', 'PriceA', 'PriceB', 'Spread',
      'YieldA', 'YieldB', 'CouponA', 'CouponB',
      'Mean', 'StDev', 'Z-Score', 'Lower', 'Upper',
      'Sector', 'HistCount',
      'AvgLiqA', 'AvgLiqB', 'AvgLiq', 'CurVolA', 'CurVolB', 'CurVol', 'VolSpike'
    ];
    wcSheet.getRange(1, 1, 1, 24).setValues([headers]);
    wcSheet.getRange(1, 1, 1, 24).setFontWeight('bold');

    if (results.length > 0) {
      // Write in chunks of 5000 rows
      var CHUNK = 5000;
      for (var c = 0; c < results.length; c += CHUNK) {
        var chunk = results.slice(c, Math.min(c + CHUNK, results.length));
        wcSheet.getRange(c + 2, 1, chunk.length, 24).setValues(chunk);
      }
    }

    // Clear leftover rows from previous run (if pair count decreased)
    var lastRow = wcSheet.getLastRow();
    if (lastRow > results.length + 1) {
      wcSheet.getRange(results.length + 2, 1, lastRow - results.length - 1, 24).clearContent();
    }

    SpreadsheetApp.flush();
    Logger.log('computeCreditCache: computed ' + results.length + ' pairs (' + validPrices + ' with valid prices), wrote to WebCacheCredit.');

  } catch (e) {
    Logger.log('computeCreditCache ERROR: ' + e.toString());
  }
}

// ============================================================
// SHARED HELPERS
// ============================================================
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
          var stdev = calcSpreads.length > 1 ? Math.sqrt(sqSum / (calcSpreads.length - 1)) : 0.001;
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
