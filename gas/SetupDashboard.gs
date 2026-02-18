/**
 * TERMINAL PRO — V23.0 SETUP SCRIPT
 *
 * WHAT'S NEW:
 *   - ZScoreAge sheet: Tracks when each pair first crossed ±1.5 threshold
 *   - CreditPairs / CreditLevels / CreditLive: Inter-company credit rating arb
 *   - setupTriggers() now installs hourly snapshot + daily credit pair refresh
 *
 * RUN ORDER:
 *   1. setupDashboard()       — Builds Levels, Live, supporting sheets
 *   2. generateCreditPairs()  — Scans Master, generates credit arb pairs
 *   3. setupCreditSheets()    — Builds CreditLevels + CreditLive from CreditPairs
 *   4. setupTriggers()        — Installs hourly + daily automation
 *
 * Or just run setupAll() to do everything in order.
 */
function setupAll() {
  setupDashboard();
  generateCreditPairs();
  setupCreditSheets();
  setupTriggers();
}
// ============================================================
// 1. INTRA-COMPANY SETUP (same as V22 + ZScoreAge)
// ============================================================
function setupDashboard() {
  var ss = SpreadsheetApp.getActive();
  var pairsSheet = ss.getSheetByName('Pairs');

  if (!pairsSheet) {
    showMsg_('ERROR: "Pairs" sheet not found.\n\nRequired: A: PairID | B: TickerA | C: TickerB | D: Sector');
    return;
  }

  var masterSheet = ss.getSheetByName('Master');
  if (!masterSheet) {
    showMsg_('WARNING: "Master" sheet not found. Yields and coupon filtering will not work. Continuing...');
  }

  var pairs = pairsSheet.getDataRange().getValues();
  var numPairs = pairs.length - 1;
  if (numPairs < 1) {
    showMsg_('ERROR: Pairs sheet is empty.');
    return;
  }

  showMsg_('Setting up ' + numPairs + ' intra-company pairs.\nClick OK to proceed.');

  // --- LEVELS ---
  var levels = getOrCreateSheet_(ss, 'Levels');
  levels.clear();
  levels.getRange(1, 1, 1, 7).setValues([['PairID', 'Mean', 'StDev', 'Lower_5', 'Upper_95', 'HistCount', 'Historical Spread Data →']]);
  levels.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');

  var statsF = [];
  for (var i = 0; i < numPairs; i++) {
    var r = i + 2;
    statsF.push([
      '=Pairs!A' + (i+2),
      '=IFERROR(AVERAGE(G' + r + ':' + r + '), 0)',
      '=IFERROR(STDEV(G' + r + ':' + r + '), 0.001)',
      '=IFERROR(PERCENTILE(G' + r + ':' + r + ', 0.05), 0)',
      '=IFERROR(PERCENTILE(G' + r + ':' + r + ', 0.95), 0)',
      '=COUNTA(G' + r + ':' + r + ')'
    ]);
  }
  levels.getRange(2, 1, statsF.length, 6).setFormulas(statsF);

  for (var i = 0; i < numPairs; i++) {
    var tA = String(pairs[i+1][1]).trim();
    var tB = String(pairs[i+1][2]).trim();
    if (!tA || !tB) continue;
    levels.getRange(i+2, 7).setFormula(
      '=IFERROR(TRANSPOSE(ARRAYFORMULA(QUERY(GOOGLEFINANCE("'+tA+'","price",TODAY()-90,TODAY()),"select Col2 offset 1",0)-QUERY(GOOGLEFINANCE("'+tB+'","price",TODAY()-90,TODAY()),"select Col2 offset 1",0))),)'
    );
  }
  levels.setFrozenRows(1);

  // --- LIVE ---
  buildLiveSheet_(ss, 'Live', 'Pairs', 'Levels', numPairs, pairs);

  // --- SUPPORTING SHEETS ---
  ensureSheet_(ss, 'OpenTrades', ['PairID','EntryZ','CostA','CostB','SizeA','SizeB','Timestamp']);
  ensureSheet_(ss, 'ClosedTrades', ['PairID','EntryZ','CostA','CostB','SizeA','SizeB','OpenDate','CloseDate','PnL']);
  ensureSheet_(ss, 'AlertsLog', ['Timestamp','PairID','Z-Score','Spread']);

  // --- ZSCORE AGE TRACKER (Task 1) ---
  var ageSheet = ss.getSheetByName('ZScoreAge');
  if (!ageSheet) {
    ageSheet = ss.insertSheet('ZScoreAge');
    ageSheet.getRange(1, 1, 1, 3).setValues([['PairID', 'FirstCrossTimestamp', 'Source']]);
    ageSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  }
  // Don't clear — we want to preserve timestamps

  SpreadsheetApp.flush();
  showMsg_('✅ Intra-company setup complete for ' + numPairs + ' pairs.\n\nNext: run generateCreditPairs() then setupCreditSheets().');
}
// ============================================================
// 2. CREDIT RATING ARBITRAGE — Pair Generation
// ============================================================
/**
 * Reads Master sheet, groups tickers by Credit Rating,
 * generates all combinations within each rating group.
 * Caps total pairs at MAX_CREDIT_PAIRS to avoid quota issues.
 * Writes results to "CreditPairs" sheet.
 */
function generateCreditPairs() {
  var ss = SpreadsheetApp.getActive();
  var master = ss.getSheetByName('Master');

  if (!master) {
    showMsg_('ERROR: "Master" sheet not found.');
    return;
  }

  var data = master.getDataRange().getValues();
  if (data.length < 2) {
    showMsg_('ERROR: Master sheet is empty.');
    return;
  }

  // -------------------------------------------------------------------
  // BUILD COMPANY MAP from Pairs sheet (definitive same-company source)
  // If two tickers appear in the same row of Pairs, they are same-company.
  // We use union-find: every ticker gets a "company label".
  // -------------------------------------------------------------------
  var companyOf = {}; // TICKER → company label
  var intraPairs = {}; // exact pair combos to skip
  var pairsSheet = ss.getSheetByName('Pairs');
  if (pairsSheet) {
    var pData = pairsSheet.getDataRange().getValues();
    for (var i = 1; i < pData.length; i++) {
      var ptA = String(pData[i][1]).trim().toUpperCase();
      var ptB = String(pData[i][2]).trim().toUpperCase();
      if (!ptA || !ptB) continue;

      // Track exact intra pair combos
      intraPairs[cleanId_(ptA) + "_" + cleanId_(ptB)] = true;
      intraPairs[cleanId_(ptB) + "_" + cleanId_(ptA)] = true;

      // Union: assign same company label
      var labelA = companyOf[ptA];
      var labelB = companyOf[ptB];
      if (labelA && labelB) {
        // Merge groups if different
        if (labelA !== labelB) {
          for (var t in companyOf) {
            if (companyOf[t] === labelB) companyOf[t] = labelA;
          }
        }
      } else if (labelA) {
        companyOf[ptB] = labelA;
      } else if (labelB) {
        companyOf[ptA] = labelB;
      } else {
        // New group — use hyphen root as label
        var label = ptA.replace(/-.*/, '');
        companyOf[ptA] = label;
        companyOf[ptB] = label;
      }
    }
  }

  // -------------------------------------------------------------------
  // GROUP TICKERS BY CREDIT RATING
  // -------------------------------------------------------------------
  var groups = {};
  for (var i = 1; i < data.length; i++) {
    var ticker = String(data[i][0]).trim();
    var coupon = data[i][2];
    var curYield = data[i][3];
    var rating = String(data[i][4]).trim();

    if (!ticker || !rating) continue;
    if (coupon === "" || coupon === null || coupon === undefined) continue;
    if (curYield === "" || curYield === null || curYield === undefined) continue;

    // Auto-assign company label if not already mapped from Pairs
    var upper = ticker.toUpperCase();
    if (!companyOf[upper]) {
      companyOf[upper] = upper.replace(/-.*/, '');
    }

    if (!groups[rating]) groups[rating] = [];
    groups[rating].push(ticker);
  }

  // -------------------------------------------------------------------
  // GENERATE COMBINATIONS — no cap, strict same-company filter
  // -------------------------------------------------------------------
  var allPairs = [];
  for (var rating in groups) {
    var tickers = groups[rating];
    if (tickers.length < 2) continue;
    tickers.sort();

    for (var a = 0; a < tickers.length; a++) {
      for (var b = a + 1; b < tickers.length; b++) {
        var tA = tickers[a];
        var tB = tickers[b];

        // Skip exact intra-company pair
        if (intraPairs[cleanId_(tA) + "_" + cleanId_(tB)]) continue;

        // Skip same company (via union-find map)
        var coA = companyOf[tA.toUpperCase()] || tA.toUpperCase().replace(/-.*/, '');
        var coB = companyOf[tB.toUpperCase()] || tB.toUpperCase().replace(/-.*/, '');
        if (coA === coB) continue;

        allPairs.push([
          tA + "|" + tB,
          tA,
          tB,
          "CreditArb:" + rating
        ]);
      }
    }
  }

  // Write to CreditPairs sheet — NO CAP
  var cpSheet = getOrCreateSheet_(ss, 'CreditPairs');
  cpSheet.clear();
  cpSheet.getRange(1, 1, 1, 4).setValues([['PairID', 'TickerA', 'TickerB', 'Sector']]);
  cpSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#2d1a2e').setFontColor('#ffffff');

  if (allPairs.length > 0) {
    cpSheet.getRange(2, 1, allPairs.length, 4).setValues(allPairs);
  }

  // Summary
  var summary = [];
  var ratingCounts = {};
  allPairs.forEach(function(p) {
    var r = p[3].replace('CreditArb:', '');
    ratingCounts[r] = (ratingCounts[r] || 0) + 1;
  });
  for (var r in ratingCounts) summary.push(r + ': ' + ratingCounts[r]);

  showMsg_(
    '✅ Generated ' + allPairs.length + ' credit arb pairs (no cap).\n\n' +
    'Breakdown:\n' + summary.join('\n') + '\n\n' +
    'Same-company pairs filtered via Pairs sheet union + hyphen-root matching.\n\n' +
    'Next: run setupCreditSheets() to build CreditLevels + CreditLive.'
  );
}
// ============================================================
// 3. CREDIT SHEETS SETUP — Levels + Live for credit pairs
// ============================================================
function setupCreditSheets() {
  var ss = SpreadsheetApp.getActive();
  var cpSheet = ss.getSheetByName('CreditPairs');

  if (!cpSheet) {
    showMsg_('ERROR: "CreditPairs" sheet not found. Run generateCreditPairs() first.');
    return;
  }

  var pairs = cpSheet.getDataRange().getValues();
  var numPairs = pairs.length - 1;
  if (numPairs < 1) {
    showMsg_('ERROR: CreditPairs is empty.');
    return;
  }

  // --- CREDIT LEVELS ---
  var cLevels = getOrCreateSheet_(ss, 'CreditLevels');
  cLevels.clear();
  cLevels.getRange(1, 1, 1, 7).setValues([['PairID', 'Mean', 'StDev', 'Lower_5', 'Upper_95', 'HistCount', 'Historical Spread Data →']]);
  cLevels.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#2d1a2e').setFontColor('#ffffff');

  var statsF = [];
  for (var i = 0; i < numPairs; i++) {
    var r = i + 2;
    statsF.push([
      '=CreditPairs!A' + (i+2),
      '=IFERROR(AVERAGE(G' + r + ':' + r + '), 0)',
      '=IFERROR(STDEV(G' + r + ':' + r + '), 0.001)',
      '=IFERROR(PERCENTILE(G' + r + ':' + r + ', 0.05), 0)',
      '=IFERROR(PERCENTILE(G' + r + ':' + r + ', 0.95), 0)',
      '=COUNTA(G' + r + ':' + r + ')'
    ]);
  }
  cLevels.getRange(2, 1, statsF.length, 6).setFormulas(statsF);

  for (var i = 0; i < numPairs; i++) {
    var tA = String(pairs[i+1][1]).trim();
    var tB = String(pairs[i+1][2]).trim();
    if (!tA || !tB) continue;
    cLevels.getRange(i+2, 7).setFormula(
      '=IFERROR(TRANSPOSE(ARRAYFORMULA(QUERY(GOOGLEFINANCE("'+tA+'","price",TODAY()-90,TODAY()),"select Col2 offset 1",0)-QUERY(GOOGLEFINANCE("'+tB+'","price",TODAY()-90,TODAY()),"select Col2 offset 1",0))),)'
    );
  }
  cLevels.setFrozenRows(1);

  // --- CREDIT LIVE ---
  buildLiveSheet_(ss, 'CreditLive', 'CreditPairs', 'CreditLevels', numPairs, pairs);

  SpreadsheetApp.flush();
  showMsg_(
    '✅ Credit sheets built!\n\n' +
    '• CreditLevels: ' + numPairs + ' pairs\n' +
    '• CreditLive: ' + numPairs + ' pairs\n\n' +
    'GOOGLEFINANCE may take 30-60s to populate.\n' +
    'Run setupTriggers() to automate hourly snapshots.'
  );
}
// ============================================================
// 4. TRIGGERS
// ============================================================
function setupTriggers() {
  // Remove existing project triggers for our functions
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === 'snapshotZScores' || fn === 'dailyCreditRefresh') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Hourly: snapshot Z-scores + update age tracking
  ScriptApp.newTrigger('snapshotZScores')
    .timeBased()
    .everyHours(1)
    .create();

  // Daily at 5 AM: regenerate credit pairs from Master
  ScriptApp.newTrigger('dailyCreditRefresh')
    .timeBased()
    .atHour(5)
    .everyDays(1)
    .create();

  showMsg_(
    '✅ Triggers installed!\n\n' +
    '• snapshotZScores(): Every hour (Z-log + age tracking)\n' +
    '• dailyCreditRefresh(): Daily at 5 AM (regenerate credit pairs)'
  );
}
// ============================================================
// 5. SNAPSHOT + AGE TRACKING (Hourly Trigger)
// ============================================================
/**
 * Called every hour. Does two things:
 * 1. Logs Z-Scores to AlertsLog (for trend ribbons)
 * 2. Updates ZScoreAge (for age column)
 *    - If |z| >= 1.5 and no existing timestamp → stamp it
 *    - If |z| < 1.5 and has timestamp → remove it (reset)
 */
function snapshotZScores() {
  try {
    var ss = SpreadsheetApp.getActive();
    var now = new Date();

    // Process both Live and CreditLive
    var sheets = [
      { name: 'Live', source: 'intra' },
      { name: 'CreditLive', source: 'credit' }
    ];

    var logSheet = ss.getSheetByName('AlertsLog');
    var ageSheet = ss.getSheetByName('ZScoreAge');
    if (!logSheet || !ageSheet) return;

    // Load current age data into map
    var ageData = ageSheet.getDataRange().getValues();
    var ageMap = {}; // cleanId → { row: N, timestamp: Date, source: str }
    for (var a = 1; a < ageData.length; a++) {
      var aid = cleanId_(String(ageData[a][0]));
      if (aid) {
        ageMap[aid] = {
          row: a + 1,
          timestamp: ageData[a][1],
          source: ageData[a][2] || ''
        };
      }
    }

    var logRows = [];
    var ageUpdates = []; // { action: 'add'|'remove', pairId, cleanId, source }

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

        var couponA = row[8];
        var couponB = row[9];
        if (couponA === "" || couponA === null || couponB === "" || couponB === null) continue;

        var zScore = parseFloat(row[12]) || 0;
        var spread = parseFloat(row[5]) || 0;
        var cid = cleanId_(String(pairId));

        // Log to AlertsLog
        logRows.push([now, String(pairId), zScore, spread]);

        // Age tracking
        var isActive = Math.abs(zScore) >= 1.5;
        var existing = ageMap[cid];

        if (isActive && !existing) {
          // New crossing — add timestamp
          ageUpdates.push({ action: 'add', pairId: String(pairId), cleanId: cid, source: sheets[s].source });
        } else if (!isActive && existing) {
          // Dropped back to neutral — remove
          ageUpdates.push({ action: 'remove', cleanId: cid, row: existing.row });
        }
      }
    }

    // Write log rows
    if (logRows.length > 0) {
      logSheet.getRange(logSheet.getLastRow() + 1, 1, logRows.length, 4).setValues(logRows);
    }

    // Process age updates — removals first (from bottom up to preserve row indices)
    var rowsToDelete = [];
    for (var u = 0; u < ageUpdates.length; u++) {
      if (ageUpdates[u].action === 'remove') {
        rowsToDelete.push(ageUpdates[u].row);
      }
    }
    rowsToDelete.sort(function(a,b){ return b - a; }); // Reverse order
    for (var d = 0; d < rowsToDelete.length; d++) {
      ageSheet.deleteRow(rowsToDelete[d]);
    }

    // Add new timestamps
    var newAgeRows = [];
    for (var u = 0; u < ageUpdates.length; u++) {
      if (ageUpdates[u].action === 'add') {
        newAgeRows.push([ageUpdates[u].pairId, now, ageUpdates[u].source]);
      }
    }
    if (newAgeRows.length > 0) {
      ageSheet.getRange(ageSheet.getLastRow() + 1, 1, newAgeRows.length, 3).setValues(newAgeRows);
    }

    // Trim AlertsLog
    var totalRows = logSheet.getLastRow();
    if (totalRows > 5500) {
      logSheet.deleteRows(2, totalRows - 5000);
    }

  } catch (e) {
    console.error("snapshotZScores error: " + e);
  }
}
// ============================================================
// 6. DAILY CREDIT REFRESH (Daily Trigger)
// ============================================================
/**
 * Called daily at 5 AM. Regenerates credit pairs from Master
 * and rebuilds CreditLevels + CreditLive.
 * Runs without UI alerts (headless).
 */
function dailyCreditRefresh() {
  try {
    var ss = SpreadsheetApp.getActive();
    var master = ss.getSheetByName('Master');
    if (!master) return;

    var data = master.getDataRange().getValues();
    if (data.length < 2) return;

    // Build company map from Pairs (same logic as generateCreditPairs)
    var companyOf = {};
    var intraPairs = {};
    var pairsSheet = ss.getSheetByName('Pairs');
    if (pairsSheet) {
      var pData = pairsSheet.getDataRange().getValues();
      for (var i = 1; i < pData.length; i++) {
        var ptA = String(pData[i][1]).trim().toUpperCase();
        var ptB = String(pData[i][2]).trim().toUpperCase();
        if (!ptA || !ptB) continue;
        intraPairs[cleanId_(ptA)+"_"+cleanId_(ptB)] = true;
        intraPairs[cleanId_(ptB)+"_"+cleanId_(ptA)] = true;
        var labelA = companyOf[ptA], labelB = companyOf[ptB];
        if (labelA && labelB) {
          if (labelA !== labelB) { for (var t in companyOf) { if (companyOf[t]===labelB) companyOf[t]=labelA; } }
        } else if (labelA) { companyOf[ptB]=labelA; }
        else if (labelB) { companyOf[ptA]=labelB; }
        else { var lbl=ptA.replace(/-.*/, ''); companyOf[ptA]=lbl; companyOf[ptB]=lbl; }
      }
    }

    // Group by rating
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
          if (intraPairs[cleanId_(tA)+"_"+cleanId_(tB)]) continue;
          var coA = companyOf[tA.toUpperCase()] || tA.toUpperCase().replace(/-.*/, '');
          var coB = companyOf[tB.toUpperCase()] || tB.toUpperCase().replace(/-.*/, '');
          if (coA === coB) continue;
          allPairs.push([tA+"|"+tB, tA, tB, "CreditArb:"+rating]);
        }
      }
    }

    // Write CreditPairs
    var cpSheet = getOrCreateSheet_(ss, 'CreditPairs');
    cpSheet.clear();
    cpSheet.getRange(1, 1, 1, 4).setValues([['PairID','TickerA','TickerB','Sector']]);
    if (allPairs.length > 0) {
      cpSheet.getRange(2, 1, allPairs.length, 4).setValues(allPairs);
    }

    // Rebuild credit levels + live (headless version)
    var numPairs = allPairs.length;
    if (numPairs < 1) return;

    // CreditLevels
    var cLevels = getOrCreateSheet_(ss, 'CreditLevels');
    cLevels.clear();
    cLevels.getRange(1,1,1,7).setValues([['PairID','Mean','StDev','Lower_5','Upper_95','HistCount','Historical →']]);
    var sf = [];
    for (var i = 0; i < numPairs; i++) {
      var r = i+2;
      sf.push([
        '=CreditPairs!A'+(i+2),
        '=IFERROR(AVERAGE(G'+r+':'+r+'),0)',
        '=IFERROR(STDEV(G'+r+':'+r+'),0.001)',
        '=IFERROR(PERCENTILE(G'+r+':'+r+',0.05),0)',
        '=IFERROR(PERCENTILE(G'+r+':'+r+',0.95),0)',
        '=COUNTA(G'+r+':'+r+')'
      ]);
    }
    cLevels.getRange(2,1,sf.length,6).setFormulas(sf);
    for (var i = 0; i < numPairs; i++) {
      var tA = String(allPairs[i][1]).trim();
      var tB = String(allPairs[i][2]).trim();
      if (!tA||!tB) continue;
      cLevels.getRange(i+2,7).setFormula(
        '=IFERROR(TRANSPOSE(ARRAYFORMULA(QUERY(GOOGLEFINANCE("'+tA+'","price",TODAY()-90,TODAY()),"select Col2 offset 1",0)-QUERY(GOOGLEFINANCE("'+tB+'","price",TODAY()-90,TODAY()),"select Col2 offset 1",0))),)'
      );
    }

    // CreditLive
    buildLiveSheet_(ss, 'CreditLive', 'CreditPairs', 'CreditLevels', numPairs, [['h']].concat(allPairs));

    SpreadsheetApp.flush();
  } catch (e) {
    console.error("dailyCreditRefresh error: " + e);
  }
}
// ============================================================
// SHARED HELPERS
// ============================================================
/**
 * Builds a Live sheet with the V23 24-column layout.
 * Reusable for both intra-company (Live) and credit (CreditLive).
 */
function buildLiveSheet_(ss, sheetName, pairsRef, levelsRef, numPairs, pairsData) {
  var live = getOrCreateSheet_(ss, sheetName);
  live.clear();

  var headers = [
    'PairID','TickerA','TickerB','PriceA','PriceB','Spread',
    'YieldA','YieldB','CouponA','CouponB',
    'Mean','StDev','Z-Score','Lower','Upper',
    'Sector','HistCount',
    'AvgLiqA','AvgLiqB','AvgLiq','CurVolA','CurVolB','CurVol','VolSpike'
  ];
  live.getRange(1,1,1,headers.length).setValues([headers]);
  live.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');

  var formulas = [];
  for (var i = 0; i < numPairs; i++) {
    var r = i + 2;
    var p = i + 2;
    formulas.push([
      '='+pairsRef+'!A'+p,
      '='+pairsRef+'!B'+p,
      '='+pairsRef+'!C'+p,
      '=IFERROR(GOOGLEFINANCE(B'+r+'),0)',
      '=IFERROR(GOOGLEFINANCE(C'+r+'),0)',
      '=D'+r+'-E'+r,
      '=IFERROR(INDEX(Master!D:D,MATCH(B'+r+',Master!A:A,0)),"")',
      '=IFERROR(INDEX(Master!D:D,MATCH(C'+r+',Master!A:A,0)),"")',
      '=IFERROR(INDEX(Master!C:C,MATCH(B'+r+',Master!A:A,0)),"")',
      '=IFERROR(INDEX(Master!C:C,MATCH(C'+r+',Master!A:A,0)),"")',
      '=IFERROR(INDEX('+levelsRef+'!B:B,MATCH(A'+r+','+levelsRef+'!A:A,0)),0)',
      '=IFERROR(INDEX('+levelsRef+'!C:C,MATCH(A'+r+','+levelsRef+'!A:A,0)),0.001)',
      '=IF(L'+r+'<=0.001,0,(F'+r+'-K'+r+')/L'+r+')',
      '=IFERROR(INDEX('+levelsRef+'!D:D,MATCH(A'+r+','+levelsRef+'!A:A,0)),0)',
      '=IFERROR(INDEX('+levelsRef+'!E:E,MATCH(A'+r+','+levelsRef+'!A:A,0)),0)',
      '='+pairsRef+'!D'+p,
      '=IFERROR(INDEX('+levelsRef+'!F:F,MATCH(A'+r+','+levelsRef+'!A:A,0)),0)',
      '=IFERROR(GOOGLEFINANCE(B'+r+',"volumeavg"),0)',
      '=IFERROR(GOOGLEFINANCE(C'+r+',"volumeavg"),0)',
      '=(R'+r+'+S'+r+')/2',
      '=IFERROR(GOOGLEFINANCE(B'+r+',"volume"),0)',
      '=IFERROR(GOOGLEFINANCE(C'+r+',"volume"),0)',
      '=(U'+r+'+V'+r+')/2',
      '=IF(AND(T'+r+'>0,W'+r+'>1.5*T'+r+'),TRUE,FALSE)'
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
/** Safe alert — shows UI dialog if available, falls back to Logger */
function showMsg_(msg) {
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    Logger.log(msg);
  }
}
