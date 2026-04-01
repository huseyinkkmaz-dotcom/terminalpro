/**
 * TERMINAL PRO — V23.0 (API MODE)
 *
 * NEW:
 *   - ?action=getData&mode=intra  → Intra-company pairs (default)
 *   - ?action=getData&mode=credit → Credit rating arb pairs
 *   - Age column now reads from ZScoreAge sheet (real crossing timestamps)
 *
 * LIVE V23 COLUMN MAP (24 cols, same for Live + CreditLive):
 *   A(0):PairID  B(1):TickerA  C(2):TickerB  D(3):PriceA  E(4):PriceB  F(5):Spread
 *   G(6):YieldA  H(7):YieldB   I(8):CouponA  J(9):CouponB K(10):Mean   L(11):StDev
 *   M(12):Z-Score N(13):Lower  O(14):Upper   P(15):Sector  Q(16):HistCount
 *   R(17):AvgLiqA S(18):AvgLiqB T(19):AvgLiq U(20):CurVolA V(21):CurVolB
 *   W(22):CurVol  X(23):VolSpike
 */
// ============================================================
// ROUTING
// ============================================================
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'getData';
  var mode = (e && e.parameter && e.parameter.mode) ? e.parameter.mode : 'intra';
  var result = {};
  try {
    if (action === 'getData' || action === 'getMacroValuation') {
      setupMacroSheet();
    }
    if (action === 'getData') {
      // WebCache/WebCacheCredit = primary data source (static values, fast).
      // For intra: falls back to Live (GOOGLEFINANCE formulas).
      // For credit: no fallback needed (computeCreditCache writes directly to WebCacheCredit).
      var cacheName = (mode === 'credit') ? 'WebCacheCredit' : 'WebCache';
      var liveName = (mode === 'credit') ? null : 'Live';
      var ss = SpreadsheetApp.getActive();
      var targetSheet = ss.getSheetByName(cacheName);
      if ((!targetSheet || targetSheet.getLastRow() <= 1) && liveName) {
        targetSheet = ss.getSheetByName(liveName);
      }
      if (!targetSheet || targetSheet.getLastRow() <= 1) {
        // No data anywhere — return valid empty response
        var setupState = PropertiesService.getScriptProperties().getProperty('SETUP_STATE');
        var phase = setupState ? JSON.parse(setupState).phase : -1;
        result = {
          ok: true,
          mode: mode,
          status: "initializing",
          statusMessage: "Data not ready. Run setupAllBatched() then updateLivePrices() in Apps Script.",
          setupPhase: phase,
          alertData: [],
          portfolioData: [],
          historyData: [],
          macroData: getMacroData(),
          _diag: { totalRows: 0, noId: 0, noPrice: 0, lowHist: 0, noCoupon: 0, lowZ: 0, passed: 0 }
        };
        sanitizeObject_(result);
        var safeJson2 = JSON.stringify(result);
        return ContentService.createTextOutput(safeJson2).setMimeType(ContentService.MimeType.JSON);
      }
      var alerts = getAlertData(mode);
      var diag = alerts._diag || {};
      delete alerts._diag;
      var allPairs = alerts._allPairs || [];
      delete alerts._allPairs;
      result = {
        ok: true,
        mode: mode,
        status: alerts.length > 0 ? "live" : "no_signals",
        alertData: alerts,
        allPairs: allPairs,
        portfolioData: getOpenTrades(),
        historyData: getClosedTrades(),
        macroData: getMacroData(),
        _diag: diag
      };
    }
    else if (action === 'saveTrade') {
      saveTradeToSheet({
        id: e.parameter.id || "",
        priceA: e.parameter.priceA || 0,
        priceB: e.parameter.priceB || 0,
        sizeA: e.parameter.sizeA || 0,
        sizeB: e.parameter.sizeB || 0
      });
      result = { ok: true, message: "Trade saved" };
    }
    else if (action === 'closeTrade') {
      closeTradeInSheet(e.parameter.id || "");
      result = { ok: true, message: "Trade closed" };
    }
    else if (action === 'partialClose') {
      partialCloseTradeInSheet(e.parameter.id || "", e.parameter.reduceA || 0, e.parameter.reduceB || 0);
      result = { ok: true, message: "Partial close recorded" };
    }
    else if (action === 'addDividend') {
      addDividendToTrade(e.parameter.id || "", e.parameter.type || "", parseFloat(e.parameter.amount) || 0);
      result = { ok: true, message: "Dividend recorded" };
    }
    else if (action === 'saveNote') {
      saveTradeNote(parseInt(e.parameter.row) || 0, e.parameter.note || "");
      result = { ok: true, message: "Note saved" };
    }
    else if (action === 'clearHistory') {
      clearHistory();
      result = { ok: true, message: "History cleared" };
    }
    else if (action === 'getWatchlist') {
      result = { ok: true, watchlistData: getWatchlistData() };
    }
    else if (action === 'saveWatchlist') {
      saveToWatchlist(e.parameter.id || "", e.parameter.mode || "intra");
      result = { ok: true, message: "Added to watchlist" };
    }
    else if (action === 'removeWatchlist') {
      removeFromWatchlist(e.parameter.id || "");
      result = { ok: true, message: "Removed from watchlist" };
    }
    else if (action === 'getMacroValuation') {
      result = { ok: true, macroValData: getMacroValuationData() };
    }
    else if (action === 'getBasketAnalytics') {
      result = { ok: true, basketData: getBasketAnalytics() };
    }
    else if (action === 'getPortfolioAnalytics') {
      var paMode = (e && e.parameter && e.parameter.paMode) ? e.parameter.paMode : 'live';
      var legsJson = (e && e.parameter && e.parameter.legs) ? e.parameter.legs : '[]';
      result = { ok: true, analyticsData: getPortfolioAnalytics(paMode, legsJson) };
    }
    else if (action === 'getScreenerData') {
      result = { ok: true, screenerData: getScreenerData_() };
    }
    else if (action === 'analyzePair') {
      var pairTa = (e && e.parameter && e.parameter.tA) ? e.parameter.tA : '';
      var pairTb = (e && e.parameter && e.parameter.tB) ? e.parameter.tB : '';
      var pairPa = parseFloat((e && e.parameter && e.parameter.pA) || 0);
      var pairPb = parseFloat((e && e.parameter && e.parameter.pB) || 0);
      var pairZ = parseFloat((e && e.parameter && e.parameter.z) || 0);
      // Optional custom windows param: comma-separated day counts (e.g. "5,10,20")
      var pairWindows = null;
      if (e && e.parameter && e.parameter.windows) {
        pairWindows = String(e.parameter.windows).split(',').map(function(v){ return parseInt(v); }).filter(function(v){ return v > 0 && !isNaN(v); });
        if (pairWindows.length === 0) pairWindows = null;
      }
      result = { ok: true, analysisData: analyzeSinglePair_(pairTa, pairTb, pairPa, pairPb, pairZ, pairWindows) };
    }
    else if (action === 'getTradeAlerts') {
      result = { ok: true, tradeAlerts: getTradeAlerts_() };
    }
    else if (action === 'getPositionSizing') {
      var psMaxLoss = parseFloat((e && e.parameter && e.parameter.maxLoss) || 500);
      result = { ok: true, sizingData: getPositionSizing_(psMaxLoss) };
    }
    else if (action === 'getBacktestResults') {
      var btZThreshold = parseFloat((e && e.parameter && e.parameter.zThreshold) || 2.0);
      var btExitZ = parseFloat((e && e.parameter && e.parameter.exitZ) || 0.5);
      var btMaxHold = parseInt((e && e.parameter && e.parameter.maxHold) || 60);
      var btMode = (e && e.parameter && e.parameter.mode) ? e.parameter.mode : 'all';
      result = { ok: true, backtestData: runBacktest_(btZThreshold, btExitZ, btMaxHold, btMode) };
    }
    else if (action === 'getSensitivityHeatmap') {
      var hmMaxHold = parseInt((e && e.parameter && e.parameter.maxHold) || 60);
      var hmMode = (e && e.parameter && e.parameter.mode) ? e.parameter.mode : 'all';
      var hmPairId = (e && e.parameter && e.parameter.pairId) ? e.parameter.pairId : null;
      result = { ok: true, heatmapData: runSensitivitySweep_(hmMaxHold, hmMode, hmPairId) };
    }
    else if (action === 'getOptimalSweep') {
      var osMaxHold = parseInt((e && e.parameter && e.parameter.maxHold) || 60);
      var osMode = (e && e.parameter && e.parameter.mode) ? e.parameter.mode : 'all';
      result = { ok: true, sweepData: runOptimalSweep_(osMaxHold, osMode) };
    }
    else if (action === 'getExitAlerts') {
      result = { ok: true, exitAlerts: getExitAlerts_() };
    }
    else if (action === 'saveExitParams') {
      var epExitZ = parseFloat((e && e.parameter && e.parameter.exitZ) || 0.5);
      var epMaxHold = parseInt((e && e.parameter && e.parameter.maxHold) || 60);
      var epStopLoss = parseFloat((e && e.parameter && e.parameter.stopLossPct) || 0);
      saveExitParams_(epExitZ, epMaxHold, epStopLoss);
      result = { ok: true, message: 'Exit params saved' };
    }
    else if (action === 'getJournalAnalytics') {
      result = { ok: true, journalAnalytics: getJournalAnalytics_() };
    }
    else if (action === 'getDividendCapture') {
      result = { ok: true, divCapture: getDividendCapture_() };
    }
    else if (action === 'getRegimeData') {
      result = { ok: true, regimeData: getRegimeData_() };
    }
    else if (action === 'getCorrelationMonitor') {
      result = { ok: true, correlationData: getCorrelationMonitor_() };
    }
    else if (action === 'saveNotificationSettings') {
      var chatId = (e && e.parameter && e.parameter.chatId) ? e.parameter.chatId : '';
      var botToken = (e && e.parameter && e.parameter.botToken) ? e.parameter.botToken : '';
      saveNotificationSettings_(chatId, botToken);
      result = { ok: true, message: 'Notification settings saved' };
    }
    else if (action === 'getNotificationSettings') {
      result = { ok: true, settings: getNotificationSettings_() };
    }
    else if (action === 'getModelPortfolios') {
      result = { ok: true, modelPortfolios: getModelPortfolios_() };
    }
    else if (action === 'runModelPortfolios') {
      var genError = null;
      try { runModelPortfolioGenerator(); } catch (genErr) { genError = genErr.toString(); }
      var mpResult = getModelPortfolios_();
      mpResult.generationError = genError;
      result = { ok: true, modelPortfolios: mpResult };
    }
    else {
      result = { ok: false, message: "Unknown action: " + action };
    }
  } catch (err) {
    result = { ok: false, message: err.toString() };
  }
  // Sanitize: replace NaN/Infinity with null iteratively (avoid replacer recursion on large results)
  sanitizeObject_(result);
  var safeJson = JSON.stringify(result);
  return ContentService
    .createTextOutput(safeJson)
    .setMimeType(ContentService.MimeType.JSON);
}
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action || 'getData';
    if (action === 'saveTrade') {
      saveTradeToSheet(payload);
      return ContentService.createTextOutput(JSON.stringify({ok:true,message:"Trade saved"})).setMimeType(ContentService.MimeType.JSON);
    } else if (action === 'closeTrade') {
      closeTradeInSheet(payload.id||"");
      return ContentService.createTextOutput(JSON.stringify({ok:true,message:"Trade closed"})).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,message:err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}
// ============================================================
// MACRO DATA
// ============================================================
function setupMacroSheet() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('MacroData');
  if (!sheet) {
    sheet = ss.insertSheet('MacroData');
    sheet.hideSheet();
    sheet.getRange("A1").setValue("TICKER"); sheet.getRange("B1").setValue("PRICE"); sheet.getRange("C1").setValue("CHANGE_PCT");
    sheet.getRange("A2").setValue("TNX"); sheet.getRange("B2").setFormula('=GOOGLEFINANCE("TNX")'); sheet.getRange("C2").setFormula('=GOOGLEFINANCE("TNX","changepct")/100');
    sheet.getRange("A3").setValue("TLT"); sheet.getRange("B3").setFormula('=GOOGLEFINANCE("TLT")'); sheet.getRange("C3").setFormula('=GOOGLEFINANCE("TLT","changepct")/100');
    sheet.getRange("A4").setValue("PFF"); sheet.getRange("B4").setFormula('=GOOGLEFINANCE("PFF")'); sheet.getRange("C4").setFormula('=GOOGLEFINANCE("PFF","changepct")/100');
  }
}
function getMacroData() {
  try {
    var sheet = SpreadsheetApp.getActive().getSheetByName('MacroData');
    if (!sheet) return {};
    var data = sheet.getRange("A2:C4").getValues();
    return {
      us10y: { val: (parseFloat(data[0][1])/10).toFixed(2)+"%", chg: parseFloat(data[0][2]) },
      tlt:   { val: "$"+parseFloat(data[1][1]).toFixed(2), chg: parseFloat(data[1][2]) },
      pff:   { val: "$"+parseFloat(data[2][1]).toFixed(2), chg: parseFloat(data[2][2]) }
    };
  } catch(e) { return {}; }
}
// ============================================================
// HELPERS
// ============================================================
function parseMoney(val) {
  if (val === null || val === undefined || val === "") return 0;
  if (typeof val === 'number') return val;
  return parseFloat(String(val).replace(',','.').replace(/[^0-9.-]/g,'')) || 0;
}
function cleanId(id) {
  return id ? String(id).toUpperCase().replace(/[^A-Z0-9]/g,'') : "";
}
// Downsample an array to maxLen by evenly picking elements (always keeps first and last)
function downsampleArray_(arr, maxLen) {
  if (!arr || arr.length <= maxLen) return arr;
  var result = [arr[0]];
  var step = (arr.length - 1) / (maxLen - 1);
  for (var i = 1; i < maxLen - 1; i++) { result.push(arr[Math.round(i * step)]); }
  result.push(arr[arr.length - 1]);
  return result;
}
// Iteratively sanitize NaN/Infinity values in an object (avoids JSON.stringify replacer stack overflow)
function sanitizeObject_(obj) {
  var stack = [obj];
  while (stack.length > 0) {
    var current = stack.pop();
    if (Array.isArray(current)) {
      for (var i = 0; i < current.length; i++) {
        if (typeof current[i] === 'number' && !isFinite(current[i])) { current[i] = null; }
        else if (current[i] && typeof current[i] === 'object') { stack.push(current[i]); }
      }
    } else if (current && typeof current === 'object') {
      var keys = Object.keys(current);
      for (var k = 0; k < keys.length; k++) {
        var v = current[keys[k]];
        if (typeof v === 'number' && !isFinite(v)) { current[keys[k]] = null; }
        else if (v && typeof v === 'object') { stack.push(v); }
      }
    }
  }
}
// ============================================================
// TICKER EXCLUSION LISTS
// ============================================================
/** Permanently excluded from ALL views */
var BLACKLIST = {
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
/** Only allowed in intra-company mode, excluded from credit arb */
var INTRA_ONLY = {
  'BHFAN':1,'BHFAO':1,'BHFAM':1,'BHFAP':1,'BHFAL':1,
  'HFRO-A':1,'HFRO-B':1
};
function isBlacklisted(ticker) {
  return !!BLACKLIST[String(ticker).toUpperCase().trim()];
}
function isIntraOnly(ticker) {
  var t = String(ticker).toUpperCase().trim();
  if (INTRA_ONLY[t]) return true;
  if (t.indexOf('SCE-') === 0) return true;
  return false;
}
// ============================================================
function parseTickerInfo(rawId) {
  var cleanRaw = String(rawId).toUpperCase().trim();
  var tA = "LEG A", tB = "LEG B", displayId = cleanRaw;
  if (cleanRaw.includes('|')) {
    var parts = cleanRaw.split('|');
    if (parts.length >= 3) { tA=parts[1]; tB=parts[2]; displayId=parts[1]+"|"+parts[2]; }
    else if (parts.length === 2) { tA=parts[0]; tB=parts[1]; displayId=cleanRaw; }
  } else if (cleanRaw.includes(' VS ')) {
    var parts = cleanRaw.split(' VS '); tA=parts[0]; tB=parts[1];
  } else { tA=cleanRaw; tB="HEDGE"; }
  return { id: displayId, tA: tA.trim(), tB: tB.trim() };
}
// ============================================================
// ALERT DATA — supports mode switching
// ============================================================
function getAlertData(mode) {
  // WebCache/WebCacheCredit = primary data source.
  // For intra: falls back to Live sheet (GOOGLEFINANCE formulas).
  // For credit: WebCacheCredit is the only source (computed by computeCreditCache).
  var cacheName = (mode === 'credit') ? 'WebCacheCredit' : 'WebCache';
  var liveName = (mode === 'credit') ? null : 'Live';

  try {
    var ss = SpreadsheetApp.getActive();
    var liveSheet = ss.getSheetByName(cacheName);
    if ((!liveSheet || liveSheet.getLastRow() <= 1) && liveName) {
      liveSheet = ss.getSheetByName(liveName);
    }
    if (!liveSheet) return [];
    var data = liveSheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    // Load AlertsLog for trend ribbons — group by DATE for daily Z-trend
    // snapshotZScores runs hourly → ~24 entries per pair per day.
    // Group by date, keep last Z per day. Today excluded (live Z appended later).
    // Final array capped at 5 most recent daily Z-scores (including today).
    var logMap = {}; // {cleanId: {dateStr: lastZForThatDay}}
    var now = new Date();
    var tMM = now.getMonth() + 1;
    var tDD = now.getDate();
    var todayKey = now.getFullYear() + '-' + (tMM < 10 ? '0' : '') + tMM + '-' + (tDD < 10 ? '0' : '') + tDD;
    var logSheet = ss.getSheetByName('AlertsLog');
    if (logSheet) {
      var lastRow = logSheet.getLastRow();
      if (lastRow > 1) {
        var startRow = Math.max(2, lastRow - 15000);
        var logData = logSheet.getRange(startRow, 1, lastRow - startRow + 1, 4).getValues();
        for (var k = 0; k < logData.length; k++) {
          var ts = logData[k][0];
          var lcid = cleanId(logData[k][1]);
          if (!lcid) continue;
          var zVal = parseFloat(logData[k][2]);
          if (isNaN(zVal)) continue;
          // Extract YYYY-MM-DD from timestamp
          var dateKey = '';
          if (ts instanceof Date) {
            var mm = ts.getMonth() + 1;
            var dd = ts.getDate();
            dateKey = ts.getFullYear() + '-' + (mm < 10 ? '0' : '') + mm + '-' + (dd < 10 ? '0' : '') + dd;
          } else if (typeof ts === 'string' && ts.length >= 10) {
            // Fallback: parse "YYYY-MM-DD..." string timestamps
            var candidate = ts.substring(0, 10);
            if (/^\d{4}-\d{2}-\d{2}$/.test(candidate)) dateKey = candidate;
          }
          if (!dateKey) continue; // skip entries without valid timestamps
          if (dateKey === todayKey) continue; // exclude today — live Z appended later
          if (!logMap[lcid]) logMap[lcid] = {};
          logMap[lcid][dateKey] = zVal; // last snapshot per day wins
        }
      }
    }
    // Load ZScoreAge for age column (Task 1)
    var ageMap = {};
    var ageSheet = ss.getSheetByName('ZScoreAge');
    if (ageSheet) {
      var ageData = ageSheet.getDataRange().getValues();
      for (var a = 1; a < ageData.length; a++) {
        var aid = cleanId(String(ageData[a][0]));
        var ts = ageData[a][1];
        if (aid && ts instanceof Date) {
          var diffMs = now.getTime() - ts.getTime();
          var diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
          ageMap[aid] = diffDays;
        }
      }
    }
    // Load DivDates for ex-dividend date column
    var divMap = {};
    var divSheet = ss.getSheetByName('DivDates');
    if (divSheet && divSheet.getLastRow() > 1) {
      var divData = divSheet.getRange(2, 1, divSheet.getLastRow() - 1, 2).getValues();
      for (var d = 0; d < divData.length; d++) {
        var dticker = String(divData[d][0]).toUpperCase().trim();
        var ddate = divData[d][1];
        if (dticker && ddate instanceof Date) {
          divMap[dticker] = ddate;
        }
      }
    }
    var output = [];
    var _diag = {totalRows: data.length - 1, noId: 0, noPrice: 0, noHistory: 0, noCoupon: 0, lowZ: 0,
                 blacklisted: 0, intraOnly: 0, divTooFar: 0, passed: 0,
                 sheetUsed: liveSheet.getName(), sampleRows: []};
    // Capture first 5 rows raw data for debugging
    for (var s = 1; s < Math.min(6, data.length); s++) {
      var sr = data[s];
      _diag.sampleRows.push({
        row: s,
        pairId: String(sr[0] || "").substring(0, 30),
        priceA: sr[3],
        priceB: sr[4],
        spread: sr[5],
        mean: sr[10],
        stdev: sr[11],
        zScore: sr[12],
        couponA: sr[8],
        couponB: sr[9],
        histCount: sr[16],
        yieldA: sr[6],
        yieldB: sr[7]
      });
    }
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rawId = row[0];
      if (!rawId) { _diag.noId++; continue; }

      // FILTER: blacklisted tickers — excluded from all views
      var tickerA = String(row[1]).toUpperCase().trim();
      var tickerB = String(row[2]).toUpperCase().trim();
      if (isBlacklisted(tickerA) || isBlacklisted(tickerB)) { _diag.blacklisted++; continue; }

      // FILTER: intra-only tickers — excluded from credit arb
      if (mode === 'credit' && (isIntraOnly(tickerA) || isIntraOnly(tickerB))) { _diag.intraOnly++; continue; }

      // FILTER: valid prices
      var priceA = parseFloat(row[3]) || 0;
      var priceB = parseFloat(row[4]) || 0;
      if (priceA <= 0 || priceB <= 0) { _diag.noPrice++; continue; }

      // FILTER: valid statistical history
      // 1) StDev must be real (not the 0.001 IFERROR default) — catches total GOOGLEFINANCE failure
      // 2) HistCount >= 40 — catches new tickers with too little data (wild Z-scores)
      var stdev = parseFloat(row[11]) || 0;
      var histCount = parseFloat(row[16]) || 0;
      if (stdev <= 0.001 || histCount < 40) { _diag.noHistory++; continue; }

      // FILTER: coupon must exist (exclude variable/reset)
      var couponA = row[8];
      var couponB = row[9];
      if (couponA === "" || couponA === null || couponA === undefined ||
          couponB === "" || couponB === null || couponB === undefined) { _diag.noCoupon++; continue; }
      var currentZ = parseFloat(row[12]) || 0;

      // Collect ALL valid pairs for sweep search (before Z-score filter)
      var sweepInfo = parseTickerInfo(rawId);
      output._allPairs = output._allPairs || [];
      output._allPairs.push({ id: sweepInfo.id, tA: String(row[1] || sweepInfo.tA).trim(), tB: String(row[2] || sweepInfo.tB).trim(), z: currentZ.toFixed(2) });

      // FILTER: |z| >= 1.8
      if (Math.abs(currentZ) < 1.8) { _diag.lowZ++; continue; }

      // FILTER: Div date proximity — exclude pairs where ex-div dates are >45 days apart
      var divA = divMap[tickerA] || null;
      var divB = divMap[tickerB] || null;
      if (divA && divB) {
        var divDiffDays = Math.abs(divA.getTime() - divB.getTime()) / 86400000;
        if (divDiffDays > 45) { _diag.divTooFar++; continue; }
      }
      _diag.passed++;
      var info = parseTickerInfo(rawId);
      var cid = cleanId(rawId);
      // TREND from AlertsLog — last 5 daily Z-scores (including today's live)
      var zDateMap = logMap[cid] || {};
      var dates = Object.keys(zDateMap).sort(); // sorted YYYY-MM-DD strings
      var historicalZScores = dates.map(function(d) { return zDateMap[d].toFixed(1); });
      historicalZScores.push(currentZ.toFixed(1)); // append today's live Z
      var trend = historicalZScores.slice(-5);     // strictly last 5 ribbons
      // AGE from ZScoreAge (Task 1)
      var ageDays = ageMap[cid] || 0;
      // Range
      var lower = parseFloat(row[13]) || 0;
      var upper = parseFloat(row[14]) || 0;
      // Yields
      var yieldA = parseFloat(row[6]) || 0;
      var yieldB = parseFloat(row[7]) || 0;
      var yA = yieldA > 1 ? yieldA.toFixed(2) : (yieldA * 100).toFixed(2);
      var yB = yieldB > 1 ? yieldB.toFixed(2) : (yieldB * 100).toFixed(2);
      // Liquidity
      var avgLiq = parseFloat(row[19]) || 0;
      var curVol = parseFloat(row[22]) || 0;
      var volSpike = (row[23] === true || row[23] === "TRUE");
      // DIV DATES — divA/divB already resolved above (proximity filter)
      // Expected Profit = |Current Spread - 90-Day Mean| (distance to mean reversion)
      var spread = parseFloat(row[5]) || 0;
      var mean = parseFloat(row[10]) || 0;
      var expProfit = Math.abs(spread - mean);
      output.push({
        id: info.id,
        tA: row[1] || info.tA,
        tB: row[2] || info.tB,
        pA: priceA,
        pB: priceB,
        stdev: parseFloat(stdev.toFixed(4)),
        rng: lower.toFixed(2) + " / " + upper.toFixed(2),
        sec: row[15] || "",
        spr: spread.toFixed(2),
        expProfit: expProfit.toFixed(2),
        yA: yA,
        yB: yB,
        z: currentZ.toFixed(2),
        age: ageDays,
        histDays: histCount,
        liq: avgLiq,
        curVol: curVol,
        volSpike: volSpike,
        zTrend: trend,
        exDivA: (divA && divA >= now) ? divA.toISOString().split('T')[0] : null,
        exDivB: (divB && divB >= now) ? divB.toISOString().split('T')[0] : null
      });
    }
    output._diag = _diag;
    return output;
  } catch (e) {
    console.error("getAlertData error: " + e);
    return [];
  }
}
// ============================================================
// PORTFOLIO — unified (reads from OpenTrades + Live)
// ============================================================
function getOpenTrades() {
  try {
    var ss = SpreadsheetApp.getActive();
    var openSheet = ss.getSheetByName('OpenTrades');
    if (!openSheet) return [];
    var openData = openSheet.getDataRange().getValues();
    if (openData.length < 2) return [];
    // Merge rows for unified lookup — prefer WebCache, fall back to Live
    var liveRows = [];
    var intra = ss.getSheetByName('WebCache');
    if (!intra || intra.getLastRow() <= 1) intra = ss.getSheetByName('Live');
    var credit = ss.getSheetByName('WebCacheCredit');
    var sources = [intra, credit];
    for (var s = 0; s < sources.length; s++) {
      var ls = sources[s];
      if (ls && ls.getLastRow() > 1) {
        var rows = ls.getRange(2, 1, ls.getLastRow()-1, 24).getValues();
        liveRows = liveRows.concat(rows);
      }
    }
    // Pre-read credit pair IDs for strategy detection
    var creditIds = {};
    if (credit && credit.getLastRow() > 1) {
      var creditPairCol = credit.getRange(2, 1, credit.getLastRow()-1, 1).getValues();
      for (var ci = 0; ci < creditPairCol.length; ci++) {
        if (creditPairCol[ci][0]) creditIds[cleanId(creditPairCol[ci][0])] = true;
      }
    }
    // Load DivDates + Master coupon yields for auto-div tracking
    var divMap = {};
    var divSheet = ss.getSheetByName('DivDates');
    if (divSheet && divSheet.getLastRow() > 1) {
      var divData = divSheet.getRange(2, 1, divSheet.getLastRow() - 1, 2).getValues();
      for (var d = 0; d < divData.length; d++) {
        var dticker = String(divData[d][0]).toUpperCase().trim();
        var ddate = divData[d][1];
        if (dticker && ddate instanceof Date) divMap[dticker] = ddate;
      }
    }
    var couponMap = {};
    var masterSheet = ss.getSheetByName('Master');
    if (masterSheet && masterSheet.getLastRow() > 1) {
      var masterData = masterSheet.getDataRange().getValues();
      for (var m = 1; m < masterData.length; m++) {
        var mTicker = String(masterData[m][0]).toUpperCase().trim();
        var mCoupon = parseFloat(masterData[m][2]) || 0; // Col C = Coupon Yield
        if (mTicker && mCoupon > 0) couponMap[mTicker] = mCoupon;
      }
    }
    var now = new Date();
    var results = [];
    for (var j = 1; j < openData.length; j++) {
      try {
        var rawId = openData[j][0];
        if (!rawId) continue;
        var info = parseTickerInfo(rawId);
        var tA_Name = info.tA, tB_Name = info.tB, displayId = info.id;
        var openAnchor = cleanId(rawId);
        var pair = null;
        for (var k = 0; k < liveRows.length; k++) {
          if (liveRows[k][0] && cleanId(liveRows[k][0]) === openAnchor) { pair = liveRows[k]; break; }
        }
        var strategy = creditIds[openAnchor] ? 'credit' : 'intra';
        var costA = parseMoney(openData[j][2]);
        var costB = parseMoney(openData[j][3]);
        var sA = parseMoney(openData[j][4]);
        var sB = parseMoney(openData[j][5]);
        var paidDiv = parseMoney(openData[j][7]);
        var rcvdDiv = parseMoney(openData[j][8]);
        var openDate = openData[j][6] instanceof Date ? openData[j][6].toISOString().split('T')[0] : String(openData[j][6] || '');
        if (pair) {
          if (pair[1] && String(pair[1]).length > 1) tA_Name = String(pair[1]);
          if (pair[2] && String(pair[2]).length > 1) tB_Name = String(pair[2]);
          var liveSpr = parseFloat(pair[5])||0;
          var meanTarget = parseFloat(pair[10])||0;
          var stdev = parseFloat(pair[11])||0;
          var livePriceA = parseFloat(pair[3])||0;
          var livePriceB = parseFloat(pair[4])||0;
          var capGains = ((livePriceA-costA)*sA) + ((livePriceB-costB)*sB);
          var netPnl = capGains + rcvdDiv - paidDiv;
          var currentZ = parseFloat(pair[12]) || 0;
          var sector = pair[15] ? String(pair[15]) : '';
          // Div dates + coupon yields for auto-div tracking
          var tkA = String(tA_Name).toUpperCase().trim();
          var tkB = String(tB_Name).toUpperCase().trim();
          var exDivA = divMap[tkA] || null;
          var exDivB = divMap[tkB] || null;
          var couponA = couponMap[tkA] || 0;
          var couponB = couponMap[tkB] || 0;
          results.push({
            id: displayId, tA: tA_Name, tB: tB_Name,
            spr: liveSpr.toFixed(2), target: meanTarget.toFixed(2),
            stdev: parseFloat(stdev.toFixed(4)),
            sA: sA, sB: sB, pA: costA.toFixed(2), pB: costB.toFixed(2),
            livePriceA: livePriceA.toFixed(2), livePriceB: livePriceB.toFixed(2),
            dollarPnL: netPnl.toFixed(2), capGains: capGains.toFixed(2),
            paidDiv: paidDiv.toFixed(2), rcvdDiv: rcvdDiv.toFixed(2),
            centGoal: (Math.abs((costA-costB)-meanTarget)*100).toFixed(0),
            centRem: (Math.abs(liveSpr-meanTarget)*100).toFixed(0),
            distToMean: (liveSpr - meanTarget).toFixed(4),
            isWinning: netPnl > 0, entryZ: openData[j][1],
            currentZ: currentZ, sector: sector, strategy: strategy,
            openDate: openDate,
            exDivA: (exDivA && exDivA >= now) ? exDivA.toISOString().split('T')[0] : null,
            exDivB: (exDivB && exDivB >= now) ? exDivB.toISOString().split('T')[0] : null,
            couponA: couponA, couponB: couponB
          });
        } else {
          results.push({
            id: displayId+" [WAITING]", tA: tA_Name, tB: tB_Name,
            spr:"0.00", target:"0.00", stdev: 0,
            sA:sA, sB:sB,
            pA:costA.toFixed(2), pB:costB.toFixed(2),
            livePriceA: "0.00", livePriceB: "0.00",
            dollarPnL:"0.00", capGains:"0.00",
            paidDiv: paidDiv.toFixed(2), rcvdDiv: rcvdDiv.toFixed(2),
            centGoal:"0", centRem:"0", distToMean: "0",
            isWinning:false, entryZ:"0",
            currentZ: 0, sector: '', strategy: strategy,
            openDate: openDate,
            exDivA: null, exDivB: null,
            couponA: 0, couponB: 0
          });
        }
      } catch(err) { console.error(err); }
    }
    return results;
  } catch(e) { return []; }
}
// ============================================================
// HISTORY
// ============================================================
function getClosedTrades() {
  try {
    var rows = SpreadsheetApp.getActive().getSheetByName('ClosedTrades').getDataRange().getValues();
    var output = [];
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i]; var info = parseTickerInfo(r[0]);
      var openDate = r[6] instanceof Date ? r[6] : null;
      var closeDate = r[7] instanceof Date ? r[7] : null;
      var holdDays = (openDate && closeDate) ? Math.floor((closeDate.getTime() - openDate.getTime()) / 86400000) : 0;
      var costA = parseFloat(r[2]) || 0;
      var costB = parseFloat(r[3]) || 0;
      var sizeA = parseFloat(r[4]) || 0;
      var sizeB = parseFloat(r[5]) || 0;
      var pnl = parseFloat(r[8]) || 0;
      var exitPriceA = parseFloat(r[9]) || 0;
      var exitPriceB = parseFloat(r[10]) || 0;
      var exitZ = parseFloat(r[11]) || 0;
      var closeType = r[12] || 'FULL';
      var paidDiv = parseFloat(r[13]) || 0;
      var rcvdDiv = parseFloat(r[14]) || 0;
      var note = (r[15] !== undefined && r[15] !== null) ? String(r[15]) : '';
      // capGains = netPnl - rcvdDiv + paidDiv (reverse the formula to extract capital gains)
      var capGains = pnl - rcvdDiv + paidDiv;
      var netDiv = rcvdDiv - paidDiv;
      var entryCost = Math.abs(costA * sizeA) + Math.abs(costB * sizeB);
      var returnPct = entryCost > 0 ? (pnl / entryCost * 100) : 0;
      output.push({
        id: info.id, tA: info.tA, tB: info.tB,
        rowIdx: i + 1, // 1-indexed sheet row for saveTradeNote
        entryZ: parseFloat(r[1]) || 0, exitZ: exitZ,
        costA: costA.toFixed(2), costB: costB.toFixed(2),
        exitPriceA: exitPriceA.toFixed(2), exitPriceB: exitPriceB.toFixed(2),
        sizeA: sizeA, sizeB: sizeB,
        openDate: openDate ? openDate.toLocaleDateString() : '---',
        closeDate: closeDate ? closeDate.toLocaleDateString() : '---',
        holdDays: holdDays,
        pnl: pnl.toFixed(2),
        capGains: capGains.toFixed(2),
        paidDiv: paidDiv.toFixed(2),
        rcvdDiv: rcvdDiv.toFixed(2),
        netDiv: netDiv.toFixed(2),
        returnPct: returnPct.toFixed(2),
        closeType: closeType,
        note: note
      });
    }
    return output.reverse();
  } catch(e) { return []; }
}
function clearHistory() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('ClosedTrades');
  if (!sheet || sheet.getLastRow() <= 1) return;
  sheet.deleteRows(2, sheet.getLastRow() - 1);
}
// ============================================================
// BASKET ANALYTICS
// ============================================================
/**
 * Computes rolling basket Z-scores for the active portfolio.
 * Reads TickerHistory for historical prices, computes daily basket values,
 * then returns 30/60/90-day rolling Z-scores.
 */
function getBasketAnalytics() {
  try {
    var ss = SpreadsheetApp.getActive();
    var openSheet = ss.getSheetByName('OpenTrades');
    if (!openSheet || openSheet.getLastRow() <= 1) return { trades: 0 };
    var openData = openSheet.getDataRange().getValues();

    // Get live pair data for current Z, sector, strategy detection
    var liveRows = [];
    var intra = ss.getSheetByName('WebCache');
    if (!intra || intra.getLastRow() <= 1) intra = ss.getSheetByName('Live');
    var credit = ss.getSheetByName('WebCacheCredit');
    var sources = [intra, credit];
    for (var s = 0; s < sources.length; s++) {
      var ls = sources[s];
      if (ls && ls.getLastRow() > 1) {
        var rows = ls.getRange(2, 1, ls.getLastRow()-1, 24).getValues();
        liveRows = liveRows.concat(rows);
      }
    }

    // Build trade list: ticker pairs + weights
    var trades = [];
    // Pre-read credit pair IDs to avoid repeated sheet reads in the loop
    var creditIds = {};
    if (credit && credit.getLastRow() > 1) {
      var creditPairCol = credit.getRange(2, 1, credit.getLastRow()-1, 1).getValues();
      for (var cr = 0; cr < creditPairCol.length; cr++) {
        if (creditPairCol[cr][0]) creditIds[cleanId(creditPairCol[cr][0])] = true;
      }
    }
    for (var j = 1; j < openData.length; j++) {
      var rawId = openData[j][0];
      if (!rawId) continue;
      var info = parseTickerInfo(rawId);
      var openAnchor = cleanId(rawId);
      var sA = parseMoney(openData[j][4]);
      var sB = parseMoney(openData[j][5]);
      var pair = null;
      for (var k = 0; k < liveRows.length; k++) {
        if (liveRows[k][0] && cleanId(liveRows[k][0]) === openAnchor) { pair = liveRows[k]; break; }
      }
      var isCredit = !!creditIds[openAnchor];
      trades.push({
        id: rawId,
        tA: info.tA,
        tB: info.tB,
        sA: sA,
        sB: sB,
        weight: Math.abs(sA),
        sector: pair ? String(pair[15] || '') : '',
        strategy: isCredit ? 'credit' : 'intra',
        currentZ: pair ? (parseFloat(pair[12]) || 0) : 0,
        liveSpread: pair ? (parseFloat(pair[5]) || 0) : 0
      });
    }
    if (trades.length === 0) return { trades: 0 };

    // Read TickerHistory for rolling Z computation
    var histSheet = ss.getSheetByName('TickerHistory');
    var tickerHist = {}; // ticker -> [{date, price}]
    if (histSheet && histSheet.getLastRow() > 1) {
      var histData = histSheet.getDataRange().getValues();
      // TickerHistory: Row 1 = headers (Ticker, Date1, Date2, ...)
      // Row N = [Ticker, Price1, Price2, ...]
      for (var h = 1; h < histData.length; h++) {
        var ticker = String(histData[h][0]).trim();
        if (!ticker) continue;
        var prices = [];
        for (var d = 1; d < histData[h].length; d++) {
          var p = parseFloat(histData[h][d]);
          if (!isNaN(p) && p > 0) prices.push({ idx: d - 1, price: p });
        }
        tickerHist[ticker] = prices;
      }
    }

    // Compute daily basket values (spread * weight for each trade, summed)
    // Find the common date range across all trades
    var maxDays = 90;
    var basketValues = [];
    if (Object.keys(tickerHist).length > 0) {
      for (var day = 0; day < maxDays; day++) {
        var bVal = 0;
        var validTrades = 0;
        for (var t = 0; t < trades.length; t++) {
          var hA = tickerHist[trades[t].tA];
          var hB = tickerHist[trades[t].tB];
          if (hA && hB && hA.length > day && hB.length > day) {
            var spread = hA[hA.length - 1 - day].price - hB[hB.length - 1 - day].price;
            bVal += trades[t].weight * spread;
            validTrades++;
          }
        }
        if (validTrades > 0) basketValues.push(bVal);
      }
      basketValues.reverse(); // oldest first
    }

    // Compute rolling stats for 30, 60, 90 day windows
    var rollingZ = {};
    var windows = [30, 60, 90];
    for (var w = 0; w < windows.length; w++) {
      var n = windows[w];
      if (basketValues.length >= n) {
        var slice = basketValues.slice(basketValues.length - n);
        var sum = 0, sq = 0;
        for (var i = 0; i < slice.length; i++) sum += slice[i];
        var mean = sum / slice.length;
        for (var i = 0; i < slice.length; i++) sq += (slice[i] - mean) * (slice[i] - mean);
        var denom = slice.length > 1 ? slice.length - 1 : 1;
        var std = Math.sqrt(sq / denom);
        var current = basketValues[basketValues.length - 1];
        rollingZ[n + 'd'] = {
          z: std > 0.0001 ? parseFloat(((current - mean) / std).toFixed(2)) : 0,
          mean: parseFloat(mean.toFixed(4)),
          std: parseFloat(std.toFixed(4)),
          dataPoints: slice.length
        };
      } else {
        rollingZ[n + 'd'] = { z: 0, mean: 0, std: 0, dataPoints: basketValues.length, insufficient: true };
      }
    }

    // Compute sector concentration
    var sectorNotional = {};
    var totalNotional = 0;
    var strategyCount = { intra: 0, credit: 0 };
    for (var t = 0; t < trades.length; t++) {
      var notional = trades[t].weight * Math.abs(trades[t].liveSpread);
      var sec = trades[t].sector || 'Unknown';
      sectorNotional[sec] = (sectorNotional[sec] || 0) + notional;
      totalNotional += notional;
      strategyCount[trades[t].strategy]++;
    }
    var sectorPct = {};
    for (var sec in sectorNotional) {
      sectorPct[sec] = totalNotional > 0 ? parseFloat((sectorNotional[sec] / totalNotional * 100).toFixed(1)) : 0;
    }

    // Weighted average Z (instant, from live data)
    var weightedZSum = 0, totalWeight = 0;
    for (var t = 0; t < trades.length; t++) {
      weightedZSum += trades[t].weight * trades[t].currentZ;
      totalWeight += trades[t].weight;
    }
    var weightedAvgZ = totalWeight > 0 ? parseFloat((weightedZSum / totalWeight).toFixed(2)) : 0;

    return {
      trades: trades.length,
      weightedAvgZ: weightedAvgZ,
      rollingZ: rollingZ,
      sectorPct: sectorPct,
      strategyMix: strategyCount,
      totalNotional: parseFloat(totalNotional.toFixed(2)),
      basketHistory: basketValues.length > 0 ? basketValues.slice(-30) : []
    };
  } catch(e) {
    return { trades: 0, error: e.message };
  }
}
// ============================================================
// PORTFOLIO ANALYTICS (dual-mode: live + sandbox)
// ============================================================
/**
 * Reads TickerHistory and returns a map of {TICKER: [price0, price1, ...]}
 * where index 0 = oldest, last index = most recent.
 */
function readTickerHistMap_(ss) {
  var histSheet = ss.getSheetByName('TickerHistory');
  var histMap = {};
  if (!histSheet || histSheet.getLastRow() <= 1) return histMap;
  var histData = histSheet.getDataRange().getValues();
  for (var h = 1; h < histData.length; h++) {
    var ticker = String(histData[h][0]).trim().toUpperCase();
    if (!ticker) continue;
    var prices = [];
    for (var d = 1; d < histData[h].length; d++) {
      var v = histData[h][d];
      // Match the proven parsing from computeCreditCache in SetupDashboard.gs:
      // handles Numbers, strings-that-are-numbers, and skips empty/null/Date/zero
      if (v === '' || v === null || v === undefined) continue;
      if (typeof v === 'object') continue; // skip Date objects from GOOGLEFINANCE
      var n = Number(v);
      if (!isNaN(n) && n > 0) prices.push(n);
    }
    if (prices.length > 0) histMap[ticker] = prices;
  }
  return histMap;
}

/**
 * Core math: builds a synthetic daily portfolio value array from legs + history,
 * then computes 30/60/90-day rolling Z-scores and expected profit.
 *
 * @param {Array} legs - [{ticker, size, direction}] where direction=+1 (long) or -1 (short)
 * @param {Object} histMap - {TICKER: [price0..priceN]} oldest-first
 * @returns {Object} {rollingZ, netSpread, grossLong, grossShort, grossExposure, dailyValues, validLegs, historyDays}
 */
function computeBasketMetrics_(legs, histMap, customWindows) {
  if (!legs || legs.length === 0) return { error: 'No legs provided', dailyValues: [] };

  // Find the minimum history length across all legs
  var minLen = Infinity;
  var validLegs = [];
  for (var i = 0; i < legs.length; i++) {
    var tk = String(legs[i].ticker).trim().toUpperCase();
    var hist = histMap[tk];
    if (!hist || hist.length === 0) continue;
    var size = parseFloat(legs[i].size) || 0;
    var dir = parseFloat(legs[i].direction) || 0;
    if (size === 0 || dir === 0) continue;
    validLegs.push({ ticker: tk, size: Math.abs(size), dir: dir > 0 ? 1 : -1, hist: hist });
    if (hist.length < minLen) minLen = hist.length;
  }
  if (validLegs.length === 0) return { error: 'No valid legs with history', dailyValues: [] };
  if (minLen === Infinity) minLen = 0;

  // Use all available history (no artificial cap — GOOGLEFINANCE provides ~100 trading days)
  var maxDays = minLen;

  // Compute total weight for normalization (sum of leg sizes / 2 = sum of pair sizes)
  // This converts raw dollar values into size-weighted average spread
  var totalWeight = 0;
  for (var j = 0; j < validLegs.length; j++) totalWeight += validLegs[j].size;
  totalWeight = totalWeight / 2;
  if (totalWeight === 0) totalWeight = 1; // prevent division by zero

  // Build daily size-weighted average spread array (oldest first)
  var dailyValues = [];
  for (var day = 0; day < maxDays; day++) {
    var val = 0;
    for (var j = 0; j < validLegs.length; j++) {
      var leg = validLegs[j];
      var legDayIdx = leg.hist.length - maxDays + day;
      val += leg.dir * leg.size * leg.hist[legDayIdx];
    }
    dailyValues.push(val / totalWeight);
  }

  var currentValue = dailyValues.length > 0 ? dailyValues[dailyValues.length - 1] : 0;

  // Compute rolling stats for 30, 60, 90 day windows
  // Allow 10% shortfall (e.g. 81+ days satisfies 90-day window) since GOOGLEFINANCE
  // may return slightly fewer trading days due to holidays or data gaps
  var rollingZ = {};
  var windows = [30, 60, 90];
  for (var w = 0; w < windows.length; w++) {
    var n = windows[w];
    var minRequired = Math.floor(n * 0.9);
    if (dailyValues.length >= minRequired) {
      var actualN = Math.min(n, dailyValues.length);
      var slice = dailyValues.slice(dailyValues.length - actualN);
      var sum = 0;
      for (var k = 0; k < slice.length; k++) sum += slice[k];
      var mean = sum / slice.length;
      var sq = 0;
      for (var k = 0; k < slice.length; k++) sq += (slice[k] - mean) * (slice[k] - mean);
      var denom = slice.length > 1 ? slice.length - 1 : 1;
      var std = Math.sqrt(sq / denom);
      var z = std > 0.0001 ? (currentValue - mean) / std : 0;
      var expectedProfit = mean - currentValue; // positive = portfolio should revert UP (per-share)
      rollingZ[n + 'd'] = {
        z: parseFloat(z.toFixed(2)),
        mean: parseFloat(mean.toFixed(2)),
        std: parseFloat(std.toFixed(2)),
        expectedProfit: parseFloat(expectedProfit.toFixed(4)),
        dataPoints: slice.length
      };
    } else {
      rollingZ[n + 'd'] = { z: 0, mean: 0, std: 0, expectedProfit: 0, dataPoints: dailyValues.length, insufficient: true };
    }
  }

  // Compute gross exposure (long $ + short $) from most recent prices
  var grossLong = 0, grossShort = 0;
  for (var j = 0; j < validLegs.length; j++) {
    var leg = validLegs[j];
    var latestPrice = leg.hist[leg.hist.length - 1];
    var legVal = leg.size * latestPrice;
    if (leg.dir > 0) grossLong += legVal; else grossShort += legVal;
  }

  // ── Historical Probability Engine ──
  // Scan dailyValues for trigger points where spread ≈ currentValue (±5%),
  // then compute forward-looking max adverse excursion and mean reversion rates.
  var probResult = computeHistoricalProbabilities_(dailyValues, currentValue, rollingZ, totalWeight, customWindows);

  // ── MAE + Recovery Days (basket-level) ──
  var maeRecovery = computeBasketMAE_(dailyValues, rollingZ, totalWeight);

  return {
    rollingZ: rollingZ,
    netSpread: parseFloat(currentValue.toFixed(2)),
    grossLong: parseFloat(grossLong.toFixed(2)),
    grossShort: parseFloat(grossShort.toFixed(2)),
    grossExposure: parseFloat((grossLong + grossShort).toFixed(2)),
    dailyValues: dailyValues.length > 0 ? dailyValues.slice(-120).map(function(v) { return parseFloat(v.toFixed(2)); }) : [],
    dailyValuesFull_: dailyValues, // internal: full precision array for probability overrides
    validLegs: validLegs.length,
    historyDays: maxDays,
    totalWeight: totalWeight,
    probabilities: probResult,
    maeRecovery: maeRecovery
  };
}

/**
 * Basket-level Max Adverse Excursion + Recovery Days.
 * Scans the full daily spread history for every point where the spread
 * deviated from the rolling mean, tracking:
 *   - The worst single deviation from mean (MAE) across all history
 *   - The average number of trading days to recover back to mean after
 *     each excursion beyond ±1σ
 *
 * @param {number[]} dailyValues - Full daily spread array (oldest first)
 * @param {Object} rollingZ - The rollingZ object with per-window mean/std
 * @param {number} totalWeight - Position size multiplier for dollar conversion
 * @return {Object} {maxMae, maxMaeDollar, avgRecoveryDays, medianRecoveryDays, excursionCount}
 */
function computeBasketMAE_(dailyValues, rollingZ, totalWeight) {
  var result = { maxMae: 0, maxMaeDollar: 0, avgRecoveryDays: 0, medianRecoveryDays: 0, excursionCount: 0 };
  var len = dailyValues.length;
  if (len < 20) return result;

  // Use the longest available window for mean/std
  var rz = null;
  var windowKeys = ['90d', '60d', '30d'];
  for (var w = 0; w < windowKeys.length; w++) {
    if (rollingZ[windowKeys[w]] && !rollingZ[windowKeys[w]].insufficient) {
      rz = rollingZ[windowKeys[w]]; break;
    }
  }
  if (!rz) return result;

  var mean = rz.mean;
  var std = rz.std;
  if (std < 0.0001) return result;

  // Track the worst deviation from mean across the entire series
  var worstMae = 0;
  for (var i = 0; i < len; i++) {
    var dev = Math.abs(dailyValues[i] - mean);
    if (dev > worstMae) worstMae = dev;
  }

  // Find excursion events: stretches where spread goes beyond ±1.5σ from mean
  // Track how many trading days each excursion takes to recover (touch the mean ±0.25σ)
  var recoveryDays = [];
  var excursionDetails = []; // {startDay, duration, peakDev, peakDevDollar, peakZ, recovered}
  var meanTolerance = std * 0.25;
  var excursionThreshold = std * 1.5;
  var windowSize = rz === rollingZ['90d'] ? 90 : rz === rollingZ['60d'] ? 60 : 30;
  var inExcursion = false;
  var excursionStart = 0;
  var excursionPeakDev = 0;
  var excursionPeakDay = 0;
  var excursionDir = '';

  for (var i = 0; i < len; i++) {
    var dev = Math.abs(dailyValues[i] - mean);
    var dir = dailyValues[i] >= mean ? 'above' : 'below';
    if (!inExcursion && dev >= excursionThreshold) {
      inExcursion = true;
      excursionStart = i;
      excursionPeakDev = dev;
      excursionPeakDay = i;
      excursionDir = dir;
    } else if (inExcursion) {
      if (dev > excursionPeakDev) { excursionPeakDev = dev; excursionPeakDay = i; }
      if (dev <= meanTolerance) {
        var days = i - excursionStart;
        if (days > 0) {
          // Compute shifted mean at recovery point: trailing window absorbs excursion
          var trailStart = Math.max(0, i - windowSize);
          var trailSum = 0;
          for (var t = trailStart; t <= i; t++) trailSum += dailyValues[t];
          var shiftedMean = trailSum / (i - trailStart + 1);
          var shiftedDev = Math.abs(dailyValues[excursionPeakDay] - shiftedMean);
          recoveryDays.push(days);
          excursionDetails.push({
            startDay: excursionStart,
            daysAgo: len - excursionStart,
            duration: days,
            peakDev: parseFloat(excursionPeakDev.toFixed(4)),
            peakDevDollar: parseFloat((excursionPeakDev * totalWeight).toFixed(2)),
            peakZ: parseFloat((excursionPeakDev / std).toFixed(2)),
            recovered: true,
            meanRevProfit: parseFloat((shiftedDev * totalWeight).toFixed(2))
          });
        }
        inExcursion = false;
        excursionPeakDev = 0;
      }
    }
  }

  // If still in excursion at end
  if (inExcursion) {
    var openDays = len - excursionStart;
    var trailStart = Math.max(0, len - 1 - windowSize);
    var trailSum = 0;
    for (var t = trailStart; t < len; t++) trailSum += dailyValues[t];
    var shiftedMean = trailSum / (len - trailStart);
    var shiftedDev = Math.abs(dailyValues[excursionPeakDay] - shiftedMean);
    excursionDetails.push({
      startDay: excursionStart,
      daysAgo: openDays,
      duration: openDays,
      peakDev: parseFloat(excursionPeakDev.toFixed(4)),
      peakDevDollar: parseFloat((excursionPeakDev * totalWeight).toFixed(2)),
      peakZ: parseFloat((excursionPeakDev / std).toFixed(2)),
      recovered: false,
      meanRevProfit: parseFloat((shiftedDev * totalWeight).toFixed(2))
    });
  }

  // Sort recovery days for median
  var sorted = recoveryDays.slice().sort(function(a, b) { return a - b; });
  var avgRec = 0;
  if (sorted.length > 0) {
    var sum = 0;
    for (var i = 0; i < sorted.length; i++) sum += sorted[i];
    avgRec = sum / sorted.length;
  }
  var medianRec = 0;
  if (sorted.length > 0) {
    var mid = Math.floor(sorted.length / 2);
    medianRec = sorted.length % 2 === 0 ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid];
  }

  result.maxMae = parseFloat(worstMae.toFixed(4));
  result.maxMaeDollar = parseFloat((worstMae * totalWeight).toFixed(2));
  result.avgRecoveryDays = parseFloat(avgRec.toFixed(1));
  result.medianRecoveryDays = parseFloat(medianRec.toFixed(1));
  result.excursionCount = excursionDetails.length;
  result.openExcursion = inExcursion;
  // Include up to 10 most recent excursions for detail dropdown
  result.excursions = excursionDetails.slice(-10).reverse();
  return result;
}

/**
 * Active Pair Correlation.
 * For each open pair, builds a daily spread series (priceA - priceB) from
 * TickerHistory, then computes daily changes (first differences). The average
 * pairwise Pearson correlation across all pairs measures how independently
 * the pairs are mean-reverting.
 *
 * @param {Array} openData - OpenTrades sheet data (row 0 = header)
 * @param {Object} histMap - {TICKER: [price0..priceN]} oldest-first
 * @return {Object} {avg, min, max, count, pairs: [{idA, idB, corr}]}
 */
function computePairCorrelation_(openData, histMap) {
  var result = { avg: 0, min: 0, max: 0, count: 0 };

  // Build per-pair daily spread change series
  var pairSeries = []; // [{id, changes: []}]
  for (var j = 1; j < openData.length; j++) {
    var rawId = openData[j][0];
    if (!rawId) continue;
    var info = parseTickerInfo(rawId);
    var hA = histMap[info.tA.toUpperCase()];
    var hB = histMap[info.tB.toUpperCase()];
    if (!hA || !hB) continue;
    var minLen = Math.min(hA.length, hB.length);
    if (minLen < 10) continue;
    // Build spread series and then daily changes
    var spreads = [];
    for (var d = 0; d < minLen; d++) {
      spreads.push(hA[hA.length - minLen + d] - hB[hB.length - minLen + d]);
    }
    var changes = [];
    for (var d = 1; d < spreads.length; d++) {
      changes.push(spreads[d] - spreads[d - 1]);
    }
    if (changes.length >= 10) {
      pairSeries.push({ id: info.id, changes: changes });
    }
  }

  if (pairSeries.length < 2) return result;

  // Compute pairwise Pearson correlation on daily changes
  var correlations = [];
  for (var i = 0; i < pairSeries.length; i++) {
    for (var k = i + 1; k < pairSeries.length; k++) {
      var a = pairSeries[i].changes;
      var b = pairSeries[k].changes;
      // Align to the shorter series (both are roughly the same length)
      var n = Math.min(a.length, b.length);
      // Use the most recent n observations
      var aSlice = a.slice(a.length - n);
      var bSlice = b.slice(b.length - n);

      var sumA = 0, sumB = 0;
      for (var d = 0; d < n; d++) { sumA += aSlice[d]; sumB += bSlice[d]; }
      var meanA = sumA / n, meanB = sumB / n;

      var covAB = 0, varA = 0, varB = 0;
      for (var d = 0; d < n; d++) {
        var da = aSlice[d] - meanA, db = bSlice[d] - meanB;
        covAB += da * db;
        varA += da * da;
        varB += db * db;
      }
      var corr = (varA > 0 && varB > 0) ? covAB / Math.sqrt(varA * varB) : 0;
      correlations.push(parseFloat(corr.toFixed(3)));
    }
  }

  if (correlations.length === 0) return result;

  // Build detail list with pair IDs
  var details = [];
  var idx = 0;
  for (var i = 0; i < pairSeries.length; i++) {
    for (var k = i + 1; k < pairSeries.length; k++) {
      details.push({
        pairA: pairSeries[i].id,
        pairB: pairSeries[k].id,
        corr: correlations[idx]
      });
      idx++;
    }
  }
  // Sort by absolute correlation descending (most correlated first)
  details.sort(function(a, b) { return Math.abs(b.corr) - Math.abs(a.corr); });

  var sum = 0, minC = correlations[0], maxC = correlations[0];
  for (var i = 0; i < correlations.length; i++) {
    sum += correlations[i];
    if (correlations[i] < minC) minC = correlations[i];
    if (correlations[i] > maxC) maxC = correlations[i];
  }

  result.avg = parseFloat((sum / correlations.length).toFixed(3));
  result.min = minC;
  result.max = maxC;
  result.count = correlations.length;
  result.pairs = pairSeries.length;
  result.details = details.slice(0, 15); // top 15 by |corr|
  return result;
}

/**
 * Historical Probability Engine.
 * Scans the daily spread array for "trigger points" where the spread was at
 * approximately the same level as the reference value (±5% of the full range).
 * From each trigger, looks forward to compute:
 *   1) Max Adverse Excursion (worst-case widening before reversion)
 *   2) Probability of continued widening vs immediate reversion
 *   3) Per-window (30/60/90d) probability of touching the mean
 *
 * @param {number[]} dailyValues - Full daily spread array (oldest first)
 * @param {number} refValue - Reference spread (current market or entry)
 * @param {Object} rollingZ - The rollingZ object with per-window mean/std
 * @param {number} totalWeight - Position size multiplier for dollar conversion
 * @return {Object} probabilities result
 */
function computeHistoricalProbabilities_(dailyValues, refValue, rollingZ, totalWeight, customWindows) {
  var emptySpreadAnalysis = { triggers: 0, refValue: 0, meanTarget: 0, invertedCount: 0, details: [] };
  var result = { triggers: 0, badScenario: null, winRates: {}, spreadAnalysis: emptySpreadAnalysis, toleranceWidened: false };
  var len = dailyValues.length;
  if (len < 20) return result; // need minimum history

  // Determine tolerance as 5% of the spread's full range
  var minVal = dailyValues[0], maxVal = dailyValues[0];
  for (var i = 1; i < len; i++) {
    if (dailyValues[i] < minVal) minVal = dailyValues[i];
    if (dailyValues[i] > maxVal) maxVal = dailyValues[i];
  }
  var fullRange = maxVal - minVal;
  if (fullRange < 0.001) return result; // flat spread, no meaningful stats
  var tolerance = fullRange * 0.05;

  // Determine if we're above or below the 90d mean to define "adverse" direction.
  // If spread > mean, adverse = spread widens further up. If spread < mean, adverse = widens further down.
  var mean90 = (rollingZ && rollingZ['90d'] && !rollingZ['90d'].insufficient) ? rollingZ['90d'].mean : 0;
  var isAboveMean = refValue >= mean90;

  // ── Two-Pass Trigger Search ──
  // Pass 1: ±5% tolerance. Pass 2 (fallback): ±10% if Pass 1 finds 0 triggers.
  var triggers = [];
  var lastTriggerDay = -10; // prevent overlapping triggers (min 5-day gap)
  for (var i = 0; i < len - 5; i++) {
    if (Math.abs(dailyValues[i] - refValue) <= tolerance && (i - lastTriggerDay) >= 5) {
      triggers.push(i);
      lastTriggerDay = i;
    }
  }

  // Pass 2: Auto-expand to 10% tolerance if Pass 1 found 0 triggers
  if (triggers.length === 0) {
    tolerance = fullRange * 0.10;
    lastTriggerDay = -10;
    for (var i = 0; i < len - 5; i++) {
      if (Math.abs(dailyValues[i] - refValue) <= tolerance && (i - lastTriggerDay) >= 5) {
        triggers.push(i);
        lastTriggerDay = i;
      }
    }
    if (triggers.length > 0) result.toleranceWidened = true;
  }

  result.triggers = triggers.length;
  if (triggers.length < 1) return result; // return with 0 triggers — UI will handle gracefully

  // ── FEATURE 1: Bad Scenario (Max Adverse Excursion) ──
  var maeValues = []; // max adverse excursion per trigger (in spread units)
  var widenCount = 0; // how many triggers saw widening before any reversion

  for (var t = 0; t < triggers.length; t++) {
    var idx = triggers[t];
    var triggerSpread = dailyValues[idx];
    var maxAdverse = 0;
    var sawWidening = false;
    // Look forward from trigger, track worst excursion
    var lookForward = Math.min(30, len - idx - 1); // 30-day worst-case window
    for (var f = 1; f <= lookForward; f++) {
      var futureSpread = dailyValues[idx + f];
      var excursion;
      if (isAboveMean) {
        // Spread is above mean → adverse = going even higher
        excursion = futureSpread - triggerSpread;
      } else {
        // Spread is below mean → adverse = going even lower
        excursion = triggerSpread - futureSpread;
      }
      if (excursion > maxAdverse) maxAdverse = excursion;
    }
    maeValues.push(maxAdverse);
  }

  // Average and P75 MAE
  var maeSum = 0;
  var sortedMae = maeValues.slice().sort(function(a, b) { return a - b; });
  for (var i = 0; i < maeValues.length; i++) maeSum += maeValues[i];
  var avgMae = maeSum / maeValues.length;
  var p75Idx = Math.floor((sortedMae.length - 1) * 0.75);
  var p75Mae = sortedMae[Math.min(p75Idx, sortedMae.length - 1)];

  // ── FEATURE 2: Mean Reversion Win Rates per Window + window-matched Widen Prob ──
  // widenProb is now: % of triggers where spread worsens by >1σ BEFORE touching mean (per window)
  var windows = (customWindows && customWindows.length > 0) ? customWindows : [30, 60, 90];
  var widenCountByWindow = {};
  for (var w = 0; w < windows.length; w++) {
    var n = windows[w];
    var wKey = n + 'd';
    // For custom windows that don't have a matching rollingZ entry, fall back to 90d or 30d mean
    var wMean;
    if (rollingZ[wKey] && !rollingZ[wKey].insufficient) {
      wMean = rollingZ[wKey].mean;
    } else if (rollingZ['30d'] && !rollingZ['30d'].insufficient) {
      wMean = rollingZ['30d'].mean;
    } else {
      wMean = mean90;
    }
    // Mean reversion tolerance: spread counts as "touched mean" if it gets within 10% of distance to mean
    var distToMean = Math.abs(refValue - wMean);
    var meanTolerance = distToMean * 0.10;
    if (meanTolerance < 0.01) meanTolerance = 0.01; // minimum floor

    // 1σ adverse threshold for widen definition
    var oneSigma = (rollingZ && rollingZ['90d'] && !rollingZ['90d'].insufficient) ? rollingZ['90d'].std : (fullRange * 0.15);

    var wins = 0, eligible = 0, widenCount = 0;
    for (var t = 0; t < triggers.length; t++) {
      var idx = triggers[t];
      var remaining = len - idx - 1;
      if (remaining < Math.floor(n * 0.5)) continue; // need at least half the window to be meaningful
      eligible++;
      var lookAhead = Math.min(n, remaining);
      var touched = false;
      var widened = false;
      var triggerSpread = dailyValues[idx];
      for (var f = 1; f <= lookAhead; f++) {
        var futVal = dailyValues[idx + f];
        // Check mean touch first
        if (Math.abs(futVal - wMean) <= meanTolerance) {
          touched = true;
          break;
        }
        // Check adverse >1σ from trigger (spread worsening)
        if (!widened) {
          var excursion = isAboveMean ? (futVal - triggerSpread) : (triggerSpread - futVal);
          if (excursion > oneSigma) widened = true;
        }
      }
      if (touched) wins++;
      if (widened && !touched) widenCount++; // widened >1σ without ever touching mean in window
    }

    widenCountByWindow[wKey] = { widenCount: widenCount, eligible: eligible };

    result.winRates[wKey] = {
      rate: eligible > 0 ? parseFloat((wins / eligible * 100).toFixed(1)) : 0,
      wins: wins,
      eligible: eligible
    };
  }

  // Compute widen prob from 30d window (or best available)
  var bestWiden = widenCountByWindow['30d'] || widenCountByWindow['60d'] || widenCountByWindow['90d'] || { widenCount: 0, eligible: 0 };
  var widenProbPct = bestWiden.eligible > 0 ? parseFloat((bestWiden.widenCount / bestWiden.eligible * 100).toFixed(1)) : 0;
  // Derive primary win rate for stagnant calc (prefer 30d)
  var primaryWR = (result.winRates['30d'] && result.winRates['30d'].eligible > 0) ? result.winRates['30d'].rate
    : (result.winRates['60d'] && result.winRates['60d'].eligible > 0) ? result.winRates['60d'].rate
    : (result.winRates['90d'] && result.winRates['90d'].eligible > 0) ? result.winRates['90d'].rate : 0;
  var stagnantRate = parseFloat(Math.max(0, 100 - primaryWR - widenProbPct).toFixed(1));

  result.badScenario = {
    avgMae: parseFloat(avgMae.toFixed(4)),
    avgMaeDollar: parseFloat((avgMae * totalWeight).toFixed(2)),
    p75Mae: parseFloat(p75Mae.toFixed(4)),
    p75MaeDollar: parseFloat((p75Mae * totalWeight).toFixed(2)),
    wideningProb: widenProbPct,
    stagnantRate: stagnantRate,
    sampleSize: triggers.length
  };

  // ── FEATURE 3: Spread-Level Analysis ──
  // Finds historical days where spread MAGNITUDE ≈ current (10% tolerance), regardless of sign.
  // This captures both same-direction triggers AND inverted (role-reversal) triggers where
  // tickers swapped long/short positions. Direction is determined per-trigger based on actual value.
  var daysToMean = [];
  var peakDevFromRef = [];
  var meanTarget = mean90;
  var meanTolRef = Math.abs(refValue - meanTarget) * 0.15;
  if (meanTolRef < 0.01) meanTolRef = 0.01;

  var saAbsRef = Math.abs(refValue);
  var saTolerance = fullRange * 0.10; // 10% of range for spread-level matching
  var saTriggers = [];
  var saLastDay = -10;
  var saInvertedCount = 0;
  for (var si = 0; si < len - 5; si++) {
    if (Math.abs(Math.abs(dailyValues[si]) - saAbsRef) <= saTolerance && (si - saLastDay) >= 5) {
      var inverted = (refValue > 0 && dailyValues[si] < 0) || (refValue < 0 && dailyValues[si] > 0);
      saTriggers.push({ idx: si, inverted: inverted });
      if (inverted) saInvertedCount++;
      saLastDay = si;
    }
  }

  var saDetails = [];
  for (var t = 0; t < saTriggers.length; t++) {
    var idx = saTriggers[t].idx;
    var remaining = len - idx - 1;
    if (remaining < 5) continue;
    var lookAhead = Math.min(90, remaining);
    var peakDev = 0;
    var revertDay = -1;
    // Determine adverse direction from this trigger's actual value vs mean
    var trigAboveMean = dailyValues[idx] >= meanTarget;
    for (var f = 1; f <= lookAhead; f++) {
      var devFromTrigger;
      if (trigAboveMean) {
        devFromTrigger = dailyValues[idx + f] - dailyValues[idx]; // further above = adverse
      } else {
        devFromTrigger = dailyValues[idx] - dailyValues[idx + f]; // further below = adverse
      }
      if (devFromTrigger > peakDev) peakDev = devFromTrigger;
      if (revertDay < 0 && Math.abs(dailyValues[idx + f] - meanTarget) <= meanTolRef) {
        revertDay = f;
      }
    }
    peakDevFromRef.push(peakDev);
    if (revertDay > 0) daysToMean.push(revertDay);
    // Compute mean reversion profit: how much spread moved toward mean from trigger
    var endIdx = revertDay > 0 ? idx + revertDay : idx + lookAhead;
    var mrProfit = trigAboveMean ? dailyValues[idx] - dailyValues[endIdx] : dailyValues[endIdx] - dailyValues[idx];
    saDetails.push({
      daysAgo: len - idx,
      peakDevDollar: parseFloat((peakDev * totalWeight).toFixed(2)),
      duration: revertDay > 0 ? revertDay : null,
      reverted: revertDay > 0,
      inverted: saTriggers[t].inverted,
      meanRevProfit: parseFloat((mrProfit * totalWeight).toFixed(2))
    });
  }

  // Compute summary stats
  var spreadAnalysis = { triggers: saTriggers.length, refValue: parseFloat(refValue.toFixed(2)), meanTarget: parseFloat(meanTarget.toFixed(2)), invertedCount: saInvertedCount, details: saDetails, toleranceWidened: result.toleranceWidened };
  if (daysToMean.length > 0) {
    var dtmSorted = daysToMean.slice().sort(function(a, b) { return a - b; });
    var dtmSum = 0;
    for (var i = 0; i < dtmSorted.length; i++) dtmSum += dtmSorted[i];
    var dtmMid = Math.floor(dtmSorted.length / 2);
    spreadAnalysis.avgDaysToMean = parseFloat((dtmSum / dtmSorted.length).toFixed(1));
    spreadAnalysis.medianDaysToMean = dtmSorted.length % 2 === 0 ? (dtmSorted[dtmMid - 1] + dtmSorted[dtmMid]) / 2 : dtmSorted[dtmMid];
    spreadAnalysis.reversionRate = parseFloat((daysToMean.length / saTriggers.length * 100).toFixed(1));
    spreadAnalysis.reverted = daysToMean.length;
  }
  if (peakDevFromRef.length > 0) {
    var pdSorted = peakDevFromRef.slice().sort(function(a, b) { return a - b; });
    var pdSum = 0;
    for (var i = 0; i < pdSorted.length; i++) pdSum += pdSorted[i];
    var p75i = Math.floor(pdSorted.length * 0.75);
    spreadAnalysis.avgPeakDev = parseFloat((pdSum / pdSorted.length).toFixed(4));
    spreadAnalysis.avgPeakDevDollar = parseFloat((spreadAnalysis.avgPeakDev * totalWeight).toFixed(2));
    spreadAnalysis.p75PeakDev = parseFloat(pdSorted[Math.min(p75i, pdSorted.length - 1)].toFixed(4));
    spreadAnalysis.p75PeakDevDollar = parseFloat((spreadAnalysis.p75PeakDev * totalWeight).toFixed(2));
  }
  result.spreadAnalysis = spreadAnalysis;

  return result;
}

/**
 * Wide-tolerance variant of computeHistoricalProbabilities_ for single-pair alert analysis.
 * Uses 10% tolerance (vs 5%), minimum 2 triggers (vs 3), and 3-day gap (vs 5-day).
 * This is necessary because alert pairs are at extreme Z-scores by definition,
 * so the standard 5% tolerance often can't find enough historical occurrences.
 */
function computeHistoricalProbabilitiesWide_(dailyValues, refValue, rollingZ, totalWeight, customWindows) {
  var emptySpreadAnalysis = { triggers: 0, refValue: 0, meanTarget: 0, invertedCount: 0, details: [] };
  var result = { triggers: 0, badScenario: null, winRates: {}, spreadAnalysis: emptySpreadAnalysis, toleranceWidened: false };
  var len = dailyValues.length;
  if (len < 20) return result;

  var minVal = dailyValues[0], maxVal = dailyValues[0];
  for (var i = 1; i < len; i++) {
    if (dailyValues[i] < minVal) minVal = dailyValues[i];
    if (dailyValues[i] > maxVal) maxVal = dailyValues[i];
  }
  var fullRange = maxVal - minVal;
  if (fullRange < 0.001) return result;
  var tolerance = fullRange * 0.10; // 10% tolerance (wider than standard 5%)

  var mean90 = (rollingZ && rollingZ['90d'] && !rollingZ['90d'].insufficient) ? rollingZ['90d'].mean : 0;
  var isAboveMean = refValue >= mean90;

  var triggers = [];
  var lastTriggerDay = -5;
  for (var i = 0; i < len - 5; i++) {
    if (Math.abs(dailyValues[i] - refValue) <= tolerance && (i - lastTriggerDay) >= 3) {
      triggers.push(i);
      lastTriggerDay = i;
    }
  }
  result.triggers = triggers.length;
  if (triggers.length < 1) return result; // return with 0 triggers — UI handles gracefully

  var maeValues = [];
  for (var t = 0; t < triggers.length; t++) {
    var idx = triggers[t];
    var triggerSpread = dailyValues[idx];
    var maxAdverse = 0;
    var lookForward = Math.min(30, len - idx - 1);
    for (var f = 1; f <= lookForward; f++) {
      var futureSpread = dailyValues[idx + f];
      var excursion;
      if (isAboveMean) { excursion = futureSpread - triggerSpread; }
      else { excursion = triggerSpread - futureSpread; }
      if (excursion > maxAdverse) maxAdverse = excursion;
    }
    maeValues.push(maxAdverse);
  }

  var maeSum = 0;
  var sortedMae = maeValues.slice().sort(function(a, b) { return a - b; });
  for (var i = 0; i < maeValues.length; i++) maeSum += maeValues[i];
  var avgMae = maeSum / maeValues.length;
  var p75Idx = Math.floor((sortedMae.length - 1) * 0.75);
  var p75Mae = sortedMae[Math.min(p75Idx, sortedMae.length - 1)];

  var windows = (customWindows && customWindows.length > 0) ? customWindows : [30, 60, 90];
  var widenCountByWindow = {};
  for (var w = 0; w < windows.length; w++) {
    var n = windows[w];
    var wKey = n + 'd';
    var wMean;
    if (rollingZ[wKey] && !rollingZ[wKey].insufficient) {
      wMean = rollingZ[wKey].mean;
    } else if (rollingZ['30d'] && !rollingZ['30d'].insufficient) {
      wMean = rollingZ['30d'].mean;
    } else {
      wMean = mean90;
    }
    var distToMean = Math.abs(refValue - wMean);
    var meanTolerance = distToMean * 0.10;
    if (meanTolerance < 0.01) meanTolerance = 0.01;

    var oneSigma = (rollingZ && rollingZ['90d'] && !rollingZ['90d'].insufficient) ? rollingZ['90d'].std : (fullRange * 0.15);

    var wins = 0, eligible = 0, widenCount = 0;
    for (var t = 0; t < triggers.length; t++) {
      var idx = triggers[t];
      var remaining = len - idx - 1;
      if (remaining < Math.floor(n * 0.4)) continue; // 40% forward data (relaxed)
      eligible++;
      var lookAhead = Math.min(n, remaining);
      var touched = false;
      var widened = false;
      var triggerSpread = dailyValues[idx];
      for (var f = 1; f <= lookAhead; f++) {
        var futVal = dailyValues[idx + f];
        if (Math.abs(futVal - wMean) <= meanTolerance) { touched = true; break; }
        if (!widened) {
          var excursion = isAboveMean ? (futVal - triggerSpread) : (triggerSpread - futVal);
          if (excursion > oneSigma) widened = true;
        }
      }
      if (touched) wins++;
      if (widened && !touched) widenCount++;
    }
    widenCountByWindow[wKey] = { widenCount: widenCount, eligible: eligible };
    result.winRates[wKey] = {
      rate: eligible > 0 ? parseFloat((wins / eligible * 100).toFixed(1)) : 0,
      wins: wins,
      eligible: eligible
    };
  }

  var bestWiden = widenCountByWindow['30d'] || widenCountByWindow['60d'] || widenCountByWindow['90d'] || { widenCount: 0, eligible: 0 };
  var widenProbPct = bestWiden.eligible > 0 ? parseFloat((bestWiden.widenCount / bestWiden.eligible * 100).toFixed(1)) : 0;
  // Derive primary win rate for stagnant calc (prefer 30d)
  var primaryWR = (result.winRates['30d'] && result.winRates['30d'].eligible > 0) ? result.winRates['30d'].rate
    : (result.winRates['60d'] && result.winRates['60d'].eligible > 0) ? result.winRates['60d'].rate
    : (result.winRates['90d'] && result.winRates['90d'].eligible > 0) ? result.winRates['90d'].rate : 0;
  var stagnantRate = parseFloat(Math.max(0, 100 - primaryWR - widenProbPct).toFixed(1));

  result.badScenario = {
    avgMae: parseFloat(avgMae.toFixed(4)),
    avgMaeDollar: parseFloat((avgMae * totalWeight).toFixed(2)),
    p75Mae: parseFloat(p75Mae.toFixed(4)),
    p75MaeDollar: parseFloat((p75Mae * totalWeight).toFixed(2)),
    wideningProb: widenProbPct,
    stagnantRate: stagnantRate,
    sampleSize: triggers.length
  };

  // ── Spread-Level Analysis (wide variant) ──
  // Matches by absolute spread magnitude (10% tolerance), captures role-reversal triggers
  var daysToMean = [];
  var peakDevFromRef = [];
  var meanTarget = mean90;
  var meanTolRef = Math.abs(refValue - meanTarget) * 0.15;
  if (meanTolRef < 0.01) meanTolRef = 0.01;

  var saAbsRef = Math.abs(refValue);
  var saTolerance = fullRange * 0.10;
  var saTriggers = [];
  var saLastDay = -5;
  var saInvertedCount = 0;
  for (var si = 0; si < len - 5; si++) {
    if (Math.abs(Math.abs(dailyValues[si]) - saAbsRef) <= saTolerance && (si - saLastDay) >= 3) {
      var inverted = (refValue > 0 && dailyValues[si] < 0) || (refValue < 0 && dailyValues[si] > 0);
      saTriggers.push({ idx: si, inverted: inverted });
      if (inverted) saInvertedCount++;
      saLastDay = si;
    }
  }

  var saDetails = [];
  for (var t = 0; t < saTriggers.length; t++) {
    var idx = saTriggers[t].idx;
    var remaining = len - idx - 1;
    if (remaining < 5) continue;
    var lookAhead = Math.min(90, remaining);
    var peakDev = 0;
    var revertDay = -1;
    var trigAboveMean = dailyValues[idx] >= meanTarget;
    for (var f = 1; f <= lookAhead; f++) {
      var devFromTrigger;
      if (trigAboveMean) { devFromTrigger = dailyValues[idx + f] - dailyValues[idx]; }
      else { devFromTrigger = dailyValues[idx] - dailyValues[idx + f]; }
      if (devFromTrigger > peakDev) peakDev = devFromTrigger;
      if (revertDay < 0 && Math.abs(dailyValues[idx + f] - meanTarget) <= meanTolRef) {
        revertDay = f;
      }
    }
    peakDevFromRef.push(peakDev);
    if (revertDay > 0) daysToMean.push(revertDay);
    var endIdx = revertDay > 0 ? idx + revertDay : idx + lookAhead;
    var mrProfit = trigAboveMean ? dailyValues[idx] - dailyValues[endIdx] : dailyValues[endIdx] - dailyValues[idx];
    saDetails.push({
      daysAgo: len - idx,
      peakDevDollar: parseFloat((peakDev * totalWeight).toFixed(2)),
      duration: revertDay > 0 ? revertDay : null,
      reverted: revertDay > 0,
      inverted: saTriggers[t].inverted,
      meanRevProfit: parseFloat((mrProfit * totalWeight).toFixed(2))
    });
  }

  var spreadAnalysis = { triggers: saTriggers.length, refValue: parseFloat(refValue.toFixed(2)), meanTarget: parseFloat(meanTarget.toFixed(2)), invertedCount: saInvertedCount, details: saDetails, toleranceWidened: result.toleranceWidened };
  if (daysToMean.length > 0) {
    var dtmSorted = daysToMean.slice().sort(function(a, b) { return a - b; });
    var dtmSum = 0;
    for (var i = 0; i < dtmSorted.length; i++) dtmSum += dtmSorted[i];
    var dtmMid = Math.floor(dtmSorted.length / 2);
    spreadAnalysis.avgDaysToMean = parseFloat((dtmSum / dtmSorted.length).toFixed(1));
    spreadAnalysis.medianDaysToMean = dtmSorted.length % 2 === 0 ? (dtmSorted[dtmMid - 1] + dtmSorted[dtmMid]) / 2 : dtmSorted[dtmMid];
    spreadAnalysis.reversionRate = parseFloat((daysToMean.length / saTriggers.length * 100).toFixed(1));
    spreadAnalysis.reverted = daysToMean.length;
  }
  if (peakDevFromRef.length > 0) {
    var pdSorted = peakDevFromRef.slice().sort(function(a, b) { return a - b; });
    var pdSum = 0;
    for (var i = 0; i < pdSorted.length; i++) pdSum += pdSorted[i];
    var p75i = Math.floor(pdSorted.length * 0.75);
    spreadAnalysis.avgPeakDev = parseFloat((pdSum / pdSorted.length).toFixed(4));
    spreadAnalysis.avgPeakDevDollar = parseFloat((spreadAnalysis.avgPeakDev * totalWeight).toFixed(2));
    spreadAnalysis.p75PeakDev = parseFloat(pdSorted[Math.min(p75i, pdSorted.length - 1)].toFixed(4));
    spreadAnalysis.p75PeakDevDollar = parseFloat((spreadAnalysis.p75PeakDev * totalWeight).toFixed(2));
  }
  result.spreadAnalysis = spreadAnalysis;

  return result;
}

/**
 * Portfolio Analytics endpoint. Two modes:
 *  - live: reads OpenTrades, resolves legs from actual positions
 *  - sandbox: uses user-provided legs JSON
 */
function getPortfolioAnalytics(mode, legsJson) {
  try {
    var ss = SpreadsheetApp.getActive();
    var histMap = readTickerHistMap_(ss);
    var legs = [];

    if (mode === 'live') {
      // Read OpenTrades and resolve legs
      var openSheet = ss.getSheetByName('OpenTrades');
      if (!openSheet || openSheet.getLastRow() <= 1) return { mode: 'live', trades: 0, metrics: { dailyValues: [] } };
      var openData = openSheet.getDataRange().getValues();

      // Read live data for current prices
      var liveRows = [];
      var intra = ss.getSheetByName('WebCache');
      if (!intra || intra.getLastRow() <= 1) intra = ss.getSheetByName('Live');
      var credit = ss.getSheetByName('WebCacheCredit');
      var sources = [intra, credit];
      for (var s = 0; s < sources.length; s++) {
        var ls = sources[s];
        if (ls && ls.getLastRow() > 1) {
          var rows = ls.getRange(2, 1, ls.getLastRow()-1, 24).getValues();
          liveRows = liveRows.concat(rows);
        }
      }

      var entrySpreadDollar = 0, entryTotalSize = 0, entryGrossLong = 0, entryGrossShort = 0;
      for (var j = 1; j < openData.length; j++) {
        var rawId = openData[j][0];
        if (!rawId) continue;
        var info = parseTickerInfo(rawId);
        var sA = parseMoney(openData[j][4]);
        var sB = parseMoney(openData[j][5]);
        var costA = parseMoney(openData[j][2]);
        var costB = parseMoney(openData[j][3]);

        // Each open trade has two legs: A and B, with signed sizes
        if (sA !== 0) {
          legs.push({ ticker: info.tA, size: Math.abs(sA), direction: sA > 0 ? 1 : -1 });
          entrySpreadDollar += sA * costA; // signed: long adds, short subtracts
          entryTotalSize += Math.abs(sA);
          if (sA > 0) entryGrossLong += Math.abs(sA) * costA; else entryGrossShort += Math.abs(sA) * costA;
        }
        if (sB !== 0) {
          legs.push({ ticker: info.tB, size: Math.abs(sB), direction: sB > 0 ? 1 : -1 });
          entrySpreadDollar += sB * costB;
          entryTotalSize += Math.abs(sB);
          if (sB > 0) entryGrossLong += Math.abs(sB) * costB; else entryGrossShort += Math.abs(sB) * costB;
        }
      }
      // Normalize entry spread to size-weighted average (divide by sum of pair sizes)
      var entryWeight = entryTotalSize / 2;
      if (entryWeight === 0) entryWeight = 1;

      var metrics = computeBasketMetrics_(legs, histMap);

      // Compute current live value for PnL
      var currentLiveValue = 0;
      for (var j = 1; j < openData.length; j++) {
        var rawId = openData[j][0];
        if (!rawId) continue;
        var openAnchor = cleanId(rawId);
        var sA = parseMoney(openData[j][4]);
        var sB = parseMoney(openData[j][5]);
        var pair = null;
        for (var k = 0; k < liveRows.length; k++) {
          if (liveRows[k][0] && cleanId(liveRows[k][0]) === openAnchor) { pair = liveRows[k]; break; }
        }
        if (pair) {
          currentLiveValue += sA * (parseFloat(pair[3]) || 0);
          currentLiveValue += sB * (parseFloat(pair[4]) || 0);
        }
      }

      // Accumulate dividends
      var totalPaidDiv = 0, totalRcvdDiv = 0;
      for (var j = 1; j < openData.length; j++) {
        totalPaidDiv += parseMoney(openData[j][7]);
        totalRcvdDiv += parseMoney(openData[j][8]);
      }

      var rawEntrySpread = entrySpreadDollar / entryWeight;
      metrics.entrySpread = isFinite(rawEntrySpread) ? parseFloat(rawEntrySpread.toFixed(2)) : 0;
      metrics.entryGrossLong = parseFloat(entryGrossLong.toFixed(2));
      metrics.entryGrossShort = parseFloat(entryGrossShort.toFixed(2));
      metrics.entryGrossExposure = parseFloat((entryGrossLong + entryGrossShort).toFixed(2));
      metrics.capitalGains = parseFloat((currentLiveValue - entrySpreadDollar).toFixed(2));
      metrics.netDividends = parseFloat((totalRcvdDiv - totalPaidDiv).toFixed(2));
      metrics.totalPnL = parseFloat((currentLiveValue - entrySpreadDollar + totalRcvdDiv - totalPaidDiv).toFixed(2));

      // Wide-tolerance fallback for probability engine (live mode)
      if (metrics.dailyValuesFull_ && metrics.dailyValuesFull_.length >= 20) {
        var stdTriggers = metrics.probabilities ? metrics.probabilities.triggers : 0;
        if (stdTriggers < 3) {
          var liveRef = metrics.dailyValuesFull_[metrics.dailyValuesFull_.length - 1];
          var wideProb = computeHistoricalProbabilitiesWide_(
            metrics.dailyValuesFull_, liveRef, metrics.rollingZ, metrics.totalWeight
          );
          if (wideProb.triggers > stdTriggers) {
            metrics.probabilities = wideProb;
          }
        }

        // ── Entry Spread Probability Analysis ──
        // Run the same probability engine but anchored to entry spread instead of current spread.
        // This tells us: "historically when spread was at MY entry level, what happened?"
        var entryRef = metrics.entrySpread;
        var entryProb = computeHistoricalProbabilities_(
          metrics.dailyValuesFull_, entryRef, metrics.rollingZ, metrics.totalWeight
        );
        if (entryProb.triggers < 3) {
          var entryWideProb = computeHistoricalProbabilitiesWide_(
            metrics.dailyValuesFull_, entryRef, metrics.rollingZ, metrics.totalWeight
          );
          if (entryWideProb.triggers > entryProb.triggers) {
            entryProb = entryWideProb;
          }
        }
        metrics.entryProbabilities = entryProb;
      }

      // ── Active Pair Correlation ──
      // Compute avg pairwise correlation of daily spread changes across all open pairs
      var pairCorrelation = computePairCorrelation_(openData, histMap);
      metrics.pairCorrelation = pairCorrelation;

      delete metrics.dailyValuesFull_; // strip internal array before response
      return { mode: 'live', trades: openData.length - 1, legs: legs.length, metrics: metrics };
    }
    else if (mode === 'sandbox') {
      // ── Cross-Matrix Sandbox: every Long × every Short ──
      // Parse user-provided legs
      var userLegs = [];
      try { userLegs = JSON.parse(legsJson); } catch(e) { return { mode: 'sandbox', error: 'Invalid legs JSON' }; }
      var longLegs = [], shortLegs = [];
      for (var i = 0; i < userLegs.length; i++) {
        var ul = userLegs[i];
        if (!ul.ticker || !ul.size || !ul.direction) continue;
        var parsed = {
          ticker: String(ul.ticker).trim().toUpperCase(),
          size: Math.abs(parseFloat(ul.size) || 0),
          direction: parseFloat(ul.direction) > 0 ? 1 : -1,
          entryPrice: parseFloat(ul.entryPrice) || 0
        };
        if (parsed.size === 0) continue;
        if (parsed.direction > 0) longLegs.push(parsed); else shortLegs.push(parsed);
        legs.push(parsed);
      }

      // Fallback: if all legs are same direction (no cross possible), run legacy single-basket
      if (longLegs.length === 0 || shortLegs.length === 0) {
        return runLegacySandbox_(legs, userLegs, histMap);
      }

      // Flag which tickers are missing history
      var missing = [];
      for (var i = 0; i < legs.length; i++) {
        if (!histMap[legs[i].ticker]) missing.push(legs[i].ticker);
      }

      // Compute total shares per side for proportional allocation
      var totalLongShares = 0, totalShortShares = 0;
      for (var i = 0; i < longLegs.length; i++) totalLongShares += longLegs[i].size;
      for (var i = 0; i < shortLegs.length; i++) totalShortShares += shortLegs[i].size;

      // Generate cross-pairs (cap at 20 by weight, sorted descending)
      var crossPairDefs = [];
      for (var li = 0; li < longLegs.length; li++) {
        for (var si = 0; si < shortLegs.length; si++) {
          var weight = (longLegs[li].size / totalLongShares) * (shortLegs[si].size / totalShortShares);
          var effLong = longLegs[li].size * (shortLegs[si].size / totalShortShares);
          var effShort = shortLegs[si].size * (longLegs[li].size / totalLongShares);
          crossPairDefs.push({
            longLeg: longLegs[li],
            shortLeg: shortLegs[si],
            weight: weight,
            effLongShares: parseFloat(effLong.toFixed(1)),
            effShortShares: parseFloat(effShort.toFixed(1))
          });
        }
      }
      // Sort by weight descending and cap at 20
      crossPairDefs.sort(function(a,b) { return b.weight - a.weight; });
      if (crossPairDefs.length > 20) crossPairDefs = crossPairDefs.slice(0, 20);

      // ── Run probability engine on each cross-pair ──
      var crossPairResults = [];
      var aggZ30 = 0, aggZ60 = 0, aggZ90 = 0;
      var aggWR30 = 0, aggWR60 = 0, aggWR90 = 0;
      var aggEP30 = 0, aggEP60 = 0, aggEP90 = 0;
      var aggBadAvg = 0, aggBadP75 = 0, aggWidenProb = 0, aggStagnant = 0;
      var aggWeightSum = 0; // for renormalization after excluding missing-ticker pairs
      // Per-metric weight sums to avoid denominator inflation when some pairs have null values
      var wrWeightSum30 = 0, wrWeightSum60 = 0, wrWeightSum90 = 0;
      var zWeightSum30 = 0, zWeightSum60 = 0, zWeightSum90 = 0;
      var widenWeightSum = 0, stagnantWeightSum = 0;
      var totalGrossLong = 0, totalGrossShort = 0;
      var worstZ = null, worstZPair = '';

      for (var cp = 0; cp < crossPairDefs.length; cp++) {
        var def = crossPairDefs[cp];
        var lTk = def.longLeg.ticker;
        var sTk = def.shortLeg.ticker;

        // Skip pairs where either ticker is missing history
        if (!histMap[lTk] || !histMap[sTk]) {
          crossPairResults.push({
            longTicker: lTk, shortTicker: sTk,
            weight: parseFloat((def.weight * 100).toFixed(2)),
            effLongShares: def.effLongShares, effShortShares: def.effShortShares,
            error: 'Missing history for ' + (!histMap[lTk] ? lTk : sTk)
          });
          continue;
        }

        // Run with 100 normalized shares per side for clean Z-scores/probabilities
        var pairLegs = [
          { ticker: lTk, size: 100, direction: 1 },
          { ticker: sTk, size: 100, direction: -1 }
        ];
        var pairMetrics = computeBasketMetrics_(pairLegs, histMap);

        // Compute entry spread for this pair (using user's entry prices, normalized)
        var entrySpreadPerShare = def.longLeg.entryPrice - def.shortLeg.entryPrice;

        // Override expected profit to use entry spread
        var pairRZ = pairMetrics.rollingZ || {};
        for (var wKey in pairRZ) {
          if (pairRZ[wKey] && !pairRZ[wKey].insufficient) {
            pairRZ[wKey].expectedProfit = parseFloat((pairRZ[wKey].mean - entrySpreadPerShare).toFixed(4));
          }
        }

        // Run probability engine anchored to entry spread
        var pairProb = { triggers: 0, badScenario: null, winRates: {} };
        if (pairMetrics.dailyValuesFull_ && pairMetrics.dailyValuesFull_.length >= 20) {
          pairProb = computeHistoricalProbabilities_(
            pairMetrics.dailyValuesFull_, entrySpreadPerShare, pairRZ, pairMetrics.totalWeight
          );
          if (pairProb.triggers < 3) {
            var widePairProb = computeHistoricalProbabilitiesWide_(
              pairMetrics.dailyValuesFull_, entrySpreadPerShare, pairRZ, pairMetrics.totalWeight
            );
            if (widePairProb.triggers > pairProb.triggers) pairProb = widePairProb;
          }
        }

        // Extract per-pair metrics
        var z30 = pairRZ['30d'] && !pairRZ['30d'].insufficient ? pairRZ['30d'].z : null;
        var z60 = pairRZ['60d'] && !pairRZ['60d'].insufficient ? pairRZ['60d'].z : null;
        var z90 = pairRZ['90d'] && !pairRZ['90d'].insufficient ? pairRZ['90d'].z : null;
        var wrObj30 = pairProb.winRates && pairProb.winRates['30d'] ? pairProb.winRates['30d'] : null;
        var wrObj60 = pairProb.winRates && pairProb.winRates['60d'] ? pairProb.winRates['60d'] : null;
        var wrObj90 = pairProb.winRates && pairProb.winRates['90d'] ? pairProb.winRates['90d'] : null;
        var wr30 = wrObj30 && wrObj30.eligible >= 2 ? wrObj30.rate : null;
        var wr60 = wrObj60 && wrObj60.eligible >= 2 ? wrObj60.rate : null;
        var wr90 = wrObj90 && wrObj90.eligible >= 2 ? wrObj90.rate : null;
        var ep30 = pairRZ['30d'] && !pairRZ['30d'].insufficient ? pairRZ['30d'].expectedProfit : null;
        var ep60 = pairRZ['60d'] && !pairRZ['60d'].insufficient ? pairRZ['60d'].expectedProfit : null;
        var ep90 = pairRZ['90d'] && !pairRZ['90d'].insufficient ? pairRZ['90d'].expectedProfit : null;

        // Dollar-scale expected profit using proportionally allocated shares
        var effTotalShares = def.effLongShares + def.effShortShares;
        var ep30Dollar = ep30 != null ? ep30 * effTotalShares : null;
        var ep60Dollar = ep60 != null ? ep60 * effTotalShares : null;
        var ep90Dollar = ep90 != null ? ep90 * effTotalShares : null;

        // Bad scenario dollar amounts
        var bs = pairProb.badScenario;
        var badAvgDollar = bs ? bs.avgMae * effTotalShares : null;
        var badP75Dollar = bs ? bs.p75Mae * effTotalShares : null;

        // Gross exposure for this cross-pair (using latest live prices)
        var latestLong = histMap[lTk][histMap[lTk].length - 1];
        var latestShort = histMap[sTk][histMap[sTk].length - 1];
        var pairGrossLong = def.effLongShares * latestLong;
        var pairGrossShort = def.effShortShares * latestShort;
        totalGrossLong += pairGrossLong;
        totalGrossShort += pairGrossShort;

        // Track worst Z-score pair (highest |Z| = most extreme deviation)
        var absZ90 = z90 != null ? Math.abs(z90) : 0;
        if (worstZ === null || absZ90 > Math.abs(worstZ)) {
          worstZ = z90;
          worstZPair = lTk + ' (L) x ' + sTk + ' (S)';
        }

        // Accumulate weighted aggregates (only add weight to denominator when metric is valid)
        var w = def.weight;
        aggWeightSum += w;
        if (z30 != null) { aggZ30 += w * z30; zWeightSum30 += w; }
        if (z60 != null) { aggZ60 += w * z60; zWeightSum60 += w; }
        if (z90 != null) { aggZ90 += w * z90; zWeightSum90 += w; }
        if (wr30 != null) { aggWR30 += w * wr30; wrWeightSum30 += w; }
        if (wr60 != null) { aggWR60 += w * wr60; wrWeightSum60 += w; }
        if (wr90 != null) { aggWR90 += w * wr90; wrWeightSum90 += w; }
        if (ep30Dollar != null) aggEP30 += ep30Dollar;
        if (ep60Dollar != null) aggEP60 += ep60Dollar;
        if (ep90Dollar != null) aggEP90 += ep90Dollar;
        if (badAvgDollar != null) aggBadAvg += badAvgDollar;
        if (badP75Dollar != null) aggBadP75 += badP75Dollar;
        if (bs && bs.wideningProb != null) { aggWidenProb += w * bs.wideningProb; widenWeightSum += w; }
        if (bs && bs.stagnantRate != null) { aggStagnant += w * bs.stagnantRate; stagnantWeightSum += w; }

        // Build result object for this cross-pair
        crossPairResults.push({
          longTicker: lTk,
          shortTicker: sTk,
          weight: parseFloat((def.weight * 100).toFixed(2)),
          effLongShares: def.effLongShares,
          effShortShares: def.effShortShares,
          entrySpread: parseFloat(entrySpreadPerShare.toFixed(4)),
          mktSpread: parseFloat(pairMetrics.netSpread.toFixed(4)),
          z30: z30, z60: z60, z90: z90,
          wr30: wr30, wr60: wr60, wr90: wr90,
          ep30: ep30Dollar != null ? parseFloat(ep30Dollar.toFixed(2)) : null,
          ep60: ep60Dollar != null ? parseFloat(ep60Dollar.toFixed(2)) : null,
          ep90: ep90Dollar != null ? parseFloat(ep90Dollar.toFixed(2)) : null,
          badAvg: badAvgDollar != null ? parseFloat(badAvgDollar.toFixed(2)) : null,
          badP75: badP75Dollar != null ? parseFloat(badP75Dollar.toFixed(2)) : null,
          widenProb: bs ? bs.wideningProb : null,
          stagnantRate: bs ? bs.stagnantRate : null,
          triggers: pairProb.triggers || 0,
          historyDays: pairMetrics.historyDays,
          grossLong: parseFloat(pairGrossLong.toFixed(2)),
          grossShort: parseFloat(pairGrossShort.toFixed(2)),
          dailyValues: pairMetrics.dailyValues || [],
          rollingZ: pairRZ,
          probabilities: pairProb,
          maeRecovery: pairMetrics.maeRecovery
        });

        // Clean up internal arrays
        delete pairMetrics.dailyValuesFull_;
      }

      // ── Build aggregate metrics ──
      // Renormalize weighted averages (aggWeightSum may be < 1.0 if some pairs were skipped)
      var nw = aggWeightSum > 0 ? aggWeightSum : 1;
      // Count converging pairs (|Z| trending toward 0 over recent history)
      var convergingCount = 0;
      for (var cp = 0; cp < crossPairResults.length; cp++) {
        var cpr = crossPairResults[cp];
        if (cpr.error) continue;
        if (cpr.z30 != null && cpr.z90 != null && Math.abs(cpr.z30) < Math.abs(cpr.z90)) convergingCount++;
      }

      var aggregateMetrics = {
        crossPairCount: crossPairResults.length,
        totalLongLegs: longLegs.length,
        totalShortLegs: shortLegs.length,
        totalLongShares: totalLongShares,
        totalShortShares: totalShortShares,
        grossLong: parseFloat(totalGrossLong.toFixed(2)),
        grossShort: parseFloat(totalGrossShort.toFixed(2)),
        grossExposure: parseFloat((totalGrossLong + totalGrossShort).toFixed(2)),
        // Weighted average Z-scores (denominator = only pairs with valid Z data)
        z30: parseFloat((aggZ30 / (zWeightSum30 || 1)).toFixed(2)),
        z60: parseFloat((aggZ60 / (zWeightSum60 || 1)).toFixed(2)),
        z90: parseFloat((aggZ90 / (zWeightSum90 || 1)).toFixed(2)),
        // Weighted average win rates (denominator = only pairs with valid WR data)
        wr30: parseFloat((aggWR30 / (wrWeightSum30 || 1)).toFixed(1)),
        wr60: parseFloat((aggWR60 / (wrWeightSum60 || 1)).toFixed(1)),
        wr90: parseFloat((aggWR90 / (wrWeightSum90 || 1)).toFixed(1)),
        // Summed dollar amounts (no double counting via proportional allocation)
        ep30: parseFloat(aggEP30.toFixed(2)),
        ep60: parseFloat(aggEP60.toFixed(2)),
        ep90: parseFloat(aggEP90.toFixed(2)),
        // Summed bad scenario
        badAvg: parseFloat(aggBadAvg.toFixed(2)),
        badP75: parseFloat(aggBadP75.toFixed(2)),
        widenProb: parseFloat((aggWidenProb / (widenWeightSum || 1)).toFixed(1)),
        stagnantRate: parseFloat((aggStagnant / (stagnantWeightSum || 1)).toFixed(1)),
        // Diagnostics
        worstZ: worstZ,
        worstZPair: worstZPair,
        convergingCount: convergingCount,
        validPairs: crossPairResults.filter(function(c) { return !c.error; }).length,
        missingTickers: missing.length > 0 ? missing : undefined
      };

      return {
        mode: 'sandbox',
        crossMatrix: true,
        legs: legs.length,
        aggregateMetrics: aggregateMetrics,
        crossPairs: crossPairResults
      };
    }
    else {
      return { error: 'Unknown mode: ' + mode };
    }
  } catch(e) {
    return { error: e.message };
  }
}
/**
 * Legacy sandbox fallback: used when all legs are the same direction (no cross-pairs possible).
 * Runs the original single-basket analysis.
 */
function runLegacySandbox_(legs, userLegs, histMap) {
  var metrics = computeBasketMetrics_(legs, histMap);

  var hypotheticalSpreadDollar = 0, hypotheticalTotalSize = 0, hypotheticalGrossLong = 0, hypotheticalGrossShort = 0;
  for (var i = 0; i < userLegs.length; i++) {
    var ul = userLegs[i];
    var sz = Math.abs(parseFloat(ul.size) || 0);
    var dir = parseFloat(ul.direction) > 0 ? 1 : -1;
    var ep = parseFloat(ul.entryPrice) || 0;
    hypotheticalSpreadDollar += dir * sz * ep;
    hypotheticalTotalSize += sz;
    if (dir > 0) hypotheticalGrossLong += sz * ep; else hypotheticalGrossShort += sz * ep;
  }
  var hypotheticalWeight = hypotheticalTotalSize / 2;
  if (hypotheticalWeight === 0) hypotheticalWeight = 1;
  metrics.hypotheticalSpread = parseFloat((hypotheticalSpreadDollar / hypotheticalWeight).toFixed(2));
  metrics.hypotheticalGrossLong = parseFloat(hypotheticalGrossLong.toFixed(2));
  metrics.hypotheticalGrossShort = parseFloat(hypotheticalGrossShort.toFixed(2));
  metrics.hypotheticalGrossExposure = parseFloat((hypotheticalGrossLong + hypotheticalGrossShort).toFixed(2));

  var rz = metrics.rollingZ || {};
  for (var wKey in rz) {
    if (rz[wKey] && !rz[wKey].insufficient) {
      rz[wKey].expectedProfit = parseFloat((rz[wKey].mean - metrics.hypotheticalSpread).toFixed(4));
    }
  }

  if (metrics.dailyValuesFull_ && metrics.dailyValuesFull_.length > 0) {
    metrics.probabilities = computeHistoricalProbabilities_(
      metrics.dailyValuesFull_, metrics.hypotheticalSpread, rz, metrics.totalWeight
    );
    var stdTriggers = metrics.probabilities ? metrics.probabilities.triggers : 0;
    if (stdTriggers < 3) {
      var wideProb = computeHistoricalProbabilitiesWide_(
        metrics.dailyValuesFull_, metrics.hypotheticalSpread, rz, metrics.totalWeight
      );
      if (wideProb.triggers > stdTriggers) metrics.probabilities = wideProb;
    }
    var mktRef = metrics.dailyValuesFull_[metrics.dailyValuesFull_.length - 1];
    var mktProb = computeHistoricalProbabilities_(
      metrics.dailyValuesFull_, mktRef, rz, metrics.totalWeight
    );
    if (mktProb.triggers < 3) {
      var mktWideProb = computeHistoricalProbabilitiesWide_(
        metrics.dailyValuesFull_, mktRef, rz, metrics.totalWeight
      );
      if (mktWideProb.triggers > mktProb.triggers) mktProb = mktWideProb;
    }
    metrics.currentProbabilities = mktProb;
  }

  var missing = [];
  for (var i = 0; i < legs.length; i++) {
    if (!histMap[legs[i].ticker]) missing.push(legs[i].ticker);
  }
  if (missing.length > 0) metrics.missingTickers = missing;
  delete metrics.dailyValuesFull_;
  return { mode: 'sandbox', crossMatrix: false, legs: legs.length, metrics: metrics };
}

// ============================================================
// WRITE OPERATIONS
// ============================================================
// Helper: look up live prices and Z-score for a pair
function getLivePairData_(ss, pairId) {
  var cleanTarget = cleanId(pairId);
  var sheets = ['WebCache', 'WebCacheCredit', 'Live'];
  for (var s = 0; s < sheets.length; s++) {
    var ls = ss.getSheetByName(sheets[s]);
    if (!ls || ls.getLastRow() <= 1) continue;
    var data = ls.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (cleanId(data[i][0]) === cleanTarget) {
        return {
          priceA: parseFloat(data[i][3]) || 0,
          priceB: parseFloat(data[i][4]) || 0,
          spread: parseFloat(data[i][5]) || 0,
          mean: parseFloat(data[i][10]) || 0,
          z: parseFloat(data[i][12]) || 0
        };
      }
    }
  }
  return null;
}
function saveTradeToSheet(trade) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('OpenTrades');
  if (!sheet) throw new Error('OpenTrades sheet not found. Run setupAllBatched() first.');
  var curZ = 0;
  var cleanTradeId = cleanId(trade.id);
  var realId = trade.id;
  var foundPair = false;
  // Check WebCache/Live (intra) + WebCacheCredit (credit computed cache)
  var liveSheets = ['WebCache','WebCacheCredit','Live'];
  for (var s = 0; s < liveSheets.length; s++) {
    if (foundPair) break;
    var ls = ss.getSheetByName(liveSheets[s]);
    if (!ls || ls.getLastRow() <= 1) continue;
    var live = ls.getDataRange().getValues();
    for (var i = 1; i < live.length; i++) {
      if (cleanId(live[i][0]) === cleanTradeId) {
        curZ = parseFloat(live[i][12]) || 0; // M: Z-Score
        realId = live[i][0];
        foundPair = true;
        break;
      }
    }
  }
  // Check if position already exists — scale into it (weighted avg prices, sum sizes)
  var openData = sheet.getDataRange().getValues();
  var targetClean = cleanId(realId);
  for (var i = 1; i < openData.length; i++) {
    if (cleanId(openData[i][0]) === targetClean) {
      var oldPA = parseMoney(openData[i][2]);
      var oldPB = parseMoney(openData[i][3]);
      var oldSA = parseMoney(openData[i][4]);
      var oldSB = parseMoney(openData[i][5]);
      var newPA = parseMoney(trade.priceA);
      var newPB = parseMoney(trade.priceB);
      var newSA = parseMoney(trade.sizeA);
      var newSB = parseMoney(trade.sizeB);
      var totalSA = oldSA + newSA;
      var totalSB = oldSB + newSB;
      // Weighted average entry prices
      var avgPA = totalSA !== 0 ? ((oldPA * oldSA) + (newPA * newSA)) / totalSA : newPA;
      var avgPB = totalSB !== 0 ? ((oldPB * oldSB) + (newPB * newSB)) / totalSB : newPB;
      // Update existing row in-place (sheet rows are 1-indexed)
      sheet.getRange(i + 1, 3, 1, 4).setValues([[avgPA, avgPB, totalSA, totalSB]]);
      return true;
    }
  }
  // No existing position — create new row (cols: PairID, EntryZ, CostA, CostB, SizeA, SizeB, Timestamp, PaidDiv, ReceivedDiv)
  sheet.appendRow([realId, curZ, trade.priceA, trade.priceB, trade.sizeA, trade.sizeB, new Date(), 0, 0]);
  return true;
}
function closeTradeInSheet(id) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('OpenTrades');
  if (!sheet || sheet.getLastRow() <= 1) return true;
  var data = sheet.getDataRange().getValues();
  var cleanTarget = cleanId(id);
  for (var i = data.length - 1; i >= 1; i--) {
    var rowId = cleanId(data[i][0]);
    if (rowId === cleanTarget) {
      var pairId = data[i][0];
      var entryZ = data[i][1];
      var costA = parseMoney(data[i][2]);
      var costB = parseMoney(data[i][3]);
      var sA = parseMoney(data[i][4]);
      var sB = parseMoney(data[i][5]);
      var openDate = data[i][6];
      var paidDiv = parseMoney(data[i][7]);
      var rcvdDiv = parseMoney(data[i][8]);
      // Look up live exit prices
      var live = getLivePairData_(ss, pairId);
      var exitA = live ? live.priceA : costA;
      var exitB = live ? live.priceB : costB;
      var exitZ = live ? live.z : 0;
      if (!live) {
        Logger.log('WARNING: closeTradeInSheet — live data unavailable for ' + pairId + '. Using entry prices as exit prices (PnL will be $0 cap gains).');
      }
      var capGains = ((exitA - costA) * sA) + ((exitB - costB) * sB);
      var pnl = capGains + rcvdDiv - paidDiv;
      // Write to ClosedTrades
      var closed = ss.getSheetByName('ClosedTrades');
      if (!closed) {
        closed = ss.insertSheet('ClosedTrades');
        closed.getRange(1, 1, 1, 16).setValues([['PairID','EntryZ','CostA','CostB','SizeA','SizeB','OpenDate','CloseDate','PnL','ExitPriceA','ExitPriceB','ExitZ','CloseType','PaidDiv','ReceivedDiv','Notes']]);
      }
      closed.appendRow([pairId, entryZ, costA, costB, sA, sB, openDate, new Date(), pnl, exitA, exitB, exitZ, 'FULL', paidDiv, rcvdDiv]);
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return true;
}
function partialCloseTradeInSheet(id, reduceA, reduceB) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('OpenTrades');
  if (!sheet || sheet.getLastRow() <= 1) return true;
  var data = sheet.getDataRange().getValues();
  var cleanTarget = cleanId(id);
  reduceA = Math.abs(parseFloat(reduceA) || 0);
  reduceB = Math.abs(parseFloat(reduceB) || 0);
  if (reduceA === 0 && reduceB === 0) return true;
  for (var i = data.length - 1; i >= 1; i--) {
    var rowId = cleanId(data[i][0]);
    if (rowId === cleanTarget) {
      var pairId = data[i][0];
      var entryZ = data[i][1];
      var costA = parseMoney(data[i][2]);
      var costB = parseMoney(data[i][3]);
      var sA = parseMoney(data[i][4]);
      var sB = parseMoney(data[i][5]);
      var openDate = data[i][6];
      var totalPaidDiv = parseMoney(data[i][7]);
      var totalRcvdDiv = parseMoney(data[i][8]);
      // Clamp reduce amounts to position size
      var closedA = Math.min(reduceA, Math.abs(sA));
      var closedB = Math.min(reduceB, Math.abs(sB));
      // Preserve sign (long=+, short=-)
      var signA = sA >= 0 ? 1 : -1;
      var signB = sB >= 0 ? 1 : -1;
      // Proportionally split div amounts between closed and remaining portions (per-leg average)
      var ratioA = Math.abs(sA) > 0 ? closedA / Math.abs(sA) : 1;
      var ratioB = Math.abs(sB) > 0 ? closedB / Math.abs(sB) : 1;
      var divRatio = (ratioA + ratioB) / 2;
      var closedPaidDiv = Math.round(totalPaidDiv * divRatio * 100) / 100;
      var closedRcvdDiv = Math.round(totalRcvdDiv * divRatio * 100) / 100;
      // Look up live exit prices
      var live = getLivePairData_(ss, pairId);
      var exitA = live ? live.priceA : costA;
      var exitB = live ? live.priceB : costB;
      var exitZ = live ? live.z : 0;
      // PnL: capital gains on closed portion + proportional dividends
      // FIX: capGains uses signed sizes (closedA*signA) for direction, not closedA*signA again
      // For long (signA=+1): profit = (exit-cost)*shares. For short (signA=-1): profit = (cost-exit)*shares = (exit-cost)*(-shares)
      var capGains = ((exitA - costA) * (closedA * signA)) + ((exitB - costB) * (closedB * signB));
      var pnl = capGains + closedRcvdDiv - closedPaidDiv;
      // Write closed portion to ClosedTrades
      var closed = ss.getSheetByName('ClosedTrades');
      if (!closed) {
        closed = ss.insertSheet('ClosedTrades');
        closed.getRange(1, 1, 1, 16).setValues([['PairID','EntryZ','CostA','CostB','SizeA','SizeB','OpenDate','CloseDate','PnL','ExitPriceA','ExitPriceB','ExitZ','CloseType','PaidDiv','ReceivedDiv','Notes']]);
      }
      closed.appendRow([pairId, entryZ, costA, costB, closedA * signA, closedB * signB, openDate, new Date(), pnl, exitA, exitB, exitZ, 'PARTIAL', closedPaidDiv, closedRcvdDiv]);
      // Update remaining position — subtract proportional div amounts
      var remainA = sA - (closedA * signA);
      var remainB = sB - (closedB * signB);
      var remainPaidDiv = totalPaidDiv - closedPaidDiv;
      var remainRcvdDiv = totalRcvdDiv - closedRcvdDiv;
      if (Math.abs(remainA) < 0.01 && Math.abs(remainB) < 0.01) {
        sheet.deleteRow(i + 1);
      } else {
        sheet.getRange(i + 1, 5, 1, 2).setValues([[remainA, remainB]]);
        sheet.getRange(i + 1, 8, 1, 2).setValues([[remainPaidDiv, remainRcvdDiv]]);
      }
      break;
    }
  }
  return true;
}
// ============================================================
// DIVIDEND TRACKING
// ============================================================
// OpenTrades columns: A(0):PairID B(1):EntryZ C(2):CostA D(3):CostB E(4):SizeA F(5):SizeB G(6):Timestamp H(7):PaidDiv I(8):ReceivedDiv
function addDividendToTrade(id, type, amount) {
  if (!id || !type || !amount || amount <= 0) throw new Error('Invalid dividend input');
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('OpenTrades');
  if (!sheet || sheet.getLastRow() <= 1) throw new Error('No open trades found');
  var data = sheet.getDataRange().getValues();
  var cleanTarget = cleanId(id);
  // Determine column: PaidDiv=col H (index 7, sheet col 8), ReceivedDiv=col I (index 8, sheet col 9)
  var colIndex = (type === 'paid') ? 7 : 8;
  var sheetCol = colIndex + 1; // 1-indexed
  for (var i = 1; i < data.length; i++) {
    if (cleanId(data[i][0]) === cleanTarget) {
      var existing = parseMoney(data[i][colIndex]);
      sheet.getRange(i + 1, sheetCol).setValue(existing + amount);
      return true;
    }
  }
  throw new Error('Trade not found: ' + id);
}
// Saves a journal note to a specific ClosedTrades row.
// row = 1-indexed sheet row number (from getClosedTrades rowIdx field)
function saveTradeNote(row, note) {
  if (!row || row < 2) throw new Error('Invalid row number');
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('ClosedTrades');
  if (!sheet) throw new Error('ClosedTrades sheet not found');
  if (row > sheet.getLastRow()) throw new Error('Row does not exist');
  // Notes column = P (col 16)
  sheet.getRange(row, 16).setValue(note);
  return true;
}
// ============================================================
// WATCHLIST
// ============================================================
function saveToWatchlist(id, mode) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Watchlist');
  if (!sheet) {
    sheet = ss.insertSheet('Watchlist');
    sheet.getRange(1, 1, 1, 6).setValues([['PairID', 'AddedDate', 'AddedZ', 'Mode', 'AddedExpProfit', 'AddedSpread']]);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  // Ensure new columns exist on old sheets
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.length < 6) {
    sheet.getRange(1, 5).setValue('AddedExpProfit');
    sheet.getRange(1, 6).setValue('AddedSpread');
  }
  // Check for duplicate
  var cleanTarget = cleanId(id);
  if (sheet.getLastRow() > 1) {
    var existing = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < existing.length; i++) {
      if (cleanId(existing[i][0]) === cleanTarget) return; // already exists
    }
  }
  // Look up current Z-score, spread, and mean from live data
  var curZ = 0, curSpread = 0, curMean = 0;
  var foundWatch = false;
  var cacheSheets = ['WebCache', 'WebCacheCredit', 'Live'];
  for (var s = 0; s < cacheSheets.length; s++) {
    if (foundWatch) break;
    var ls = ss.getSheetByName(cacheSheets[s]);
    if (!ls || ls.getLastRow() <= 1) continue;
    var data = ls.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (cleanId(data[i][0]) === cleanTarget) {
        curZ = parseFloat(data[i][12]) || 0;
        curSpread = parseFloat(data[i][5]) || 0;
        curMean = parseFloat(data[i][10]) || 0;
        foundWatch = true;
        break;
      }
    }
  }
  var expProfit = Math.abs(curSpread - curMean);
  sheet.appendRow([id, new Date(), curZ, mode, expProfit, curSpread]);
}
function removeFromWatchlist(id) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Watchlist');
  if (!sheet || sheet.getLastRow() <= 1) return;
  var data = sheet.getDataRange().getValues();
  var cleanTarget = cleanId(id);
  for (var i = data.length - 1; i >= 1; i--) {
    if (cleanId(data[i][0]) === cleanTarget) { sheet.deleteRow(i + 1); break; }
  }
}
function getWatchlistData() {
  try {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName('Watchlist');
    if (!sheet || sheet.getLastRow() <= 1) return [];
    var colCount = Math.max(sheet.getLastColumn(), 6);
    var watchRows = sheet.getRange(2, 1, sheet.getLastRow() - 1, colCount).getValues();

    // Build unified live lookup from both caches
    var liveMap = {}; // cleanId → row array
    var cacheSheets = [
      { name: 'WebCache', fallback: 'Live' },
      { name: 'WebCacheCredit', fallback: null }
    ];
    for (var c = 0; c < cacheSheets.length; c++) {
      var ls = ss.getSheetByName(cacheSheets[c].name);
      if (!ls || ls.getLastRow() <= 1) ls = ss.getSheetByName(cacheSheets[c].fallback);
      if (!ls || ls.getLastRow() <= 1) continue;
      var rows = ls.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        var cid = cleanId(rows[i][0]);
        if (cid && !liveMap[cid]) liveMap[cid] = rows[i];
      }
    }

    // Load AlertsLog for 5-day Z history — pre-build map for O(1) lookup per pair
    var logSheet = ss.getSheetByName('AlertsLog');
    var logMap = {}; // cleanId → [{ts: Date, z: number}, ...]
    if (logSheet && logSheet.getLastRow() > 1) {
      var lastRow = logSheet.getLastRow();
      var startRow = Math.max(2, lastRow - 2000);
      var logData = logSheet.getRange(startRow, 1, lastRow - startRow + 1, 4).getValues();
      for (var lg = 0; lg < logData.length; lg++) {
        var lcid = cleanId(logData[lg][1]);
        if (!lcid) continue;
        if (!logMap[lcid]) logMap[lcid] = [];
        logMap[lcid].push({ ts: logData[lg][0], z: parseFloat(logData[lg][2]) || 0 });
      }
    }

    // Load DivDates
    var divMap = {};
    var divSheet = ss.getSheetByName('DivDates');
    if (divSheet && divSheet.getLastRow() > 1) {
      var divData = divSheet.getRange(2, 1, divSheet.getLastRow() - 1, 2).getValues();
      for (var d = 0; d < divData.length; d++) {
        var dticker = String(divData[d][0]).toUpperCase().trim();
        var ddate = divData[d][1];
        if (dticker && ddate instanceof Date) divMap[dticker] = ddate;
      }
    }

    var now = new Date();
    var output = [];
    for (var w = 0; w < watchRows.length; w++) {
      var pairId = watchRows[w][0];
      if (!pairId) continue;
      var addedDate = watchRows[w][1];
      var addedZ = parseFloat(watchRows[w][2]) || 0;
      var wMode = watchRows[w][3] || 'intra';
      var addedExpProfit = parseFloat(watchRows[w][4]) || 0;
      var addedSpread = parseFloat(watchRows[w][5]) || 0;
      var cid = cleanId(pairId);
      var row = liveMap[cid];

      if (!row) {
        // Pair not found in live data — show minimal card
        output.push({
          id: String(pairId), tA: '', tB: '', sec: '', mode: wMode,
          priceA: 0, priceB: 0, spr: '0.00', z: '0.00',
          addedZ: addedZ.toFixed(2), zChange: '0.00', mean: '0.00',
          distToMean: '0.00', yA: '0', yB: '0', yieldSpread: '0.00',
          liq: 0, volSpike: false, daysWatching: 0,
          addedDate: addedDate instanceof Date ? addedDate.toISOString().split('T')[0] : '',
          zHistory: [], converging: false, exDivA: null, exDivB: null,
          addedExpProfit: addedExpProfit.toFixed(2), curExpProfit: '0.00',
          addedSpread: addedSpread.toFixed(2),
          target: 'Data unavailable — pair may have been removed'
        });
        continue;
      }

      // Extract live data (same column indices as getAlertData)
      var tA = String(row[1]).trim();
      var tB = String(row[2]).trim();
      var priceA = parseFloat(row[3]) || 0;
      var priceB = parseFloat(row[4]) || 0;
      var spread = parseFloat(row[5]) || 0;
      var yieldA = parseFloat(row[6]) || 0;
      var yieldB = parseFloat(row[7]) || 0;
      var mean = parseFloat(row[10]) || 0;
      var stdev = parseFloat(row[11]) || 0;
      var currentZ = parseFloat(row[12]) || 0;
      var sector = row[15] || '';
      var avgLiq = parseFloat(row[19]) || 0;
      var volSpike = (row[23] === true || row[23] === 'TRUE');

      var yA = yieldA > 1 ? yieldA.toFixed(2) : (yieldA * 100).toFixed(2);
      var yB = yieldB > 1 ? yieldB.toFixed(2) : (yieldB * 100).toFixed(2);
      var yieldSpread = Math.abs(yieldA - yieldB);
      yieldSpread = yieldA > 1 ? yieldSpread.toFixed(2) : (yieldSpread * 100).toFixed(2);
      var distToMean = spread - mean;
      var zChange = currentZ - addedZ;
      var daysWatching = addedDate instanceof Date ? Math.floor((now.getTime() - addedDate.getTime()) / 86400000) : 0;

      // 5-day Z history from AlertsLog (pre-built map)
      var dailyZ = {}; // dateStr → last z
      var pairLog = logMap[cid] || [];
      for (var k = 0; k < pairLog.length; k++) {
        var ts = pairLog[k].ts;
        if (ts instanceof Date) {
          var dateStr = ts.toISOString().split('T')[0];
          dailyZ[dateStr] = pairLog[k].z;
        }
      }
      var sortedDays = Object.keys(dailyZ).sort();
      var last5 = sortedDays.slice(-5);
      var zHistory = [];
      var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
      for (var h = 0; h < last5.length; h++) {
        var parts = last5[h].split('-');
        var label = months[parseInt(parts[1], 10) - 1] + ' ' + parseInt(parts[2], 10);
        zHistory.push({ date: label, z: dailyZ[last5[h]].toFixed(2) });
      }
      var converging = false;
      if (zHistory.length >= 2) {
        var firstZ = Math.abs(parseFloat(zHistory[0].z));
        var lastZ = Math.abs(parseFloat(zHistory[zHistory.length - 1].z));
        converging = lastZ < firstZ;
      }

      // Ex-div dates
      var tickerAUp = tA.toUpperCase();
      var tickerBUp = tB.toUpperCase();
      var divA = divMap[tickerAUp] || null;
      var divB = divMap[tickerBUp] || null;

      // Target string
      var target = '';
      if (mean >= 0) {
        target = tA + ' should be $' + Math.abs(mean).toFixed(2) + ' higher than ' + tB;
      } else {
        target = tA + ' should be $' + Math.abs(mean).toFixed(2) + ' lower than ' + tB;
      }

      var curExpProfit = Math.abs(spread - mean);
      output.push({
        id: String(pairId),
        tA: tA, tB: tB, sec: sector, mode: wMode,
        priceA: priceA, priceB: priceB,
        spr: spread.toFixed(2),
        z: currentZ.toFixed(2),
        addedZ: addedZ.toFixed(2),
        zChange: zChange.toFixed(2),
        mean: mean.toFixed(2),
        distToMean: distToMean.toFixed(2),
        yA: yA, yB: yB,
        yieldSpread: yieldSpread,
        liq: avgLiq,
        volSpike: volSpike,
        daysWatching: daysWatching,
        addedDate: addedDate instanceof Date ? addedDate.toISOString().split('T')[0] : '',
        zHistory: zHistory,
        converging: converging,
        exDivA: (divA && divA >= now) ? divA.toISOString().split('T')[0] : null,
        exDivB: (divB && divB >= now) ? divB.toISOString().split('T')[0] : null,
        addedExpProfit: addedExpProfit.toFixed(2),
        curExpProfit: curExpProfit.toFixed(2),
        addedSpread: addedSpread.toFixed(2),
        target: target
      });
    }
    return output;
  } catch (e) {
    console.error('getWatchlistData error: ' + e);
    return [];
  }
}
// ============================================================
// MACRO VALUATION — Preferreds vs. US Treasuries
// ============================================================
function getMacroValuationData() {
  try {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName('MacroCache');
    if (!sheet || sheet.getLastRow() <= 1) {
      sheet = ss.getSheetByName('MacroCalc');
    }
    if (!sheet || sheet.getLastRow() <= 1) {
      return { status: 'no_data', data: [], sectorAvg: {} };
    }
    var data = sheet.getDataRange().getValues();
    var bmKeys = ['US2Y', 'US5Y', 'US7Y', 'US10Y', 'US30Y'];

    // First pass: collect data and credit rating z-scores
    var creditZMap = {};
    var results = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var ticker = String(row[0]).trim();
      if (!ticker) continue;

      var item = {
        ticker: ticker,
        coupon: parseFloat(row[1]) || 0,
        curPrice: parseFloat(row[2]) || 0,
        curYield: parseFloat(row[3]) || 0,
        sector: String(row[4] || ''),
        credit: String(row[5] || ''),
        benchmarks: {},
        avgZ: parseFloat(row[41]) || 0,
        signal: String(row[42] || 'FAIR'),
        bestBM: String(row[43] || 'N/A'),
        sectorAvgZ: 0
      };

      // Parse benchmark data (7 columns each, starting at index 6)
      for (var b = 0; b < bmKeys.length; b++) {
        var base = 6 + (b * 7);
        item.benchmarks[bmKeys[b]] = {
          spread: parseFloat(row[base]) || 0,
          mean: parseFloat(row[base + 1]) || 0,
          z: parseFloat(row[base + 2]) || 0,
          pct: parseFloat(row[base + 3]) || 0,
          hi90: parseFloat(row[base + 4]) || 0,
          lo90: parseFloat(row[base + 5]) || 0,
          dir: String(row[base + 6] || '→')
        };
      }

      // Track credit rating averages
      if (item.credit && item.avgZ !== 0) {
        if (!creditZMap[item.credit]) creditZMap[item.credit] = [];
        creditZMap[item.credit].push(item.avgZ);
      }

      results.push(item);
    }

    // Compute credit rating averages
    var creditAvg = {};
    for (var cr in creditZMap) {
      var arr = creditZMap[cr];
      var sum = 0;
      for (var j = 0; j < arr.length; j++) sum += arr[j];
      creditAvg[cr] = parseFloat((sum / arr.length).toFixed(2));
    }

    // Attach credit avg to each item
    for (var i = 0; i < results.length; i++) {
      results[i].creditAvgZ = creditAvg[results[i].credit] || 0;
    }

    // Read latest treasury yields from TreasuryHist for the banner
    var treasuryYields = {};
    try {
      var thSheet = ss.getSheetByName('TreasuryHist');
      if (thSheet && thSheet.getLastRow() > 1) {
        var thHeaders = thSheet.getRange(1, 1, 1, thSheet.getLastColumn()).getValues()[0];
        var lastRow = thSheet.getRange(thSheet.getLastRow(), 1, 1, thSheet.getLastColumn()).getValues()[0];
        for (var c = 1; c < thHeaders.length; c++) {
          var key = String(thHeaders[c]).trim();
          var val = parseFloat(lastRow[c]);
          if (key && !isNaN(val)) treasuryYields[key] = val;
        }
      }
    } catch (e) { /* treasury banner is non-critical */ }

    return { status: 'ok', data: results, creditAvg: creditAvg, treasuryYields: treasuryYields };
  } catch (e) {
    console.error('getMacroValuationData error: ' + e);
    return { status: 'error', data: [], creditAvg: {} };
  }
}

// ═══════════════════════════════════════════════════════════════════
// SINGLE-PAIR ANALYSIS — lightweight endpoint for inline alert analysis
// ═══════════════════════════════════════════════════════════════════
function analyzeSinglePair_(tA, tB, priceA, priceB, currentZ, customWindows) {
  try {
    var ss = SpreadsheetApp.getActive();
    var histMap = readTickerHistMap_(ss);

    // Debug diagnostics
    var _debug = {};
    _debug.histMapSize = Object.keys(histMap).length;
    var tkA = String(tA).toUpperCase().trim();
    var tkB = String(tB).toUpperCase().trim();
    _debug.tA = tkA;
    _debug.tB = tkB;
    _debug.histA = histMap[tkA] ? histMap[tkA].length : 0;
    _debug.histB = histMap[tkB] ? histMap[tkB].length : 0;
    _debug.pA = priceA;
    _debug.pB = priceB;
    _debug.z = currentZ;

    // If no history at all, return diagnostic error
    if (_debug.histMapSize === 0) {
      return { error: 'TickerHistory sheet is empty or not found. Run setupAllBatched() first.', _debug: _debug };
    }
    if (!histMap[tkA] && !histMap[tkB]) {
      // Show first 10 tickers in histMap so we can diagnose format mismatch
      _debug.sampleTickers = Object.keys(histMap).slice(0, 10);
      return { error: 'Neither ticker found in TickerHistory. Tickers: ' + tkA + ', ' + tkB, _debug: _debug };
    }

    // Build two legs: if Z > 0 spread is above mean → short A, long B
    var dirA = currentZ > 0 ? -1 : 1;
    var dirB = currentZ > 0 ? 1 : -1;
    var legs = [
      { ticker: tkA, size: 100, direction: dirA },
      { ticker: tkB, size: 100, direction: dirB }
    ];
    var metrics = computeBasketMetrics_(legs, histMap, customWindows);
    if (metrics.error) {
      _debug.metricsError = metrics.error;
      return { error: metrics.error, _debug: _debug };
    }

    _debug.validLegs = metrics.validLegs;
    _debug.historyDays = metrics.historyDays;
    _debug.dailyValuesLen = metrics.dailyValuesFull_ ? metrics.dailyValuesFull_.length : 0;
    _debug.triggers = metrics.probabilities ? metrics.probabilities.triggers : 0;

    // computeBasketMetrics_ already ran computeHistoricalProbabilities_ with
    // currentValue (the last daily spread value). If it found enough triggers, great.
    // If not, try the wide-tolerance version which relaxes constraints.
    var dailyVals = metrics.dailyValuesFull_;
    if (dailyVals && dailyVals.length >= 20) {
      var currentValue = dailyVals[dailyVals.length - 1];
      _debug.currentValue = parseFloat(currentValue.toFixed(4));

      var wideProb = computeHistoricalProbabilitiesWide_(dailyVals, currentValue, metrics.rollingZ, metrics.totalWeight || 100, customWindows);
      _debug.wideTriggersFound = wideProb.triggers;

      // Use whichever found more triggers
      if (wideProb.triggers > (metrics.probabilities ? metrics.probabilities.triggers : 0)) {
        metrics.probabilities = wideProb;
      }
    }

    // Strip internal fields
    delete metrics.dailyValuesFull_;
    return {
      tA: tA, tB: tB, dirA: dirA, dirB: dirB,
      metrics: metrics,
      windows: customWindows || [30, 60, 90],
      _debug: _debug
    };
  } catch (e) {
    return { error: e.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════
// NIGHTLY SCREENER — pre-computes probability analysis for top alerts
// ═══════════════════════════════════════════════════════════════════
function runNightlyScreener() {
  var startTime = new Date().getTime();
  var MAX_MS = 300000; // 5 min safety (GAS limit = 6 min)
  var MAX_PAIRS = 20;
  try {
    var ss = SpreadsheetApp.getActive();
    var histMap = readTickerHistMap_(ss);
    // Get current alerts from both modes
    var allAlerts = [];
    try { var intra = getAlertData('intra'); if (intra && intra.length) allAlerts = allAlerts.concat(intra.map(function(a){ a._mode='intra'; return a; })); } catch(e){}
    try { var credit = getAlertData('credit'); if (credit && credit.length) allAlerts = allAlerts.concat(credit.map(function(a){ a._mode='credit'; return a; })); } catch(e){}
    if (allAlerts.length === 0) { Logger.log('Screener: no alerts to process'); return; }
    // Sort by |Z| descending, take top N
    allAlerts.sort(function(a,b){ return Math.abs(parseFloat(b.z)||0) - Math.abs(parseFloat(a.z)||0); });
    var top = allAlerts.slice(0, MAX_PAIRS);
    var results = [];
    for (var i = 0; i < top.length; i++) {
      if (new Date().getTime() - startTime > MAX_MS) { Logger.log('Screener: timeout after ' + i + ' pairs'); break; }
      var a = top[i];
      try {
        var analysis = analyzeSinglePair_(a.tA, a.tB, parseFloat(a.pA) || 0, parseFloat(a.pB) || 0, parseFloat(a.z), [15, 30, 60, 90]);
        if (analysis && !analysis.error && analysis.metrics) {
          var prob = analysis.metrics.probabilities || {};
          var wr15 = (prob.winRates && prob.winRates['15d']) ? prob.winRates['15d'].rate : null;
          var wr30 = (prob.winRates && prob.winRates['30d']) ? prob.winRates['30d'].rate : null;
          var wr60 = (prob.winRates && prob.winRates['60d']) ? prob.winRates['60d'].rate : null;
          var wr90 = (prob.winRates && prob.winRates['90d']) ? prob.winRates['90d'].rate : null;
          var bs = prob.badScenario || {};
          results.push({
            id: a.id, tA: a.tA, tB: a.tB, mode: a._mode,
            z: parseFloat(a.z), expProfit: parseFloat(a.expProfit),
            wr15: wr15, wr30: wr30, wr60: wr60, wr90: wr90,
            avgMae: (bs.avgMae != null) ? bs.avgMae : null, p75Mae: (bs.p75Mae != null) ? bs.p75Mae : null,
            wideningProb: (bs.wideningProb != null) ? bs.wideningProb : null,
            stagnantRate: (bs.stagnantRate != null) ? bs.stagnantRate : null, triggers: prob.triggers || 0,
            ep15: (analysis.metrics.rollingZ && analysis.metrics.rollingZ['15d']) ? analysis.metrics.rollingZ['15d'].expectedProfit : null,
            ep30: (analysis.metrics.rollingZ && analysis.metrics.rollingZ['30d']) ? analysis.metrics.rollingZ['30d'].expectedProfit : null,
            ep60: (analysis.metrics.rollingZ && analysis.metrics.rollingZ['60d']) ? analysis.metrics.rollingZ['60d'].expectedProfit : null,
            ep90: (analysis.metrics.rollingZ && analysis.metrics.rollingZ['90d']) ? analysis.metrics.rollingZ['90d'].expectedProfit : null,
            ts: new Date().toISOString()
          });
        }
      } catch(e) { Logger.log('Screener: failed ' + a.id + ': ' + e); }
    }
    // Write to ScreenerCache sheet
    var sheet = ss.getSheetByName('ScreenerCache');
    if (!sheet) {
      sheet = ss.insertSheet('ScreenerCache');
      sheet.getRange(1,1,1,20).setValues([['PairID','TickerA','TickerB','Mode','Z','ExpProfit','WR30','WR60','WR90','AvgMAE','P75MAE','WidenProb','Triggers','EP30','EP60','UpdatedAt','WR15','EP15','EP90','StagnantRate']]);
    } else {
      if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 20).clearContent();
    }
    if (results.length > 0) {
      var rows = results.map(function(r){
        return [r.id, r.tA, r.tB, r.mode, r.z, r.expProfit, r.wr30, r.wr60, r.wr90, r.avgMae, r.p75Mae, r.wideningProb, r.triggers, r.ep30, r.ep60, r.ts, r.wr15, r.ep15, r.ep90, r.stagnantRate];
      });
      sheet.getRange(2, 1, rows.length, 20).setValues(rows);
    }
    Logger.log('Screener: processed ' + results.length + '/' + top.length + ' pairs in ' + ((new Date().getTime()-startTime)/1000).toFixed(1) + 's');
  } catch(e) {
    Logger.log('Screener error: ' + e);
  }
}

// ═══════════════════════════════════════════════════════════════════
// JOURNAL ANALYTICS — aggregated stats from closed trades
// ═══════════════════════════════════════════════════════════════════
function getJournalAnalytics_() {
  try {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName('ClosedTrades');
    if (!sheet || sheet.getLastRow() <= 1) return { trades: 0 };
    var rows = sheet.getDataRange().getValues();

    var wins = 0, losses = 0, totalPnl = 0, totalReturn = 0;
    var holdDays = [], pnls = [], returns = [];
    var byMonth = {}, bySector = {}, byStrategy = {};
    var streak = 0, maxWinStreak = 0, maxLossStreak = 0, curStreakType = null;
    var biggestWin = 0, biggestLoss = 0;

    // Pre-read credit pair IDs for strategy detection
    var creditIds = {};
    var creditSheet = ss.getSheetByName('WebCacheCredit') || ss.getSheetByName('CreditLive');
    if (creditSheet && creditSheet.getLastRow() > 1) {
      var cpCol = creditSheet.getRange(2, 1, creditSheet.getLastRow()-1, 1).getValues();
      for (var ci = 0; ci < cpCol.length; ci++) {
        if (cpCol[ci][0]) creditIds[cleanId(cpCol[ci][0])] = true;
      }
    }

    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0]) continue;
      var pnl = parseFloat(r[8]) || 0;
      var costA = parseFloat(r[2]) || 0;
      var costB = parseFloat(r[3]) || 0;
      var sizeA = parseFloat(r[4]) || 0;
      var sizeB = parseFloat(r[5]) || 0;
      var entryCost = Math.abs(costA * sizeA) + Math.abs(costB * sizeB);
      var retPct = entryCost > 0 ? (pnl / entryCost * 100) : 0;
      var openDate = r[6] instanceof Date ? r[6] : null;
      var closeDate = r[7] instanceof Date ? r[7] : null;
      var hold = (openDate && closeDate) ? Math.floor((closeDate.getTime() - openDate.getTime()) / 86400000) : 0;
      var paidDiv = parseFloat(r[13]) || 0;
      var rcvdDiv = parseFloat(r[14]) || 0;
      var entryZ = parseFloat(r[1]) || 0;
      var exitZ = parseFloat(r[11]) || 0;
      var sector = '';
      var cid = cleanId(r[0]);
      var strategy = creditIds[cid] ? 'credit' : 'intra';

      // Try to get sector from Live/WebCache
      // (lightweight — just use the pair ID to determine strategy)

      totalPnl += pnl;
      totalReturn += retPct;
      pnls.push(pnl);
      returns.push(retPct);
      holdDays.push(hold);

      if (pnl > 0) {
        wins++;
        if (pnl > biggestWin) biggestWin = pnl;
        if (curStreakType === 'win') { streak++; }
        else { streak = 1; curStreakType = 'win'; }
        if (streak > maxWinStreak) maxWinStreak = streak;
      } else {
        losses++;
        if (pnl < biggestLoss) biggestLoss = pnl;
        if (curStreakType === 'loss') { streak++; }
        else { streak = 1; curStreakType = 'loss'; }
        if (streak > maxLossStreak) maxLossStreak = streak;
      }

      // Monthly breakdown
      if (closeDate) {
        var mk = closeDate.getFullYear() + '-' + ('0' + (closeDate.getMonth()+1)).slice(-2);
        if (!byMonth[mk]) byMonth[mk] = { pnl: 0, trades: 0, wins: 0 };
        byMonth[mk].pnl += pnl;
        byMonth[mk].trades++;
        if (pnl > 0) byMonth[mk].wins++;
      }

      // Strategy breakdown
      if (!byStrategy[strategy]) byStrategy[strategy] = { pnl: 0, trades: 0, wins: 0 };
      byStrategy[strategy].pnl += pnl;
      byStrategy[strategy].trades++;
      if (pnl > 0) byStrategy[strategy].wins++;
    }

    var totalTrades = wins + losses;
    var winRate = totalTrades > 0 ? (wins / totalTrades * 100) : 0;
    var avgPnl = totalTrades > 0 ? totalPnl / totalTrades : 0;
    var avgReturn = totalTrades > 0 ? totalReturn / totalTrades : 0;
    var avgHold = holdDays.length > 0 ? holdDays.reduce(function(a,b){return a+b;},0) / holdDays.length : 0;
    var avgWin = wins > 0 ? pnls.filter(function(p){return p>0;}).reduce(function(a,b){return a+b;},0) / wins : 0;
    var avgLoss = losses > 0 ? pnls.filter(function(p){return p<=0;}).reduce(function(a,b){return a+b;},0) / losses : 0;
    var grossProfits = pnls.filter(function(p){return p>0;}).reduce(function(a,b){return a+b;},0);
    var grossLosses = Math.abs(pnls.filter(function(p){return p<=0;}).reduce(function(a,b){return a+b;},0));
    var profitFactor = grossLosses > 0 ? parseFloat((grossProfits / grossLosses).toFixed(2)) : (grossProfits > 0 ? 99.99 : 0);

    // Expectancy = (WinRate × AvgWin) - (LossRate × |AvgLoss|)
    var expectancy = (winRate/100 * avgWin) - ((1 - winRate/100) * Math.abs(avgLoss));

    // Sortino-like: downside deviation
    var negReturns = returns.filter(function(r){return r < 0;});
    var downsideDev = 0;
    if (negReturns.length > 0) {
      var sumSq = negReturns.reduce(function(a,r){return a + r*r;}, 0);
      downsideDev = Math.sqrt(sumSq / negReturns.length);
    }
    var sortinoLike = downsideDev > 0 ? (avgReturn / downsideDev) : 0;

    // Monthly array sorted
    var monthlyArr = Object.keys(byMonth).sort().map(function(k) {
      return { month: k, pnl: parseFloat(byMonth[k].pnl.toFixed(2)), trades: byMonth[k].trades, winRate: byMonth[k].trades > 0 ? parseFloat((byMonth[k].wins/byMonth[k].trades*100).toFixed(1)) : 0 };
    });

    return {
      trades: totalTrades,
      wins: wins,
      losses: losses,
      winRate: parseFloat(winRate.toFixed(1)),
      totalPnl: parseFloat(totalPnl.toFixed(2)),
      avgPnl: parseFloat(avgPnl.toFixed(2)),
      avgReturn: parseFloat(avgReturn.toFixed(2)),
      avgHold: parseFloat(avgHold.toFixed(1)),
      avgWin: parseFloat(avgWin.toFixed(2)),
      avgLoss: parseFloat(avgLoss.toFixed(2)),
      biggestWin: parseFloat(biggestWin.toFixed(2)),
      biggestLoss: parseFloat(biggestLoss.toFixed(2)),
      profitFactor: parseFloat(profitFactor.toFixed(2)),
      expectancy: parseFloat(expectancy.toFixed(2)),
      sortinoLike: parseFloat(sortinoLike.toFixed(2)),
      maxWinStreak: maxWinStreak,
      maxLossStreak: maxLossStreak,
      monthly: monthlyArr,
      byStrategy: byStrategy
    };
  } catch(e) {
    console.error('getJournalAnalytics_ error: ' + e);
    return { trades: 0, error: e.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════
// DIVIDEND CAPTURE — pairs where div timing + Z-score align
// ═══════════════════════════════════════════════════════════════════
function getDividendCapture_() {
  try {
    var ss = SpreadsheetApp.getActive();
    var now = new Date();

    // Load dividend dates
    var divMap = {};
    var divSheet = ss.getSheetByName('DivDates');
    if (divSheet && divSheet.getLastRow() > 1) {
      var divData = divSheet.getRange(2, 1, divSheet.getLastRow() - 1, 3).getValues();
      for (var d = 0; d < divData.length; d++) {
        var tk = String(divData[d][0]).toUpperCase().trim();
        var dt = divData[d][1];
        if (tk && dt instanceof Date) divMap[tk] = dt;
      }
    }

    // Load all pairs from both WebCache + WebCacheCredit
    var allPairs = [];
    var sheets = [
      { name: 'WebCache', fallback: 'Live', mode: 'intra' },
      { name: 'WebCacheCredit', fallback: null, mode: 'credit' }
    ];
    for (var s = 0; s < sheets.length; s++) {
      var sh = ss.getSheetByName(sheets[s].name);
      if ((!sh || sh.getLastRow() <= 1) && sheets[s].fallback) {
        sh = ss.getSheetByName(sheets[s].fallback);
      }
      if (!sh || sh.getLastRow() <= 1) continue;
      var data = sh.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        if (!row[0]) continue;
        var tA = String(row[1]).toUpperCase().trim();
        var tB = String(row[2]).toUpperCase().trim();
        var priceA = parseFloat(row[3]) || 0;
        var priceB = parseFloat(row[4]) || 0;
        if (priceA <= 0 || priceB <= 0) continue;
        var mean = parseFloat(row[10]) || 0;
        var stdev = parseFloat(row[11]) || 0;
        if (stdev <= 0.001) continue;
        var z = parseFloat(row[12]) || 0;
        var yieldA = parseFloat(row[6]) || 0;
        var yieldB = parseFloat(row[7]) || 0;

        // Z-score must be |Z| >= 1.8
        if (Math.abs(z) < 1.8) continue;

        // Spread capture = distance from current spread to mean
        var spread = priceA - priceB;
        var spreadCapture = Math.abs(spread - mean);

        // Must have at least $0.30 spread capture potential
        if (spreadCapture < 0.30) continue;

        // Determine long/short legs from Z-score direction
        // Z > 0 → spread above mean → short A / long B
        // Z < 0 → spread below mean → long A / short B
        var longLeg = z < 0 ? 'A' : 'B';
        var shortLeg = z < 0 ? 'B' : 'A';

        // Get div dates for both legs
        var divA = divMap[tA] || null;
        var divB = divMap[tB] || null;
        var daysA = divA ? Math.floor((divA.getTime() - now.getTime()) / 86400000) : 999;
        var daysB = divB ? Math.floor((divB.getTime() - now.getTime()) / 86400000) : 999;

        var longDays = longLeg === 'A' ? daysA : daysB;
        var shortDays = shortLeg === 'A' ? daysA : daysB;

        // Long side must have div coming up in 3-20 days
        if (longDays < 3 || longDays > 20) continue;

        // Short side div must be further away than long side (or no div at all)
        if (shortDays <= longDays) continue;

        var info = parseTickerInfo(row[0]);
        allPairs.push({
          id: info.id, tA: info.tA, tB: info.tB,
          mode: sheets[s].mode,
          z: parseFloat(z.toFixed(2)),
          pA: parseFloat(priceA.toFixed(2)),
          pB: parseFloat(priceB.toFixed(2)),
          yA: yieldA > 1 ? parseFloat(yieldA.toFixed(2)) : parseFloat((yieldA*100).toFixed(2)),
          yB: yieldB > 1 ? parseFloat(yieldB.toFixed(2)) : parseFloat((yieldB*100).toFixed(2)),
          longLeg: longLeg,
          longDivDays: longDays,
          shortDivDays: shortDays,
          spreadCapture: parseFloat(spreadCapture.toFixed(2)),
          divDateA: divA ? divA.toISOString().split('T')[0] : null,
          divDateB: divB ? divB.toISOString().split('T')[0] : null,
          sec: row[15] || ''
        });
      }
    }

    // Sort by spread capture descending (best opportunities first)
    allPairs.sort(function(a, b) {
      return b.spreadCapture - a.spreadCapture;
    });

    // Summary stats
    var avgDays = allPairs.length > 0 ? allPairs.reduce(function(s,p){return s+p.longDivDays;},0) / allPairs.length : 0;
    var avgCapture = allPairs.length > 0 ? allPairs.reduce(function(s,p){return s+p.spreadCapture;},0) / allPairs.length : 0;

    return {
      pairs: allPairs,
      summary: {
        total: allPairs.length,
        avgDaysToDiv: parseFloat(avgDays.toFixed(1)),
        avgSpreadCapture: parseFloat(avgCapture.toFixed(2))
      }
    };
  } catch(e) {
    console.error('getDividendCapture_ error: ' + e);
    return { pairs: [], summary: {}, error: e.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════
// BACKTEST ENGINE — historical "what if" simulation
// ═══════════════════════════════════════════════════════════════════
function runBacktest_(zThreshold, exitZ, maxHold, mode) {
  try {
    var ss = SpreadsheetApp.getActive();
    var histMap = readTickerHistMap_(ss);
    if (Object.keys(histMap).length === 0) return { error: 'No history data' };

    // Load all pair definitions
    var pairDefs = [];
    var pairSources = [];
    if (mode === 'all' || mode === 'intra') pairSources.push({ sheet: 'Pairs', mode: 'intra' });
    if (mode === 'all' || mode === 'credit') pairSources.push({ sheet: 'CreditPairs', mode: 'credit' });

    for (var s = 0; s < pairSources.length; s++) {
      var pSheet = ss.getSheetByName(pairSources[s].sheet);
      if (!pSheet || pSheet.getLastRow() <= 1) continue;
      var pData = pSheet.getDataRange().getValues();
      for (var p = 1; p < pData.length; p++) {
        if (!pData[p][0]) continue;
        var info = parseTickerInfo(pData[p][0]);
        pairDefs.push({ id: info.id, tA: info.tA, tB: info.tB, sector: pData[p][3] || '', mode: pairSources[s].mode });
      }
    }

    // Also load Levels/CreditLevels for mean/stdev
    var statsMap = {}; // cleanId → {mean, stdev}
    var lvlSheets = ['Levels', 'CreditLevels'];
    for (var ls = 0; ls < lvlSheets.length; ls++) {
      var lvl = ss.getSheetByName(lvlSheets[ls]);
      if (!lvl || lvl.getLastRow() <= 1) continue;
      var lvlData = lvl.getDataRange().getValues();
      // Levels columns: A=PairID, B=Mean, C=StDev, D=Lower, E=Upper, F=HistCount
      for (var li = 1; li < lvlData.length; li++) {
        if (!lvlData[li][0]) continue;
        var cid = cleanId(lvlData[li][0]);
        var mn = parseFloat(lvlData[li][1]) || 0;
        var sd = parseFloat(lvlData[li][2]) || 0;
        if (sd > 0.001) statsMap[cid] = { mean: mn, stdev: sd };
      }
    }

    var allTrades = [];
    var startTime = new Date().getTime();
    var MAX_MS = 240000; // 4 min safety — must finish before GAS 6-min hard limit
    var pairsProcessed = 0;
    var pairsSkippedTime = 0;
    var pairsSkippedHist = 0;

    var pairsWithTrades = 0;
    for (var pi = 0; pi < pairDefs.length; pi++) {
      if (new Date().getTime() - startTime > MAX_MS) { pairsSkippedTime += (pairDefs.length - pi); break; }
      var pair = pairDefs[pi];
      var tA = pair.tA.toUpperCase().trim();
      var tB = pair.tB.toUpperCase().trim();

      // FILTER: blacklisted tickers — excluded from all backtest modes
      if (isBlacklisted(tA) || isBlacklisted(tB)) { pairsSkippedHist++; continue; }

      // FILTER: intra-only tickers — excluded from credit backtest
      if (pair.mode === 'credit' && (isIntraOnly(tA) || isIntraOnly(tB))) { pairsSkippedHist++; continue; }

      var histA = histMap[tA];
      var histB = histMap[tB];
      if (!histA || !histB) { pairsSkippedHist++; continue; }

      // Align from end (most recent prices align)
      var len = Math.min(histA.length, histB.length);
      if (len < 30) { pairsSkippedHist++; continue; } // need minimum history
      var pricesA = histA.slice(histA.length - len);
      var pricesB = histB.slice(histB.length - len);

      // Compute rolling 30-day mean and stdev for Z-scores
      var WINDOW = 30;
      if (len < WINDOW + 5) { pairsSkippedHist++; continue; }

      pairsProcessed++;
      var tradesBefore = allTrades.length;
      var openTrade = null;
      for (var day = WINDOW; day < len; day++) {
        // Rolling window stats
        var sumSpr = 0, sumSprSq = 0;
        for (var w = day - WINDOW; w < day; w++) {
          var spr = pricesA[w] - pricesB[w];
          sumSpr += spr;
          sumSprSq += spr * spr;
        }
        var rollMean = sumSpr / WINDOW;
        var rollVar = (sumSprSq / WINDOW) - (rollMean * rollMean);
        var rollStdev = rollVar > 0 ? Math.sqrt(rollVar * WINDOW / (WINDOW - 1)) : 0.001;
        var spread = pricesA[day] - pricesB[day];
        var zScore = (spread - rollMean) / rollStdev;

        if (!openTrade) {
          // Entry signal
          if (Math.abs(zScore) >= zThreshold) {
            openTrade = {
              entryDay: day,
              entryZ: zScore,
              entrySpread: spread,
              entryPriceA: pricesA[day],
              entryPriceB: pricesB[day],
              dirA: zScore > 0 ? -1 : 1, // Z>0 → short spread (short A, long B)
              dirB: zScore > 0 ? 1 : -1
            };
          }
        } else {
          // Exit conditions: Z crosses back inside exitZ band, or max hold reached
          var holdDays = day - openTrade.entryDay;
          var exitNow = Math.abs(zScore) <= exitZ || holdDays >= maxHold;
          if (exitNow) {
            var exitSpread = spread;
            var pnlA = (pricesA[day] - openTrade.entryPriceA) * openTrade.dirA;
            var pnlB = (pricesB[day] - openTrade.entryPriceB) * openTrade.dirB;
            var tradePnl = (pnlA + pnlB) * 100; // per 100 shares
            allTrades.push({
              id: pair.id, tA: pair.tA, tB: pair.tB,
              mode: pair.mode, sector: pair.sector,
              entryDay: openTrade.entryDay, exitDay: day,
              holdDays: holdDays,
              entryZ: parseFloat(openTrade.entryZ.toFixed(2)),
              exitZ: parseFloat(zScore.toFixed(2)),
              entrySpread: parseFloat(openTrade.entrySpread.toFixed(4)),
              exitSpread: parseFloat(exitSpread.toFixed(4)),
              pnl: parseFloat(tradePnl.toFixed(2)),
              exitReason: holdDays >= maxHold ? 'MAX_HOLD' : 'Z_REVERT'
            });
            openTrade = null;
          }
        }
      }
      if (allTrades.length > tradesBefore) pairsWithTrades++;
    }

    // Aggregate stats
    var elapsedMs = new Date().getTime() - startTime;
    var totalTrades = allTrades.length;
    var intraPairCount = 0, creditPairCount = 0;
    for (var mi = 0; mi < pairDefs.length; mi++) { if (pairDefs[mi].mode === 'intra') intraPairCount++; else creditPairCount++; }
    var _meta = { pairsDefined: pairDefs.length, intraPairs: intraPairCount, creditPairs: creditPairCount, pairsProcessed: pairsProcessed, pairsWithTrades: pairsWithTrades, pairsSkippedTime: pairsSkippedTime, pairsSkippedHist: pairsSkippedHist, elapsedMs: elapsedMs, modeReceived: mode, histMapSize: Object.keys(histMap).length };
    if (totalTrades === 0) return { trades: 0, params: { zThreshold: zThreshold, exitZ: exitZ, maxHold: maxHold, mode: mode }, _meta: _meta };

    var winTrades = allTrades.filter(function(t){return t.pnl > 0;});
    var lossTrades = allTrades.filter(function(t){return t.pnl <= 0;});
    var totalPnl = allTrades.reduce(function(s,t){return s+t.pnl;},0);
    var avgPnl = totalPnl / totalTrades;
    var avgHold = allTrades.reduce(function(s,t){return s+t.holdDays;},0) / totalTrades;
    var avgWin = winTrades.length > 0 ? winTrades.reduce(function(s,t){return s+t.pnl;},0) / winTrades.length : 0;
    var avgLoss = lossTrades.length > 0 ? lossTrades.reduce(function(s,t){return s+t.pnl;},0) / lossTrades.length : 0;
    var maxDrawdown = 0, peak = 0, cumPnl = 0;
    var equityCurve = [];
    for (var ec = 0; ec < allTrades.length; ec++) {
      cumPnl += allTrades[ec].pnl;
      equityCurve.push(parseFloat(cumPnl.toFixed(2)));
      if (cumPnl > peak) peak = cumPnl;
      var dd = peak - cumPnl;
      if (dd > maxDrawdown) maxDrawdown = dd;
    }

    var zRevertExits = allTrades.filter(function(t){return t.exitReason==='Z_REVERT';}).length;

    // By-pair breakdown (top 15 by trade count)
    var pairStats = {};
    for (var bt = 0; bt < allTrades.length; bt++) {
      var pid = allTrades[bt].id;
      if (!pairStats[pid]) pairStats[pid] = { trades: 0, pnl: 0, wins: 0, losses: 0, winPnl: 0, lossPnl: 0, holdSum: 0, entryZSum: 0, exitZSum: 0, entrySprSum: 0, exitSprSum: 0, tA: allTrades[bt].tA, tB: allTrades[bt].tB, mode: allTrades[bt].mode, allPairTrades: [] };
      pairStats[pid].trades++;
      pairStats[pid].pnl += allTrades[bt].pnl;
      pairStats[pid].holdSum += allTrades[bt].holdDays;
      pairStats[pid].entryZSum += Math.abs(allTrades[bt].entryZ);
      pairStats[pid].exitZSum += Math.abs(allTrades[bt].exitZ);
      pairStats[pid].entrySprSum += allTrades[bt].entrySpread;
      pairStats[pid].exitSprSum += allTrades[bt].exitSpread;
      if (allTrades[bt].pnl > 0) { pairStats[pid].wins++; pairStats[pid].winPnl += allTrades[bt].pnl; }
      else { pairStats[pid].losses++; pairStats[pid].lossPnl += Math.abs(allTrades[bt].pnl); }
      pairStats[pid].allPairTrades.push(allTrades[bt]);
    }
    var allPairsSorted = Object.keys(pairStats).map(function(k) {
      var ps = pairStats[k];
      var recent = ps.allPairTrades.slice(-5).map(function(t){ return { entryZ: t.entryZ, exitZ: t.exitZ, holdDays: t.holdDays, pnl: t.pnl, exitReason: t.exitReason, entrySpread: t.entrySpread, exitSpread: t.exitSpread }; });
      var wr = parseFloat((ps.wins/ps.trades*100).toFixed(1));
      var aWin = ps.wins > 0 ? parseFloat((ps.winPnl/ps.wins).toFixed(2)) : 0;
      var aLoss = ps.losses > 0 ? parseFloat((ps.lossPnl/ps.losses).toFixed(2)) : 0;
      var pf = aLoss > 0 ? parseFloat((ps.winPnl / ps.lossPnl).toFixed(2)) : 0;
      var zRevertCount = ps.allPairTrades.filter(function(t){return t.exitReason==='Z_REVERT';}).length;
      var rob = computeRobustnessScore_(ps.trades, wr, aWin, aLoss, { profitFactor: pf });
      return { id: k, tA: ps.tA, tB: ps.tB, mode: ps.mode, trades: ps.trades, pnl: parseFloat(ps.pnl.toFixed(2)), winRate: wr, avgPnl: parseFloat((ps.pnl/ps.trades).toFixed(2)), avgWin: aWin, avgLoss: aLoss, profitFactor: pf, avgHold: parseFloat((ps.holdSum/ps.trades).toFixed(1)), avgEntryZ: parseFloat((ps.entryZSum/ps.trades).toFixed(2)), avgExitZ: parseFloat((ps.exitZSum/ps.trades).toFixed(2)), avgEntrySpr: parseFloat((ps.entrySprSum/ps.trades).toFixed(4)), avgExitSpr: parseFloat((ps.exitSprSum/ps.trades).toFixed(4)), recentTrades: recent, robustness: rob.score, robustnessGrade: rob.grade, robustnessFlags: rob.flags };
    }).sort(function(a,b){return b.pnl - a.pnl;});

    // Return top 30 pairs by PnL (no balanced quota — frontend handles mode labels)
    var topPairs = allPairsSorted.slice(0, 30);

    var overallWR = parseFloat((winTrades.length / totalTrades * 100).toFixed(1));
    var overallPF = (function(){ var gp=winTrades.reduce(function(s,t){return s+t.pnl;},0); var gl=Math.abs(lossTrades.reduce(function(s,t){return s+t.pnl;},0)); return gl>0?parseFloat((gp/gl).toFixed(2)):0; })();
    var overallRob = computeRobustnessScore_(totalTrades, overallWR, parseFloat(avgWin.toFixed(2)), parseFloat(Math.abs(avgLoss).toFixed(2)), { profitFactor: overallPF });

    return {
      trades: totalTrades,
      wins: winTrades.length,
      losses: lossTrades.length,
      winRate: overallWR,
      totalPnl: parseFloat(totalPnl.toFixed(2)),
      avgPnl: parseFloat(avgPnl.toFixed(2)),
      avgWin: parseFloat(avgWin.toFixed(2)),
      avgLoss: parseFloat(avgLoss.toFixed(2)),
      avgHold: parseFloat(avgHold.toFixed(1)),
      profitFactor: overallPF,
      maxDrawdown: parseFloat(maxDrawdown.toFixed(2)),
      zRevertPct: parseFloat((zRevertExits / totalTrades * 100).toFixed(1)),
      equityCurve: downsampleArray_(equityCurve, 500),
      topPairs: topPairs,
      sampleTrades: allTrades.slice(-50),
      params: { zThreshold: zThreshold, exitZ: exitZ, maxHold: maxHold, mode: mode },
      robustness: overallRob.score,
      robustnessGrade: overallRob.grade,
      robustnessFlags: overallRob.flags,
      _meta: _meta
    };
  } catch(e) {
    console.error('runBacktest_ error: ' + e);
    return { trades: 0, error: e.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════
// TRADE ALERTS — smart entry signals from current alert data
// ═══════════════════════════════════════════════════════════════════
function getTradeAlerts_() {
  try {
    var alerts = [];
    try { var intra = getAlertData('intra'); if (intra && intra.length) alerts = alerts.concat(intra.map(function(a){a.mode='intra';return a;})); } catch(e){}
    try { var credit = getAlertData('credit'); if (credit && credit.length) alerts = alerts.concat(credit.map(function(a){a.mode='credit';return a;})); } catch(e){}

    // Score and rank alerts
    var scored = [];
    for (var i = 0; i < alerts.length; i++) {
      var a = alerts[i];
      var absZ = Math.abs(parseFloat(a.z) || 0);
      var age = parseInt(a.age) || 0;
      var liq = parseFloat(a.liq) || 0;
      var expProfit = parseFloat(a.expProfit) || 0;
      var volSpike = a.volSpike ? true : false;

      // Trend convergence: check if Z is moving toward 0
      var trend = a.zTrend || [];
      var converging = false;
      if (trend.length >= 2) {
        var last = Math.abs(parseFloat(trend[trend.length-1]) || 0);
        var prev = Math.abs(parseFloat(trend[trend.length-2]) || 0);
        converging = last < prev;
      }

      // Score (similar to frontend composite but server-side)
      var score = 0;
      score += Math.min(absZ / 3.5 * 25, 25); // Z magnitude (25pts)
      score += Math.min(expProfit / 2.0 * 20, 20); // Expected profit (20pts)
      score += Math.min(liq / 200000 * 15, 15); // Liquidity (15pts)
      if (volSpike) score += 5; // Volume spike (5pts)
      if (converging) score += 10; // Trend convergence (10pts)
      // Age sweet spot: 3-14 days is ideal
      if (age >= 3 && age <= 14) score += 10;
      else if (age > 14 && age <= 30) score += 5;
      score += Math.min((parseFloat(a.yA)||0) / 10 * 5, 5); // Yield (5pts)

      // Div proximity penalty
      var divA = a.exDivA ? Math.floor((new Date(a.exDivA).getTime() - new Date().getTime()) / 86400000) : 999;
      var divB = a.exDivB ? Math.floor((new Date(a.exDivB).getTime() - new Date().getTime()) / 86400000) : 999;
      if (Math.min(divA, divB) <= 7) score -= 5;

      scored.push({
        id: a.id, tA: a.tA, tB: a.tB, mode: a.mode,
        z: a.z, age: age, liq: liq, expProfit: a.expProfit,
        volSpike: volSpike, converging: converging,
        score: parseFloat(Math.max(0, Math.min(100, score)).toFixed(1)),
        yA: a.yA, yB: a.yB, sec: a.sec,
        exDivA: a.exDivA, exDivB: a.exDivB
      });
    }

    scored.sort(function(a,b){ return b.score - a.score; });
    return scored.slice(0, 30);
  } catch(e) {
    console.error('getTradeAlerts_ error: ' + e);
    return [];
  }
}

// ═══════════════════════════════════════════════════════════════════
// POSITION SIZING — capital allocation calculator
// ═══════════════════════════════════════════════════════════════════
function getPositionSizing_(maxLossPerTrade) {
  try {
    if (!maxLossPerTrade || maxLossPerTrade <= 0) return { error: 'Max loss per trade must be positive' };
    var sigma = 2; // hardcoded 2σ stop width

    // Get current alerts for sizing recommendations
    var alerts = [];
    try { var intra = getAlertData('intra'); if (intra && intra.length) alerts = alerts.concat(intra); } catch(e){}
    try { var credit = getAlertData('credit'); if (credit && credit.length) alerts = alerts.concat(credit); } catch(e){}

    // Get open trades for exposure tracking
    var openTrades = getOpenTrades();
    var currentExposure = 0;
    for (var ot = 0; ot < openTrades.length; ot++) {
      var t = openTrades[ot];
      currentExposure += Math.abs(parseFloat(t.pA) * t.sA) + Math.abs(parseFloat(t.pB) * t.sB);
    }

    // Load screener cache for Kelly suggestions
    var screenerMap = {};
    try {
      var scrData = getScreenerData_();
      if (scrData && scrData.length) {
        for (var s = 0; s < scrData.length; s++) {
          var key = String(scrData[s].id || '').replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
          screenerMap[key] = scrData[s];
        }
      }
    } catch(e){}

    var recommendations = [];
    for (var i = 0; i < Math.min(alerts.length, 20); i++) {
      var a = alerts[i];
      var pA = parseFloat(a.pA) || 0;
      var pB = parseFloat(a.pB) || 0;
      var stdev = parseFloat(a.stdev) || 0;
      if (pA <= 0 || pB <= 0 || stdev <= 0) continue;

      // Core sizing: shares = maxLoss / (stopSigma * stdev)
      // StDev is the spread standard deviation (price-space)
      var spreadRisk = sigma * stdev;
      var shares = Math.floor(maxLossPerTrade / spreadRisk);
      if (shares <= 0) continue;

      var notional = shares * (pA + pB);
      var maxLossActual = shares * spreadRisk;

      // Kelly-informed suggestion
      var kellyShares = null;
      var kellyFraction = null;
      var winRate = null;
      var scrKey = String(a.id || '').replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
      var scr = screenerMap[scrKey];
      if (scr && scr.wr30 !== undefined && scr.ep30 !== undefined) {
        winRate = parseFloat(scr.wr30) / 100;
        var avgWin = parseFloat(scr.ep30) || 0;
        var avgLoss = (scr.avgMae != null && parseFloat(scr.avgMae) > 0) ? parseFloat(scr.avgMae) : spreadRisk;
        if (avgWin > 0 && avgLoss > 0 && winRate > 0 && winRate <= 1) {
          // Kelly fraction: f* = (p * b - q) / b where b = avgWin/avgLoss, p = winRate, q = 1-p
          var b = avgWin / avgLoss;
          var kelly = b > 0 ? Math.min((winRate * b - (1 - winRate)) / b, 1.0) : 0;
          if (kelly > 0) {
            kellyFraction = parseFloat((kelly * 100).toFixed(1)); // as percentage
            // Half-Kelly shares (conservative): kelly/2 * maxLoss-based shares
            kellyShares = Math.floor(shares * Math.min(kelly * 0.5, 1.0));
          }
        }
      }

      recommendations.push({
        id: a.id, tA: a.tA, tB: a.tB,
        z: a.z, pA: pA, pB: pB,
        stdev: parseFloat(stdev.toFixed(4)),
        shares: shares,
        notional: parseFloat(notional.toFixed(2)),
        maxLoss: parseFloat(maxLossActual.toFixed(2)),
        kellyShares: kellyShares,
        kellyFraction: kellyFraction,
        winRate: winRate !== null ? parseFloat((winRate * 100).toFixed(1)) : null,
        expProfit: a.expProfit
      });
    }

    return {
      maxLossPerTrade: maxLossPerTrade,
      openPositions: openTrades.length,
      currentExposure: parseFloat(currentExposure.toFixed(2)),
      recommendations: recommendations
    };
  } catch(e) {
    console.error('getPositionSizing_ error: ' + e);
    return { error: e.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════
// REGIME DETECTION — yield curve analysis + pair Z dispersion
// ═══════════════════════════════════════════════════════════════════
function getRegimeData_() {
  try {
    var ss = SpreadsheetApp.getActive();

    // 1. Read TreasuryHist for yield curve analysis
    var thSheet = ss.getSheetByName('TreasuryHist');
    var yieldSlope = null; // 30Y - 2Y spread (yield curve slope)
    var rateChange = null; // 30-day change in 10Y yield
    var yieldVol = null;   // 30-day volatility of 10Y yield
    var curveHistory = [];

    if (thSheet && thSheet.getLastRow() > 2) {
      var thData = thSheet.getDataRange().getValues();
      var headers = thData[0];
      // Find column indices
      var col2Y = -1, col5Y = -1, col10Y = -1, col30Y = -1;
      for (var c = 0; c < headers.length; c++) {
        var h = String(headers[c]).trim().toUpperCase();
        if (h === 'US2Y') col2Y = c;
        if (h === 'US5Y') col5Y = c;
        if (h === 'US10Y') col10Y = c;
        if (h === 'US30Y') col30Y = c;
      }

      // Get last 60 rows for analysis
      var startIdx = Math.max(1, thData.length - 60);
      var yields10Y = [];
      for (var i = startIdx; i < thData.length; i++) {
        var y10 = col10Y >= 0 ? parseFloat(thData[i][col10Y]) : NaN;
        if (!isNaN(y10) && y10 > 0) yields10Y.push(y10);
      }

      // Current yield curve slope (30Y - 2Y)
      var lastRow = thData[thData.length - 1];
      var cur2Y = col2Y >= 0 ? parseFloat(lastRow[col2Y]) || 0 : 0;
      var cur30Y = col30Y >= 0 ? parseFloat(lastRow[col30Y]) || 0 : 0;
      if (cur2Y > 0 && cur30Y > 0) {
        yieldSlope = (cur30Y - cur2Y) * 100; // in basis points
      }

      // 30-day rate change
      if (yields10Y.length >= 30) {
        var recent10Y = yields10Y[yields10Y.length - 1];
        var ago10Y = yields10Y[yields10Y.length - 30];
        rateChange = (recent10Y - ago10Y) * 100; // basis points
      }

      // 30-day volatility of 10Y
      if (yields10Y.length >= 20) {
        var last20 = yields10Y.slice(-20);
        var changes = [];
        for (var j = 1; j < last20.length; j++) {
          changes.push((last20[j] - last20[j-1]) * 100); // daily bp change
        }
        var avgChg = changes.reduce(function(a,b){return a+b;},0) / changes.length;
        var sumSq = changes.reduce(function(a,c){return a + (c-avgChg)*(c-avgChg);},0);
        yieldVol = Math.sqrt(sumSq / changes.length); // std dev of daily bp changes
      }
    }

    // 2. Read MacroData for PFF change (preferred stock ETF sentiment)
    var pffChange = null;
    var macroSheet = ss.getSheetByName('MacroData');
    if (macroSheet) {
      try {
        var macroVals = macroSheet.getRange('A2:C4').getValues();
        for (var m = 0; m < macroVals.length; m++) {
          if (String(macroVals[m][0]).toUpperCase() === 'PFF') {
            pffChange = parseFloat(macroVals[m][2]) * 100 || 0; // convert decimal to %
          }
        }
      } catch(e) {}
    }

    // 3. Z-score dispersion across active alerts
    var zDispersion = null;
    var zScores = [];
    try {
      var intra = getAlertData('intra');
      if (intra && intra.length) {
        for (var a = 0; a < intra.length; a++) {
          var zv = parseFloat(intra[a].z) || 0;
          zScores.push(zv);
        }
      }
    } catch(e) {}
    try {
      var credit = getAlertData('credit');
      if (credit && credit.length) {
        for (var a = 0; a < credit.length; a++) {
          var zv = parseFloat(credit[a].z) || 0;
          zScores.push(zv);
        }
      }
    } catch(e) {}

    if (zScores.length >= 3) {
      var zMean = zScores.reduce(function(a,b){return a+b;},0) / zScores.length;
      var zSumSq = zScores.reduce(function(a,z){return a + (z-zMean)*(z-zMean);},0);
      zDispersion = Math.sqrt(zSumSq / zScores.length);
    }

    // 4. Determine regime
    // Logic:
    //   MEAN_REVERT: low rate vol + flat-ish curve + moderate Z dispersion
    //   TRENDING:    large rate change (>30bp/month) + expanding Z dispersion
    //   VOLATILE:    high yield vol (>5bp/day) or extreme Z dispersion (>1.5)
    //   NEUTRAL:     everything else
    var regime = 'NEUTRAL';
    var confidence = 0;
    var signals = [];

    // Yield volatility signal
    if (yieldVol !== null) {
      if (yieldVol > 6) { signals.push('HIGH_VOL'); }
      else if (yieldVol < 3) { signals.push('LOW_VOL'); }
    }

    // Rate direction signal
    if (rateChange !== null) {
      if (Math.abs(rateChange) > 30) { signals.push('RATE_TREND'); }
      else if (Math.abs(rateChange) < 10) { signals.push('RATE_STABLE'); }
    }

    // Z dispersion signal
    if (zDispersion !== null) {
      if (zDispersion > 1.5) { signals.push('Z_EXTREME'); }
      else if (zDispersion < 0.8) { signals.push('Z_TIGHT'); }
    }

    // PFF sentiment
    if (pffChange !== null) {
      if (pffChange < -1.0) { signals.push('PFF_WEAK'); }
      else if (pffChange > 0.5) { signals.push('PFF_STRONG'); }
    }

    // Classify
    if (signals.indexOf('HIGH_VOL') !== -1 || signals.indexOf('Z_EXTREME') !== -1) {
      regime = 'VOLATILE';
      confidence = 0.7;
      if (signals.indexOf('HIGH_VOL') !== -1 && signals.indexOf('Z_EXTREME') !== -1) confidence = 0.9;
    } else if (signals.indexOf('RATE_TREND') !== -1) {
      regime = 'TRENDING';
      confidence = 0.6;
      if (signals.indexOf('PFF_WEAK') !== -1) confidence = 0.8;
    } else if (signals.indexOf('LOW_VOL') !== -1 || signals.indexOf('RATE_STABLE') !== -1) {
      regime = 'MEAN_REVERT';
      confidence = 0.6;
      if (signals.indexOf('LOW_VOL') !== -1 && signals.indexOf('RATE_STABLE') !== -1) confidence = 0.85;
      if (signals.indexOf('Z_TIGHT') !== -1) confidence = Math.min(confidence + 0.1, 1.0);
    }

    return {
      regime: regime,
      confidence: parseFloat(confidence.toFixed(2)),
      signals: signals,
      yieldSlope: yieldSlope !== null ? parseFloat(yieldSlope.toFixed(1)) : null,
      rateChange: rateChange !== null ? parseFloat(rateChange.toFixed(1)) : null,
      yieldVol: yieldVol !== null ? parseFloat(yieldVol.toFixed(2)) : null,
      zDispersion: zDispersion !== null ? parseFloat(zDispersion.toFixed(2)) : null,
      pffChange: pffChange !== null ? parseFloat(pffChange.toFixed(2)) : null,
      alertCount: zScores.length
    };
  } catch(e) {
    console.error('getRegimeData_ error: ' + e);
    return { regime: 'NEUTRAL', confidence: 0, error: e.toString() };
  }
}

// Read pre-computed screener results
function getScreenerData_() {
  try {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName('ScreenerCache');
    if (!sheet || sheet.getLastRow() <= 1) return [];
    var data = sheet.getDataRange().getValues();
    var results = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      if (!r[0]) continue;
      var scrTriggers = parseInt(r[12]) || 0;
      var scrWr30 = parseFloat(r[6]) || 0;
      var scrEp30 = parseFloat(r[13]) || 0;
      var scrMae = parseFloat(r[9]) || 0;
      // Compute robustness: use triggers as trade count proxy, wr30 as win rate, ep30 as reward, avgMae as loss
      var scrRob = computeRobustnessScore_(scrTriggers, scrWr30, Math.max(0, scrEp30), scrMae > 0 ? scrMae : 0.5, {});
      results.push({
        id: r[0], tA: r[1], tB: r[2], mode: r[3],
        z: r[4], expProfit: r[5],
        wr30: r[6], wr60: r[7], wr90: r[8],
        avgMae: r[9], p75Mae: r[10], wideningProb: r[11],
        triggers: r[12], ep30: r[13], ep60: r[14],
        ts: r[15],
        wr15: r[16] != null ? r[16] : null,
        ep15: r[17] != null ? r[17] : null,
        ep90: r[18] != null ? r[18] : null,
        stagnantRate: r[19] != null ? r[19] : null,
        robustness: scrRob.score,
        robustnessGrade: scrRob.grade,
        robustnessFlags: scrRob.flags
      });
    }
    return results;
  } catch(e) { return []; }
}

// ═══════════════════════════════════════════════════════════════════
// SENSITIVITY HEATMAP — parameter sweep across Entry Z × Exit Z
// Pre-computes Z-series once, then simulates trades for each combo
// ═══════════════════════════════════════════════════════════════════
function runSensitivitySweep_(maxHold, mode, pairId) {
  try {
    var ss = SpreadsheetApp.getActive();
    var histMap = readTickerHistMap_(ss);
    if (Object.keys(histMap).length === 0) return { error: 'No history data' };

    // Load pair definitions — if pairId specified, only load that pair
    var pairDefs = [];
    if (pairId) {
      var info = parseTickerInfo(pairId);
      if (info && info.tA && info.tB) {
        pairDefs.push({ id: info.id, tA: info.tA, tB: info.tB, mode: mode === 'credit' ? 'credit' : 'intra' });
      } else {
        return { error: 'Invalid pairId: ' + pairId };
      }
    } else {
      var pairSources = [];
      if (mode === 'all' || mode === 'intra') pairSources.push({ sheet: 'Pairs', mode: 'intra' });
      if (mode === 'all' || mode === 'credit') pairSources.push({ sheet: 'CreditPairs', mode: 'credit' });
      for (var s = 0; s < pairSources.length; s++) {
        var pSheet = ss.getSheetByName(pairSources[s].sheet);
        if (!pSheet || pSheet.getLastRow() <= 1) continue;
        var pData = pSheet.getDataRange().getValues();
        for (var p = 1; p < pData.length; p++) {
          if (!pData[p][0]) continue;
          var info = parseTickerInfo(pData[p][0]);
          pairDefs.push({ id: info.id, tA: info.tA, tB: info.tB, mode: pairSources[s].mode });
        }
      }
    }

    // Phase 1: Pre-compute rolling Z-score series for each pair
    var WINDOW = 30;
    var startTime = new Date().getTime();
    var MAX_MS = 240000;
    var zSeriesList = []; // [{zScores: [floats], pricesA: [], pricesB: []}]
    var pairsSkipped = 0;

    for (var pi = 0; pi < pairDefs.length; pi++) {
      if (new Date().getTime() - startTime > MAX_MS) { pairsSkipped += (pairDefs.length - pi); break; }
      var pair = pairDefs[pi];
      var histA = histMap[pair.tA.toUpperCase().trim()];
      var histB = histMap[pair.tB.toUpperCase().trim()];
      if (!histA || !histB) continue;
      var len = Math.min(histA.length, histB.length);
      if (len < WINDOW + 5) continue;
      var pA = histA.slice(histA.length - len);
      var pB = histB.slice(histB.length - len);

      // Compute Z-score for each day from WINDOW onward
      var zArr = new Array(len);
      for (var day = WINDOW; day < len; day++) {
        var sumSpr = 0, sumSprSq = 0;
        for (var w = day - WINDOW; w < day; w++) {
          var spr = pA[w] - pB[w];
          sumSpr += spr;
          sumSprSq += spr * spr;
        }
        var rollMean = sumSpr / WINDOW;
        var rollVar = (sumSprSq / WINDOW) - (rollMean * rollMean);
        var rollStdev = rollVar > 0 ? Math.sqrt(rollVar * WINDOW / (WINDOW - 1)) : 0.001;
        zArr[day] = (pA[day] - pB[day] - rollMean) / rollStdev;
      }
      zSeriesList.push({ z: zArr, pA: pA, pB: pB, len: len, start: WINDOW });
    }

    // Phase 2: Sweep parameter grid
    var entryZValues = [1.5, 1.8, 2.0, 2.2, 2.5, 2.8, 3.0];
    var exitZValues = [0.0, 0.25, 0.5, 0.75, 1.0, 1.25, 1.5];
    var grid = [];

    for (var ei = 0; ei < entryZValues.length; ei++) {
      var row = [];
      var zThreshold = entryZValues[ei];
      for (var xi = 0; xi < exitZValues.length; xi++) {
        var exitZ = exitZValues[xi];
        if (exitZ >= zThreshold) { row.push({ trades: 0, winRate: 0, avgPnl: 0, profitFactor: 0 }); continue; }
        var trades = 0, wins = 0, totalPnl = 0, grossWin = 0, grossLoss = 0;

        for (var si = 0; si < zSeriesList.length; si++) {
          var series = zSeriesList[si];
          var openEntry = null;
          for (var day = series.start; day < series.len; day++) {
            var z = series.z[day];
            if (z === undefined) continue;
            if (!openEntry) {
              if (Math.abs(z) >= zThreshold) {
                openEntry = { day: day, z: z, dirA: z > 0 ? -1 : 1, dirB: z > 0 ? 1 : -1, pA: series.pA[day], pB: series.pB[day] };
              }
            } else {
              var hold = day - openEntry.day;
              if (Math.abs(z) <= exitZ || hold >= maxHold) {
                var pnl = ((series.pA[day] - openEntry.pA) * openEntry.dirA + (series.pB[day] - openEntry.pB) * openEntry.dirB) * 100;
                trades++;
                totalPnl += pnl;
                if (pnl > 0) { wins++; grossWin += pnl; } else { grossLoss += Math.abs(pnl); }
                openEntry = null;
              }
            }
          }
        }

        var cellWR = trades > 0 ? parseFloat((wins / trades * 100).toFixed(1)) : 0;
        var cellAvgPnl = trades > 0 ? parseFloat((totalPnl / trades).toFixed(2)) : 0;
        var cellPF = grossLoss > 0 ? parseFloat((grossWin / grossLoss).toFixed(2)) : 0;
        var cellAvgWin = wins > 0 ? parseFloat((grossWin / wins).toFixed(2)) : 0;
        var cellAvgLoss = (trades - wins) > 0 ? parseFloat((grossLoss / (trades - wins)).toFixed(2)) : 0;
        var cellRob = trades > 0 ? computeRobustnessScore_(trades, cellWR, cellAvgWin, cellAvgLoss, { profitFactor: cellPF }) : { score: 0, grade: 'F', flags: [] };
        row.push({
          trades: trades,
          winRate: cellWR,
          avgPnl: cellAvgPnl,
          profitFactor: cellPF,
          totalPnl: parseFloat(totalPnl.toFixed(2)),
          robustness: cellRob.score,
          robustnessGrade: cellRob.grade
        });
      }
      grid.push(row);
    }

    return {
      entryZValues: entryZValues,
      exitZValues: exitZValues,
      grid: grid,
      maxHold: maxHold,
      mode: mode,
      pairId: pairId || null,
      pairsUsed: zSeriesList.length,
      pairsSkipped: pairsSkipped,
      elapsedMs: new Date().getTime() - startTime
    };
  } catch(e) {
    console.error('runSensitivitySweep_ error: ' + e);
    return { error: e.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════
// OPTIMAL ENTRY/EXIT SWEEP — per-pair parameter optimization
// Scans only dislocated pairs (|Z| >= 1.8), sweeps 7×7 param grid
// per pair, returns Top 25 ranked by maximized Avg PnL
// ═══════════════════════════════════════════════════════════════════
function runOptimalSweep_(maxHold, mode) {
  try {
    var startTime = new Date().getTime();
    var MAX_MS = 240000; // 4-min safety

    // Step 1: Get only dislocated pairs from live alerts
    var alertPairs = [];
    var modes = [];
    if (mode === 'all' || mode === 'intra') modes.push('intra');
    if (mode === 'all' || mode === 'credit') modes.push('credit');
    for (var mi = 0; mi < modes.length; mi++) {
      var alerts = getAlertData(modes[mi]);
      if (!alerts || !alerts.length) continue;
      for (var ai = 0; ai < alerts.length; ai++) {
        var a = alerts[ai];
        if (!a || !a.tA || !a.tB) continue;
        alertPairs.push({
          id: a.id, tA: a.tA, tB: a.tB, mode: modes[mi],
          sector: a.sec || '', currentZ: parseFloat(a.z) || 0,
          pA: parseFloat(a.pA) || 0, pB: parseFloat(a.pB) || 0
        });
      }
    }
    if (alertPairs.length === 0) return { results: [], pairsScanned: 0, elapsedMs: new Date().getTime() - startTime };

    // Step 2: Load ticker history once
    var ss = SpreadsheetApp.getActive();
    var histMap = readTickerHistMap_(ss);
    if (Object.keys(histMap).length === 0) return { error: 'No history data' };

    // Step 3: Per-pair parameter sweep
    var WINDOW = 30;
    var entryGrid = [1.5, 1.8, 2.0, 2.2, 2.5, 2.8, 3.0];
    var exitGrid = [0.0, 0.25, 0.5, 0.75, 1.0, 1.25, 1.5];
    var MIN_TRADES = 3; // minimum trades for a combo to be considered
    var results = [];
    var pairsSkippedTime = 0;
    var pairsSkippedHist = 0;

    for (var pi = 0; pi < alertPairs.length; pi++) {
      if (new Date().getTime() - startTime > MAX_MS) { pairsSkippedTime += (alertPairs.length - pi); break; }
      var pair = alertPairs[pi];
      var tA = pair.tA.toUpperCase().trim();
      var tB = pair.tB.toUpperCase().trim();
      var histA = histMap[tA];
      var histB = histMap[tB];
      if (!histA || !histB) { pairsSkippedHist++; continue; }
      var len = Math.min(histA.length, histB.length);
      if (len < WINDOW + 5) { pairsSkippedHist++; continue; }
      var pA = histA.slice(histA.length - len);
      var pB = histB.slice(histB.length - len);

      // Compute rolling Z-score series
      var zArr = new Array(len);
      for (var day = WINDOW; day < len; day++) {
        var sumSpr = 0, sumSprSq = 0;
        for (var w = day - WINDOW; w < day; w++) {
          var spr = pA[w] - pB[w];
          sumSpr += spr;
          sumSprSq += spr * spr;
        }
        var rollMean = sumSpr / WINDOW;
        var rollVar = (sumSprSq / WINDOW) - (rollMean * rollMean);
        var rollStdev = rollVar > 0 ? Math.sqrt(rollVar * WINDOW / (WINDOW - 1)) : 0.001;
        zArr[day] = (pA[day] - pB[day] - rollMean) / rollStdev;
      }

      // Sweep all entry/exit combos, find the one with best avgPnl
      var bestAvgPnl = -Infinity;
      var bestCombo = null;

      for (var ei = 0; ei < entryGrid.length; ei++) {
        var zThreshold = entryGrid[ei];
        for (var xi = 0; xi < exitGrid.length; xi++) {
          var exitZ = exitGrid[xi];
          if (exitZ >= zThreshold) continue;

          var trades = 0, wins = 0, totalPnl = 0, grossWin = 0, grossLoss = 0, holdSum = 0;
          var openEntry = null;
          for (var day = WINDOW; day < len; day++) {
            var z = zArr[day];
            if (z === undefined) continue;
            if (!openEntry) {
              if (Math.abs(z) >= zThreshold) {
                openEntry = { day: day, z: z, dirA: z > 0 ? -1 : 1, dirB: z > 0 ? 1 : -1, pA: pA[day], pB: pB[day] };
              }
            } else {
              var hold = day - openEntry.day;
              if (Math.abs(z) <= exitZ || hold >= maxHold) {
                var pnl = ((pA[day] - openEntry.pA) * openEntry.dirA + (pB[day] - openEntry.pB) * openEntry.dirB) * 100;
                trades++;
                totalPnl += pnl;
                holdSum += hold;
                if (pnl > 0) { wins++; grossWin += pnl; } else { grossLoss += Math.abs(pnl); }
                openEntry = null;
              }
            }
          }

          if (trades >= MIN_TRADES) {
            var avgPnl = totalPnl / trades;
            if (avgPnl > bestAvgPnl) {
              bestAvgPnl = avgPnl;
              bestCombo = {
                optimalEntryZ: zThreshold,
                optimalExitZ: exitZ,
                avgPnl: parseFloat(avgPnl.toFixed(2)),
                winRate: parseFloat((wins / trades * 100).toFixed(1)),
                trades: trades,
                avgHold: parseFloat((holdSum / trades).toFixed(1)),
                totalPnl: parseFloat(totalPnl.toFixed(2)),
                profitFactor: grossLoss > 0 ? parseFloat((grossWin / grossLoss).toFixed(2)) : (grossWin > 0 ? 99.9 : 0)
              };
            }
          }
        }
      }

      if (bestCombo) {
        bestCombo.id = pair.id;
        bestCombo.tA = pair.tA;
        bestCombo.tB = pair.tB;
        bestCombo.mode = pair.mode;
        bestCombo.sector = pair.sector;
        bestCombo.currentZ = parseFloat(pair.currentZ.toFixed(2));
        results.push(bestCombo);
      }
    }

    // Step 4: Rank by avgPnl descending, return top 25
    results.sort(function(a, b) { return b.avgPnl - a.avgPnl; });
    var top25 = results.slice(0, 25);

    return {
      results: top25,
      pairsScanned: alertPairs.length,
      pairsOptimized: results.length,
      pairsSkippedHist: pairsSkippedHist,
      pairsSkippedTime: pairsSkippedTime,
      elapsedMs: new Date().getTime() - startTime,
      params: { maxHold: maxHold, mode: mode, entryGrid: entryGrid, exitGrid: exitGrid }
    };
  } catch(e) {
    console.error('runOptimalSweep_ error: ' + e);
    return { error: e.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════
// EXIT ALERTS — proactive exit signal detection for open trades
// Three triggers: Statistical Target, Time Stop, Stop Loss
// ═══════════════════════════════════════════════════════════════════

/**
 * Check all open trades for exit signals. Called hourly by trigger
 * and on-demand via getExitAlerts API endpoint.
 * @param {Object} [params] - Optional override thresholds {exitZ, maxHold, stopLossPct}
 * @returns {Array} Array of exit signal objects
 */
function checkExitSignals_(params) {
  try {
    var ss = SpreadsheetApp.getActive();
    var openSheet = ss.getSheetByName('OpenTrades');
    if (!openSheet || openSheet.getLastRow() < 2) return [];

    var openData = openSheet.getDataRange().getValues();

    // Merge live rows from WebCache + WebCacheCredit (same as getOpenTrades)
    var liveRows = [];
    var intra = ss.getSheetByName('WebCache');
    if (!intra || intra.getLastRow() <= 1) intra = ss.getSheetByName('Live');
    var credit = ss.getSheetByName('WebCacheCredit');
    var sources = [intra, credit];
    for (var s = 0; s < sources.length; s++) {
      var ls = sources[s];
      if (ls && ls.getLastRow() > 1) {
        var rows = ls.getRange(2, 1, ls.getLastRow() - 1, 24).getValues();
        liveRows = liveRows.concat(rows);
      }
    }

    // Load screener cache for P75 MAE data
    var screenerMap = {};
    var scrSheet = ss.getSheetByName('ScreenerCache');
    if (scrSheet && scrSheet.getLastRow() > 1) {
      var scrData = scrSheet.getDataRange().getValues();
      for (var si = 1; si < scrData.length; si++) {
        var scrId = String(scrData[si][0] || '').toUpperCase().replace(/[^A-Z0-9|]/g, '');
        if (scrId) {
          screenerMap[scrId] = {
            p75Mae: parseFloat(scrData[si][10]) || 0,
            avgMae: parseFloat(scrData[si][9]) || 0
          };
        }
      }
    }

    // Default thresholds — can be overridden by params
    var exitZ = (params && params.exitZ) ? parseFloat(params.exitZ) : 0.5;
    var maxHold = (params && params.maxHold) ? parseInt(params.maxHold) : 60;
    var stopLossPct = (params && params.stopLossPct) ? parseFloat(params.stopLossPct) : 0;
    var now = new Date();
    var signals = [];

    for (var j = 1; j < openData.length; j++) {
      var rawId = openData[j][0];
      if (!rawId) continue;
      var info = parseTickerInfo(rawId);
      var openAnchor = cleanId(rawId);

      // Find live data for this pair
      var pair = null;
      for (var k = 0; k < liveRows.length; k++) {
        if (liveRows[k][0] && cleanId(liveRows[k][0]) === openAnchor) { pair = liveRows[k]; break; }
      }
      if (!pair) continue;

      var costA = parseMoney(openData[j][2]);
      var costB = parseMoney(openData[j][3]);
      var sA = parseMoney(openData[j][4]);
      var sB = parseMoney(openData[j][5]);
      var paidDiv = parseMoney(openData[j][7]);
      var rcvdDiv = parseMoney(openData[j][8]);
      var openDate = openData[j][6] instanceof Date ? openData[j][6] : new Date(openData[j][6]);

      var livePriceA = parseFloat(pair[3]) || 0;
      var livePriceB = parseFloat(pair[4]) || 0;
      var currentZ = parseFloat(pair[12]) || 0;
      var stdev = parseFloat(pair[11]) || 0;
      var mean = parseFloat(pair[10]) || 0;

      var capGains = ((livePriceA - costA) * sA) + ((livePriceB - costB) * sB);
      var netPnl = capGains + rcvdDiv - paidDiv;
      var daysHeld = Math.floor((now - openDate) / 86400000);
      var entryNotional = Math.abs(costA * sA) + Math.abs(costB * sB);

      // --- Trigger 1: Statistical Target Reached ---
      if (Math.abs(currentZ) <= exitZ) {
        signals.push({
          pairId: info.id,
          tA: info.tA,
          tB: info.tB,
          trigger: 'STATISTICAL_TARGET',
          severity: 'profit',
          message: 'Z-score reverted to ' + currentZ.toFixed(2) + 'σ (target: ≤' + exitZ.toFixed(1) + 'σ) — take profit',
          currentZ: currentZ,
          daysHeld: daysHeld,
          unrealizedPnl: parseFloat(netPnl.toFixed(2)),
          timestamp: now.toISOString()
        });
      }

      // --- Trigger 2: Time Stop ---
      if (daysHeld >= maxHold) {
        signals.push({
          pairId: info.id,
          tA: info.tA,
          tB: info.tB,
          trigger: 'TIME_STOP',
          severity: 'warning',
          message: 'Day ' + daysHeld + ' of ' + maxHold + '-day max hold — consider closing',
          currentZ: currentZ,
          daysHeld: daysHeld,
          unrealizedPnl: parseFloat(netPnl.toFixed(2)),
          timestamp: now.toISOString()
        });
      }

      // --- Trigger 3: Stop Loss ---
      var scrKey = openAnchor;
      var scrInfo = screenerMap[scrKey];
      var lossThreshold = 0;

      if (stopLossPct > 0 && entryNotional > 0) {
        // Hard dollar/pct stop loss
        lossThreshold = entryNotional * (stopLossPct / 100);
      } else if (scrInfo && scrInfo.p75Mae > 0) {
        // P75 MAE-based stop loss (from screener)
        lossThreshold = scrInfo.p75Mae;
      } else if (stdev > 0) {
        // Fallback: 2σ spread risk × avg position size
        var avgShares = (Math.abs(sA) + Math.abs(sB)) / 2;
        lossThreshold = 2 * stdev * avgShares;
      }

      if (lossThreshold > 0 && netPnl < 0 && Math.abs(netPnl) >= lossThreshold) {
        var lossSource = (stopLossPct > 0) ? (stopLossPct + '% of notional') : (scrInfo && scrInfo.p75Mae > 0 ? 'P75 MAE' : '2σ spread risk');
        signals.push({
          pairId: info.id,
          tA: info.tA,
          tB: info.tB,
          trigger: 'STOP_LOSS',
          severity: 'danger',
          message: 'Loss $' + Math.abs(netPnl).toFixed(2) + ' exceeds ' + lossSource + ' threshold ($' + lossThreshold.toFixed(2) + ')',
          currentZ: currentZ,
          daysHeld: daysHeld,
          unrealizedPnl: parseFloat(netPnl.toFixed(2)),
          lossThreshold: parseFloat(lossThreshold.toFixed(2)),
          timestamp: now.toISOString()
        });
      }
    }

    return signals;
  } catch (e) {
    console.error('checkExitSignals_ error: ' + e);
    return [];
  }
}

/**
 * Hourly trigger — checks exit signals and logs them to ExitAlerts sheet.
 */
function runExitAlertCheck() {
  try {
    var ss = SpreadsheetApp.getActive();
    // Read user-configured thresholds from ExitParams sheet (if exists)
    var params = {};
    var paramSheet = ss.getSheetByName('ExitParams');
    if (paramSheet && paramSheet.getLastRow() > 1) {
      var paramData = paramSheet.getDataRange().getValues();
      for (var i = 1; i < paramData.length; i++) {
        var key = String(paramData[i][0]).trim();
        var val = paramData[i][1];
        if (key === 'exitZ') params.exitZ = parseFloat(val);
        else if (key === 'maxHold') params.maxHold = parseInt(val);
        else if (key === 'stopLossPct') params.stopLossPct = parseFloat(val);
      }
    }

    var signals = checkExitSignals_(params);
    if (signals.length === 0) return;

    // Write to ExitAlerts sheet (append-only log)
    var alertSheet = ss.getSheetByName('ExitAlerts');
    if (!alertSheet) {
      alertSheet = ss.insertSheet('ExitAlerts');
      alertSheet.getRange(1, 1, 1, 9).setValues([['Timestamp', 'PairID', 'TickerA', 'TickerB', 'Trigger', 'Severity', 'Message', 'CurrentZ', 'UnrealizedPnL']]);
      alertSheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
    }

    var rows = [];
    for (var i = 0; i < signals.length; i++) {
      var sig = signals[i];
      rows.push([sig.timestamp, sig.pairId, sig.tA, sig.tB, sig.trigger, sig.severity, sig.message, sig.currentZ, sig.unrealizedPnl]);
    }
    alertSheet.getRange(alertSheet.getLastRow() + 1, 1, rows.length, 9).setValues(rows);

    // Trim to last 500 rows to prevent sheet bloat
    var totalRows = alertSheet.getLastRow();
    if (totalRows > 501) {
      alertSheet.deleteRows(2, totalRows - 501);
    }

    Logger.log('Exit alert check: ' + signals.length + ' signals logged.');
  } catch (e) {
    Logger.log('runExitAlertCheck error: ' + e);
  }
}

/**
 * API endpoint — returns recent exit alerts + live check.
 * Combines: fresh live signals + last 24h of logged signals.
 */
function getExitAlerts_() {
  try {
    var ss = SpreadsheetApp.getActive();
    // Read user thresholds
    var params = {};
    var paramSheet = ss.getSheetByName('ExitParams');
    if (paramSheet && paramSheet.getLastRow() > 1) {
      var paramData = paramSheet.getDataRange().getValues();
      for (var i = 1; i < paramData.length; i++) {
        var key = String(paramData[i][0]).trim();
        var val = paramData[i][1];
        if (key === 'exitZ') params.exitZ = parseFloat(val);
        else if (key === 'maxHold') params.maxHold = parseInt(val);
        else if (key === 'stopLossPct') params.stopLossPct = parseFloat(val);
      }
    }

    // Fresh live check
    var liveSignals = checkExitSignals_(params);

    // Read recent logged signals (last 24h) from ExitAlerts sheet
    var recentLog = [];
    var alertSheet = ss.getSheetByName('ExitAlerts');
    if (alertSheet && alertSheet.getLastRow() > 1) {
      var logData = alertSheet.getDataRange().getValues();
      var cutoff = new Date(new Date().getTime() - 86400000); // 24h ago
      for (var i = logData.length - 1; i >= 1; i--) {
        var ts = new Date(logData[i][0]);
        if (ts < cutoff) break;
        recentLog.push({
          timestamp: logData[i][0],
          pairId: logData[i][1],
          trigger: logData[i][4],
          severity: logData[i][5],
          message: logData[i][6],
          currentZ: logData[i][7],
          unrealizedPnl: logData[i][8]
        });
      }
    }

    return {
      liveSignals: liveSignals,
      recentLog: recentLog,
      params: { exitZ: params.exitZ || 0.5, maxHold: params.maxHold || 60, stopLossPct: params.stopLossPct || 0 }
    };
  } catch (e) {
    console.error('getExitAlerts_ error: ' + e);
    return { liveSignals: [], recentLog: [], params: {} };
  }
}

/**
 * Save exit alert parameters to ExitParams sheet.
 * Called via ?action=saveExitParams&exitZ=0.5&maxHold=60&stopLossPct=5
 */
function saveExitParams_(exitZ, maxHold, stopLossPct) {
  try {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName('ExitParams');
    if (!sheet) {
      sheet = ss.insertSheet('ExitParams');
      sheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]);
      sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    }
    // Clear and rewrite
    if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
    sheet.getRange(2, 1, 3, 2).setValues([
      ['exitZ', exitZ],
      ['maxHold', maxHold],
      ['stopLossPct', stopLossPct]
    ]);
    return true;
  } catch (e) {
    console.error('saveExitParams_ error: ' + e);
    return false;
  }
}

// ═══════════════════════════════════════════════════════════════════
// CORRELATION MONITOR — stub (planned feature)
// ═══════════════════════════════════════════════════════════════════
function getCorrelationMonitor_() {
  try {
    var ss = SpreadsheetApp.getActive();
    var openSheet = ss.getSheetByName('OpenTrades');
    if (!openSheet || openSheet.getLastRow() <= 1) return { pairs: 0 };
    var openData = openSheet.getDataRange().getValues();
    var histMap = readTickerHistMap_(ss);
    return computePairCorrelation_(openData, histMap);
  } catch (e) {
    return { pairs: 0, error: e.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════
// NOTIFICATION SETTINGS — stub (planned feature)
// ═══════════════════════════════════════════════════════════════════
function saveNotificationSettings_(chatId, botToken) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('NotificationSettings');
  if (!sheet) {
    sheet = ss.insertSheet('NotificationSettings');
    sheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]);
  }
  if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
  sheet.getRange(2, 1, 2, 2).setValues([['chatId', chatId], ['botToken', botToken]]);
}

function getNotificationSettings_() {
  try {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName('NotificationSettings');
    if (!sheet || sheet.getLastRow() <= 1) return { chatId: '', botToken: '' };
    var data = sheet.getDataRange().getValues();
    var settings = {};
    for (var i = 1; i < data.length; i++) {
      settings[String(data[i][0]).trim()] = String(data[i][1] || '');
    }
    return settings;
  } catch (e) {
    return { chatId: '', botToken: '' };
  }
}

// ═══════════════════════════════════════════════════════════════════
// MODEL PORTFOLIO GENERATOR
// ═══════════════════════════════════════════════════════════════════
// 3-stage pipeline:
//   Stage 1: Candidate selection from ScreenerCache + live alerts
//   Stage 2: Pairwise spread correlation matrix
//   Stage 3: Greedy diversified portfolio construction (5 portfolios × 5 pairs)
// ═══════════════════════════════════════════════════════════════════

/**
 * Read cached model portfolios from ModelPortfolioCache sheet.
 */
function getModelPortfolios_() {
  try {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName('ModelPortfolioCache');
    if (!sheet || sheet.getLastRow() <= 1) return { portfolios: [], updatedAt: null };
    var data = sheet.getDataRange().getValues();
    var portfolios = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      if (!r[0] && r[0] !== 0) continue;
      try {
        var m = JSON.parse(r[8] || '{}');
        portfolios.push({
          rank: parseInt(r[0]) || (i),
          pairs: JSON.parse(r[1] || '[]'),
          basketZ: parseFloat(r[2]) || 0,
          blendedWR: parseFloat(r[3]) || 0,
          expectedProfit: parseFloat(r[4]) || 0,
          widenProb: parseFloat(r[5]) || 0,
          sectorMix: JSON.parse(r[6] || '{}'),
          avgCorrelation: parseFloat(r[7]) || 0,
          metrics: m,
          profitCapture: m.profitCapture != null ? m.profitCapture : null,
          kellySizing: m.kellySizing || null,
          updatedAt: r[9] || ''
        });
      } catch (e) { continue; }
    }
    // Enrich pairs with live prices from WebCache if pA/pB are missing or zero.
    // This fixes stale cached portfolios that were generated before the price fix.
    if (portfolios.length > 0) {
      var priceMap = buildWebCachePriceMap_(ss);
      for (var pi = 0; pi < portfolios.length; pi++) {
        var pp = portfolios[pi].pairs;
        if (!pp) continue;
        for (var qi = 0; qi < pp.length; qi++) {
          var pair = pp[qi];
          if ((!pair.pA || pair.pA === 0) && priceMap[String(pair.tA).toUpperCase()]) {
            pair.pA = priceMap[String(pair.tA).toUpperCase()];
          }
          if ((!pair.pB || pair.pB === 0) && priceMap[String(pair.tB).toUpperCase()]) {
            pair.pB = priceMap[String(pair.tB).toUpperCase()];
          }
        }
      }
    }
    var ts = portfolios.length > 0 ? portfolios[0].updatedAt : null;
    return { portfolios: portfolios, updatedAt: ts };
  } catch (e) {
    return { portfolios: [], updatedAt: null, error: e.toString() };
  }
}

/**
 * Build a ticker→price map from WebCache + WebCacheCredit sheets.
 * Used by getModelPortfolios_() to enrich pairs with live prices.
 */
function buildWebCachePriceMap_(ss) {
  var priceMap = {};
  var cacheNames = ['WebCache', 'WebCacheCredit'];
  for (var ci = 0; ci < cacheNames.length; ci++) {
    var sheet = ss.getSheetByName(cacheNames[ci]);
    if (!sheet || sheet.getLastRow() <= 1) continue;
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var tA = String(data[i][1] || '').trim().toUpperCase(); // TickerA col B
      var tB = String(data[i][2] || '').trim().toUpperCase(); // TickerB col C
      var pA = parseFloat(data[i][3]) || 0; // PriceA col D
      var pB = parseFloat(data[i][4]) || 0; // PriceB col E
      if (tA && pA > 0 && !priceMap[tA]) priceMap[tA] = pA;
      if (tB && pB > 0 && !priceMap[tB]) priceMap[tB] = pB;
    }
  }
  return priceMap;
}

/**
 * Main generator — called by daily trigger or on-demand via API.
 * Writes 5 model portfolios to ModelPortfolioCache sheet.
 */
function runModelPortfolioGenerator() {
  var startTime = new Date().getTime();
  var MAX_MS = 240000; // 4 min safety
  try {
    var ss = SpreadsheetApp.getActive();
    var histMap = readTickerHistMap_(ss);

    // ── STAGE 1: Candidate selection ──
    var candidates = buildCandidatePool_(ss, histMap, startTime, MAX_MS);
    if (candidates.length < 3) {
      Logger.log('ModelPortfolio: Only ' + candidates.length + ' candidates with strict gates, retrying with relaxed gates...');
      candidates = buildCandidatePool_(ss, histMap, startTime, MAX_MS, true);
    }
    if (candidates.length < 2) {
      Logger.log('ModelPortfolio: Only ' + candidates.length + ' candidates even with relaxed gates. Aborting.');
      return;
    }
    // Cap at 15 candidates — C(15,5) = 3003 combos, tractable with pre-computed metrics
    if (candidates.length > 15) candidates = candidates.slice(0, 15);
    Logger.log('ModelPortfolio: Stage 1 complete — ' + candidates.length + ' candidates in ' + ((new Date().getTime() - startTime) / 1000).toFixed(1) + 's');

    // ── STAGE 2a: Correlation matrix ──
    var corrMatrix = buildCorrelationMatrix_(candidates, histMap);
    Logger.log('ModelPortfolio: Stage 2a complete — correlation matrix in ' + ((new Date().getTime() - startTime) / 1000).toFixed(1) + 's');

    // ── STAGE 2b: Pre-compute cross-pair metrics for all N×N long×short combinations ──
    // This is the expensive step — O(N²) calls to computeBasketMetrics_ + probability engine.
    // But it's done ONCE, and all C(n,k) combo evaluations then use O(1) lookups.
    // Budget: up to 50% of MAX_MS for Stage 2b, leaving the rest for Stage 3+ sweeps
    var STAGE2B_DEADLINE = MAX_MS * 0.5;
    var crossMetrics = precomputeCrossPairMetrics_(candidates, histMap, startTime, STAGE2B_DEADLINE);
    Logger.log('ModelPortfolio: Stage 2b complete — ' + candidates.length + '×' + candidates.length + ' cross-pair metrics pre-computed in ' + ((new Date().getTime() - startTime) / 1000).toFixed(1) + 's');

    // ── STAGE 3: Exhaustive combinatorial search ──
    // Reserve 30s for mini-sweeps (Stage 4)
    var SWEEP_RESERVE_MS = 30000;
    var SEARCH_DEADLINE = MAX_MS - SWEEP_RESERVE_MS;
    var PAIRS_PER_PORTFOLIO = Math.min(5, candidates.length);
    var NUM_PORTFOLIOS = 5;

    var scoredCombos = exhaustivePortfolioSearch_(candidates, crossMetrics, corrMatrix, PAIRS_PER_PORTFOLIO, startTime, SEARCH_DEADLINE);
    Logger.log('ModelPortfolio: Stage 3 complete — ' + scoredCombos.length + ' valid combos from C(' + candidates.length + ',' + PAIRS_PER_PORTFOLIO + ') in ' + ((new Date().getTime() - startTime) / 1000).toFixed(1) + 's');

    // ── STAGE 3b: Select diverse top-K portfolios ──
    var topCombos = diverseTopK_(scoredCombos, NUM_PORTFOLIOS, 2);
    Logger.log('ModelPortfolio: Stage 3b complete — ' + topCombos.length + ' diverse portfolios selected');

    // ── Build portfolio objects + STAGE 4-6: mini-sweeps, profit capture, Kelly ──
    var portfolios = [];
    for (var p = 0; p < topCombos.length; p++) {
      var portfolio = buildPortfolioFromCombo_(topCombos[p], candidates, corrMatrix, histMap);

      // STAGE 4: Per-pair mini-sweeps for optimal entry/exit
      for (var mp = 0; mp < portfolio.pairs.length; mp++) {
        if (new Date().getTime() - startTime > MAX_MS) break;
        var mpPair = portfolio.pairs[mp];
        var sweep = miniSweepSinglePair_(mpPair.tA, mpPair.tB, histMap, 60);
        if (sweep) {
          mpPair.optEntry = sweep.optEntry;
          mpPair.optExit = sweep.optExit;
          mpPair.optWR = sweep.optWR;
          mpPair.optAvgPnl = sweep.optAvgPnl;
          mpPair.optTrades = sweep.optTrades;
          mpPair.optProfitFactor = sweep.optProfitFactor;
        } else {
          var absZ = Math.abs(parseFloat(mpPair.z) || 2.0);
          mpPair.optEntry = Math.max(1.5, Math.round(absZ * 10) / 10);
          mpPair.optExit = 0.5;
          mpPair.optWR = parseFloat(mpPair.wr30) || 0;
          mpPair.optAvgPnl = 0;
          mpPair.optTrades = 0;
          mpPair.optProfitFactor = 0;
        }
      }

      // STAGE 5: Basket-level profit capture %
      portfolio.profitCapture = computeProfitCapture_(portfolio.pairs, histMap);

      // STAGE 6: Basket-level Kelly sizing
      portfolio.kellySizing = computeBasketKelly_(portfolio);

      portfolios.push(portfolio);
    }

    Logger.log('ModelPortfolio: Stages 3-6 complete — ' + portfolios.length + ' portfolios with sweeps in ' + ((new Date().getTime() - startTime) / 1000).toFixed(1) + 's');

    // Write to ModelPortfolioCache (portfolios are already sorted by exhaustive search score)
    if (portfolios.length > 0) {
      writeModelPortfolioCache_(ss, portfolios);
      Logger.log('ModelPortfolio: Complete — ' + portfolios.length + ' portfolios generated in ' + ((new Date().getTime() - startTime) / 1000).toFixed(1) + 's');
    } else {
      Logger.log('ModelPortfolio: WARNING — 0 portfolios passed filters in ' + ((new Date().getTime() - startTime) / 1000).toFixed(1) + 's. Keeping existing cache.');
    }
  } catch (e) {
    Logger.log('ModelPortfolio error: ' + e);
  }
}

/**
 * Stage 1: Build candidate pool from ScreenerCache + live alerts.
 * Returns up to 15 candidates sorted by composite quality score.
 */
function buildCandidatePool_(ss, histMap, startTime, maxMs, relaxed) {
  var candidates = [];
  var seen = {};
  var CANDIDATE_TIME_BUDGET = maxMs * 0.4; // Use at most 40% of total budget for candidates

  // ── Anti-overfitting constants for candidate screening ──
  var SUSPICIOUS_WR_THRESHOLD = 95; // WR above this gets dampened (likely overfit)
  var SUSPICIOUS_WR_MULT = 0.6;     // dampening multiplier for suspicious WR
  var MIN_TRIGGERS = 3;             // minimum historical trigger signals required

  // Source 1: ScreenerCache (pre-analyzed with win rates + MAE)
  var scrSheet = ss.getSheetByName('ScreenerCache');
  if (scrSheet && scrSheet.getLastRow() > 1) {
    var scrData = scrSheet.getDataRange().getValues();
    for (var i = 1; i < scrData.length; i++) {
      var r = scrData[i];
      var id = String(r[0] || '').trim();
      if (!id) continue;
      var tA = String(r[1] || '').trim();
      var tB = String(r[2] || '').trim();
      // Need both tickers in histMap
      if (!histMap[tA.toUpperCase()] || !histMap[tB.toUpperCase()]) continue;
      var cid = cleanId(id);
      if (seen[cid]) continue;
      seen[cid] = true;
      var wr30 = parseFloat(r[6]) || 0;
      var wr60 = parseFloat(r[7]) || 0;
      var wr90 = parseFloat(r[8]) || 0;
      var avgMae = parseFloat(r[9]) || 0;
      var p75Mae = parseFloat(r[10]) || 0;
      var widenProb = parseFloat(r[11]) || 0;
      var ep30 = parseFloat(r[13]) || 0;
      var ep60 = parseFloat(r[14]) || 0;
      var z = parseFloat(r[4]) || 0;
      var triggers = parseInt(r[12]) || 0;
      // Anti-overfitting: reject candidates with too few historical trigger signals
      if (triggers < MIN_TRIGGERS) continue;
      // Hard gate: reject candidates with poor win rates (relaxed mode lowers threshold)
      var wrGate = relaxed ? 15 : 30;
      if (wr30 < wrGate) continue;
      // Hard gate: reject candidates with negative 30d expected profit (relaxed allows ~zero)
      if (!relaxed && ep30 <= 0) continue;
      if (relaxed && ep30 < -0.5) continue;
      var mode = String(r[3] || 'intra');
      var sector = '';

      // Look up sector from WebCache/WebCacheCredit
      var cacheName = (mode === 'credit') ? 'WebCacheCredit' : 'WebCache';
      var cacheSheet = ss.getSheetByName(cacheName);
      if (cacheSheet && cacheSheet.getLastRow() > 1) {
        var cacheData = cacheSheet.getDataRange().getValues();
        for (var ci = 1; ci < cacheData.length; ci++) {
          if (cleanId(cacheData[ci][0]) === cid) {
            sector = String(cacheData[ci][15] || '');
            break;
          }
        }
      }

      // Composite quality score for ranking (uses 30d win rate and expected profit)
      var qualityScore = computeCandidateScore_(wr30, ep30, widenProb, z, avgMae, p75Mae);
      // Anti-overfitting: dampen suspiciously high win rates (likely curve-fitted)
      if (wr30 > SUSPICIOUS_WR_THRESHOLD) qualityScore *= SUSPICIOUS_WR_MULT;

      candidates.push({
        id: id, tA: tA, tB: tB, mode: mode, sector: sector,
        z: z, wr30: wr30, wr60: wr60, wr90: wr90,
        ep30: ep30, ep60: ep60, widenProb: widenProb,
        avgMae: avgMae, p75Mae: p75Mae,
        qualityScore: qualityScore
      });
    }
  }

  // Source 2: Live alerts (fill up to 15 if screener didn't have enough)
  // IMPORTANT: analyzeSinglePair_ is expensive (~2-5s each), so enforce time budget
  if (candidates.length < 15) {
    var modes = ['intra', 'credit'];
    var wrGate2 = relaxed ? 15 : 30;
    for (var m = 0; m < modes.length; m++) {
      if (new Date().getTime() - startTime > CANDIDATE_TIME_BUDGET) {
        Logger.log('ModelPortfolio: Candidate pool hit time budget at ' + candidates.length + ' candidates (Source 2, mode=' + modes[m] + ')');
        break;
      }
      try {
        var alerts = getAlertData(modes[m]);
        delete alerts._diag;
        delete alerts._allPairs;
        for (var a = 0; a < alerts.length; a++) {
          // Check time budget before each expensive analyzeSinglePair_ call
          if (new Date().getTime() - startTime > CANDIDATE_TIME_BUDGET) break;
          var alert = alerts[a];
          var aid = cleanId(alert.id);
          if (seen[aid]) continue;
          var aTa = String(alert.tA || '').trim().toUpperCase();
          var aTb = String(alert.tB || '').trim().toUpperCase();
          if (!histMap[aTa] || !histMap[aTb]) continue;
          seen[aid] = true;
          var aZ = parseFloat(alert.z) || 0;
          var aEp = parseFloat(alert.expProfit) || 0;
          // No screener data → run live analysis instead of fabricating estimates
          var pA = parseFloat(alert.pA) || 0;
          var pB = parseFloat(alert.pB) || 0;
          if (pA > 0 && pB > 0) {
            try {
              var liveAnalysis = analyzeSinglePair_(alert.tA, alert.tB, pA, pB, aZ);
              if (liveAnalysis && liveAnalysis.metrics) {
                var lm = liveAnalysis.metrics;
                var lwr30 = (lm.winRates && lm.winRates['30d']) ? lm.winRates['30d'].rate : 0;
                var lwr60 = (lm.winRates && lm.winRates['60d']) ? lm.winRates['60d'].rate : 0;
                var lwr90 = (lm.winRates && lm.winRates['90d']) ? lm.winRates['90d'].rate : 0;
                var lep30 = (lm.rollingZ && lm.rollingZ['30d']) ? lm.rollingZ['30d'].expectedProfit : 0;
                var lep60 = (lm.rollingZ && lm.rollingZ['60d']) ? lm.rollingZ['60d'].expectedProfit : 0;
                var lWiden = (lm.badScenario) ? lm.badScenario.wideningProb : 0;
                var lMae = (lm.badScenario) ? lm.badScenario.avgMae : 0;
                var lp75 = (lm.badScenario) ? lm.badScenario.p75Mae : 0;
                if (lwr30 < wrGate2) continue; // Apply hard gate (relaxed-aware)
                if (!relaxed && lep30 <= 0) continue; // Reject negative EP
                if (relaxed && lep30 < -0.5) continue;
                var aScore = computeCandidateScore_(lwr30, lep30, lWiden, aZ, lMae, lp75);
                // Anti-overfitting: dampen suspiciously high win rates
                if (lwr30 > SUSPICIOUS_WR_THRESHOLD) aScore *= SUSPICIOUS_WR_MULT;
                candidates.push({
                  id: alert.id, tA: alert.tA, tB: alert.tB,
                  mode: modes[m], sector: alert.sec || '',
                  z: aZ, wr30: lwr30, wr60: lwr60, wr90: lwr90,
                  ep30: lep30 || 0, ep60: lep60 || 0, widenProb: lWiden,
                  avgMae: lMae, p75Mae: lp75,
                  qualityScore: aScore
                });
              }
            } catch (e) { /* skip pair if analysis fails */ }
          }
          if (candidates.length >= 15) break;
        }
      } catch (e) { /* skip mode if error */ }
      if (candidates.length >= 15) break;
    }
  }

  // ── Ticker deduplication: each ticker may only appear in ONE pair ──
  // When multiple pairs share a ticker, keep only the best pair for that ticker.
  // "Best" = composite of highest win rate, highest expected profit, lowest bad scenario.
  candidates.sort(function(a, b) { return b.qualityScore - a.qualityScore; });

  var tickerBestPair = {};  // ticker → index of best pair that uses it
  var dedupedIndices = [];  // indices into candidates[] that survive dedup
  var rejected = {};        // index → true if pair was rejected

  for (var di = 0; di < candidates.length; di++) {
    var c = candidates[di];
    var ta = String(c.tA).toUpperCase();
    var tb = String(c.tB).toUpperCase();

    // Check if either ticker is already claimed by a better pair
    var conflictA = tickerBestPair[ta] !== undefined;
    var conflictB = tickerBestPair[tb] !== undefined;

    if (!conflictA && !conflictB) {
      // Both tickers are free — claim them
      tickerBestPair[ta] = di;
      tickerBestPair[tb] = di;
      dedupedIndices.push(di);
    } else {
      // At least one ticker is already taken — compare against the existing pair(s)
      // and only replace if this pair is strictly better for that ticker
      var conflictIdx = conflictA ? tickerBestPair[ta] : tickerBestPair[tb];
      var existing = candidates[conflictIdx];

      // Composite comparison: WR (40%), EP (40%), bad scenario (20%, inverted)
      var scoreNew = (c.wr30 || 0) * 0.4 + (c.ep30 || 0) * 10 * 0.4 + Math.max(0, 5 - (c.avgMae || 0)) * 0.2;
      var scoreOld = (existing.wr30 || 0) * 0.4 + (existing.ep30 || 0) * 10 * 0.4 + Math.max(0, 5 - (existing.avgMae || 0)) * 0.2;

      if (scoreNew > scoreOld) {
        // New pair is better — evict the old pair and reclaim its tickers
        var oldTa = String(existing.tA).toUpperCase();
        var oldTb = String(existing.tB).toUpperCase();
        rejected[conflictIdx] = true;
        dedupedIndices = dedupedIndices.filter(function(idx) { return idx !== conflictIdx; });

        // Release old pair's tickers (only if they pointed to the evicted pair)
        if (tickerBestPair[oldTa] === conflictIdx) delete tickerBestPair[oldTa];
        if (tickerBestPair[oldTb] === conflictIdx) delete tickerBestPair[oldTb];

        // Claim tickers for new pair
        tickerBestPair[ta] = di;
        tickerBestPair[tb] = di;
        dedupedIndices.push(di);
      }
      // else: existing pair is better, skip this candidate
    }
  }

  var deduped = dedupedIndices.map(function(idx) { return candidates[idx]; });
  Logger.log('ModelPortfolio: Deduped candidates from ' + candidates.length + ' to ' + deduped.length + ' (unique tickers)');

  // Sort deduped by quality score descending
  deduped.sort(function(a, b) { return b.qualityScore - a.qualityScore; });
  return deduped;
}

/**
 * Compute a 0-100 candidate quality score.
 * Weights: WR30(25%), EP30(25%), low widen(15%), |Z| magnitude(10%), low AvgMAE(10%), low P75MAE(15%)
 * All 5 target metrics are used: Win Rate, Expected Profit, Widen Prob (holding proxy), Avg MAE, P75 MAE.
 * P75 MAE (tail risk) is weighted heavily because it represents the worst-case drawdown scenario.
 */
function computeCandidateScore_(wr30, ep30, widenProb, z, avgMae, p75Mae) {
  // Normalize win rate: 50% = 0, 100% = 25, below 50% goes negative (penalty)
  var wrScore = Math.min(25, (wr30 - 50) * 0.5);
  // Normalize expected profit: negative EP = negative score (penalty), $0 = 0, $2+ = 25
  // CRITICAL: Use signed ep30 — negative EP must penalize, not reward
  var epScore = ep30 >= 0 ? Math.min(25, ep30 * 12.5) : Math.max(-25, ep30 * 12.5);
  // Widen penalty: 0% widen = 15, 50%+ = 0 (proxy for holding period / stagnation risk)
  var widenScore = Math.max(0, 15 - (widenProb * 0.3));
  // Z magnitude: |Z| of 1.8 = 3, |Z| of 3.0 = 10 (reduced weight — extreme Z can be structural)
  var zScore = Math.min(10, Math.max(0, (Math.abs(z) - 1.5) * 6.67));
  // Avg MAE bonus: lower is better. 0 = 10, 2+ = 0
  var maeScore = avgMae > 0 ? Math.max(0, 10 - avgMae * 5) : 5;
  // P75 MAE penalty (tail risk): lower is better. 0 = 15, 3+ = 0
  // This is the 75th percentile worst drawdown — critical for risk management
  var p75 = parseFloat(p75Mae) || 0;
  var p75Score = p75 > 0 ? Math.max(0, 15 - p75 * 5) : 7.5;
  return wrScore + epScore + widenScore + zScore + maeScore + p75Score;
}

/**
 * Robustness Score (0-100): Anti-overfitting metric that rewards the sweet spot
 * where trade count, win rate, AND expected reward ALL converge.
 *
 * Philosophy: A pair is only truly tradeable when it has enough history to trust,
 * wins consistently within that history, and pays real money per trade.
 * If any one of these three is weak, the signal is likely overfit or noise.
 *
 * Inputs:
 *   trades    — number of completed trades (backtest or historical triggers)
 *   winRate   — win rate as percentage (0-100)
 *   avgReward — average profit per winning trade (dollars, per 100 shares for backtest)
 *   avgLoss   — average loss per losing trade (dollars, absolute value)
 *   opts      — optional: { profitFactor, zRevertPct, maxDrawdown, totalPnl }
 *
 * Returns: { score: 0-100, grade: 'A'|'B'|'C'|'D'|'F', flags: [...], details: {...} }
 */
function computeRobustnessScore_(trades, winRate, avgReward, avgLoss, opts) {
  opts = opts || {};
  var flags = [];
  var details = {};

  // ── Tier 1: Trade Count Score (0-30) ──
  // Requires meaningful sample. Sweet spot = 8-30 trades. Below 5 = unreliable.
  // Above 30 = diminishing returns but still good.
  var tradeScore = 0;
  if (trades >= 30) tradeScore = 30;
  else if (trades >= 20) tradeScore = 25 + (trades - 20) / 10 * 5;
  else if (trades >= 10) tradeScore = 18 + (trades - 10) / 10 * 7;
  else if (trades >= 8) tradeScore = 14 + (trades - 8) / 2 * 4;
  else if (trades >= 5) tradeScore = 6 + (trades - 5) / 3 * 8;
  else if (trades >= 3) tradeScore = 2 + (trades - 3) / 2 * 4;
  else tradeScore = trades * 0.5; // 0-2 trades = almost zero
  if (trades < 5) flags.push('LOW_SAMPLE');
  if (trades < 3) flags.push('UNRELIABLE');
  details.tradeScore = parseFloat(tradeScore.toFixed(1));

  // ── Tier 2: Win Rate Score (0-30) ──
  // But ONLY meaningful with sufficient trades. Discounted by sample confidence.
  // Sweet spot: 55-75%. Below 45% = bad. Above 85% = suspicious (possible overfit).
  var wrScore = 0;
  if (winRate >= 55 && winRate <= 75) wrScore = 25 + (winRate - 55) / 20 * 5; // sweet spot: 25-30
  else if (winRate > 75 && winRate <= 85) wrScore = 25 - (winRate - 75) / 10 * 5; // declining: 25-20
  else if (winRate > 85) {
    wrScore = 15 - (winRate - 85) / 15 * 10; // suspicious: 15-5
    if (winRate > 90) flags.push('SUSPICIOUS_WR');
    if (winRate >= 100 && trades > 1) flags.push('PERFECT_WR');
  }
  else if (winRate >= 45) wrScore = 15 + (winRate - 45) / 10 * 10; // marginal: 15-25
  else wrScore = Math.max(0, winRate / 45 * 15); // poor: 0-15
  // Confidence discount: WR score is less trustworthy with fewer trades
  var confidence = trades >= 20 ? 1.0 : trades >= 10 ? 0.85 : trades >= 5 ? 0.65 : 0.35;
  wrScore *= confidence;
  details.wrScore = parseFloat(wrScore.toFixed(1));
  details.confidence = confidence;

  // ── Tier 3: Expected Reward Score (0-30) ──
  // Must be meaningful in dollar terms. $0.50+ avg reward per 100 shares is decent.
  // $2+ is excellent. $0.10 is noise / spread cost territory.
  var rewardScore = 0;
  if (avgReward > 0) {
    if (avgReward >= 3.0) rewardScore = 30;
    else if (avgReward >= 2.0) rewardScore = 24 + (avgReward - 2.0) / 1.0 * 6;
    else if (avgReward >= 1.0) rewardScore = 16 + (avgReward - 1.0) / 1.0 * 8;
    else if (avgReward >= 0.50) rewardScore = 8 + (avgReward - 0.50) / 0.50 * 8;
    else if (avgReward >= 0.25) rewardScore = 3 + (avgReward - 0.25) / 0.25 * 5;
    else rewardScore = avgReward / 0.25 * 3; // near-zero reward
  }
  if (avgReward < 0.25 && trades > 5) flags.push('LOW_REWARD');
  if (avgReward < 0.10) flags.push('NOISE_LEVEL');
  details.rewardScore = parseFloat(rewardScore.toFixed(1));

  // ── Tier 4: Synergy Bonus/Penalty (0-10) ──
  // All three must converge. Penalize if one is great but others are weak.
  var tiers = [details.tradeScore / 30, details.wrScore / 30, details.rewardScore / 30];
  var tierMin = Math.min.apply(null, tiers);
  var tierMax = Math.max.apply(null, tiers);
  var tierSpread = tierMax - tierMin;
  var synergyScore = 0;
  if (tierSpread < 0.2 && tierMin > 0.4) synergyScore = 10; // all three balanced + strong
  else if (tierSpread < 0.3 && tierMin > 0.3) synergyScore = 7;
  else if (tierSpread < 0.4 && tierMin > 0.2) synergyScore = 4;
  else if (tierSpread >= 0.5) {
    synergyScore = -3; // heavy imbalance = likely overfit to one dimension
    flags.push('IMBALANCED');
  }
  // Extra bonus for risk/reward ratio
  if (avgReward > 0 && avgLoss > 0) {
    var rr = avgReward / avgLoss;
    if (rr >= 1.5) synergyScore += 2;
    else if (rr < 0.75) { synergyScore -= 2; flags.push('BAD_RR'); }
    details.riskReward = parseFloat(rr.toFixed(2));
  }
  // Profit factor bonus (if provided)
  if (opts.profitFactor > 0) {
    if (opts.profitFactor >= 1.5 && opts.profitFactor <= 4.0) synergyScore += 1;
    else if (opts.profitFactor > 5.0 && trades < 20) {
      synergyScore -= 2;
      flags.push('EXTREME_PF');
    }
  }
  details.synergyScore = parseFloat(Math.max(-5, Math.min(10, synergyScore)).toFixed(1));

  // ── Final Score ──
  var raw = tradeScore + wrScore + rewardScore + details.synergyScore;
  var score = Math.max(0, Math.min(100, Math.round(raw)));

  // Grade assignment
  var grade = 'F';
  if (score >= 75) grade = 'A';
  else if (score >= 60) grade = 'B';
  else if (score >= 45) grade = 'C';
  else if (score >= 25) grade = 'D';

  return {
    score: score,
    grade: grade,
    flags: flags,
    details: details
  };
}

/**
 * Stage 2: Build pairwise correlation matrix between candidate spread series.
 * Returns a 2D array: corrMatrix[i][j] = Pearson correlation of daily spreads.
 */
function buildCorrelationMatrix_(candidates, histMap) {
  var n = candidates.length;
  // Pre-compute daily spread series for each candidate
  var spreadSeries = [];
  for (var i = 0; i < n; i++) {
    var c = candidates[i];
    var tA = String(c.tA).toUpperCase();
    var tB = String(c.tB).toUpperCase();
    var histA = histMap[tA] || [];
    var histB = histMap[tB] || [];
    var len = Math.min(histA.length, histB.length);
    var series = [];
    for (var d = 0; d < len; d++) {
      var idxA = histA.length - len + d;
      var idxB = histB.length - len + d;
      series.push(histA[idxA] - histB[idxB]);
    }
    spreadSeries.push(series);
  }

  // Compute Pearson correlation for each pair of candidates
  var matrix = [];
  for (var i = 0; i < n; i++) {
    matrix[i] = [];
    for (var j = 0; j < n; j++) {
      if (i === j) { matrix[i][j] = 1.0; continue; }
      if (j < i) { matrix[i][j] = matrix[j][i]; continue; } // symmetric
      matrix[i][j] = pearsonCorrelation_(spreadSeries[i], spreadSeries[j]);
    }
  }
  return matrix;
}

/**
 * Pearson correlation between two arrays.
 * Aligns on the shorter length. Returns 0 if insufficient data.
 */
function pearsonCorrelation_(a, b) {
  var n = Math.min(a.length, b.length);
  if (n < 10) return 0;
  // Use last n values from each
  var offA = a.length - n;
  var offB = b.length - n;
  var sumA = 0, sumB = 0;
  for (var i = 0; i < n; i++) { sumA += a[offA + i]; sumB += b[offB + i]; }
  var meanA = sumA / n, meanB = sumB / n;
  var num = 0, denA = 0, denB = 0;
  for (var i = 0; i < n; i++) {
    var da = a[offA + i] - meanA;
    var db = b[offB + i] - meanB;
    num += da * db;
    denA += da * da;
    denB += db * db;
  }
  var den = Math.sqrt(denA * denB);
  return den > 0 ? num / den : 0;
}

/**
 * Stage 2b: Pre-compute cross-pair metrics for all N×N long×short ticker combinations.
 * For each candidate pair (i), Z-sign determines long/short legs. We then compute metrics
 * for every combination of longLeg[i] × shortLeg[j] across all candidates.
 * Returns a 2D map: crossMetrics[i][j] = { wr30, wr60, wr90, ep30, ep60, ep90, badAvg, badP75, widenProb }
 * where i is the long-leg source candidate and j is the short-leg source candidate.
 * Self-pairs (i===j) are included since they represent the original pair's own cross-pair.
 */
function precomputeCrossPairMetrics_(candidates, histMap, startTime, deadlineMs) {
  var n = candidates.length;
  // Decompose each candidate into its long and short tickers based on Z-sign
  var legs = [];
  for (var i = 0; i < n; i++) {
    var c = candidates[i];
    var z = parseFloat(c.z) || 0;
    var longTk = z > 0 ? c.tB : c.tA;
    var shortTk = z > 0 ? c.tA : c.tB;
    var histL = histMap[String(longTk).toUpperCase().trim()];
    var histS = histMap[String(shortTk).toUpperCase().trim()];
    var longPrice = (histL && histL.length > 0) ? histL[histL.length - 1] : 0;
    var shortPrice = (histS && histS.length > 0) ? histS[histS.length - 1] : 0;
    legs.push({ longTk: longTk, shortTk: shortTk, longPrice: longPrice, shortPrice: shortPrice });
  }

  // Pre-compute metrics for every long[i] × short[j] cross-pair
  var crossMetrics = [];
  var computed = 0, skippedTimeout = 0;
  for (var i = 0; i < n; i++) {
    crossMetrics[i] = [];
    for (var j = 0; j < n; j++) {
      // Timeout check — leave at least 90s for Stage 3+ (exhaustive search + sweeps)
      if (deadlineMs && new Date().getTime() - startTime > deadlineMs) {
        crossMetrics[i][j] = null;
        skippedTimeout++;
        continue;
      }
      var lTk = legs[i].longTk;
      var sTk = legs[j].shortTk;
      if (!histMap[String(lTk).toUpperCase()] || !histMap[String(sTk).toUpperCase()]) {
        crossMetrics[i][j] = null;
        continue;
      }

      var pairLegs = [
        { ticker: lTk, size: 100, direction: 1 },
        { ticker: sTk, size: 100, direction: -1 }
      ];
      var pm = computeBasketMetrics_(pairLegs, histMap);
      if (!pm || pm.error) { crossMetrics[i][j] = null; continue; }

      // Entry spread from current market prices (matches Sandbox methodology)
      var entrySpread = (legs[i].longPrice || 0) - (legs[j].shortPrice || 0);
      if (entrySpread === 0 && pm.netSpread != null) entrySpread = pm.netSpread;

      // Override EP with entry-anchored values
      var pairRZ = pm.rollingZ || {};
      for (var wKey in pairRZ) {
        if (pairRZ[wKey] && !pairRZ[wKey].insufficient) {
          pairRZ[wKey].expectedProfit = parseFloat((pairRZ[wKey].mean - entrySpread).toFixed(4));
        }
      }

      // Run probability engine anchored to entry spread
      var pairProb = { triggers: 0, winRates: {}, badScenario: null };
      if (pm.dailyValuesFull_ && pm.dailyValuesFull_.length >= 20) {
        pairProb = computeHistoricalProbabilities_(pm.dailyValuesFull_, entrySpread, pairRZ, pm.totalWeight);
        if (pairProb.triggers < 3) {
          var widePairProb = computeHistoricalProbabilitiesWide_(pm.dailyValuesFull_, entrySpread, pairRZ, pm.totalWeight);
          if (widePairProb.triggers > pairProb.triggers) pairProb = widePairProb;
        }
      }

      var wr30 = (pairProb.winRates && pairProb.winRates['30d'] && pairProb.winRates['30d'].eligible >= 2) ? pairProb.winRates['30d'].rate : null;
      var wr60 = (pairProb.winRates && pairProb.winRates['60d'] && pairProb.winRates['60d'].eligible >= 2) ? pairProb.winRates['60d'].rate : null;
      var wr90 = (pairProb.winRates && pairProb.winRates['90d'] && pairProb.winRates['90d'].eligible >= 2) ? pairProb.winRates['90d'].rate : null;

      var ep30 = (pairRZ['30d'] && !pairRZ['30d'].insufficient) ? pairRZ['30d'].expectedProfit * 200 : null;
      var ep60 = (pairRZ['60d'] && !pairRZ['60d'].insufficient) ? pairRZ['60d'].expectedProfit * 200 : null;
      var ep90 = (pairRZ['90d'] && !pairRZ['90d'].insufficient) ? pairRZ['90d'].expectedProfit * 200 : null;

      var z30 = (pairRZ['30d'] && !pairRZ['30d'].insufficient) ? pairRZ['30d'].z : null;
      var z60 = (pairRZ['60d'] && !pairRZ['60d'].insufficient) ? pairRZ['60d'].z : null;
      var z90 = (pairRZ['90d'] && !pairRZ['90d'].insufficient) ? pairRZ['90d'].z : null;

      var badAvg = null, badP75 = null, widenProb = null;
      var cpBs = pairProb.badScenario;
      if (cpBs) {
        badAvg = (cpBs.avgMae != null) ? cpBs.avgMae * 200 : null;
        badP75 = (cpBs.p75Mae != null) ? cpBs.p75Mae * 200 : null;
        widenProb = (cpBs.wideningProb != null) ? cpBs.wideningProb : null;
      }

      crossMetrics[i][j] = {
        wr30: wr30, wr60: wr60, wr90: wr90,
        ep30: ep30, ep60: ep60, ep90: ep90,
        z30: z30, z60: z60, z90: z90,
        badAvg: badAvg, badP75: badP75, widenProb: widenProb
      };
      computed++;
    }
  }
  if (skippedTimeout > 0) {
    Logger.log('ModelPortfolio: Stage 2b computed ' + computed + '/' + (n * n) + ' cross-pairs (' + skippedTimeout + ' skipped due to timeout)');
  }
  return crossMetrics;
}

/**
 * Stage 3: Exhaustive portfolio search — enumerates all C(n, k) combinations of candidates,
 * scores each using pre-computed cross-pair metrics, and returns all scored combos sorted by score.
 *
 * For each combo, the portfolio score is computed from the weighted-average cross-matrix of its
 * constituent pairs (same methodology as Sandbox getPortfolioAnalytics). This is O(C(n,k) × k²)
 * since cross-pair lookups are O(1) from pre-computed metrics.
 *
 * @param {Array} candidates - candidate pool (max ~15)
 * @param {Array} crossMetrics - pre-computed N×N cross-pair metrics from precomputeCrossPairMetrics_
 * @param {Array} corrMatrix - pairwise correlation matrix
 * @param {number} k - pairs per portfolio (typically 5)
 * @param {number} startTime - for timeout checking
 * @param {number} deadlineMs - max elapsed ms before aborting
 * @returns {Array} scored combos [{indices: [], score, metrics: {wr30, wr60, wr90, ep30, ep60, ep90, basketZ, badAvg, badP75, widenProb, avgCorr}}] sorted by score desc
 */
function exhaustivePortfolioSearch_(candidates, crossMetrics, corrMatrix, k, startTime, deadlineMs) {
  var n = candidates.length;
  k = Math.min(k, n);
  var results = [];
  var evalCount = 0;

  // ── Tunable constants ──
  var MAX_TICKER_EXPOSURE = 1;  // each ticker can only appear once per portfolio
  var PORTFOLIO_WR_CEILING = 90; // blended WR above this gets skepticism discount

  // Generate all C(n, k) combinations iteratively using an index array
  var combo = [];
  for (var i = 0; i < k; i++) combo[i] = i;

  while (true) {
    // Check timeout every 500 evaluated combos (not just accepted ones)
    if (evalCount % 500 === 0 && evalCount > 0 && new Date().getTime() - startTime > deadlineMs) break;
    evalCount++;

    // ── Ticker exposure gate: skip combos where any ticker appears > MAX_TICKER_EXPOSURE times ──
    var tickerFreq = {};
    var exposureViolation = false;
    for (var te = 0; te < k; te++) {
      var cTa = String(candidates[combo[te]].tA).toUpperCase();
      var cTb = String(candidates[combo[te]].tB).toUpperCase();
      tickerFreq[cTa] = (tickerFreq[cTa] || 0) + 1;
      tickerFreq[cTb] = (tickerFreq[cTb] || 0) + 1;
      if (tickerFreq[cTa] > MAX_TICKER_EXPOSURE || tickerFreq[cTb] > MAX_TICKER_EXPOSURE) {
        exposureViolation = true;
        break;
      }
    }

    if (!exposureViolation) {
    // ── Score this combination using pre-computed cross-metrics ──
    // Each pair contributes a long leg and a short leg. The cross-matrix evaluates
    // every longLeg[i] × shortLeg[j] combination with equal weight (uniform 100 shares each).
    var totalWeight = 0;
    var aggWR30 = 0, aggWR60 = 0, aggWR90 = 0;
    var wrWS30 = 0, wrWS60 = 0, wrWS90 = 0;
    var aggEP30 = 0, aggEP60 = 0, aggEP90 = 0;
    var aggZ30 = 0, zWeightSum = 0;
    var aggZ60 = 0, z60WeightSum = 0;
    var aggZ90 = 0, z90WeightSum = 0;
    var aggBadAvg = 0, aggBadP75 = 0, aggWidenProb = 0;
    var badAvgWS = 0, badP75WS = 0, widenWS = 0;

    var kk = combo.length;
    var w = 1.0 / (kk * kk); // equal weight for each cross-pair

    for (var li = 0; li < kk; li++) {
      for (var si = 0; si < kk; si++) {
        var cm = crossMetrics[combo[li]][combo[si]];
        if (!cm) continue;
        totalWeight += w;

        if (cm.wr30 != null) { aggWR30 += w * cm.wr30; wrWS30 += w; }
        if (cm.wr60 != null) { aggWR60 += w * cm.wr60; wrWS60 += w; }
        if (cm.wr90 != null) { aggWR90 += w * cm.wr90; wrWS90 += w; }
        if (cm.ep30 != null) aggEP30 += cm.ep30;
        if (cm.ep60 != null) aggEP60 += cm.ep60;
        if (cm.ep90 != null) aggEP90 += cm.ep90;
        if (cm.z30 != null) { aggZ30 += w * cm.z30; zWeightSum += w; }
        if (cm.z60 != null) { aggZ60 += w * cm.z60; z60WeightSum += w; }
        if (cm.z90 != null) { aggZ90 += w * cm.z90; z90WeightSum += w; }
        if (cm.badAvg != null) { aggBadAvg += w * cm.badAvg; badAvgWS += w; }
        if (cm.badP75 != null) { aggBadP75 += w * cm.badP75; badP75WS += w; }
        if (cm.widenProb != null) { aggWidenProb += w * cm.widenProb; widenWS += w; }
      }
    }

    var cmWR30 = wrWS30 > 0 ? parseFloat((aggWR30 / wrWS30).toFixed(1)) : 0;
    var cmWR60 = wrWS60 > 0 ? parseFloat((aggWR60 / wrWS60).toFixed(1)) : 0;
    var cmWR90 = wrWS90 > 0 ? parseFloat((aggWR90 / wrWS90).toFixed(1)) : 0;
    // EP is additive (total dollars across all k² cross-pairs, each scaled to 200 shares)
    var cmEP30 = parseFloat(aggEP30.toFixed(2));
    var cmEP60 = parseFloat(aggEP60.toFixed(2));
    var cmEP90 = parseFloat(aggEP90.toFixed(2));
    var basketZ = zWeightSum > 0 ? parseFloat((aggZ30 / zWeightSum).toFixed(2)) : 0;
    var basketZ60 = z60WeightSum > 0 ? parseFloat((aggZ60 / z60WeightSum).toFixed(2)) : 0;
    var basketZ90 = z90WeightSum > 0 ? parseFloat((aggZ90 / z90WeightSum).toFixed(2)) : 0;
    var cmBadAvg = badAvgWS > 0 ? parseFloat((aggBadAvg / badAvgWS).toFixed(2)) : 0;
    var cmBadP75 = badP75WS > 0 ? parseFloat((aggBadP75 / badP75WS).toFixed(2)) : 0;
    var cmWidenProb = widenWS > 0 ? parseFloat((aggWidenProb / widenWS).toFixed(1)) : 0;

    // Hard gates: reject combos that fail at any horizon
    var reject = false;
    if ((cmWR30 > 0 && cmWR30 < 50) || (cmWR60 > 0 && cmWR60 < 50) || (cmWR90 > 0 && cmWR90 < 50)) reject = true;
    if (!reject && (cmEP30 <= 0 || cmEP60 <= 0 || cmEP90 <= 0)) reject = true;

    if (!reject) {
      // Avg pairwise correlation within the combo
      var avgCorrSum = 0, corrCount = 0;
      for (var ci = 0; ci < kk; ci++) {
        for (var cj = ci + 1; cj < kk; cj++) {
          avgCorrSum += Math.abs(corrMatrix[combo[ci]][combo[cj]]);
          corrCount++;
        }
      }
      var avgCorr = corrCount > 0 ? parseFloat((avgCorrSum / corrCount).toFixed(3)) : 0;

      // Risk-adjusted composite score (same formula as final ranking)
      var wrPts = Math.min(25, (cmWR30 - 50) * 0.5);
      var epNorm = cmEP30 / (kk * kk * 200); // normalize to per-share across k² cross-pairs
      var epPts = Math.min(25, Math.max(-25, epNorm * 20));
      var badAvgPts = cmBadAvg > 0 ? Math.max(0, 15 - cmBadAvg * 0.05) : 7.5;
      var badP75Pts = cmBadP75 > 0 ? Math.max(0, 20 - cmBadP75 * 0.05) : 10;
      var widenPts = Math.max(0, 15 - (cmWidenProb * 0.3));
      // Bonus for low correlation (diversification reward, up to 5 pts)
      var corrBonus = Math.max(0, 5 * (1 - avgCorr));
      var score = wrPts + epPts + badAvgPts + badP75Pts + widenPts + corrBonus;

      // ── Anti-overfitting: penalize suspiciously high blended win rates ──
      if (cmWR30 > PORTFOLIO_WR_CEILING) score *= 0.85;

      results.push({
        indices: combo.slice(), // copy
        score: parseFloat(score.toFixed(2)),
        metrics: {
          wr30: cmWR30, wr60: cmWR60, wr90: cmWR90,
          ep30: cmEP30, ep60: cmEP60, ep90: cmEP90,
          basketZ: basketZ, basketZ60: basketZ60, basketZ90: basketZ90,
          badAvg: cmBadAvg, badP75: cmBadP75, widenProb: cmWidenProb,
          avgCorr: avgCorr
        }
      });
    }
    } // end if (!exposureViolation)

    // ── Advance to next combination (lexicographic order) ──
    var pos = k - 1;
    while (pos >= 0 && combo[pos] === n - k + pos) pos--;
    if (pos < 0) break; // all combinations exhausted
    combo[pos]++;
    for (var fill = pos + 1; fill < k; fill++) combo[fill] = combo[fill - 1] + 1;
  }

  // Sort by score descending
  results.sort(function(a, b) { return b.score - a.score; });
  return results;
}

/**
 * Stage 3b: Select top-K diverse portfolios from scored combinations.
 * Avoids picking overlapping portfolios by penalizing pair reuse.
 * Uses a greedy pass over the score-sorted results, skipping combos that share
 * too many pairs with already-selected portfolios.
 *
 * @param {Array} scoredCombos - from exhaustivePortfolioSearch_, sorted by score desc
 * @param {number} topK - number of portfolios to select
 * @param {number} maxOverlap - max shared pairs between any two selected portfolios (default: 2)
 * @returns {Array} top-K diverse combos (same shape as input elements)
 */
function diverseTopK_(scoredCombos, topK, maxOverlap) {
  maxOverlap = maxOverlap || 2;
  var selected = [];

  for (var i = 0; i < scoredCombos.length && selected.length < topK; i++) {
    var combo = scoredCombos[i];
    var tooSimilar = false;

    for (var s = 0; s < selected.length; s++) {
      // Count shared indices between this combo and an already-selected one
      var overlap = 0;
      var selIndices = selected[s].indices;
      for (var a = 0; a < combo.indices.length; a++) {
        for (var b = 0; b < selIndices.length; b++) {
          if (combo.indices[a] === selIndices[b]) { overlap++; break; }
        }
      }
      if (overlap > maxOverlap) { tooSimilar = true; break; }
    }

    if (!tooSimilar) selected.push(combo);
  }

  return selected;
}

/**
 * Build a portfolio object from a scored combo and candidate pool.
 * Assembles the pair list, sector mix, correlation stats, and cross-matrix metrics
 * into the same shape expected by downstream stages (mini-sweep, profit capture, Kelly, cache writer).
 */
function buildPortfolioFromCombo_(combo, candidates, corrMatrix, histMap) {
  var pairs = [];
  var sectorMix = {};
  var totalWr = 0, totalEp = 0, totalWiden = 0, totalStagnant = 0;

  for (var i = 0; i < combo.indices.length; i++) {
    var c = candidates[combo.indices[i]];
    var histA = histMap[String(c.tA).toUpperCase().trim()];
    var histB = histMap[String(c.tB).toUpperCase().trim()];
    var priceA = (histA && histA.length > 0) ? histA[histA.length - 1] : 0;
    var priceB = (histB && histB.length > 0) ? histB[histB.length - 1] : 0;
    pairs.push({
      id: c.id, tA: c.tA, tB: c.tB,
      mode: c.mode, sector: c.sector,
      z: c.z, wr30: c.wr30, wr60: c.wr60, ep30: c.ep30, ep60: c.ep60,
      widenProb: c.widenProb, qualityScore: c.qualityScore,
      avgMae: c.avgMae, p75Mae: c.p75Mae,
      pA: parseFloat(priceA.toFixed(2)), pB: parseFloat(priceB.toFixed(2))
    });
    totalWr += c.wr30;
    totalEp += c.ep30;
    totalWiden += c.widenProb;
    totalStagnant += Math.max(0, 100 - c.wr30 - c.widenProb);
    var sec = c.sector || 'Other';
    sectorMix[sec] = (sectorMix[sec] || 0) + 1;
  }

  var m = combo.metrics;
  var nn = pairs.length || 1;
  return {
    pairs: pairs,
    blendedWR: m.wr30,
    expectedProfit: parseFloat((m.ep30 / (nn * nn * 200)).toFixed(4)),
    widenProb: m.widenProb,
    stagnantRate: parseFloat((totalStagnant / nn).toFixed(1)),
    avgCorrelation: m.avgCorr,
    sectorMix: sectorMix,
    basketZ: m.basketZ,
    rollingZ: {
      '30d': { z: m.basketZ, expectedProfit: parseFloat((m.ep30 / (nn * nn * 200)).toFixed(4)) },
      '60d': { z: m.basketZ60 || 0, expectedProfit: parseFloat((m.ep60 / (nn * nn * 200)).toFixed(4)) },
      '90d': { z: m.basketZ90 || 0, expectedProfit: parseFloat((m.ep90 / (nn * nn * 200)).toFixed(4)) }
    },
    crossMatrixWR: { wr30: m.wr30, wr60: m.wr60, wr90: m.wr90 },
    crossMatrixEP: { ep30: m.ep30, ep60: m.ep60, ep90: m.ep90 },
    badScenario: { avgMaeDollar: m.badAvg, p75MaeDollar: m.badP75 }
  };
}

/**
 * Mini parameter sweep for a single pair — lightweight version of runSensitivitySweep_.
 * Finds the optimal entry/exit Z thresholds based on 30-day rolling Z-score history.
 * Returns { optEntry, optExit, optWR, optAvgPnl, optTrades, optProfitFactor } or null.
 */
function miniSweepSinglePair_(tA, tB, histMap, maxHold) {
  maxHold = maxHold || 60;
  var WINDOW = 30;
  var histA = histMap[String(tA).toUpperCase().trim()];
  var histB = histMap[String(tB).toUpperCase().trim()];
  if (!histA || !histB) return null;
  var len = Math.min(histA.length, histB.length);
  if (len < WINDOW + 5) return null;
  var pA = histA.slice(histA.length - len);
  var pB = histB.slice(histB.length - len);

  // Compute rolling Z-score series
  var zArr = new Array(len);
  for (var day = WINDOW; day < len; day++) {
    var sumSpr = 0, sumSprSq = 0;
    for (var w = day - WINDOW; w < day; w++) {
      var spr = pA[w] - pB[w];
      sumSpr += spr;
      sumSprSq += spr * spr;
    }
    var rollMean = sumSpr / WINDOW;
    var rollVar = (sumSprSq / WINDOW) - (rollMean * rollMean);
    var rollStdev = rollVar > 0 ? Math.sqrt(rollVar * WINDOW / (WINDOW - 1)) : 0.001;
    zArr[day] = (pA[day] - pB[day] - rollMean) / rollStdev;
  }

  // ── Anti-overfitting constants ──
  var MAX_PROFIT_FACTOR = 4.0;   // cap PF to prevent curve-fitted outliers from dominating
  var MIN_SWEEP_TRADES = 3;      // minimum trades for statistical relevance
  var LOW_SAMPLE_THRESHOLD = 5;  // below this, apply shrinkage penalty

  // Sweep entry × exit grid — same values as full heatmap
  var entryZValues = [1.5, 1.8, 2.0, 2.2, 2.5, 2.8, 3.0];
  var exitZValues = [0.0, 0.25, 0.5, 0.75, 1.0, 1.25, 1.5];
  var best = null;

  for (var ei = 0; ei < entryZValues.length; ei++) {
    var zThreshold = entryZValues[ei];
    for (var xi = 0; xi < exitZValues.length; xi++) {
      var exitZ = exitZValues[xi];
      if (exitZ >= zThreshold) continue;
      var trades = 0, wins = 0, totalPnl = 0, grossWin = 0, grossLoss = 0;
      var openEntry = null;

      for (var day = WINDOW; day < len; day++) {
        var z = zArr[day];
        if (z === undefined) continue;
        if (!openEntry) {
          if (Math.abs(z) >= zThreshold) {
            openEntry = { day: day, z: z, dirA: z > 0 ? -1 : 1, dirB: z > 0 ? 1 : -1, pA: pA[day], pB: pB[day] };
          }
        } else {
          var hold = day - openEntry.day;
          if (Math.abs(z) <= exitZ || hold >= maxHold) {
            var pnl = ((pA[day] - openEntry.pA) * openEntry.dirA + (pB[day] - openEntry.pB) * openEntry.dirB) * 100;
            trades++;
            totalPnl += pnl;
            if (pnl > 0) { wins++; grossWin += pnl; } else { grossLoss += Math.abs(pnl); }
            openEntry = null;
          }
        }
      }

      // Require minimum trades for statistical relevance
      if (trades < MIN_SWEEP_TRADES) continue;
      var wr = parseFloat((wins / trades * 100).toFixed(1));
      var avgPnl = parseFloat((totalPnl / trades).toFixed(2));
      var pf = grossLoss > 0 ? parseFloat((grossWin / grossLoss).toFixed(2)) : 0;

      // Anti-overfitting: cap profit factor to prevent curve-fitted outliers
      var pfCapped = Math.min(pf, MAX_PROFIT_FACTOR);

      // Score: prioritize win rate, then capped profit factor, then avg PnL
      // Low sample shrinkage: scale score toward zero when trades < LOW_SAMPLE_THRESHOLD
      var score = wr * 2 + pfCapped * 10 + (avgPnl > 0 ? avgPnl * 0.5 : avgPnl * 2);
      if (trades < LOW_SAMPLE_THRESHOLD) score *= (trades / LOW_SAMPLE_THRESHOLD);
      if (!best || score > best.score) {
        best = {
          score: score,
          optEntry: zThreshold,
          optExit: exitZ,
          optWR: wr,
          optAvgPnl: avgPnl,
          optTrades: trades,
          optProfitFactor: pf
        };
      }
    }
  }

  if (!best) return null;
  return {
    optEntry: best.optEntry,
    optExit: best.optExit,
    optWR: best.optWR,
    optAvgPnl: best.optAvgPnl,
    optTrades: best.optTrades,
    optProfitFactor: best.optProfitFactor
  };
}

/**
 * Compute Basket Profit Capture % for a model portfolio.
 * Measures what % of the theoretical mean-reversion profit the basket historically captures.
 * Theoretical profit = sum of |entryZ × stdev| per pair (full mean reversion to Z=0).
 * Actual expected profit = sum of expected profit from 30d rolling Z.
 * Profit Capture % = actual / theoretical × 100.
 */
function computeProfitCapture_(pairs, histMap) {
  var WINDOW = 30;
  var theoreticalTotal = 0;
  var actualTotal = 0;

  for (var i = 0; i < pairs.length; i++) {
    var pair = pairs[i];
    var z = Math.abs(parseFloat(pair.z) || 0);
    var ep30 = parseFloat(pair.ep30) || 0;

    // Compute the pair's current stdev from history for theoretical profit calculation
    var histA = histMap[String(pair.tA).toUpperCase().trim()];
    var histB = histMap[String(pair.tB).toUpperCase().trim()];
    if (!histA || !histB || histA.length < WINDOW || histB.length < WINDOW) continue;
    var len = Math.min(histA.length, histB.length);
    // Use last WINDOW days for rolling stdev
    var sumSpr = 0, sumSprSq = 0;
    for (var d = len - WINDOW; d < len; d++) {
      var spr = histA[d] - histB[d];
      sumSpr += spr;
      sumSprSq += spr * spr;
    }
    var rollMean = sumSpr / WINDOW;
    var rollVar = (sumSprSq / WINDOW) - (rollMean * rollMean);
    var rollStdev = rollVar > 0 ? Math.sqrt(rollVar * WINDOW / (WINDOW - 1)) : 0;
    if (rollStdev <= 0) continue;

    // Theoretical profit per share if spread fully reverts to mean (Z→0)
    // Per 100 shares (matching the legs size used in basket computation)
    var theoreticalPerShare = z * rollStdev;
    theoreticalTotal += theoreticalPerShare * 100;

    // Actual expected profit from 30d analysis (already per-share × 100 in ep30 from screener)
    // ep30 from screener is per-share, so multiply by 100 shares to match
    actualTotal += Math.abs(ep30) * 100;
  }

  if (theoreticalTotal <= 0) return null;
  return parseFloat((actualTotal / theoreticalTotal * 100).toFixed(1));
}

/**
 * Compute basket-level Kelly criterion for an entire model portfolio.
 * Uses the blended 30d win rate and portfolio-level expected profit/loss.
 * Returns { kellyPct, halfKellyPct, suggestedCapital, riskRewardRatio } or null.
 */
function computeBasketKelly_(portfolio) {
  var wr = parseFloat(portfolio.blendedWR) || 0;
  if (wr <= 0 || wr >= 100) return null;

  var winRate = wr / 100;
  var avgWin = Math.abs(parseFloat(portfolio.expectedProfit) || 0);
  var bs = portfolio.badScenario || {};
  var avgLoss = parseFloat(bs.avgMaeDollar) || parseFloat(bs.avgMae) || 0;
  if (avgWin <= 0 || avgLoss <= 0) return null;

  // Kelly fraction: f* = (p × b - q) / b where b = avgWin/avgLoss
  var b = avgWin / avgLoss;
  var kelly = b > 0 ? (winRate * b - (1 - winRate)) / b : 0;
  if (kelly <= 0) return null;

  var halfKelly = kelly * 0.5;
  var pairs = (portfolio.pairs || []).length || 1;
  // Suggested capital = (avgLoss per pair × pairs) / halfKelly
  // This gives the total capital needed so that max expected loss = halfKelly fraction
  var totalRisk = avgLoss * pairs;
  var suggestedCapital = totalRisk > 0 ? parseFloat((totalRisk / halfKelly).toFixed(0)) : 0;
  var riskRewardRatio = avgLoss > 0 ? parseFloat((avgWin / avgLoss).toFixed(1)) : 0;

  return {
    kellyPct: parseFloat((kelly * 100).toFixed(1)),
    halfKellyPct: parseFloat((halfKelly * 100).toFixed(1)),
    suggestedCapital: suggestedCapital,
    riskRewardRatio: riskRewardRatio
  };
}

/**
 * Write model portfolios to ModelPortfolioCache sheet.
 */
function writeModelPortfolioCache_(ss, portfolios) {
  var sheet = ss.getSheetByName('ModelPortfolioCache');
  if (!sheet) {
    sheet = ss.insertSheet('ModelPortfolioCache');
    sheet.getRange(1, 1, 1, 10).setValues([['Rank', 'Pairs', 'BasketZ', 'BlendedWR', 'ExpProfit', 'WidenProb', 'SectorMix', 'AvgCorrelation', 'Metrics', 'UpdatedAt']]);
  } else {
    if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).clearContent();
  }
  if (portfolios.length === 0) return;
  var ts = new Date().toISOString();
  var rows = portfolios.map(function(p, idx) {
    return [
      idx + 1,
      JSON.stringify(p.pairs),
      p.basketZ,
      p.blendedWR,
      p.expectedProfit,
      p.widenProb,
      JSON.stringify(p.sectorMix),
      p.avgCorrelation,
      JSON.stringify({ rollingZ: p.rollingZ || {}, badScenario: p.badScenario || {}, profitCapture: p.profitCapture, kellySizing: p.kellySizing || null, crossMatrixWR: p.crossMatrixWR || {}, crossMatrixEP: p.crossMatrixEP || {} }),
      ts
    ];
  });
  sheet.getRange(2, 1, rows.length, 10).setValues(rows);
}
