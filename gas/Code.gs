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
        var safeJson2 = JSON.stringify(result, function(key, val) {
          if (typeof val === 'number' && !isFinite(val)) return null;
          return val;
        });
        return ContentService.createTextOutput(safeJson2).setMimeType(ContentService.MimeType.JSON);
      }
      var alerts = getAlertData(mode);
      var diag = alerts._diag || {};
      delete alerts._diag;
      result = {
        ok: true,
        mode: mode,
        status: alerts.length > 0 ? "live" : "no_signals",
        alertData: alerts,
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
      result = { ok: true, analysisData: analyzeSinglePair_(pairTa, pairTb, pairPa, pairPb, pairZ) };
    }
    else {
      result = { ok: false, message: "Unknown action: " + action };
    }
  } catch (err) {
    result = { ok: false, message: err.toString() };
  }
  // Sanitize: replace NaN/Infinity with null so JSON.stringify doesn't silently fail
  var safeJson = JSON.stringify(result, function(key, val) {
    if (typeof val === 'number' && !isFinite(val)) return null;
    return val;
  });
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
            dateKey = ts.substring(0, 10);
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
        if (pair) {
          if (pair[1] && String(pair[1]).length > 1) tA_Name = String(pair[1]);
          if (pair[2] && String(pair[2]).length > 1) tB_Name = String(pair[2]);
          var liveSpr = parseFloat(pair[5])||0;
          var meanTarget = parseFloat(pair[10])||0;
          var livePriceA = parseFloat(pair[3])||0;
          var livePriceB = parseFloat(pair[4])||0;
          var capGains = ((livePriceA-costA)*sA) + ((livePriceB-costB)*sB);
          var netPnl = capGains + rcvdDiv - paidDiv;
          var currentZ = parseFloat(pair[12]) || 0;
          var sector = pair[15] ? String(pair[15]) : '';
          results.push({
            id: displayId, tA: tA_Name, tB: tB_Name,
            spr: liveSpr.toFixed(2), target: meanTarget.toFixed(2),
            sA: sA, sB: sB, pA: costA.toFixed(2), pB: costB.toFixed(2),
            dollarPnL: netPnl.toFixed(2), capGains: capGains.toFixed(2),
            paidDiv: paidDiv.toFixed(2), rcvdDiv: rcvdDiv.toFixed(2),
            centGoal: (Math.abs((costA-costB)-meanTarget)*100).toFixed(0),
            centRem: (Math.abs(liveSpr-meanTarget)*100).toFixed(0),
            isWinning: netPnl > 0, entryZ: openData[j][1],
            currentZ: currentZ, sector: sector, strategy: strategy
          });
        } else {
          results.push({
            id: displayId+" [WAITING]", tA: tA_Name, tB: tB_Name,
            spr:"0.00", target:"0.00", sA:sA, sB:sB,
            pA:costA.toFixed(2), pB:costB.toFixed(2),
            dollarPnL:"0.00", capGains:"0.00",
            paidDiv: paidDiv.toFixed(2), rcvdDiv: rcvdDiv.toFixed(2),
            centGoal:"0", centRem:"0",
            isWinning:false, entryZ:"0",
            currentZ: 0, sector: '', strategy: strategy
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
        var std = Math.sqrt(sq / (slice.length - 1));
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
function computeBasketMetrics_(legs, histMap) {
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
      var std = Math.sqrt(sq / (slice.length - 1));
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
  var probResult = computeHistoricalProbabilities_(dailyValues, currentValue, rollingZ, totalWeight);

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
function computeHistoricalProbabilities_(dailyValues, refValue, rollingZ, totalWeight) {
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
  var mean90 = (rollingZ['90d'] && !rollingZ['90d'].insufficient) ? rollingZ['90d'].mean : 0;
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
    // Did the spread widen on the very next day? (immediate direction check)
    if (lookForward >= 1) {
      var nextDay = dailyValues[idx + 1];
      if (isAboveMean && nextDay > triggerSpread + tolerance * 0.1) sawWidening = true;
      else if (!isAboveMean && nextDay < triggerSpread - tolerance * 0.1) sawWidening = true;
    }
    if (sawWidening) widenCount++;
  }

  // Average and P75 MAE
  var maeSum = 0;
  var sortedMae = maeValues.slice().sort(function(a, b) { return a - b; });
  for (var i = 0; i < maeValues.length; i++) maeSum += maeValues[i];
  var avgMae = maeSum / maeValues.length;
  var p75Idx = Math.floor(sortedMae.length * 0.75);
  var p75Mae = sortedMae[Math.min(p75Idx, sortedMae.length - 1)];

  result.badScenario = {
    avgMae: parseFloat(avgMae.toFixed(4)),           // avg max adverse excursion (per share)
    avgMaeDollar: parseFloat((avgMae * totalWeight).toFixed(2)),  // total $ loss
    p75Mae: parseFloat(p75Mae.toFixed(4)),            // 75th percentile MAE (per share)
    p75MaeDollar: parseFloat((p75Mae * totalWeight).toFixed(2)),  // 75th percentile $ loss
    wideningProb: parseFloat((widenCount / triggers.length * 100).toFixed(1)), // % chance of widening next day
    sampleSize: triggers.length
  };

  // ── FEATURE 2: Mean Reversion Win Rates per Window ──
  var windows = [30, 60, 90];
  for (var w = 0; w < windows.length; w++) {
    var n = windows[w];
    var wKey = n + 'd';
    var wMean = (rollingZ[wKey] && !rollingZ[wKey].insufficient) ? rollingZ[wKey].mean : mean90;
    // Mean reversion tolerance: spread counts as "touched mean" if it gets within 10% of distance to mean
    var distToMean = Math.abs(refValue - wMean);
    var meanTolerance = distToMean * 0.10;
    if (meanTolerance < 0.01) meanTolerance = 0.01; // minimum floor

    var wins = 0, eligible = 0;
    for (var t = 0; t < triggers.length; t++) {
      var idx = triggers[t];
      var remaining = len - idx - 1;
      if (remaining < Math.floor(n * 0.5)) continue; // need at least half the window to be meaningful
      eligible++;
      var lookAhead = Math.min(n, remaining);
      var touched = false;
      for (var f = 1; f <= lookAhead; f++) {
        if (Math.abs(dailyValues[idx + f] - wMean) <= meanTolerance) {
          touched = true;
          break;
        }
      }
      if (touched) wins++;
    }

    result.winRates[wKey] = {
      rate: eligible > 0 ? parseFloat((wins / eligible * 100).toFixed(1)) : 0,
      wins: wins,
      eligible: eligible
    };
  }

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
function computeHistoricalProbabilitiesWide_(dailyValues, refValue, rollingZ, totalWeight) {
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

  var mean90 = (rollingZ['90d'] && !rollingZ['90d'].insufficient) ? rollingZ['90d'].mean : 0;
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
  var widenCount = 0;
  for (var t = 0; t < triggers.length; t++) {
    var idx = triggers[t];
    var triggerSpread = dailyValues[idx];
    var maxAdverse = 0;
    var sawWidening = false;
    var lookForward = Math.min(30, len - idx - 1);
    for (var f = 1; f <= lookForward; f++) {
      var futureSpread = dailyValues[idx + f];
      var excursion;
      if (isAboveMean) { excursion = futureSpread - triggerSpread; }
      else { excursion = triggerSpread - futureSpread; }
      if (excursion > maxAdverse) maxAdverse = excursion;
    }
    maeValues.push(maxAdverse);
    if (lookForward >= 1) {
      var nextDay = dailyValues[idx + 1];
      if (isAboveMean && nextDay > triggerSpread + tolerance * 0.1) sawWidening = true;
      else if (!isAboveMean && nextDay < triggerSpread - tolerance * 0.1) sawWidening = true;
    }
    if (sawWidening) widenCount++;
  }

  var maeSum = 0;
  var sortedMae = maeValues.slice().sort(function(a, b) { return a - b; });
  for (var i = 0; i < maeValues.length; i++) maeSum += maeValues[i];
  var avgMae = maeSum / maeValues.length;
  var p75Idx = Math.floor(sortedMae.length * 0.75);
  var p75Mae = sortedMae[Math.min(p75Idx, sortedMae.length - 1)];

  result.badScenario = {
    avgMae: parseFloat(avgMae.toFixed(4)),
    avgMaeDollar: parseFloat((avgMae * totalWeight).toFixed(2)),
    p75Mae: parseFloat(p75Mae.toFixed(4)),
    p75MaeDollar: parseFloat((p75Mae * totalWeight).toFixed(2)),
    wideningProb: parseFloat((widenCount / triggers.length * 100).toFixed(1)),
    sampleSize: triggers.length
  };

  var windows = [30, 60, 90];
  for (var w = 0; w < windows.length; w++) {
    var n = windows[w];
    var wKey = n + 'd';
    var wMean = (rollingZ[wKey] && !rollingZ[wKey].insufficient) ? rollingZ[wKey].mean : mean90;
    var distToMean = Math.abs(refValue - wMean);
    var meanTolerance = distToMean * 0.10;
    if (meanTolerance < 0.01) meanTolerance = 0.01;

    var wins = 0, eligible = 0;
    for (var t = 0; t < triggers.length; t++) {
      var idx = triggers[t];
      var remaining = len - idx - 1;
      if (remaining < Math.floor(n * 0.4)) continue; // 40% forward data (relaxed)
      eligible++;
      var lookAhead = Math.min(n, remaining);
      var touched = false;
      for (var f = 1; f <= lookAhead; f++) {
        if (Math.abs(dailyValues[idx + f] - wMean) <= meanTolerance) { touched = true; break; }
      }
      if (touched) wins++;
    }
    result.winRates[wKey] = {
      rate: eligible > 0 ? parseFloat((wins / eligible * 100).toFixed(1)) : 0,
      wins: wins,
      eligible: eligible
    };
  }

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
      var aggBadAvg = 0, aggBadP75 = 0, aggWidenProb = 0;
      var aggWeightSum = 0; // for renormalization after excluding missing-ticker pairs
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

        // Accumulate weighted aggregates
        var w = def.weight;
        aggWeightSum += w;
        if (z30 != null) aggZ30 += w * z30;
        if (z60 != null) aggZ60 += w * z60;
        if (z90 != null) aggZ90 += w * z90;
        if (wr30 != null) aggWR30 += w * wr30;
        if (wr60 != null) aggWR60 += w * wr60;
        if (wr90 != null) aggWR90 += w * wr90;
        if (ep30Dollar != null) aggEP30 += ep30Dollar;
        if (ep60Dollar != null) aggEP60 += ep60Dollar;
        if (ep90Dollar != null) aggEP90 += ep90Dollar;
        if (badAvgDollar != null) aggBadAvg += badAvgDollar;
        if (badP75Dollar != null) aggBadP75 += badP75Dollar;
        if (bs && bs.widenProb != null) aggWidenProb += w * bs.widenProb;

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
          widenProb: bs ? bs.widenProb : null,
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
        // Weighted average Z-scores
        z30: parseFloat((aggZ30 / nw).toFixed(2)),
        z60: parseFloat((aggZ60 / nw).toFixed(2)),
        z90: parseFloat((aggZ90 / nw).toFixed(2)),
        // Weighted average win rates
        wr30: parseFloat((aggWR30 / nw).toFixed(1)),
        wr60: parseFloat((aggWR60 / nw).toFixed(1)),
        wr90: parseFloat((aggWR90 / nw).toFixed(1)),
        // Summed dollar amounts (no double counting via proportional allocation)
        ep30: parseFloat(aggEP30.toFixed(2)),
        ep60: parseFloat(aggEP60.toFixed(2)),
        ep90: parseFloat(aggEP90.toFixed(2)),
        // Summed bad scenario
        badAvg: parseFloat(aggBadAvg.toFixed(2)),
        badP75: parseFloat(aggBadP75.toFixed(2)),
        widenProb: parseFloat((aggWidenProb / nw).toFixed(1)),
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
  // Check WebCache/Live (intra) + WebCacheCredit (credit computed cache)
  var liveSheets = ['WebCache','WebCacheCredit','Live'];
  for (var s = 0; s < liveSheets.length; s++) {
    var ls = ss.getSheetByName(liveSheets[s]);
    if (!ls || ls.getLastRow() <= 1) continue;
    var live = ls.getDataRange().getValues();
    for (var i = 1; i < live.length; i++) {
      if (cleanId(live[i][0]) === cleanTradeId) {
        curZ = live[i][12]; // M: Z-Score
        realId = live[i][0];
        break;
      }
    }
    if (curZ !== 0) break;
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
      // Proportionally split div amounts between closed and remaining portions
      var totalSize = Math.abs(sA) + Math.abs(sB);
      var closedSize = closedA + closedB;
      var divRatio = totalSize > 0 ? closedSize / totalSize : 1;
      var closedPaidDiv = Math.round(totalPaidDiv * divRatio * 100) / 100;
      var closedRcvdDiv = Math.round(totalRcvdDiv * divRatio * 100) / 100;
      // Look up live exit prices
      var live = getLivePairData_(ss, pairId);
      var exitA = live ? live.priceA : costA;
      var exitB = live ? live.priceB : costB;
      var exitZ = live ? live.z : 0;
      // PnL: capital gains on closed portion + proportional dividends
      var capGains = ((exitA - costA) * closedA * signA) + ((exitB - costB) * closedB * signB);
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
  var cacheSheets = ['WebCache', 'WebCacheCredit', 'Live'];
  for (var s = 0; s < cacheSheets.length; s++) {
    var ls = ss.getSheetByName(cacheSheets[s]);
    if (!ls || ls.getLastRow() <= 1) continue;
    var data = ls.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (cleanId(data[i][0]) === cleanTarget) {
        curZ = parseFloat(data[i][12]) || 0;
        curSpread = parseFloat(data[i][5]) || 0;
        curMean = parseFloat(data[i][10]) || 0;
        break;
      }
    }
    if (curZ !== 0) break;
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
function analyzeSinglePair_(tA, tB, priceA, priceB, currentZ) {
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
    var metrics = computeBasketMetrics_(legs, histMap);
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

      var wideProb = computeHistoricalProbabilitiesWide_(dailyVals, currentValue, metrics.rollingZ, metrics.totalWeight || 100);
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
        var analysis = analyzeSinglePair_(a.tA, a.tB, a.pA || parseFloat(a.spr), a.pB || 0, parseFloat(a.z));
        if (analysis && !analysis.error && analysis.metrics) {
          var prob = analysis.metrics.probabilities || {};
          var wr30 = (prob.winRates && prob.winRates['30d']) ? prob.winRates['30d'].rate : null;
          var wr60 = (prob.winRates && prob.winRates['60d']) ? prob.winRates['60d'].rate : null;
          var wr90 = (prob.winRates && prob.winRates['90d']) ? prob.winRates['90d'].rate : null;
          var bs = prob.badScenario || {};
          results.push({
            id: a.id, tA: a.tA, tB: a.tB, mode: a._mode,
            z: parseFloat(a.z), expProfit: parseFloat(a.expProfit),
            wr30: wr30, wr60: wr60, wr90: wr90,
            avgMae: bs.avgMae || null, p75Mae: bs.p75Mae || null,
            wideningProb: bs.wideningProb || null, triggers: prob.triggers || 0,
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
      sheet.getRange(1,1,1,16).setValues([['PairID','TickerA','TickerB','Mode','Z','ExpProfit','WR30','WR60','WR90','AvgMAE','P75MAE','WidenProb','Triggers','EP30','EP60','UpdatedAt']]);
    } else {
      if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 16).clearContent();
    }
    if (results.length > 0) {
      var rows = results.map(function(r){
        return [r.id, r.tA, r.tB, r.mode, r.z, r.expProfit, r.wr30, r.wr60, r.wr90, r.avgMae, r.p75Mae, r.wideningProb, r.triggers, r.ep30, r.ep60, r.ts];
      });
      sheet.getRange(2, 1, rows.length, 16).setValues(rows);
    }
    Logger.log('Screener: processed ' + results.length + '/' + top.length + ' pairs in ' + ((new Date().getTime()-startTime)/1000).toFixed(1) + 's');
  } catch(e) {
    Logger.log('Screener error: ' + e);
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
      results.push({
        id: r[0], tA: r[1], tB: r[2], mode: r[3],
        z: r[4], expProfit: r[5],
        wr30: r[6], wr60: r[7], wr90: r[8],
        avgMae: r[9], p75Mae: r[10], wideningProb: r[11],
        triggers: r[12], ep30: r[13], ep60: r[14],
        ts: r[15]
      });
    }
    return results;
  } catch(e) { return []; }
}
