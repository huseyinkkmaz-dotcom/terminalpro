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
      var p = parseFloat(histData[h][d]);
      if (!isNaN(p) && p > 0) prices.push(p);
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
      var expectedProfit = mean - currentValue; // positive = portfolio should revert UP
      rollingZ[n + 'd'] = {
        z: parseFloat(z.toFixed(2)),
        mean: parseFloat(mean.toFixed(2)),
        std: parseFloat(std.toFixed(2)),
        expectedProfit: parseFloat(expectedProfit.toFixed(2)),
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
    probabilities: probResult
  };
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
  var result = { triggers: 0, badScenario: null, winRates: {} };
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

  // Find trigger points: days where spread ≈ refValue (±tolerance)
  // Exclude the last day (that's "today") and leave room for at least 5 days of forward data
  var triggers = [];
  var lastTriggerDay = -10; // prevent overlapping triggers (min 5-day gap)
  for (var i = 0; i < len - 5; i++) {
    if (Math.abs(dailyValues[i] - refValue) <= tolerance && (i - lastTriggerDay) >= 5) {
      triggers.push(i);
      lastTriggerDay = i;
    }
  }
  result.triggers = triggers.length;
  if (triggers.length < 3) return result; // not enough samples for meaningful statistics

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

      metrics.entrySpread = parseFloat((entrySpreadDollar / entryWeight).toFixed(2));
      metrics.entryGrossLong = parseFloat(entryGrossLong.toFixed(2));
      metrics.entryGrossShort = parseFloat(entryGrossShort.toFixed(2));
      metrics.entryGrossExposure = parseFloat((entryGrossLong + entryGrossShort).toFixed(2));
      metrics.capitalGains = parseFloat((currentLiveValue - entrySpreadDollar).toFixed(2));
      metrics.netDividends = parseFloat((totalRcvdDiv - totalPaidDiv).toFixed(2));
      metrics.totalPnL = parseFloat((currentLiveValue - entrySpreadDollar + totalRcvdDiv - totalPaidDiv).toFixed(2));

      delete metrics.dailyValuesFull_; // strip internal array before response
      return { mode: 'live', trades: openData.length - 1, legs: legs.length, metrics: metrics };
    }
    else if (mode === 'sandbox') {
      // Parse user-provided legs
      var userLegs = [];
      try { userLegs = JSON.parse(legsJson); } catch(e) { return { mode: 'sandbox', error: 'Invalid legs JSON' }; }
      for (var i = 0; i < userLegs.length; i++) {
        var ul = userLegs[i];
        if (ul.ticker && ul.size && ul.direction) {
          legs.push({ ticker: String(ul.ticker).trim().toUpperCase(), size: Math.abs(parseFloat(ul.size) || 0), direction: parseFloat(ul.direction) > 0 ? 1 : -1 });
        }
      }
      if (legs.length === 0) return { mode: 'sandbox', error: 'No valid legs', metrics: { dailyValues: [] } };

      var metrics = computeBasketMetrics_(legs, histMap);

      // Compute hypothetical entry spread + gross exposure from provided prices
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
      // Normalize to size-weighted average spread
      var hypotheticalWeight = hypotheticalTotalSize / 2;
      if (hypotheticalWeight === 0) hypotheticalWeight = 1;
      metrics.hypotheticalSpread = parseFloat((hypotheticalSpreadDollar / hypotheticalWeight).toFixed(2));
      metrics.hypotheticalGrossLong = parseFloat(hypotheticalGrossLong.toFixed(2));
      metrics.hypotheticalGrossShort = parseFloat(hypotheticalGrossShort.toFixed(2));
      metrics.hypotheticalGrossExposure = parseFloat((hypotheticalGrossLong + hypotheticalGrossShort).toFixed(2));

      // Override expectedProfit in each rolling window to use entry spread:
      // Sandbox user wants to know "if spread reverts to mean from MY ENTRY, how much do I profit?"
      // Default computation uses market spread vs mean (irrelevant to user's entry position).
      var rz = metrics.rollingZ || {};
      for (var wKey in rz) {
        if (rz[wKey] && !rz[wKey].insufficient) {
          rz[wKey].expectedProfit = parseFloat((rz[wKey].mean - metrics.hypotheticalSpread).toFixed(2));
        }
      }

      // Re-run probability engine with entry spread as reference point
      // (default used market spread — sandbox user wants probabilities from THEIR entry level)
      if (metrics.dailyValuesFull_ && metrics.dailyValuesFull_.length > 0) {
        metrics.probabilities = computeHistoricalProbabilities_(
          metrics.dailyValuesFull_, metrics.hypotheticalSpread, rz, metrics.totalWeight
        );
      }

      // Flag which tickers are missing history
      var missing = [];
      for (var i = 0; i < legs.length; i++) {
        if (!histMap[legs[i].ticker]) missing.push(legs[i].ticker);
      }
      if (missing.length > 0) metrics.missingTickers = missing;

      delete metrics.dailyValuesFull_; // strip internal array before response
      return { mode: 'sandbox', legs: legs.length, metrics: metrics };
    }
    else {
      return { error: 'Unknown mode: ' + mode };
    }
  } catch(e) {
    return { error: e.message };
  }
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
