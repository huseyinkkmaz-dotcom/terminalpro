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
    setupMacroSheet();
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
        var costA = parseMoney(openData[j][2]);
        var costB = parseMoney(openData[j][3]);
        var sA = parseMoney(openData[j][4]);
        var sB = parseMoney(openData[j][5]);
        if (pair) {
          if (pair[1] && String(pair[1]).length > 1) tA_Name = String(pair[1]);
          if (pair[2] && String(pair[2]).length > 1) tB_Name = String(pair[2]);
          var liveSpr = parseFloat(pair[5])||0;
          var meanTarget = parseFloat(pair[10])||0;
          var livePriceA = parseFloat(pair[3])||0;
          var livePriceB = parseFloat(pair[4])||0;
          var pnl = ((livePriceA-costA)*sA) + ((livePriceB-costB)*sB);
          results.push({
            id: displayId, tA: tA_Name, tB: tB_Name,
            spr: liveSpr.toFixed(2), target: meanTarget.toFixed(2),
            sA: sA, sB: sB, pA: costA.toFixed(2), pB: costB.toFixed(2),
            dollarPnL: pnl.toFixed(2),
            centGoal: (Math.abs((costA-costB)-meanTarget)*100).toFixed(0),
            centRem: (Math.abs(liveSpr-meanTarget)*100).toFixed(0),
            isWinning: pnl > 0, entryZ: openData[j][1]
          });
        } else {
          results.push({
            id: displayId+" [WAITING]", tA: tA_Name, tB: tB_Name,
            spr:"0.00", target:"0.00", sA:sA, sB:sB,
            pA:costA.toFixed(2), pB:costB.toFixed(2),
            dollarPnL:"0.00", centGoal:"0", centRem:"0",
            isWinning:false, entryZ:"0"
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
      output.push({ id:info.id, closeDate: r[7] instanceof Date ? r[7].toLocaleDateString() : "---", pnl: r[8]||0 });
    }
    return output.reverse();
  } catch(e) { return []; }
}
// ============================================================
// WRITE OPERATIONS
// ============================================================
function saveTradeToSheet(trade) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('OpenTrades');
  var curZ = 0;
  var cleanTradeId = cleanId(trade.id);
  var realId = trade.id;
  // Check Live (intra formulas) + WebCacheCredit (credit computed cache)
  var sheets = ['Live','WebCacheCredit'];
  for (var s = 0; s < sheets.length; s++) {
    var ls = ss.getSheetByName(sheets[s]);
    if (!ls) continue;
    var live = ls.getDataRange().getValues();
    for (var i = 1; i < live.length; i++) {
      if (cleanId(live[i][0]).includes(cleanTradeId) || cleanTradeId.includes(cleanId(live[i][0]))) {
        curZ = live[i][12]; // M: Z-Score
        realId = live[i][0];
        break;
      }
    }
    if (curZ !== 0) break;
  }
  sheet.appendRow([realId, curZ, trade.priceA, trade.priceB, trade.sizeA, trade.sizeB, new Date()]);
  return true;
}
function closeTradeInSheet(id) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('OpenTrades');
  var data = sheet.getDataRange().getValues();
  var cleanTarget = cleanId(id);
  for (var i = data.length-1; i >= 1; i--) {
    var rowId = cleanId(data[i][0]);
    if (rowId.includes(cleanTarget) || cleanTarget.includes(rowId)) { sheet.deleteRow(i+1); break; }
  }
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
    sheet.getRange(1, 1, 1, 4).setValues([['PairID', 'AddedDate', 'AddedZ', 'Mode']]);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  }
  // Check for duplicate
  var cleanTarget = cleanId(id);
  if (sheet.getLastRow() > 1) {
    var existing = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < existing.length; i++) {
      if (cleanId(existing[i][0]) === cleanTarget) return; // already exists
    }
  }
  // Look up current Z-score
  var curZ = 0;
  var sheets = ['WebCache', 'WebCacheCredit', 'Live'];
  for (var s = 0; s < sheets.length; s++) {
    var ls = ss.getSheetByName(sheets[s]);
    if (!ls || ls.getLastRow() <= 1) continue;
    var data = ls.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (cleanId(data[i][0]) === cleanTarget) { curZ = parseFloat(data[i][12]) || 0; break; }
    }
    if (curZ !== 0) break;
  }
  sheet.appendRow([id, new Date(), curZ, mode]);
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
    var watchRows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();

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
