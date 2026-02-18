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
      // Failsafe: check WebCache first (fast snapshot), then Live (GOOGLEFINANCE formulas)
      var cacheName = (mode === 'credit') ? 'WebCacheCredit' : 'WebCache';
      var liveName = (mode === 'credit') ? 'CreditLive' : 'Live';
      var ss = SpreadsheetApp.getActive();
      var targetSheet = ss.getSheetByName(cacheName);
      if (!targetSheet || targetSheet.getLastRow() <= 1) {
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
  // Prefer WebCache (static snapshot, fast) over Live (GOOGLEFINANCE formulas)
  var cacheName = (mode === 'credit') ? 'WebCacheCredit' : 'WebCache';
  var liveName = (mode === 'credit') ? 'CreditLive' : 'Live';

  try {
    var ss = SpreadsheetApp.getActive();
    var liveSheet = ss.getSheetByName(cacheName);
    if (!liveSheet || liveSheet.getLastRow() <= 1) {
      liveSheet = ss.getSheetByName(liveName);
    }
    if (!liveSheet) return [];
    var data = liveSheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    // Load AlertsLog for trend ribbons
    var logData = [];
    var logSheet = ss.getSheetByName('AlertsLog');
    if (logSheet) {
      var lastRow = logSheet.getLastRow();
      if (lastRow > 1) {
        var startRow = Math.max(2, lastRow - 2000);
        logData = logSheet.getRange(startRow, 1, lastRow - startRow + 1, 4).getValues();
      }
    }
    // Load ZScoreAge for age column (Task 1)
    var ageMap = {};
    var ageSheet = ss.getSheetByName('ZScoreAge');
    if (ageSheet) {
      var ageData = ageSheet.getDataRange().getValues();
      var now = new Date();
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
    var output = [];
    var _diag = {totalRows: data.length - 1, noId: 0, noPrice: 0, lowHist: 0, noCoupon: 0, lowZ: 0, passed: 0};
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rawId = row[0];
      if (!rawId) { _diag.noId++; continue; }
      // FILTER: valid prices
      var priceA = parseFloat(row[3]) || 0;
      var priceB = parseFloat(row[4]) || 0;
      if (priceA <= 0 || priceB <= 0) { _diag.noPrice++; continue; }

      // FILTER: history >= 60 trading days
      var histCount = parseFloat(row[16]) || 0;
      if (histCount < 60) { _diag.lowHist++; continue; }

      // FILTER: coupon must exist (exclude variable/reset)
      var couponA = row[8];
      var couponB = row[9];
      if (couponA === "" || couponA === null || couponA === undefined ||
          couponB === "" || couponB === null || couponB === undefined) { _diag.noCoupon++; continue; }
      var currentZ = parseFloat(row[12]) || 0;

      // FILTER: |z| >= 1.5
      if (Math.abs(currentZ) < 1.5) { _diag.lowZ++; continue; }
      _diag.passed++;
      var info = parseTickerInfo(rawId);
      var cid = cleanId(rawId);
      // TREND from AlertsLog
      var zHistory = [];
      for (var k = 0; k < logData.length; k++) {
        if (cleanId(logData[k][1]) === cid) {
          var zVal = parseFloat(logData[k][2]);
          if (!isNaN(zVal)) zHistory.push(zVal.toFixed(1));
        }
      }
      var trend = zHistory.slice(-7);
      trend.push(currentZ.toFixed(1));
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
      output.push({
        id: info.id,
        tA: row[1] || info.tA,
        tB: row[2] || info.tB,
        rng: lower.toFixed(2) + " / " + upper.toFixed(2),
        sec: row[15] || "",
        spr: (parseFloat(row[5]) || 0).toFixed(2),
        yA: yA,
        yB: yB,
        z: currentZ.toFixed(2),
        age: ageDays,        // Real age from ZScoreAge
        histDays: histCount, // History depth for reference
        liq: avgLiq,
        curVol: curVol,
        volSpike: volSpike,
        zTrend: trend
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
    if (!credit || credit.getLastRow() <= 1) credit = ss.getSheetByName('CreditLive');
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
  // Check both Live and CreditLive
  var sheets = ['Live','CreditLive'];
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
