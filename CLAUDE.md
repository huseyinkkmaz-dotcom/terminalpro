# Terminal Pro Vol 2

## What This Is

A real-time **statistical arbitrage dashboard** for **Preferred Stock Pair Trading**. It tracks 100+ pairs using 90-day mean reversion Z-scores, yields, liquidity, volume spike detection, and upcoming ex-dividend dates.

**Three strategies:**
- **Intra-Company Pairs** — Same issuer, different series (e.g., BAC-B vs BAC-M)
- **Credit Rating Arbitrage** — Different issuers, same credit rating (e.g., BBB vs BBB cross-company)
- **Macro Valuation** — Individual preferreds vs. US Treasuries (yield spread Z-scores against 5 benchmarks: US2Y, US5Y, US7Y, US10Y, US30Y)

## Architecture

```
Google Sheets (Data Layer)
    ├── Pairs sheet         → Manual: intra-company pair definitions
    ├── Master sheet        → Manual: ticker, yield, coupon, credit rating
    ├── Levels sheet        → Auto: 90-day stats via GOOGLEFINANCE formulas
    ├── Live sheet          → Auto: real-time prices, Z-scores, liquidity
    ├── CreditPairs sheet   → Auto: generated inter-company pairs by rating
    ├── TickerData sheet    → Auto: per-ticker GOOGLEFINANCE prices/volumes
    ├── TickerHistory sheet → Auto: per-ticker 90-day historical data
    ├── WebCache sheet      → Auto: static snapshot of Live (updated every 10 min)
    ├── WebCacheCredit sheet→ Auto: computed credit pair cache (from TickerData + TickerHistory)
    ├── DivDates sheet      → Auto: ex-dividend dates fetched from Yahoo Finance
    ├── ZScoreAge sheet     → Auto: timestamps when Z-scores first cross ±1.8
    ├── AlertsLog sheet     → Auto: hourly Z-score snapshots (trend ribbons)
    ├── OpenTrades sheet    → Trade journal (active)
    ├── ClosedTrades sheet  → Trade journal (closed)
    ├── MacroData sheet     → Auto: US10Y, TLT, PFF macro indicators
    ├── TreasuryHist sheet  → Auto: 90-day FRED treasury yields (US2Y-US30Y)
    ├── MacroCalc sheet     → Auto: computed macro valuations per ticker
    ├── MacroCache sheet    → Auto: static snapshot of MacroCalc
    ├── BasketCache sheet   → Auto: pre-computed portfolio basket analytics
    └── ScreenerCache sheet → Auto: pre-computed probability analysis for top alert pairs

Google Apps Script (API Backend — ~7000 lines)
    ├── Code.gs             → doGet/doPost routing, data reading, trade operations, backtest, sizer, screener
    └── SetupDashboard.gs   → Sheet builder, credit pair generator, triggers, dividend fetcher

Vercel (Frontend — password-gated)
    ├── api/index.js        → Serverless auth gate (HMAC cookie, login page)
    ├── _app.html           → Single-file dashboard (HTML + CSS + JS, ~4400 lines)
    └── vercel.json         → Routes all traffic through auth gate
```

## Key File Descriptions

### `gas/Code.gs` — API Backend (V25, ~7000 lines)

**Security layer** — All requests pass through `verifyAuth_(e)` → `checkRateLimit_(action)` → `auditLog_()` for write actions. All write actions wrapped in `LockService.getDocumentLock()`.

**doGet()** routes 35+ actions with `?action=X&mode=intra|credit`:
- **Data**: `getData`, `getMacroValuation`, `getBasketAnalytics`, `getPortfolioAnalytics`, `getScreenerData`, `getRegimeData`, `getCorrelationMonitor`
- **Analysis**: `analyzePair` (supports custom `windows` param), `sweepPair`, `getOptimalSweep`, `getSensitivityHeatmap`, `getBacktestResults`, `getPositionSizing`
- **Trades**: `saveTrade`, `closeTrade`, `partialClose`, `addDividend`, `saveNote`, `getTradeEvents`, `getTradeAlerts`, `getExitAlerts`, `saveExitParams`, `backfillTargets`
- **Portfolio**: `getModelPortfolios`, `runModelPortfolios`, `getJournalAnalytics`, `getDividendCapture`
- **Other**: `getWatchlist`, `saveWatchlist`, `removeWatchlist`, `clearHistory`, `deleteClosedTrade`, `saveNotificationSettings`, `getNotificationSettings`

**doPost()** routes: `saveTrade`, `closeTrade`, `partialClose`, `addDividend` (JSON payload).

- **WebCache failsafe** — API prefers `WebCache` / `WebCacheCredit` (static snapshots) over `Live` / `CreditLive` (GOOGLEFINANCE formulas). Faster responses, decoupled from formula recalculation.
- **getAlertData(mode)** reads WebCache (or Live fallback) sheet. Filters:
  - Skips pairs with missing/zero prices
  - Skips pairs with < 40 trading days of history (HistCount column Q/index 16)
  - Skips pairs where either ticker has empty Coupon Yield (columns I,J / index 8,9)
  - Only returns pairs with |Z-Score| >= 1.8
  - Enriches each pair with `divDate`, `divLeg`, `halfLife`, `halfLifeValid`, `hedgeRatio`, `callRiskA`, `callRiskB`, `pA`, `pB`
- **getOpenTrades()** merges Live + CreditLive via hash-map lookup, returns PaidDiv/ReceivedDiv totals, `currentZ`, `sector`
- **saveTradeToSheet()** — validates positive prices, non-zero size, $10k price / 100k share caps. Stamps `tradeGroupId` for partial-close grouping.
- **addDividendToTrade(id, type, amount)** accumulates dividend cash flows on open positions. `type` is `paid` (short leg owes) or `received` (long leg earns). Amounts are additive (accumulator pattern).
- **PnL formula** — `Net PnL = Capital Gains + Received Div - Paid Div`. Applied in both open trade display and closed trade history.
- **getBasketAnalytics()** computes portfolio-level metrics: weighted avg Z, rolling 30/60/90-day Z-scores, sector concentration, strategy mix, max drawdown, last 30 daily basket values for sparkline.
- **analyzeSinglePair_(tA, tB, priceA, priceB, currentZ, windows, mode)** — builds 2-leg sandbox pair, runs probability engine, returns win rates (30/60d), expected profit, bad scenario, rolling Z-scores.
- **runBacktest_(zThreshold, exitZ, maxHold, mode)** — full backtest engine with HL-dynamic hold cap (`pairMaxHold = min(3×HL, maxHold)`), exit reasons: `Z_REVERT`, `HL_CAP`, `MAX_HOLD`. Soft-blacklisted tickers included for survivorship bias avoidance.
- **getPositionSizing_(maxLossPerTrade)** — HL-adjusted sizing (`hlMult = clamp(21/HL, 0.5, 2.0)`), optional Kelly sizing from ScreenerCache.
- **checkExitSignals_(params)** — 6 exit triggers: STATISTICAL_TARGET, TIME_STOP, STOP_LOSS, PARTIAL_PROFIT, FULL_TARGET, STAGNATION. Supports per-trade parameter overrides from OpenTrades cols J-O.
- **runNightlyScreener()** — daily 9 AM trigger, pre-computes probability analysis for top 20 alert pairs to `ScreenerCache`.
- **runModelPortfolioGenerator()** — generates 3 model portfolios (Aggressive/Balanced/Diversified) from candidate pool.
- **Notification settings** — credentials stored in `PropertiesService.getScriptProperties()` (not in sheets). Telegram send not yet active.

**Key statistical functions** (line numbers):
- `getAdfCriticalValues_(n)` — line 38 (MacKinnon 1994 response surface)
- `adfTest_(series)` — line 76 (ADF with min 30 observations)
- `ouHalfLife_(series)` — line 189 (returns `{halfLife, lambda, isValid}`)
- `engleGrangerTest_(seriesA, seriesB)` — line 238 (2-step cointegration)
- `computeHedgeRatio_(seriesA, seriesB)` — line 306 (OLS β, clamped [0.3, 3.0])

### `gas/SetupDashboard.gs` — Setup & Automation (~2400 lines)

**Batched setup system** — uses `PropertiesService` to save/resume progress across GAS timeouts.

- **setupAllBatched()** — 7-phase execution:
  - Phase 0: Intra Levels
  - Phase 1: Intra Live
  - Phase 2: Supporting sheets (OpenTrades, ClosedTrades, AlertsLog, ZScoreAge, DivDates)
  - Phase 3: Generate credit pairs
  - Phase 4: Credit Levels
  - Phase 5: Credit Live
  - Phase 6: Install triggers
- **generateCreditPairs()** reads Master, groups by credit rating, generates inter-company combinations
  - Uses union-find on Pairs sheet to detect same-company tickers (handles both hyphenated like BAC-M and non-hyphenated like AGNCP)
  - No pair cap — generates all valid combinations
- **createAutoTrigger()** installs all recurring triggers at once:
  - Every 10 min: `updateLivePrices()` — snapshots Live/CreditLive values to WebCache/WebCacheCredit
  - Hourly: `snapshotZScores()` — logs to AlertsLog + updates ZScoreAge
  - Daily at 5 AM: `dailyCreditRefresh()` — regenerates credit pairs from Master
  - Daily at 6 AM: `fetchDividendDates()` — refreshes ex-dividend dates
  - Daily at 8 AM: `updateBasketAnalytics()` — pre-computes portfolio basket Z-scores to BasketCache
  - Daily at 9 AM: `runNightlyScreener()` — pre-computes probability analysis for top 20 alert pairs to ScreenerCache
- **fetchDividendDates()** — two-phase Yahoo Finance fetch:
  - Phase 1: v7/finance/quote API (50 tickers per batch, requires crumb auth)
  - Phase 2: v8/finance/chart fallback (no auth, uses 1-year history for dividend events)
  - Projects quarterly dividends forward if last date is in the past (+91 days)
  - 20-hour cache window to avoid rate limiting
- **updateLivePrices()** — reads computed values from Live/CreditLive, writes static snapshots to WebCache/WebCacheCredit every 10 minutes

### `frontend/_app.html` — Dashboard UI (~4400 lines)

Single-file dashboard (HTML + CSS + JS). Served through Vercel auth gate — never directly accessible.

**Key constants** (set before deploying):
- `GAS_URL` — line 1078 (deployed GAS URL)
- `API_KEY` — line 1081 (must match GAS Script Property)
- CSP meta tag — line 6 (restricts `connect-src` to `script.google.com`)

**Security helpers**:
- `escHtml()` — line 1157 (HTML entity encoding for user-facing interpolation)
- `escAttr()` — line 1158 (attribute-safe escaping for onclick strings)
- `cleanIdFE()` — line 1156 (replaces `|` with `_`, strips non-alphanumeric except `-`)
- `apiGet()` — line 1413 (retry logic: 2 retries, exponential backoff 2s/4s, on 5xx/network errors)

**8 tabs**: Alerts (line 712), Portfolio (line 733), Macro (line 797), Risk Lab (line 823), History (line 856), Dividend Capture (line 904, hidden), Models (line 921), Backtest (line 948)

**Strategy pill switch**: "Intra-Company" (blue accent) / "Credit Arb" (purple accent)

**Alerts tab features**:
- Quality column with ADF badges, HL:Xd display, CALLABLE warnings, call risk tracking
- Z-score trend ribbons from AlertsLog (hourly snapshots)
- Volume spike tags, sortable columns
- **Inline Analyze** — gear icon per alert → expands row with 30/60d win rates, expected profit, bad scenario. Checks screener cache first, falls back to live `?action=analyzePair`.
- **Inline Sizer** — HL-adjusted position sizing with `hlMult = clamp(21/HL, 0.5, 2.0)`. Shows HL Mult card (line ~3085). `recalcInlineSizer()` (line 3155) and `buildInlineSizerCards()` (line 3132) also use hlMult.
- **Pre-Screened Badges** — "XX% WR" badges from nightly screener cache
- **CSV Export** — `csvCell()`, `csvRows()`, `downloadCsv()` (lines 1160-1162). UTF-8 BOM for Excel.
- Div Date column with color coding: Red (<=7d), Yellow (<=30d), Gray (>30d), with "EST" badge for projected dates

**Trade entry modal** (`openModal()` — line 1438):
- **β auto-fill for credit pairs**: reads `hedgeRatio` from `_alertDataMap`. For β≠1.0, shows yellow `betaHint` div (line 571) and auto-fills SIZE B = round(SIZE A × β) on input.

**Portfolio tab**:
- Portfolio Health Bar with weighted avg Z, rolling 30/60/90d Z-scores, sector concentration, strategy mix, max drawdown
- Color-coded: Green (|Z|<0.5), Blue (mild), Yellow (elevated), Red (|Z|>=1.8)
- Data staleness indicator in header ("Data: 5m ago") with green/yellow/red coding
- Concentration warning when >50% in one sector
- Portfolio sizer (`runSizer()` line 2945 / `renderSizer()` line 3009) with HL Mult column

**History tab**:
- **Group Partials toggle** (line 878, default ON) — `groupHistoryByTrade()` (line 2047) consolidates partial closes into one display row. Groups by `tradeGroupId` with composite key fallback for legacy rows. Stats count groups not rows. Purple expand/collapse for individual partials.
- CSV Export button for history data

**Other tabs**:
- Macro: Treasury yield comparison, macro valuation (`loadMacroValuation()` line 2848), rate regime badge
- Risk Lab: Cross-pair correlation, portfolio analytics
- Models: 3 model portfolios (Aggressive/Balanced/Diversified)
- Backtest: Full backtest engine (`runBacktest()` line 3486), optimal sweep (`runOptimalSweep()` line 3416)

## Live Sheet Column Map (24 columns, 0-indexed)

Both `Live` and `CreditLive` share this layout:

| Index | Col | Name | Source |
|-------|-----|------|--------|
| 0 | A | PairID | Pairs/CreditPairs |
| 1 | B | TickerA | Pairs/CreditPairs |
| 2 | C | TickerB | Pairs/CreditPairs |
| 3 | D | PriceA | GOOGLEFINANCE live |
| 4 | E | PriceB | GOOGLEFINANCE live |
| 5 | F | Spread | =D-E |
| 6 | G | YieldA | INDEX/MATCH from Master!D |
| 7 | H | YieldB | INDEX/MATCH from Master!D |
| 8 | I | CouponA | INDEX/MATCH from Master!C |
| 9 | J | CouponB | INDEX/MATCH from Master!C |
| 10 | K | Mean | INDEX/MATCH from Levels!B |
| 11 | L | StDev | INDEX/MATCH from Levels!C |
| 12 | M | Z-Score | =(Spread-Mean)/StDev |
| 13 | N | Lower | INDEX/MATCH from Levels!D |
| 14 | O | Upper | INDEX/MATCH from Levels!E |
| 15 | P | Sector | Pairs/CreditPairs |
| 16 | Q | HistCount | INDEX/MATCH from Levels!F |
| 17 | R | AvgLiqA | GOOGLEFINANCE volumeavg |
| 18 | S | AvgLiqB | GOOGLEFINANCE volumeavg |
| 19 | T | AvgLiq | =(R+S)/2 |
| 20 | U | CurVolA | GOOGLEFINANCE volume |
| 21 | V | CurVolB | GOOGLEFINANCE volume |
| 22 | W | CurVol | =(U+V)/2 |
| 23 | X | VolSpike | =IF(CurVol > 1.5*AvgLiq) |

## DivDates Sheet (auto-populated)

| A: Ticker | B: NextDivDate | C: LastFetched | D: IsEstimated |
|-----------|---------------|----------------|----------------|
| BAC-B | 2026-03-15 | 2026-02-20T06:00:00 | CONFIRMED |
| PSA-F | 2026-06-15 | 2026-02-20T06:00:00 | ESTIMATED |

Populated by `fetchDividendDates()` via Yahoo Finance. 20-hour cache — tickers fetched within the last 20 hours are skipped. `IsEstimated`: `CONFIRMED` (direct from Yahoo) vs `ESTIMATED` (+91 day projection). Frontend shows yellow "EST" badge next to projected dates.

## BasketCache Sheet (auto-populated)

| A: Key | B: Value | C: UpdatedAt |
|--------|----------|-------------|
| trades | 12 | 2026-03-05T08:00:00 |
| weightedAvgZ | -0.42 | 2026-03-05T08:00:00 |
| rollingZ_30d_z | -0.38 | 2026-03-05T08:00:00 |
| sector_Finance | 45.2 | 2026-03-05T08:00:00 |

Key-value store for pre-computed basket analytics. Updated daily at 8 AM by `updateBasketAnalytics()` trigger. The `getBasketAnalytics()` endpoint in Code.gs computes fresh values on each API call (using TickerHistory for rolling Z); BasketCache serves as a daily snapshot backup.

## ScreenerCache Sheet (auto-populated)

| A: PairID | B: TickerA | C: TickerB | D: Mode | E: Z | F: ExpProfit | G: WR30 | H: WR60 | I: WR90 | J: AvgMAE | K: P75MAE | L: WidenProb | M: Triggers | N: EP30 | O: EP60 | P: UpdatedAt |
|-----------|-----------|-----------|---------|------|-------------|---------|---------|---------|----------|----------|-------------|------------|---------|---------|-------------|

Pre-computed probability analysis for top 20 alert pairs. Updated daily at 9 AM by `runNightlyScreener()` trigger. Frontend loads on page init and displays win rate badges + feeds composite score.

## Critical Rules When Editing

1. **Preserve ticker hyphens** — Preferred stock tickers use hyphens (PSA-F, BAC-M). Never strip them. The `cleanId()` function removes hyphens only for matching purposes, never for display or storage.
2. **Column indices matter** — Code.gs uses 0-indexed column references to read Live sheet arrays. If you add/remove/reorder columns in `buildLiveSheet_()`, you MUST update every index reference in `getAlertData()`, `getOpenTrades()`, `saveTradeToSheet()`, and `snapshotZScores()`.
3. **Two parallel sheet systems** — Intra-company uses `Pairs → Levels → Live`. Credit arb uses `CreditPairs → CreditLevels → CreditLive`. They share the same 24-column layout but reference different source sheets. The `buildLiveSheet_()` function is parameterized for this.
4. **WebCache layer** — API reads from WebCache/WebCacheCredit (static values), not directly from Live/CreditLive (formulas). The `updateLivePrices()` trigger snapshots formula results every 10 minutes. If you change Live sheet structure, WebCache must match.
5. **GAS deployment** — Every code change requires: Manage Deployments → New Version → Deploy. The URL stays the same but the version must increment.
6. **GOOGLEFINANCE limits** — ~1000 GOOGLEFINANCE calls per sheet. Each pair uses ~6 calls in Live (price×2, volumeavg×2, volume×2) + 2 in Levels (historical). With 500+ credit pairs, sheets may load slowly (30-60s).
7. **Frontend is a single file** — All HTML, CSS, and JS live in `_app.html`. No build step. Deploy via `vercel --prod` from the `frontend/` directory. The file is served through the Vercel auth gate (`api/index.js`), never directly.
8. **ZScoreAge tracking** — The hourly trigger manages this. If |z| >= 1.8 and no timestamp exists → creates one. If |z| < 1.8 and timestamp exists → deletes it. Age = days since first crossing.
9. **DivDates fetch** — Yahoo Finance has rate limits. The 20-hour cache in LastFetched prevents hammering. Phase 1 (batch API) handles most tickers; Phase 2 (chart fallback) catches the rest.

## OpenTrades Sheet Column Map (15 columns, 0-indexed)

| Index | Col | Name | Description |
|-------|-----|------|-------------|
| 0 | A | PairID | Pair identifier |
| 1 | B | EntryZ | Z-score at entry |
| 2 | C | CostA | Entry price leg A |
| 3 | D | CostB | Entry price leg B |
| 4 | E | SizeA | Quantity leg A (+ long, - short) |
| 5 | F | SizeB | Quantity leg B (+ long, - short) |
| 6 | G | Timestamp | Trade open date |
| 7 | H | PaidDiv | Accumulated dividends paid (short leg) |
| 8 | I | ReceivedDiv | Accumulated dividends received (long leg) |
| 9 | J | TargetExitZ | Optimal exit Z-score from model portfolio sweep |
| 10 | K | ProfitCapturePct | Portfolio profit capture % (e.g., 60) |
| 11 | L | TargetPnL | Dollar profit target = theoretical max x capture% |
| 12 | M | PartialAtPct | % of TargetPnL for first partial (e.g., 50) |
| 13 | N | SourcePortfolio | Which model portfolio (e.g., "MP#1", "backtest") |
| 14 | O | MaxHoldDays | Expected avg hold days from sweep |

**Exit target columns (J-O)** are stamped at trade entry when entering from a model portfolio or backtest signal. Empty for manually entered trades. Per-trade targets override global exit params in `checkExitSignals_`.

## ClosedTrades Sheet Column Map (21 columns, 0-indexed)

| Index | Col | Name | Description |
|-------|-----|------|-------------|
| 0-6 | A-G | Same as OpenTrades | PairID through OpenDate |
| 7 | H | CloseDate | Trade close date |
| 8 | I | PnL | Net PnL (CapGains + RcvdDiv - PaidDiv) |
| 9 | J | ExitPriceA | Exit price leg A |
| 10 | K | ExitPriceB | Exit price leg B |
| 11 | L | ExitZ | Z-score at exit |
| 12 | M | CloseType | FULL or PARTIAL |
| 13 | N | PaidDiv | Total dividends paid |
| 14 | O | ReceivedDiv | Total dividends received |
| 15 | P | Notes | Trade journal notes (free text) |
| 16 | Q | TargetExitZ | Original exit Z target (carried from OpenTrades) |
| 17 | R | TargetPnL | Original PnL target (carried from OpenTrades) |
| 18 | S | SourcePortfolio | Source portfolio (carried from OpenTrades) |
| 19 | T | CloseReason | Free-text close classification (Target Hit / Stop Loss / Time Stop / manual) |
| 20 | U | TradeGroupID | UUID-like identifier linking all partial-close rows of the same trade lifecycle. Stamped at trade open (`cleanId(realId) + '_' + Date.now()`) and preserved across partial closes. Used by the History tab's "Group Partials" toggle to consolidate multiple partial closes into one display row and make stats (trades, wins, win rate) count trades not rows. |

## Deployment

### GAS (Backend) — done in browser

1. Open the Google Sheet → Extensions → Apps Script
2. Replace/update `Code.gs` and `SetupDashboard.gs`
3. Run `setupAllBatched()` (or individual steps for partial updates)
4. Run `createAutoTrigger()` to install all recurring triggers
5. Run `fetchDividendDates()` to populate DivDates sheet
6. Deploy → Manage Deployments → New Version → Deploy
7. The deployment URL stays the same (format: `https://script.google.com/macros/s/.../exec`)
8. **Set API_KEY** in Project Settings → Script Properties → add `API_KEY` = same value as in frontend `_app.html`

### Vercel (Frontend) — done in command prompt

1. Set `GAS_URL` and `API_KEY` in `_app.html` (not `index.html` — the app file was renamed to `_app.html`)
2. Set Vercel environment variables: `SITE_PASSWORD` (the login password) and optionally `AUTH_SECRET` (random string for cookie signing; auto-generated if not set)
3. From the `frontend/` directory, run `vercel --prod`
4. Or use: `npm run deploy:frontend` from the project root

### Frontend file structure

```
frontend/
  api/
    index.js       ← Vercel serverless function (password gate)
  _app.html        ← The actual dashboard (renamed from index.html)
  vercel.json      ← Rewrites all routes through auth gate
```

The serverless function (`api/index.js`) checks a `tp_auth` cookie. If valid → serves `_app.html`. If not → shows a login page. The `SITE_PASSWORD` env var controls the password. If not set, the site is open (backwards compat).

### npm Scripts (from project root)

- `npm run push:gas` — push GAS files via clasp
- `npm run deploy:frontend` — deploy frontend to Vercel
- `npm run deploy` — push GAS + deploy frontend

## Security Architecture

### Authentication Layers

1. **Website password** — Vercel serverless function at `api/index.js` checks `SITE_PASSWORD` env var. Login sets an `HttpOnly; Secure; SameSite=Strict` cookie valid 7 days. No one sees the HTML source, GAS URL, or API key without the password.
2. **API key** — Every GAS request includes a `key` parameter checked by `verifyAuth_()` against the `API_KEY` Script Property. Protects against direct API access even if someone discovers the GAS URL.
3. **Rate limiting** — `checkRateLimit_()` via CacheService: 60 calls/min general, 20/min for trade ops, 5/min for destructive ops (clearHistory, deleteClosedTrade).
4. **Audit logging** — All write actions logged to `AuditLog` sheet with timestamp + action + params. Auto-trims to 2000 rows.
5. **LockService** — All write actions wrapped in `LockService.getDocumentLock()` with 10s timeout.
6. **CSP** — Content-Security-Policy meta tag: `connect-src https://script.google.com` blocks data exfiltration.
7. **XSS prevention** — `escHtml()` for HTML content, `escAttr()` for onclick attribute strings.

### Credential Storage

- **Telegram bot token** — stored in `PropertiesService.getScriptProperties()` (encrypted at rest), NOT in sheet cells
- **API_KEY** — stored in GAS Script Properties (encrypted at rest) + frontend `_app.html` constant (protected by site password)
- **SITE_PASSWORD** — stored as Vercel environment variable (marked as "Sensitive")

## V25 Statistical Improvements (2026-04)

### Completed Features

1. **OLS Hedge Ratio for Credit Pairs** — `computeHedgeRatio_()` computes β via OLS, clamped [0.3, 3.0]. Credit pairs use adjusted spread `priceA - β·priceB`. Applied across all engines (backtest, sizer, screener).
2. **Half-Life as Trading Signal** — OU half-life (`ouHalfLife_()`) integrated into:
   - Quality column (sortable by HL, click "Quality" header)
   - Position sizing: `hlMult = clamp(21/HL, 0.5×, 2.0×)` in inlineSizer, recalcInlineSizer, runSizer, getPositionSizing_
   - Backtest: per-pair `pairMaxHold = min(3×HL, maxHold)` with `HL_CAP` exit reason
   - New "HL Mult" card in inline sizer and column in portfolio sizer
3. **β-Adjusted Size in Trade Modal** — `openModal()` reads `hedgeRatio` from `_alertDataMap`. For credit pairs (β≠1.0), shows yellow hint below SIZE B. Typing SIZE A auto-fills SIZE B = round(SIZE A × β).
4. **History Tab Grouping** — `groupHistoryByTrade()` consolidates partial closes of the same trade into one display row. Groups by `tradeGroupId` (UUID set at trade open) with fallback to composite key for legacy rows. Stats count groups not rows. "Group Partials" toggle (default ON) with purple expand/collapse for individual partials.
5. **30d/60d Win Rate Windows Only** — Dropped 15d and 90d windows from screener, display, and aggregate stats.
6. **3 Model Portfolios** — Reduced from 10, labeled Aggressive/Balanced/Diversified.
7. **Rate Regime Badge** — Shows "RATE MOVE +/-Xbp" when 30d 10Y yield change exceeds 25bp.
8. **Screener Cache for Inline Analysis** — `analyzeAlert()` checks screener cache first; renders cached result instantly with "Full Analysis" fallback button.

### Key Formulas

- **HL multiplier**: `clamp(21 / HL, 0.5, 2.0)` — 21 trading days baseline. HL=10d → 2.0×, HL=21d → 1.0×, HL=42d → 0.5×.
- **HL backtest cap**: `3 × HL` — 3 half-lives ≈ 87.5% expected reversion.
- **β source of truth**: `_alertDataMap[key].hedgeRatio` (scalar, clamped [0.3, 3.0]).

## Google Sheets Required

### Pairs (manual)

| A: PairID | B: TickerA | C: TickerB | D: Sector |
|-----------|-----------|-----------|----------|
| BAC-B\|BAC-M | BAC-B | BAC-M | Finance |

### Master (manual)

| A: Ticker | B: Last Close | C: Coupon Yield | D: Current Yield | E: Credit Rating |
|-----------|--------------|----------------|-----------------|-----------------|
| BAC-B | 25.50 | 6.00 | 5.88 | BBB+ |

**Coupon Yield empty = variable/reset rate → excluded from dashboard.**
**Credit Rating empty = excluded from credit arb pairing.**

## ADF Backlog

**Tier 1 — DONE**: Full-history lookback (no 90d cap), tri-state ADF (pass/weak/fail at 5%/10%), MacKinnon (1994) response surface for size-adjusted critical values, min 30 observations.

**Tier 2 — DONE (V25)**:
1. ~~Engle-Granger cointegration test~~ — `engleGrangerTest_()` at Code.gs:238. OLS priceA = α + β·priceB, ADF on residuals with EG-specific critical values.
2. ~~Hedge ratio β for sizing~~ — `computeHedgeRatio_()` at Code.gs:306. OLS β clamped [0.3, 3.0], used in credit pair spread, backtest, sizer.
3. ~~HL as secondary check~~ — OU half-life integrated into quality column, sizer, and backtest. HL 5–25d treated as tradeable signal.

**Tier 3 — regime awareness (pending):**
1. Rolling ADF: run on the last 60d in addition to the full window. Surface "was cointegrated, isn't anymore" pairs as a danger badge.

## Quant Audit Improvements (V24-V25)

### Statistical Foundation
- **ADF Cointegration Test** — Each alert pair runs an Augmented Dickey-Fuller test on the spread series. Non-stationary pairs (ADF p > 5%) display a red "✗ ADF" badge. Function: `adfTest_()` in Code.gs line 76. Size-adjusted critical values via `getAdfCriticalValues_()` (MacKinnon 1994).
- **Engle-Granger Cointegration** — Two-step test for credit pairs: OLS regression + ADF on residuals. Function: `engleGrangerTest_()` in Code.gs line 238.
- **OU Half-Life** — Ornstein-Uhlenbeck half-life per pair (`ouHalfLife_()` in Code.gs line 189). Displayed as "HL:Xd" in the Quality column. Drives position sizing multiplier and backtest hold caps.

### Call Risk Tracking
- Each alert includes `callRiskA` / `callRiskB` objects when a leg trades above par ($25). The Quality column shows red "CALLABLE" badges.

### Composite Score — REMOVED
- The `computeSetupScore()` function and associated UI (score bars, score rings, score badges) have been removed. The Quality column now surfaces ADF/half-life/call-risk as standalone badges instead.

### Survivorship Bias Fix
- **BLACKLIST** split into `BLACKLIST` (hard — data quality) and `SOFT_BLACKLIST` (poor performers). Backtests now use `isHardBlacklisted()` — only data-quality exclusions. Poor performers are included to prevent inflated backtest results.

### Adaptive Lookback Window
- Credit pair cache (`computeCreditCache`) now computes OU half-life first, then sets lookback = 3× half-life (clamped 30-180 days) instead of fixed 90. Falls back to 90 if half-life is not computable.

### Portfolio Risk Metrics
- **Max Drawdown** — `getBasketAnalytics()` now computes peak-to-trough drawdown from basket history. Displayed in the Health Bar's new "Risk Metrics" section.
- **Concentration Warning** — Alerts when >50% of portfolio notional is in one sector.

### Data Staleness
- Header now shows time since last WebCache update (e.g., "Data: 5m ago") with color coding: green (<15m), yellow (15-30m), red (>30m stale).
- `MIN_WIN_RATE_SAMPLES` raised from 5 to 20 for statistical rigor.

## Hardening Priority Matrix (session progress tracker)

Running audit of the 18-item consolidated priority matrix. Check this section first on resume so you know where you left off. Update statuses in place as items are finished.

| #  | Issue | Sev | Status | Commit(s) | Notes |
|----|---|---|---|---|---|
| 1  | XSS via innerHTML (frontend) | CRIT | DONE | 7c55041, (current) | `escHtml()` helper added; applied to all high-risk API error / diagnostic / user-facing interpolation (loadLiveAnalytics, renderCrossMatrix, renderAnalytics, renderSizer, runOptimalSweep, runBacktest, renderActionableSignals, exit banner s.message, healthStatus label, sector labels, concentration warning, loadMacroValuation). Remaining innerHTML usages (~119) are either static strings or safe numeric interpolation. |
| 2  | Trade race condition (no locking) | CRIT | DONE | b6c6e57 | `LockService.getDocumentLock()` wraps saveTradeToSheet / closeTradeInSheet / partialCloseTradeInSheet / addDividendToTrade. |
| 3  | No input validation on trades | CRIT | DONE | b6c6e57, 7c55041 | Frontend: price/size range validation in doSubmit / mgDoAdd / mgDoReduce / mgDoClose. Backend: saveTradeToSheet validates positive prices, non-zero size, $10k price / 100k share caps. |
| 4  | ADF over-detection (small n) | HIGH | DONE | b6c6e57 | MacKinnon (1994) response-surface critical values via `getAdfCriticalValues_()`, minimum sample raised 20 → 30. |
| 5  | Silent API failures | HIGH | DONE | 7c55041 | `apiGet()` retries with exponential backoff, toast notifications added (error/warn/ok). refreshData shows toast on failure. |
| 6  | N² portfolio lookup | HIGH | DONE | b6c6e57 | `getOpenTrades()` now uses hash-map lookup (was O(N) per trade). |
| 7  | Setup failure without recovery | HIGH | DONE | b6c6e57 | Prerequisite errors (missing Pairs sheet) now halt setup rather than looping the retry trigger. |
| 8  | Z-score returns 0 for low stdev | HIGH | DONE | b6c6e57, (current) | Backend: `lowVolatility: true` flag on rollingZ objects when std ≤ MIN_STDEV. Frontend: health bar + portfolio analytics windows now surface "low vol" badge and tooltip instead of silent 0σ. |
| 9  | No responsive breakpoints | MED | DONE | 7c55041 | CSS breakpoints at 1200px / 768px / 480px for tablet / mobile. |
| 10 | Probability engine spread bias | MED | DONE | b6c6e57 | `computeHistoricalProbabilitiesWide_` now only matches triggers on the same side of the mean as the current signal. |
| 11 | TNX division by 10 (verify) | MED | DONE | (current) | `getMacroData()` now autodetects: values ≥ 10 treated as CBOE index form (÷10), values < 10 treated as raw yield. Future-proofs against GOOGLEFINANCE API changes. |
| 12 | Memory leaks (timers) | MED | DONE | (current) | `toggleLive()` skips refresh when `document.hidden`; `beforeunload` handler clears `liveTimer` + `_titleFlashTimer`. Audit confirmed `_searchTimer`, `safetyTimer`, `_titleFlashTimer` were already properly managed. |
| 13 | `cleanIdFE` collision risk | MED | DONE | 7c55041 | Replaces `\|` with `_`, preserves hyphen (`-`), strips other non-alphanumeric. `BAC-B\|BAC-M` → `BAC-B_BAC-M` for safe DOM IDs. |
| 14 | No data export (CSV) | MED | DONE | (current) | CSV helpers: `csvCell`, `csvRows`, `downloadCsv`. Buttons: "⬇ CSV" in alerts filter bar (exports visible filtered rows from `_alertDataMap`), "⬇ Export CSV" in History tab (exports `_historyData`). UTF-8 BOM for Excel, proper quoting of embedded commas/quotes/newlines. |
| 15 | No retry logic on API calls | MED | DONE | 7c55041 | `apiGet()` wraps the fetch in a retry loop, up to 2 retries, exponential backoff (2s / 4s), only on AbortError / NetworkError / HTTP 5xx. |
| 16 | doPost incomplete | LOW | DONE | b6c6e57 | `doPost` now routes partialClose / addDividend and rejects unknown actions with a clear error. |
| 17 | Composite score undocumented | LOW | DONE | (current) | `computeSetupScore` fully removed. Quality column now surfaces ADF/half-life/call-risk as standalone badges. |
| 18 | Dead code / hidden features | LOW | PARTIAL | (current) | `computeSetupScore` function removed. Orphaned CSS classes (`.score-bar*`, `.score-ring`, `.score-badge`, `.score-cell`) still present — removal deferred. `loadTreasuryBanner` / `renderTreasuryBanner` confirmed active. All timers properly managed. |

**All 18 matrix items DONE except #18 (PARTIAL — orphaned CSS cleanup deferred).**

### Follow-up cleanup TODOs (safe to tackle next session)

- [ ] Remove dead `.score-bar`, `.score-bar-fill`, `.score-cell`, `.score-badge`, `.score-ring*` CSS rules from `_app.html` (orphaned after `computeSetupScore` removal).
- [ ] Audit the ~100 remaining `innerHTML =` assignments for lower-risk interpolations (static/numeric) and migrate to `textContent` where trivially safe.
- [ ] Add integration tests for the CSV export helpers (`csvCell`, `csvRows`).
- [ ] Consider rolling ADF surveillance (Tier 3 from ADF backlog above).
