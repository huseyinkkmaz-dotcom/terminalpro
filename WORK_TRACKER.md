# Work Tracker — Statistical Fixes & Improvements (V25)

**Created**: 2026-04-13
**Last Updated**: 2026-04-13
**Branch**: `claude/claude-md-mlqo1joq4yiiu2wc-b9axf`

## Session Resume Guide

When resuming, check this file FIRST. Each item has a status. Pick up the first `IN PROGRESS` or `TODO` item.

---

## DO NOW — Active Sprint

### 1. Backtest Rolling Stdev Biased Estimator
- **Status**: DONE
- **Priority**: CRITICAL (money bug)
- **Commit**: a309e41
- **Fix**: Changed to sample variance `(sumSprSq - n*mean²)/(n-1)` at all rolling variance sites.

### 2. Credit Pairs Hedge-Ratio-Adjusted Spread
- **Status**: DONE
- **Priority**: CRITICAL (biggest statistical weakness)
- **Commit**: a309e41
- **Fix**: Credit pairs now use OLS hedge-ratio-adjusted spread (`priceA - β·priceB`). β computed via `computeHedgeRatio_()` with [0.3, 3.0] clamping. Applied across all engines.

### 3. Dividend Proportional Split (Partial Closes)
- **Status**: DONE
- **Priority**: HIGH
- **Commit**: a309e41
- **Fix**: Partial close dividend split now weighted by dollar exposure per leg.

### 4. Mean Reversion Win Detection Too Loose
- **Status**: DONE
- **Priority**: HIGH
- **Commit**: a309e41
- **Fix**: Minimum tolerance raised from $0.01 to $0.10.

### 5. Rate Regime Badge on Alerts
- **Status**: DONE
- **Priority**: MEDIUM
- **Fix**: Backend computes 30d 10Y yield change from TreasuryHist. Frontend shows "RATE MOVE +/-Xbp" badge when >25bp change. Purely informational — doesn't filter.

### 6. Delete Dead Composite Score Code
- **Status**: DONE (already cleaned up in prior commit 8548542)
- **Priority**: LOW

### 7. Reduce to 3 Model Portfolios
- **Status**: DONE
- **Priority**: MEDIUM
- **Fix**: `NUM_PORTFOLIOS` 10→3, `MAX_PAIR_APPEARANCES` 4→2. Portfolios labeled "Aggressive", "Balanced", "Diversified" in frontend.

### 8. Keep Only 30d and 60d Win Rate Windows
- **Status**: DONE
- **Priority**: MEDIUM
- **Fix**: Backend screener computes [30, 60] only. ScreenerCache trimmed from 20→16 columns. `getScreenerData_` column indices updated. Frontend cross-matrix drops WR-90/EP-90 columns and totals. Aggregate banner uses wr60 instead of wr90. Inline analyze defaults to [30, 60]. Pair detail mini cards show 30d/60d only. renderAnalytics window cards show 30d/60d.

### 9. Inline Analyze Uses Screener Cache When Available
- **Status**: DONE
- **Priority**: MEDIUM
- **Fix**: `analyzeAlert()` checks screener cache first. If data <24h old, renders compact cached result immediately with "Full Analysis (API call)" fallback button. New `analyzeAlertFull()` function bypasses cache.

### 10. Half-Life as First-Class Trading Signal
- **Status**: DONE
- **Priority**: HIGH
- **Fix**: Added `data-hl` attribute to alert rows, `sortHL()` function, Quality column header is now sortable by half-life. Invalid HL sorts to bottom (999). HL now integrated into sizing: `hlMult = clamp(21/HL, 0.5x, 2.0x)` applied in `inlineSizer()`, `recalcInlineSizer()`, `buildInlineSizerCards()` (new HL Mult card), `runSizer()` (portfolio-level) with new HL Mult column, and backend `getPositionSizing_()`. HL also integrated into backtest: per-pair `pairMaxHold = min(3*HL, maxHold)` with new `HL_CAP` exit reason.

### 11. β-Adjusted Size Suggestion in Trade Modal
- **Status**: DONE
- **Priority**: MEDIUM
- **Fix**: `openModal()` now reads `hedgeRatio` from `_alertDataMap`. For credit pairs (β≠1.0), shows yellow hint below SIZE B with β value. Typing into SIZE A auto-fills SIZE B = round(SIZE A × β). Override manually if desired.

---

## SAVE FOR LATER — Backlog

### B1. Bid-Ask Spread Tracking
- **Status**: BACKLOG
- **Priority**: MEDIUM
- **Issue**: Volume without bid-ask context is incomplete.
- **Fix**: During fetchDividendDates, capture bid/ask from Yahoo. Store in TickerSpread column. Use in sizer and backtest friction.

### B2. Ex-Dividend Risk Quantification
- **Status**: BACKLOG
- **Priority**: MEDIUM
- **Issue**: Div dates shown but risk not quantified.
- **Fix**: Show div amount (coupon * par / 4). In trade modal, if either leg goes ex-div within expected hold period (from HL), show estimated div impact on PnL.

### B3. Yahoo Finance Failsafe Should Be Louder
- **Status**: BACKLOG
- **Priority**: LOW
- **Issue**: Yahoo price corrections are silent.
- **Fix**: Add "PRICE ADJ" badge to pair in WebCache. Surface in alerts view.

### B4. Half-Life in Sizer + Backtest
- **Status**: DONE (see item #10 above)
- **Priority**: HIGH
- **Fix**: HL multiplier in sizers (inline, portfolio, backend) + per-pair max hold cap in backtest.

---

## Completion Log

| Item | Completed | Commit | Notes |
|------|-----------|--------|-------|
| #1 Rolling Stdev | 2026-04-13 | a309e41 | Sample variance fix at all rolling sites |
| #2 Hedge Ratio | 2026-04-13 | a309e41 | OLS β for credit pairs |
| #3 Div Split | 2026-04-13 | a309e41 | Dollar-weighted div allocation |
| #4 Win Tolerance | 2026-04-13 | a309e41 | $0.10 min tolerance |
| #5 Rate Badge | 2026-04-13 | (current) | 30d 10Y yield change badge |
| #6 Dead Code | prior | 8548542 | computeSetupScore removed |
| #7 3 Portfolios | 2026-04-13 | (current) | Aggressive/Balanced/Diversified |
| #8 30/60d Windows | 2026-04-13 | (current) | Dropped 15d/90d from screener + display |
| #9 Cache Analyze | 2026-04-13 | (current) | Screener cache → instant inline results |
| #10 HL Sortable | 2026-04-13 | (current) | Quality column sortable by half-life |
| #10b HL Sizer+BT | 2026-04-14 | (current) | HL multiplier in sizers; per-pair max hold in backtest |
| #11 β Auto-fill | 2026-04-14 | (current) | Trade modal auto-fills SIZE B from SIZE A × β |
