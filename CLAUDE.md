# Terminal Pro Vol 2

## What This Is

A real-time **statistical arbitrage dashboard** for **Preferred Stock Pair Trading**. It tracks 100+ pairs using 90-day mean reversion Z-scores, yields, liquidity, and volume spike detection.

**Two strategies:**
- **Intra-Company Pairs** — Same issuer, different series (e.g., BAC-B vs BAC-M)
- **Credit Rating Arbitrage** — Different issuers, same credit rating (e.g., BBB vs BBB cross-company)

## Architecture

```
Google Sheets (Data Layer)
    ├── Pairs sheet         → Manual: intra-company pair definitions
    ├── Master sheet        → Manual: ticker, yield, coupon, credit rating
    ├── Levels sheet        → Auto: 90-day stats via GOOGLEFINANCE formulas
    ├── Live sheet          → Auto: real-time prices, Z-scores, liquidity
    ├── CreditPairs sheet   → Auto: generated inter-company pairs by rating
    ├── CreditLevels sheet  → Auto: 90-day stats for credit pairs
    ├── CreditLive sheet    → Auto: real-time data for credit pairs
    ├── ZScoreAge sheet     → Auto: timestamps when Z-scores first cross ±1.5
    ├── AlertsLog sheet     → Auto: hourly Z-score snapshots (trend ribbons)
    ├── OpenTrades sheet    → Trade journal (active)
    ├── ClosedTrades sheet  → Trade journal (closed)
    └── MacroData sheet     → Auto: US10Y, TLT, PFF macro indicators

Google Apps Script (API Backend)
    ├── Code.gs             → doGet/doPost routing, data reading, trade operations
    └── SetupDashboard.gs   → Sheet builder, credit pair generator, hourly/daily triggers

Netlify (Frontend)
    └── index.html          → Single-file dashboard (HTML + CSS + JS)
```

## Key File Descriptions

### `gas/Code.gs` — API Backend (V23)

- **doGet()** routes `?action=getData|saveTrade|closeTrade` with optional `&mode=intra|credit`
- **getAlertData(mode)** reads Live or CreditLive sheet. Filters:
  - Skips pairs with missing/zero prices
  - Skips pairs with < 60 trading days of history (HistCount column Q/index 16)
  - Skips pairs where either ticker has empty Coupon Yield (columns I,J / index 8,9)
  - Only returns pairs with |Z-Score| >= 1.5
- **getOpenTrades()** merges Live + CreditLive for unified portfolio lookup
- **saveTradeToSheet() / closeTradeInSheet()** check both Live and CreditLive

### `gas/SetupDashboard.gs` — Setup & Automation (V23)

- **setupDashboard()** builds Levels + Live sheets from Pairs sheet
- **generateCreditPairs()** reads Master, groups by credit rating, generates inter-company combinations
  - Uses union-find on Pairs sheet to detect same-company tickers (handles both hyphenated like BAC-M and non-hyphenated like AGNCP)
  - No pair cap — generates all valid combinations
- **setupCreditSheets()** builds CreditLevels + CreditLive from CreditPairs
- **setupTriggers()** installs:
  - Hourly: `snapshotZScores()` — logs to AlertsLog + updates ZScoreAge
  - Daily at 5AM: `dailyCreditRefresh()` — regenerates credit pairs from Master
- **setupAll()** convenience function: runs all four setup steps in order

### `frontend/index.html` — Dashboard UI

- Standalone HTML file deployed to Netlify
- Strategy pill switch: "Intra-Company" (blue accent) / "Credit Arb" (purple accent)
- Fetches from GAS via `?action=getData&mode=intra|credit`
- Features: Z-score trend ribbons, volume spike tags, sortable columns, trade entry modal
- **CRITICAL**: Line 1 of `<script>` has `GAS_URL` constant — must be set to deployed GAS URL

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

## Critical Rules When Editing

1. **Preserve ticker hyphens** — Preferred stock tickers use hyphens (PSA-F, BAC-M). Never strip them. The `cleanId()` function removes hyphens only for matching purposes, never for display or storage.
2. **Column indices matter** — Code.gs uses 0-indexed column references to read Live sheet arrays. If you add/remove/reorder columns in `buildLiveSheet_()`, you MUST update every index reference in `getAlertData()`, `getOpenTrades()`, `saveTradeToSheet()`, and `snapshotZScores()`.
3. **Two parallel sheet systems** — Intra-company uses `Pairs → Levels → Live`. Credit arb uses `CreditPairs → CreditLevels → CreditLive`. They share the same 24-column layout but reference different source sheets. The `buildLiveSheet_()` function is parameterized for this.
4. **GAS deployment** — Every code change requires: Manage Deployments → New Version → Deploy. The URL stays the same but the version must increment.
5. **GOOGLEFINANCE limits** — ~1000 GOOGLEFINANCE calls per sheet. Each pair uses ~6 calls in Live (price×2, volumeavg×2, volume×2) + 2 in Levels (historical). With 500+ credit pairs, sheets may load slowly (30-60s).
6. **Frontend is a single file** — All HTML, CSS, and JS live in `index.html`. No build step. Drag the `frontend/` folder to Netlify to deploy.
7. **ZScoreAge tracking** — The hourly trigger manages this. If |z| >= 1.5 and no timestamp exists → creates one. If |z| < 1.5 and timestamp exists → deletes it. Age = days since first crossing.

## Deployment

### GAS (Backend)

1. Open Apps Script editor for the Google Sheet
2. Replace/update `Code.gs` and `SetupDashboard.gs`
3. Run `setupAll()` (or individual steps: setupDashboard → generateCreditPairs → setupCreditSheets → setupTriggers)
4. Deploy → Manage Deployments → New Version → Deploy
5. Copy the deployment URL (format: `https://script.google.com/macros/s/.../exec`)

### Netlify (Frontend)

1. Set `GAS_URL` in index.html to the GAS deployment URL
2. Drag `frontend/` folder to Netlify deploy
3. Live at: https://terminal-pairs.netlify.app/

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
