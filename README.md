# ⚠️ STALE REPO — DO NOT DEPLOY FROM HERE

This is the **V25 (2026-04) snapshot** of Terminal Pro. Active development moved to
**`terminal-pro-v2`** (21k+-line backend, paper bot, self-learning pipeline — none of it
is in this repo).

**CRITICAL:** this repo's `gas/.clasp.json` pointed at the **SAME live Google Apps Script
project** as terminal-pro-v2. A `clasp push` (or `npm run push:gas` / `deploy`) from this
checkout would have **overwritten the live production backend with this 7k-line V25
snapshot**. The config has been renamed to `gas/.clasp.json.DISABLED` to make that
impossible by accident.

If you ever genuinely need to push from here (you almost certainly don't):
1. Stop. Check `terminal-pro-v2` first — it is the live codebase.
2. If still needed, rename `.clasp.json.DISABLED` back, and point `scriptId` at a
   THROWAWAY Apps Script project, never the production one.
