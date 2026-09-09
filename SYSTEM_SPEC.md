# Retail-Contrarian Swing System — Reference Spec

Single source of truth. Every run is checked against this file. Rules are fixed and
defined by the operator; Claude interprets the feed against them and never deviates,
never places orders, never contacts the dealer, gives no discretionary view.

## Book
- Universe: NSE F&O stock futures.
- Direction: LONG-ONLY.
- Notional cap: ₹65 crore (hard).

## Premise
Retail in Indian F&O is structurally net-long and systematically loses. A multi-day
unwind of retail net-long positioning during a trend is a contrarian BUY.

## Data feed
- Excel workbook, sheet `pos`, one row per F&O symbol.
- Two files/day, ~11:30 and ~14:30 IST. **The 14:30 file drives decisions.**
- Column map (0-indexed):
  | col | meaning |
  |----|---------|
  | 0 | symbol |
  | 1 | sector |
  | 2 | LTP |
  | 3 / 4 | retail FUT long / short — current |
  | 5 / 6 | retail FUT long / short — previous |
  | 7 / 8 | retail CALL long / short — current |
  | 9 / 10 | retail CALL long / short — previous |
  | 11 / 12 | retail PUT buy / sell — current |
  | 13 / 14 | retail PUT buy / sell — previous |
  | 15 | retail NET POSITION (current) = `(3−4) + (7−8) + (12−11)` |
  | 16 | net position (previous) |
  | 17 | DIFF — **ignore** |
  | 18–34 | net-position history, newest→oldest, ~2 readings/trading day |
  | >34 | **ignore** |

## Parser — `pos_parser.py`
`parse_pos(path, sheet="pos", header_rows="auto", tol=1e-6) -> PosFeed`
- Loads `pos`, auto-detects header rows.
- Verifies col 15 against the formula on every row; failures collected in
  `PosFeed.mismatches` (blank, non-numeric, or value ≠ formula beyond `tol`).
- Per symbol (`SymbolRecord`): `symbol`, `ltp`, `net_series` = `[col 15]` then cols
  18–34, newest first, blanks dropped. Diagnostics also kept: `row`, `sector`,
  `net_current`, `net_current_expected`, `verified`.
- `PosFeed.ok` is True only when there are zero mismatches. CLI exit code mirrors it.
- Reads cached formula results only (openpyxl `data_only`); warns if col 15 is
  empty everywhere.
- Header auto-detect: skips a leading row only if col 0 is a known label
  (`symbol`/`ticker`/…/blank) or ≥2 of {LTP, 15, 16, 18} carry text labels. A
  single corrupt cell in the first data row will NOT drop that row.
- LTP coercion: `n/a`, `-`, `#N/A`, `nil`, blank → `ltp = None` (no warning).
  Other non-numeric text (e.g. `12x`) → `ltp = None` + a warning.
- Numeric coercion (all cells): strips thousands separators, unicode minus;
  reads `(1,020)` as `-1020`.
- Verified 2026-09-07 against `sample_pos.xlsx` (built by `make_sample.py`):
  6 symbols, 3 intended mismatches (stale value / wrong value / blank col 15),
  duplicate-symbol + corrupt-LTP + all-blank-col15 warnings, trailing blank row
  skipped. `make_sample.py` + `sample_pos.xlsx` kept as the regression fixture.

### Runtime note (this machine)
Python 3.14 (`python` on PATH) + `openpyxl` 3.1.5 installed. The Claude app is an
MSIX package: `...\AppData\Roaming\Claude\...` is the app's virtualised view;
external Python resolves the same location as
`...\AppData\Local\Packages\Claude_pzs8sxrjxfjjc\LocalCache\Roaming\Claude\...`.
Consequence: `cd` into the working dir and call the parser with **relative**
paths (`python pos_parser.py feed.xlsx`); an absolute `C:\...\Roaming\Claude\...`
path passed to Python will not open.

## Daily net-position rebuild (`daily_netpos.py`) — runs on every new feed file
Recomputes retail NET POSITION per symbol from the `strike` sheet with a
moneyness-weighted option book: **futures leg weight 1.0**; **option positions
weight 0.1 (deep OTM) → 1.0 (ATM/ITM)**, linear on the OTM side (`posw()`,
floor 0.10, OTM span 0.15 — tunable). Same bullish=+ convention as pos col 15
(`(fut_b-fut_s) + Σ(call_b-call_s)·w + Σ(put_s-put_b)·w`); with all weights 1.0
it reproduces col 15 exactly (`--check`: 0/217 mismatches).
Outputs `output/netpos_<reading>.{xlsx,csv}`: symbol·sector·LTP·fut_net·
call_net_wtd·put_net_wtd·NEW_NET·NEW_NET_prev·**DIFF**·**fut_d·callW_d·putW_d**·
pos_col15·reconcile + NEW_NET history newest-first. DIFF = NEW_NET − previous
reading's NEW_NET; fut_d/callW_d/putW_d = same diff per leg (which part moved).
`archive/netpos_history.csv` stores fut/call/put/new per reading; re-running the
same reading tag is safe (its own row is excluded when picking "previous").
Model does NOT consume this — it stays on raw pos col-15 (operator confirmed).
**Filter columns** (from `model.run()`, joined by symbol) + Excel autofilter +
frozen header/symbol col: `net_side` (LONG/SHORT/FLAT by NEW_NET sign) · `ban` ·
`fut_eligible` · `opt_veto` · `opt_reducing` · `in_book` · `book_lots` ·
`fut_score` · `opt_score` · `conviction`. `--no-status` skips them.
KNOWN ISSUE: `opt_veto` fires on ~103/217 today — the -0.25 threshold + activity
floor are too loose on thin chains; recalibrate once a week of strike data
exists.

## Model (implemented in `model.py`)
Per symbol from `net_series` (newest first): `np_now=s[0]`, `np_old=s[-1]`,
`score=(np_now-np_old)/|np_old|`, `R²` = OLS fit oldest→newest (flat/degenerate → 0),
`recent=(np_now-s[k])/|s[k]|` with `k=min(8,len-1)`, `medlev=median(|s|)`.
Eligible long (ALL): `score≤-0.15`, `|np_old|≥10`, `recent≤+0.15`, `len≥3`.
`score≥+0.15` ⇒ retail building ⇒ exclude / EXIT if held.
`medlev<10` ⇒ low-participation (keep, size cap 20).
**Conviction weights — options data ~50% (operator 2026-09-07):**
price OFF: `conviction = 0.50·score_c + 0.50·opt_c`;
price ON: `0.35·score_c + 0.50·opt_c + 0.15·price_c` (split TBD).
`score_c = min(|score|,1)` (futures unwind depth). `opt_c = max(0, opt_score)`
where `opt_score = retail_dir / (opt_activity + 1)` ∈ ~[-1,1] from
`strike_analysis.band()` (moneyness-weighted, fully contrarian). Only positive
(confirming) opt_score adds conviction.
**F&O ban filter:** symbols in `config/ban_list.txt` (NSE F&O ban period =
reduce-only) are never eligible for a new BUY — excluded from the book. Refresh
that file from the exchange ban list every trading day. A *held* banned name:
HOLD (cannot add / resize up); EXIT still allowed (reducing is permitted in ban).
CLI `--ban SYM,SYM` overrides the file.

**Options veto:** a futures-eligible name with `opt_score ≤ -0.25` is DROPPED
unless the strike archive shows its adverse retail position (weighted long calls
+ short puts) *reducing* reading-over-reading. Suppressed when
`opt_activity < 3` (chain too thin to weigh). CAVEAT: the "reducing" check needs
a multi-day strike archive to be meaningful — with 1 day it is noise; thresholds
(-0.25 veto, 3.0 activity floor, 50/50 split) all provisional pending a week of
data.
Price/position gate (`price_gate()`, live only when price is on; price trend
from the trailing 10-trading-day window ≈ 20 readings — 'up' = LTP > window avg
AND slope > 0; 'down' = LTP < avg AND slope < 0; 'flat' = mixed):
  · price 'down' + retail position **increasing** (`score ≥ +0.15`) → **eliminate**
    (exclude from target; EXIT if held)
  · price 'down' + retail position **flat** (`|score| < 0.15`) → **hold-only**
    (no fresh BUY; if held, keep — do not EXIT, do not resize up)
  · price 'down' + retail **actively unwinding** (`score ≤ -0.15`) → OK, `price_c=0`
  · price 'up' or 'flat' → OK
This replaces the earlier "any downtrend = HARD eliminate".
Lots: `R²<0.40 or conv<0.25`→10; `conv 0.25–0.55`→`20+(conv-0.25)/0.30·30`;
`conv>0.55`→`50+min((conv-0.55)/0.30,1)·30`; low-part→cap 20;
SUPER-BULL→80 ONLY when a confirmed positive price uptrend AND huge retail
unwinding coincide: {price_c live AND price_c > 0} AND {not low-part · np_now<5 ·
np_old-np_now≥10 · score≤-0.60}. "Positive uptrend" = price_c > 0 (operator
confirmed 2026-09-07); any strictly-positive reading counts, no strength floor.
While the price component is OFF, price_c is not live, so SUPER-BULL never fires
and those names size on their conviction tier instead.
Regime from NIFTY row score: `≤-0.15`→bull, cap ₹65cr; `≥+0.50`→extreme-bear,
cap ₹15cr; else neutral, cap ₹32.5cr.
Rank: super-bull first, then conviction desc. Fill at ₹7,00,000/lot to the cap;
last name gets leftover room only if ≥10 lots, then stop; drop the rest.
Diff vs positions sheet: in target only→BUY; held only→EXIT full; Δlots≥10→RESIZE;
else HOLD. Tickets: buy `SYM 20+` · trim `SYM -10` · close `SYM sq 20`.

## Feed reality — observed 2026-09-07 (Drive folder "position")
- Two Google Sheets per day: **`excel 7 sep-1` = 11:30 file**, **`excel 7 sep-n` =
  14:30 file** (decision file). Basis: `sep-n`'s "previous" columns (hdr dates
  col 5/9/13/16) all = same day and equal `sep-1`'s "current"; its history is
  shifted one reading newer; created 9 s later. If a future drop names files
  differently, re-confirm by this lineage, not by the name.
- `pos` sheet has **2 header rows** (row 1 = segment labels incl. `#REF!`;
  row 2 = `stock` + per-column dates). Parse with `header_rows=2`.
- Symbol column header is `stock`. Sectors include junk values `0.0`, `None`.
- Junk rows inside the block: a `0.0` symbol row, `FOCIT`, two trailing `Total`
  rows — `model._is_tradable_symbol()` drops them.
- Index rows present as data: BANKNIFTY FINNIFTY MIDCPNIFTY NIFTYNXT50 BANKEX
  SENSEX — dropped via `INDEX_SYMBOLS`. `FORTIS` is a real stock mis-tagged
  sector `INDEX` (filter is by symbol, not sector, so it survives).
- Cols 35–43 are populated despite "ignore >34": ~3 more history readings
  (35–37), an extra diff pair (38–39), `#REF!`/`#ERROR!` (40–41), then a
  **reference price (42) + return fraction (43)** e.g. NIFTY `24397.0 / -0.0212`.
  Ignored per spec — candidate wiring for `price_c` later.
- col-15 formula verified on all 223 rows of the 14:30 file: **0 mismatches**.
- `sep-1` `pos` is 43 cols wide, `sep-n` is 44 — trailing-column drift.

## Strike analysis (`strike_analysis.py` — standalone, not yet wired to `model.py`)
Source: `strike` sheet (Excel pivot "Count of Net Qty."). Row label `SYMBOL  STRIKE`
(strike 0 = FUT leg). Cols: 1/2 FUT b/s · 4/5 CALL b/s · 7/8 PUT b/s · 10 grand
total. Values are **counts** (# retail long vs short), not lots.
The `pos` net position already accounts for total retail exposure; this overlay
is SHAPE/range only. Per strike: `netC = callB-callS`, `netP = putB-putS`. Fade:
  · net **long** calls above spot → CEILING (their OTM calls decay)
  · net **short** calls above spot → likely BREAKS (shorts run over → breakout fuel)
  · net **long** puts below spot → FLOOR
  · net **short** puts below spot → likely BREAKS DOWN (no real support)

**Moneyness weighting (operator confirmed 2026-09-07):** ITM strikes weighted
HIGHER because an ITM retail position carries intrinsic value → bigger retail
P&L outflow when the fade plays out → bigger contrarian signal; OTM = cheap,
small outflow, decays to noise. Each strike's net count × `mny_weight()`:
signed depth `x` (call: `(S-K)/S`; put: `(K-S)/S`); `x≥0` → `1 + ITM_BONUS·
min(x, ITM_CAP)` (1.0 ATM … 2.0 deep ITM); `x<0` → `exp(x / OTM_TAU)`
(≈0.43 at 5% OTM, 0.19 at 10%). Defaults BONUS 2.0 / CAP 0.50 / TAU 0.06 —
tunable, monotonic highest-ITM → lowest-OTM.
Band WALLS come from retail LONG options (their bought OTM options pin/decay):
ceiling = nearest strike > spot with weighted `netC ≥ MIN_NET`; floor = nearest
strike < spot with weighted `netP ≥ MIN_NET`. up-skew if nearest strike above is
weighted net short calls; down-skew if nearest below is weighted net short puts.

**Directional lean — `retail_dir` (FULLY CONTRARIAN, operator confirmed
2026-09-07):** fade every retail option position, no neutral "pin" bucket.
`retail_dir = Σ(wNetP) − Σ(wNetC)` (moneyness-weighted):
  · retail long calls (bullish bet)        → fade → BEARISH
  · retail short calls (capping bet)        → fade → BULLISH
  · retail long puts (bearish bet)          → fade → BULLISH
  · retail short puts (floor / complacent)  → fade → BEARISH
`retail_dir > 0` → contrarian BULLISH (confirms / can upsize a long);
`retail_dir < 0` → contrarian BEARISH (caution / veto a long).
**Broker-call filter:** concentrated single-strike activity ≈ a broker tip, not
dispersed retail. Flag a strike when gross count (b+s across C&P) is
≥ `BROKER_MULT`×median strike AND ≥ `BROKER_SHARE` of the whole chain AND
≥ `BROKER_MIN_GROSS` absolute AND not within `ATM_PCT` of spot (ATM
concentration is normal). Defaults 3× / 30% / 8 / 2% — tunable.
A flagged cluster is **OMITTED ENTIRELY** (operator confirmed 2026-09-07):
dropped from the band walls and carries NO signal — no directional/confirm/veto
read. Printed as a footnote for transparency only.
Tracking: every file → `archive/strike_archive/<reading>.csv` (symbol×strike
counts); diff strike-by-strike across readings. A sudden one-reading spike at a
single strike is a stronger broker-call signal than a persistent one.

**Held-long band-shift (operator confirmed 2026-09-07):** when a held long's
contrarian band moves against it — ceiling drops toward/below spot or entry,
floor breaks, or a fresh down-skew appears — **flag it AND RESIZE down** (not a
full exit). Resize goes through the normal path (only ticketed when the change
is ≥ 10 lots). Down-size magnitude TBD (proposed: cut to the conviction-tier
lots for the reduced band, or halve — pick one).

**Expiry (operator confirmed 2026-09-07):** the `strike` sheet blends expiries
and carries no expiry field; ~75% of the activity is near-month. No expiry
filter is possible from the data — use the chain as-is and treat band edges as
carrying ~25% cross-expiry noise.

OPEN: (a) does a band also affect a *new* BUY (gate / initial size / veto), or
only held-position resizing? (b-mag) the down-size magnitude on an adverse band
shift. (e) MIN_NET wall threshold — now applied to the *weighted* net, so raw
counts of 1–2 at 5%+ OTM no longer form a wall (thin names show "? – ?"); keep
at 1.0 weighted, lower it, or apply MIN_NET to raw count and use the weight only
for `retail_dir`?
RESOLVED: (b) adverse band shift on a held long → flag + resize down.
(c) broker-call cluster omitted entirely (no signal). (d) no expiry filter,
chain used as-is (~75% near-month).

## Daily loop
11:30 & 14:30 files land in Drive · 14:30 drives · 16:00 operator sends positions
sheet · Claude returns analysis + tickets · operator reports fills, Claude notes.
Archive col 2 (LTP) every file → `archive/price_archive.csv`; once the trailing
10-trading-day window (~20 readings) is filled, build per-symbol
LTP-vs-window-avg + slope and switch `price_c` on. Until then price is OFF and
the operator manually vetoes names not trending up.
**Not trading yet — validating the model on a week of files first.**

### Price component — design intent (not built; needs final price_c magnitude)
- Trend window = trailing **10 trading days ≈ 20 readings** (2/day). Judged on
  the **overall** window trend, NOT a recent sub-window — a short-term bounce
  inside an overall downtrend is not an uptrend.
- `price_state`: **up** = LTP > window avg AND slope > 0; **down** = LTP < avg
  AND slope < 0; **flat** = mixed (only one up).
- **up** ⇒ `price_c > 0` (any strictly-positive reading; no strength floor) —
  enables SUPER-BULL when the huge-unwind conditions also hold.
- **flat** ⇒ `price_c = 0`, name kept, no SUPER-BULL.
- **down** ⇒ not an automatic eliminate anymore; resolved by `price_gate()`
  against the retail-position move (see Model section): increasing→eliminate,
  flat→hold-only, unwinding→OK with `price_c = 0`.
- `price_c` magnitude (0..1 strength curve for the 'up' case) is still TBD
  before go-live.

## Delivery (as of 2026-09-09)
`daily_run.py FEED.xlsx --tag <tag>` is the single entry point. Per reading it
writes, into `output/`, named in the operator's convention (`9 sep-1` = 11:30,
`9 sep-n` = 14:30):

| file | size | goes to |
|---|---|---|
| `netpos <name>.xlsx` (3 sheets: netpos · strikes weighted · option strategy) | ~280 KB | git only |
| `strategy <name>.pdf` | ~24 KB | git only |
| `netpos <name>.csv` | ~33 KB | git only |
| `strategy <name>.csv` | ~46 KB | git only |
| **`summary <name>.csv`** | ~15 KB | **Google Drive** + git |
| **`preview <name>.txt`** | ~1.5 KB | **Google Drive** + git |

**Drive gets text only, and only the two compact files.** The connector's
`create_file` passes content as base64 through the model's context: a 33 KB
xlsx becomes 47 KB of base64 and hangs the run, and even the two full CSVs
(81 KB combined) hung runs on 2026-09-08/09. `summary <name>.csv` replaces both
for Drive — one row per symbol that has either a recommended structure or a
live book position, columns: symbol · LTP · NEW_NET · DIFF · fut_score ·
opt_verdict · opt_conv · structure · legs · alternative · caution · in_book ·
lots. The full detail stays in the repo.

Drive output folder: `retail-contrarian-system`, id
`1yl6JGsAx7TOi34-2Lme10p3aoUmPpvKy`. Feed folder: `position`, id
`19krMYOT8MaziQAbkOPhr8Y9OrjUBD85y`.

### Automation
Two claude.ai routines poll the feed folder on alternating half-hours, weekdays:
watcher A `15 6-10 * * 1-5` UTC (11:45/12:45/13:45/14:45/15:45 IST) and
watcher B `45 5-9 * * 1-5` UTC (11:15/12:15/13:15/14:15/15:15 IST). One hour is
the platform minimum per routine, so two offset routines give ~30-minute
latency from upload to Drive. Each fire does a cheap check first — list both
Drive folders, and if every feed already has a matching `summary <name>` it
reports "nothing new" and exits before cloning. Otherwise it clones the public
repo, reprocesses EVERY feed file (no saved state, self-healing), and delivers
the missing readings oldest-first.
