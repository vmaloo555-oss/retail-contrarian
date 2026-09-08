# retail-contrarian

Deterministic retail-contrarian swing system for NSE F&O stock futures.
Interprets a twice-daily retail-positioning feed and produces a long-only target
book. **Claude interprets; the operator places every order.** See
[SYSTEM_SPEC.md](SYSTEM_SPEC.md) for the full ruleset — it is the source of truth.

## Status: accumulation week (started 2026-09-07)
Not trading. Collecting ~20 feed files (~10 trading days) to (a) validate the
model and (b) build the price archive so the price third of conviction can switch
on. Each run currently produces the **weighted net-position file** + a quiet
**model preview** (not for trading).

## Pipeline
```
feed .xlsx  ->  daily_run.py  ->  archive/*  +  output/netpos_<tag>.xlsx  +  output/preview_<tag>.txt
```
Every run ends with a persistence check listing each per-reading artifact as
OK/MISSING. **Both steps below are part of the run, not optional** — the cloud
container is ephemeral, so a reading that is neither committed nor uploaded is
lost, and a missing `archive/strike_archive/<tag>.csv` silently disables the
options-veto "reducing" check on the next reading (measured on 2026-09-07: 2 of
13 book names change).
- `pos_parser.py`    — parse & verify the `pos` sheet (col-15 formula check)
- `strike_analysis.py` — retail option positioning by strike, moneyness-weighted,
  fully-contrarian; broker-call filter; snapshots to `archive/strike_archive/`
- `daily_netpos.py`  — weighted net position (futures x1.0, options x0.1-1.0 by
  moneyness), pos-sheet style, DIFF + per-leg diffs + model filter columns
- `model.py`         — eligibility, conviction (options ~50%: 50/50 netpos/options
  while price is off, 0.35/0.50/0.15 once price is on — see SYSTEM_SPEC.md),
  options veto, F&O ban filter, NIFTY regime cap, rank-and-fill
- `daily_run.py`     — one entry point per feed (called by the cloud routine)

## Run one feed
```
python daily_run.py path/to/FEED.xlsx --tag 2026-09-07_1430
```
Tag = `YYYY-MM-DD_HHMM` (`1130` = 11:30 file, `1430` = 2:30 decision file).

## After each run
1. Commit the accumulation record — `archive/` (price archive, netpos history,
   per-reading strike snapshots) and `output/`.
2. Upload `output/netpos_<tag>.xlsx` and `output/preview_<tag>.txt` to the
   Drive folder "retail-contrarian-system". `daily_run.py` prints the exact list;
   it has no Drive credentials of its own, so the upload is a separate step.

## Config
- `config/ban_list.txt` — NSE F&O ban list, refresh every trading day.

## Requires
Python 3.11+, `openpyxl`.
