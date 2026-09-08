"""
daily_run.py -- one entry point per feed file. Called by the cloud routine.

For each feed .xlsx it:
  1. snapshots the `strike` sheet          -> archive/strike_archive/<tag>.csv
  2. archives col-2 LTP                     -> archive/price_archive.csv (append)
  3. builds the weighted net-position file  -> output/netpos_<tag>.{xlsx,csv}
       (in the pos-sheet style, with DIFF + per-leg diffs + model filter columns)
  4. writes a quiet model preview           -> output/preview_<tag>.txt
       (the model's would-be list -- NOT for trading during the accumulation week)

Usage:
    python daily_run.py FEED.xlsx [--tag YYYY-MM-DD_HHMM]
`--tag` defaults to <feed mtime date>_<1130|1430> guessed from the mtime hour,
or pass it explicitly.
"""
from __future__ import annotations
import argparse, csv, glob, io, os, sys, datetime as dt

import openpyxl
import strike_analysis as SA
import daily_netpos as DN
import model as M

PRICE_CSV = "archive/price_archive.csv"


def guess_tag(feed_path):
    ts = dt.datetime.fromtimestamp(os.path.getmtime(feed_path))
    slot = "1130" if ts.hour < 13 else "1430"
    return f"{ts:%Y-%m-%d}_{slot}"


def archive_ltp(feed_path, tag):
    wb = openpyxl.load_workbook(feed_path, data_only=True, read_only=True)
    ws = wb["pos"]
    rows = []
    for r in list(ws.iter_rows(values_only=True))[2:]:
        if not r or r[0] in (None, "") or str(r[0]).strip().lower() == "total":
            continue
        if isinstance(r[2], (int, float)):
            rows.append((tag, str(r[0]).strip(), r[2]))
    wb.close()
    os.makedirs(os.path.dirname(PRICE_CSV), exist_ok=True)
    new = not os.path.exists(PRICE_CSV)
    # de-dup: drop any existing rows for this tag, then rewrite
    existing = []
    if not new:
        with open(PRICE_CSV, newline="") as f:
            existing = [row for row in csv.reader(f)
                        if row and row[0] not in ("reading", tag)]
    with open(PRICE_CSV, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["reading", "symbol", "ltp"])
        w.writerows(existing)
        w.writerows(rows)
    return len(rows)


def prev_strike_tag(tag):
    tags = sorted(os.path.splitext(os.path.basename(p))[0]
                  for p in glob.glob(os.path.join(SA.ARCHIVE_DIR, "*.csv")))
    tags = [t for t in tags if t < tag]
    return tags[-1] if tags else None


def price_readings_count():
    if not os.path.exists(PRICE_CSV):
        return 0
    with open(PRICE_CSV, newline="") as f:
        return len({row[0] for row in csv.reader(f)
                    if row and row[0] not in ("reading",)})


def write_preview(feed_path, tag, prev_tag):
    path = os.path.join(DN.OUT_DIR, f"preview {DN.display_name(tag)}.txt")
    buf = io.StringIO()
    try:
        out = M.run(feed_path, strike_feed=feed_path, prev_strike_tag=prev_tag)
    except Exception as e:
        buf.write(f"model preview unavailable: {e}\n")
        open(path, "w").write(buf.getvalue())
        return path
    nt = out["nifty_trend"]
    buf.write(f"# QUIET MODEL PREVIEW {DN.display_name(tag)}  ({tag})"
              f" -- NOT FOR TRADING (accumulation week)\n")
    buf.write(f"# NIFTY score {nt['score']:+.3f} -> regime {out['regime']} "
              f"cap Rs {out['regime_cap']:,}\n")
    if out["ban_list"]:
        buf.write(f"# ban: {', '.join(out['ban_list'])}\n")
    buf.write(f"# eligible {len(out['eligible'])}   book {len(out['book'])} names   "
              f"deployed Rs {out['deployed']:,}\n\n")
    buf.write(f"{'sym':<14}{'lots':>5}{'score':>9}{'opt_s':>8}{'conv':>8}   Rs value\n")
    for x in out["book"]:
        buf.write(f"{x['symbol']:<14}{x['lots']:>5}{x['score']:>+9.3f}"
                  f"{x['opt_score']:>+8.2f}{x['conviction']:>8.3f}   "
                  f"Rs {x['value']:,}\n")
    open(path, "w").write(buf.getvalue())
    return path, out


def run(feed_path, tag=None):
    tag = tag or guess_tag(feed_path)
    prev_tag = prev_strike_tag(tag)

    snap = SA.snapshot(feed_path, tag)
    nltp = archive_ltp(feed_path, tag)
    rows, xlsx, csvp = DN.run(feed_path, tag, do_archive=True,
                              prev_strike_tag=prev_tag, with_status=True)
    prev = write_preview(feed_path, tag, prev_tag)
    preview_path = prev[0] if isinstance(prev, tuple) else prev

    pdf_path = "(skipped)"
    try:
        from strategy_pdf import build as build_pdf
        pdf_path = build_pdf(xlsx, preview_txt=preview_path)
    except Exception as e:                    # never let the PDF break the run
        pdf_path = f"(failed: {e})"

    n_price = price_readings_count()
    print(f"[{tag}]  prev_strike={prev_tag}")
    print(f"  strike snapshot : {snap}")
    print(f"  LTP archived    : {nltp} rows  ({n_price} readings total)")
    print(f"  weighted netpos : {xlsx}")
    print(f"  model preview   : {preview_path}")
    print(f"  strategy PDF    : {pdf_path}")
    print(f"  price component : {'READY - >=20 readings, switch price_c on' if n_price >= 20 else f'accumulating ({n_price}/20 readings)'}")
    return tag


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("feed")
    ap.add_argument("--tag", default=None)
    a = ap.parse_args()
    run(a.feed, a.tag)
