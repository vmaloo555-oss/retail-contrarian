"""
daily_netpos.py -- rebuild retail NET POSITION per symbol with a moneyness-
weighted option book, and emit a daily file in the `pos`-sheet style.
Run once per new feed file.

Weights
    futures leg      : 1.0
    option positions : 0.1 (deep OTM) .. 1.0 (ATM / ITM), linear on the OTM
                       side -> posw()

new_net = (fut_b - fut_s) * 1.0                         # futures, from strike 0
        + SUM_k (call_b[k] - call_s[k]) * posw(k,'C')   # net LONG calls  (+)
        + SUM_k (put_s[k]  - put_b[k])  * posw(k,'P')   # net SHORT puts  (+)

Same bullish = positive convention as the pos-sheet col-15 formula
((3-4)+(7-8)+(12-11)); the only change is that option legs are moneyness-weighted
instead of counted flat. With every weight forced to 1.0 this reproduces col 15
exactly (see --check).

Inputs : the feed's `strike` sheet (per-strike b/s) + `pos` sheet (sector, LTP,
         original col-15 for reconciliation).
Outputs: output/netpos_<reading>.xlsx  (+ .csv)
         archive/netpos_history.csv  (appended every run -> previous / DIFF /
         history build up over time, newest-first like cols 15..34)
"""
from __future__ import annotations
import argparse, csv, os
import openpyxl

from strike_analysis import load_strike, _n
from pos_parser import parse_pos

FUT_WEIGHT = 1.0
POSW_FLOOR = 0.10        # weight of a deep-OTM option position
POSW_OTM_SPAN = 0.15     # OTM distance (fraction of spot) at which the floor is hit
HIST_KEEP = 20           # readings of new_net history to carry (cols 15..34 = 20)
OUT_DIR = "output"
HIST_CSV = "archive/netpos_history.csv"

_MONTHS = ["jan", "feb", "mar", "apr", "may", "jun",
           "jul", "aug", "sep", "oct", "nov", "dec"]


def display_name(tag):
    """Reading tag -> the operator's own feed naming convention.
       2026-09-08_1130 -> '8 sep-1'   (the 11:30 file)
       2026-09-08_1430 -> '8 sep-n'   (the 2:30 file)
    Archive files stay tag-named (they must sort); only delivered files use this.
    """
    try:
        d, slot = tag.split("_")
        _, mm, dd = d.split("-")
        return f"{int(dd)} {_MONTHS[int(mm) - 1]}-{'1' if slot == '1130' else 'n'}"
    except Exception:
        return tag


def posw(strike, spot, typ, flat=False):
    """Option moneyness weight in [POSW_FLOOR, 1.0].  typ 'C'/'P'.  x>0 = ITM."""
    if flat:
        return 1.0
    if not spot or strike <= 0:
        return 1.0
    x = (spot - strike) / spot if typ == "C" else (strike - spot) / spot
    if x >= 0:                       # ATM / ITM -> full weight
        return 1.0
    frac = max(0.0, 1.0 + x / POSW_OTM_SPAN)      # 1.0 at ATM -> 0.0 at the span
    return POSW_FLOOR + (1.0 - POSW_FLOOR) * frac


def _pos_meta(feed_path):
    """symbol -> (sector, ltp, col15) from the pos sheet."""
    feed = parse_pos(feed_path, header_rows=2)
    return {r.symbol.upper(): (r.sector, r.ltp, r.net_current) for r in feed.records}


def compute(feed_path, flat=False):
    strikes = load_strike(feed_path)
    meta = _pos_meta(feed_path)
    rows = []
    for sym, ks in strikes.items():
        sector, ltp, col15 = meta.get(sym, (None, None, None))
        fut_b = fut_s = 0.0
        call_w = put_w = 0.0
        for k, d in ks.items():
            if k == 0:
                fut_b, fut_s = d["fut_b"], d["fut_s"]
                continue
            call_w += (d["co_b"] - d["co_s"]) * posw(k, ltp, "C", flat)
            put_w += (d["po_s"] - d["po_b"]) * posw(k, ltp, "P", flat)
        fut_net = (fut_b - fut_s) * FUT_WEIGHT
        new_net = fut_net + call_w + put_w
        rows.append(dict(symbol=sym, sector=sector, ltp=ltp,
                         fut_net=round(fut_net, 2),
                         call_net_w=round(call_w, 2), put_net_w=round(put_w, 2),
                         new_net=round(new_net, 2), pos_col15=col15))
    rows.sort(key=lambda r: r["symbol"])
    return rows


# ---- history ----------------------------------------------------------
def _load_hist():
    """symbol -> list[(reading, {fut,call,put,new})] oldest-first."""
    h = {}
    if os.path.exists(HIST_CSV):
        with open(HIST_CSV, newline="") as f:
            for row in csv.DictReader(f):
                g = lambda k: float(row[k]) if row.get(k) not in (None, "") else None
                h.setdefault(row["symbol"].upper(), []).append(
                    (row["reading"], dict(fut=g("fut_net"), call=g("call_net_w"),
                                          put=g("put_net_w"), new=g("new_net"))))
    return h


HIST_FIELDS = ["reading", "symbol", "fut_net", "call_net_w", "put_net_w", "new_net"]


def _append_hist(reading, rows):
    os.makedirs(os.path.dirname(HIST_CSV), exist_ok=True)
    new = not os.path.exists(HIST_CSV)
    with open(HIST_CSV, "a", newline="") as f:
        w = csv.writer(f)
        if new:
            w.writerow(HIST_FIELDS)
        for r in rows:
            w.writerow([reading, r["symbol"], r["fut_net"], r["call_net_w"],
                        r["put_net_w"], r["new_net"]])


# ---- model status filters -------------------------------------------
FILTER_COLS = ["net_side", "ban", "fut_eligible", "opt_veto", "opt_reducing",
               "in_book", "book_lots", "fut_score", "opt_score", "conviction"]


def _model_status(feed_path, prev_strike_tag):
    """symbol -> filter/flag dict from model.run(); {} if model unavailable."""
    try:
        import model
        m = model.run(feed_path, strike_feed=feed_path,
                      prev_strike_tag=prev_strike_tag)
    except Exception as e:                       # keep the file working regardless
        print(f"  (model status skipped: {e})")
        return {}
    book = {b["symbol"].upper(): b for b in m["book"]}
    out = {}
    for x in m["rows"]:
        s = x["symbol"].upper()
        out[s] = dict(ban="Y" if x.get("banned") else "",
                      fut_eligible="Y" if x.get("elig_fut") else "",
                      opt_veto="Y" if x.get("opt_veto") else "",
                      opt_reducing=("Y" if x.get("opt_reducing") is True
                                    else "N" if x.get("opt_reducing") is False
                                    else ""),
                      in_book="Y" if s in book else "",
                      book_lots=book[s]["lots"] if s in book else "",
                      fut_score=x.get("score"), opt_score=x.get("opt_score"),
                      conviction=x.get("conviction"))
    return out


# ---- output file ----------------------------------------------------
HEADERS = ["symbol", "sector", "LTP", "fut_net(w1)", "call_net(wtd)",
           "put_net(wtd)", "NEW_NET", "NEW_NET_prev", "DIFF",
           "fut_d", "callW_d", "putW_d", "pos_col15", "reconcile(NEW-col15)"]


def write_files(reading, rows, hist, status=None, feed_path=None):
    status = status or {}
    os.makedirs(OUT_DIR, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "netpos"
    hist_cols = [f"h{i}" for i in range(HIST_KEEP - 1)]   # NEW_NET history, newest-first
    head = HEADERS + FILTER_COLS + hist_cols
    ws.append(head)

    name = display_name(reading)
    csv_path = os.path.join(OUT_DIR, f"netpos {name}.csv")
    cf = open(csv_path, "w", newline="")
    cw = csv.writer(cf)
    cw.writerow(head)

    def _d(now, was):
        return round(now - was, 2) if was is not None else None

    for r in rows:
        prev_list = [(rd, v) for rd, v in hist.get(r["symbol"], [])
                     if rd != reading]                    # safe re-runs
        pv = prev_list[-1][1] if prev_list else None
        prev = pv["new"] if pv else None
        recon = (round(r["new_net"] - r["pos_col15"], 2)
                 if r["pos_col15"] is not None else None)
        st = status.get(r["symbol"], {})
        nn = r["new_net"]
        net_side = "LONG" if nn > 0.5 else "SHORT" if nn < -0.5 else "FLAT"
        filt = [net_side, st.get("ban", ""), st.get("fut_eligible", ""),
                st.get("opt_veto", ""), st.get("opt_reducing", ""),
                st.get("in_book", ""), st.get("book_lots", ""),
                st.get("fut_score", ""), st.get("opt_score", ""),
                st.get("conviction", "")]
        series = [nn] + [v["new"] for _, v in reversed(prev_list)]
        series = series[:HIST_KEEP]
        line = ([r["symbol"], r["sector"], r["ltp"], r["fut_net"],
                 r["call_net_w"], r["put_net_w"], nn, prev, _d(nn, prev),
                 _d(r["fut_net"], pv["fut"] if pv else None),
                 _d(r["call_net_w"], pv["call"] if pv else None),
                 _d(r["put_net_w"], pv["put"] if pv else None),
                 r["pos_col15"], recon]
                + filt + series[1:])
        line += [None] * (len(head) - len(line))
        ws.append(line)
        cw.writerow(line)
    cf.close()
    ws.auto_filter.ref = ws.dimensions          # Excel dropdown filters on header
    ws.freeze_panes = "B2"                       # keep symbol col + header visible

    _add_strike_sheets(wb, feed_path)

    xlsx_path = os.path.join(OUT_DIR, f"netpos {name}.xlsx")
    wb.save(xlsx_path)
    return xlsx_path, csv_path


def _add_strike_sheets(wb, feed_path):
    """Sheet 2: every strike, moneyness-weighted.  Sheet 3: the contrarian
    option structure each stock's strike map implies."""
    if not feed_path:
        return
    import strike_analysis as SA
    strikes = SA.load_strike(feed_path)
    ltp = SA.load_ltp(feed_path)

    ws2 = wb.create_sheet("strikes weighted")
    ws2.append(SA.WEIGHTED_HEADERS)
    ws3 = wb.create_sheet("option strategy")
    ws3.append(SA.STRATEGY_HEADERS)

    for sym in sorted(strikes):
        spot = ltp.get(sym)
        if not spot:
            continue
        for row in SA.weighted_rows(sym, strikes[sym], spot):
            ws2.append(row)
        s = SA.suggest_strategy(sym, strikes[sym], spot)
        ws3.append([s["symbol"], s["spot"], s["floor"], s["ceiling"],
                    s["retail_dir"], s["activity"], s["read"], s["strategy"],
                    s["legs"], s["alt"], s["caution"]])

    for w in (ws2, ws3):
        w.auto_filter.ref = w.dimensions
        w.freeze_panes = "B2"


def run(feed_path, reading, do_archive=True, prev_strike_tag=None,
        with_status=True):
    rows = compute(feed_path)
    hist = _load_hist()
    status = _model_status(feed_path, prev_strike_tag) if with_status else {}
    xlsx, csvp = write_files(reading, rows, hist, status, feed_path)
    if do_archive:
        _append_hist(reading, rows)
    return rows, xlsx, csvp


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("feed", help="feed .xlsx (the newly dropped file)")
    ap.add_argument("reading", help="reading tag, e.g. 2026-09-07_1430")
    ap.add_argument("--no-archive", action="store_true",
                    help="compute + write the file but do NOT append to history")
    ap.add_argument("--no-status", action="store_true",
                    help="skip the model status/filter columns")
    ap.add_argument("--prev-strike-tag", default=None,
                    help="prior reading tag for the model's opt-reducing check")
    ap.add_argument("--check", action="store_true",
                    help="verify flat (all-weights=1) new_net == pos col 15")
    ap.add_argument("--show", nargs="*", default=[],
                    help="symbols to print")
    a = ap.parse_args()

    if a.check:
        flat = {r["symbol"]: r for r in compute(a.feed, flat=True)}
        bad = [(s, r["new_net"], r["pos_col15"]) for s, r in flat.items()
               if r["pos_col15"] is not None
               and abs(r["new_net"] - r["pos_col15"]) > 1e-6]
        print(f"flat new_net vs pos col15: {len(bad)} mismatches "
              f"of {len(flat)} symbols")
        for s, nn, c in bad[:20]:
            print(f"  {s:<12} flat={nn}  col15={c}  d={nn - c:+.2f}")
        raise SystemExit(0 if not bad else 1)

    rows, xlsx, csvp = run(a.feed, a.reading, do_archive=not a.no_archive,
                           prev_strike_tag=a.prev_strike_tag,
                           with_status=not a.no_status)
    print(f"wrote {xlsx}  and  {csvp}   ({len(rows)} symbols)"
          + ("   [history NOT updated]" if a.no_archive else ""))
    show = {s.upper() for s in a.show}
    if show:
        print(f"\n{'symbol':<12}{'LTP':>9}{'fut':>8}{'callW':>8}{'putW':>8}"
              f"{'NEW_NET':>9}{'col15':>8}{'d':>8}")
        for r in rows:
            if r["symbol"] in show:
                d = (r["new_net"] - r["pos_col15"]) if r["pos_col15"] is not None else None
                print(f"{r['symbol']:<12}{r['ltp']:>9}{r['fut_net']:>8.1f}"
                      f"{r['call_net_w']:>8.1f}{r['put_net_w']:>8.1f}"
                      f"{r['new_net']:>9.1f}{str(r['pos_col15']):>8}"
                      f"{(f'{d:+.1f}' if d is not None else '-'):>8}")
