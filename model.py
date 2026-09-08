"""
model.py -- runs the retail-contrarian model on one 2:30 feed.
Reads: pos_parser.parse_pos()  ->  per-symbol net_series (= [col15] + cols18..34, newest first).
Price component is OFF (no price archive yet): price_c = 0, list is SCORE-ONLY,
and price_gate() cannot run -> operator vetoes non-uptrending names.
conviction = 0.65*score_c + 0.35*price_c.
"""
from __future__ import annotations
import argparse, json, statistics as st
from pos_parser import parse_pos

LOT_RUPEES = 700_000
PRICE_WINDOW_DAYS = 10   # trailing trading days for the price trend (~20 readings, 2/day)
FLAT_BAND = 0.15         # |score| <= this => retail position treated as flat

# --- conviction weights: 1/3 each (operator 2026-09-07) ----------------
#   conviction = (netpos_trend + options + price_uptrend) / 3
# Price component is OFF until the archive is built -> its third contributes 0,
# so conviction is capped at ~0.667 until price goes live. Not renormalised.
CONV_W_NETPOS = 1/3     # futures retail-unwind depth (score_c)
CONV_W_OPT = 1/3        # moneyness-weighted contrarian options signal (opt_c)
CONV_W_PRICE = 1/3      # price-uptrend strength (price_c)
CONV_W_SCORE = CONV_W_NETPOS   # back-compat alias

# Options veto: a name whose contrarian options signal is bearish is dropped
# UNLESS the strike history shows the adverse retail position (long calls /
# short puts) reducing reading-over-reading.  Only applied when the name has
# enough weighted option activity for the signal to mean something.
OPT_VETO = -0.25
OPT_MIN_ACT = 3.0      # |call_bias| + |put_bias| floor for opt_score to count

# NSE F&O ban period: a banned name is reduce-only -> never a new BUY. Refresh
# config/ban_list.txt from the exchange ban list every trading day.
BAN_FILE = "config/ban_list.txt"


def load_ban_list(path=BAN_FILE):
    import os
    out = set()
    if os.path.exists(path):
        with open(path) as f:
            for ln in f:
                ln = ln.split("#", 1)[0].strip().upper()
                if ln:
                    out.add(ln)
    return out

INDEX_SYMBOLS = {
    "NIFTY", "BANKNIFTY", "FINNIFTY", "MIDCPNIFTY", "MIDCAPNIFTY", "NIFTYNXT50",
    "NIFTYNEXT50", "BANKEX", "SENSEX", "SENSEX50", "NIFTYIT", "NIFTYINFRA",
}


def price_gate(price_state, score):
    """Price/position interaction. Only meaningful when the price component is live.
    price_state in {'up','down','flat'} from the 10-day window
      ('up' = LTP > window avg AND slope > 0; 'down' = LTP < avg AND slope < 0;
       'flat' = mixed).
    Returns 'ok' | 'hold_only' | 'eliminate':
      price down + retail position increasing  (score >= +FLAT_BAND) -> eliminate
      price down + retail position flat        (|score| <  FLAT_BAND) -> hold_only
                                                 (keep if held, no fresh BUY)
      price down + retail actively unwinding   (score <= -FLAT_BAND) -> ok (price_c=0)
      price up / flat                                               -> ok
    """
    if price_state != "down":
        return "ok"
    if score >= FLAT_BAND:
        return "eliminate"
    if score <= -FLAT_BAND:
        return "ok"
    return "hold_only"


def _is_tradable_symbol(sym: str) -> bool:
    s = (sym or "").strip().upper()
    if not s or s in INDEX_SYMBOLS or s == "TOTAL":
        return False
    if not s[0].isalpha():
        return False
    return all(c.isalnum() or c in "&-." for c in s) and 2 <= len(s) <= 15


def lsq_r2(y):
    """R^2 of an OLS line fit, y given oldest->newest. Flat/degenerate -> 0.0."""
    n = len(y)
    if n < 3:
        return 0.0
    xs = list(range(n))
    mx = sum(xs) / n
    my = sum(y) / n
    sxx = sum((x - mx) ** 2 for x in xs)
    syy = sum((v - my) ** 2 for v in y)
    sxy = sum((xs[i] - mx) * (y[i] - my) for i in range(n))
    if sxx == 0 or syy == 0:
        return 0.0
    return (sxy * sxy) / (sxx * syy)


def trend(series):
    """series newest-first. Returns the spec's measures."""
    n = len(series)
    np_now = series[0]
    np_old = series[-1]
    score = (np_now - np_old) / abs(np_old) if abs(np_old) > 1e-9 else 0.0
    r2 = lsq_r2(list(reversed(series)))  # fit oldest->newest
    k = min(8, n - 1)
    base = series[k]
    recent = (np_now - base) / abs(base) if abs(base) > 1e-9 else 0.0
    medlev = st.median([abs(v) for v in series])
    return dict(n=n, np_now=np_now, np_old=np_old, score=score, r2=r2,
               recent=recent, medlev=medlev)


def regime(nifty_score):
    if nifty_score <= -0.15:
        return "bullish", 65_00_00_000
    if nifty_score >= 0.50:
        return "extreme-bearish", 15_00_00_000
    return "neutral", 32_50_00_000


def lots_for(score, r2, conviction, low_part, np_now, np_old,
             price_c=0.0, price_live=False):
    # SUPER-BULL needs BOTH a confirmed positive price uptrend AND huge retail
    # unwinding. With the price component OFF, "positive uptrend" is unprovable,
    # so SUPER-BULL cannot fire at all until price_c is live.
    huge_unwind = (not low_part and np_now < 5 and (np_old - np_now) >= 10
                   and score <= -0.60)
    super_bull = huge_unwind and price_live and price_c > 0
    if super_bull:
        return 80, True
    if r2 < 0.40 or conviction < 0.25:
        base = 10
    elif conviction <= 0.55:
        base = 20 + (conviction - 0.25) / 0.30 * 30
    else:
        base = 50 + min((conviction - 0.55) / 0.30, 1) * 30
    base = int(round(base))
    if low_part:
        base = min(base, 20)
    return base, False


def _opt_layer(feed_path, prev_tag=None):
    """Per-symbol contrarian options signal from the `strike` sheet.
      opt_score = retail_dir / (|call_bias| + |put_bias| + 1)   ~ [-1, 1]
                  > 0 options CONFIRM the contrarian long; < 0 against it
      reducing  = is the adverse retail position (long calls + short puts)
                  smaller than in prev_tag's snapshot?  (None if no history)
    """
    import strike_analysis as SA
    strikes = SA.load_strike(feed_path)
    ltp = SA.load_ltp(feed_path)
    prev = SA._load_snapshot(prev_tag) if prev_tag else {}

    def adverse(strk, spot):
        tot = 0.0
        for k, d in strk.items():
            if k == 0:
                continue
            nc = (d["co_b"] - d["co_s"]) * SA.mny_weight(k, spot, "C")
            npp = (d["po_b"] - d["po_s"]) * SA.mny_weight(k, spot, "P")
            if nc > 0:
                tot += nc          # retail long calls  (bullish bet -> bearish fade)
            if npp < 0:
                tot += -npp        # retail short puts   (-> bearish fade)
        return tot

    out = {}
    for sym, ks in strikes.items():
        spot = ltp.get(sym)
        if not spot:
            continue
        bc = SA.flag_broker_calls(ks, spot)
        b = SA.band(sym, ks, spot, exclude=set(bc))
        activity = abs(b["call_bias"]) + abs(b["put_bias"])
        opt_score = b["retail_dir"] / (activity + 1.0)
        reducing = None
        if sym in prev:
            reducing = adverse(ks, spot) < adverse(prev[sym], spot) - 1e-9
        out[sym] = dict(opt_score=opt_score, retail_dir=b["retail_dir"],
                        activity=activity, ceiling=b["ceiling"], floor=b["floor"],
                        reducing=reducing)
    return out


def run(path, sheet="pos", header_rows=2, price_c_on=False,
        strike_feed=None, prev_strike_tag=None, ban_list=None):
    feed = parse_pos(path, sheet=sheet, header_rows=header_rows)
    recs = {r.symbol.upper(): r for r in feed.records}
    opt = _opt_layer(strike_feed or path, prev_strike_tag)
    ban = {s.upper() for s in ban_list} if ban_list is not None else load_ban_list()

    # ---- regime from NIFTY ----
    nifty = recs.get("NIFTY")
    if not nifty or len(nifty.net_series) < 3:
        raise SystemExit("NIFTY row missing or too short -- cannot set regime")
    nt = trend(nifty.net_series)
    reg_name, reg_cap = regime(nt["score"])

    rows = []
    for r in feed.records:
        if not _is_tradable_symbol(r.symbol):
            continue
        s = r.net_series
        if len(s) < 1:
            continue
        t = trend(s)
        elig_fut = (t["score"] <= -0.15 and abs(t["np_old"]) >= 10
                    and t["recent"] <= 0.15 and t["n"] >= 3)
        score_c = min(abs(t["score"]) / 1.0, 1.0)
        price_c = 0.0 if not price_c_on else None

        o = opt.get(r.symbol.upper())
        opt_score = o["opt_score"] if o else 0.0
        opt_reducing = o["reducing"] if o else None
        opt_act = o["activity"] if o else 0.0
        thin_opt = opt_act < OPT_MIN_ACT           # option chain too thin to weigh
        opt_c = 0.0 if thin_opt else max(0.0, opt_score)   # confirmation only
        # options veto: bearish options signal kills the name unless the adverse
        # retail position (long calls / short puts) is reducing over readings.
        # Suppressed for thin chains.
        opt_veto = (not thin_opt and opt_score <= OPT_VETO
                    and opt_reducing is not True)
        banned = r.symbol.upper() in ban
        eligible = elig_fut and not opt_veto and not banned

        pc = price_c if price_c_on else 0.0
        conv = (CONV_W_NETPOS * score_c + CONV_W_OPT * opt_c
                + CONV_W_PRICE * pc)
        low_part = t["medlev"] < 10
        lots, sb = lots_for(t["score"], t["r2"], conv, low_part,
                            t["np_now"], t["np_old"],
                            price_c=price_c, price_live=price_c_on)
        rows.append(dict(
            symbol=r.symbol, sector=r.sector, ltp=r.ltp,
            n=t["n"], np_now=t["np_now"], np_old=t["np_old"],
            score=round(t["score"], 4), r2=round(t["r2"], 4),
            recent=round(t["recent"], 4), medlev=t["medlev"],
            score_c=round(score_c, 4),
            opt_score=round(opt_score, 3), opt_c=round(opt_c, 3),
            opt_reducing=opt_reducing, opt_veto=opt_veto, banned=banned,
            conviction=round(conv, 4),
            low_part=low_part, super_bull=sb, lots_raw=lots,
            eligible=eligible, elig_fut=elig_fut, verified=r.verified))

    elig = [x for x in rows if x["eligible"]]
    elig.sort(key=lambda x: (not x["super_bull"], -x["conviction"], x["score"]))

    # ---- rank-and-fill to regime cap ----
    book, running = [], 0
    for x in elig:
        want = x["lots_raw"] * LOT_RUPEES
        if running + want <= reg_cap:
            book.append({**x, "lots": x["lots_raw"], "value": want})
            running += want
        else:
            leftover = reg_cap - running
            lft_lots = leftover // LOT_RUPEES
            if lft_lots >= 10:
                book.append({**x, "lots": int(lft_lots),
                             "value": int(lft_lots) * LOT_RUPEES})
                running += int(lft_lots) * LOT_RUPEES
            break

    return dict(
        feed=feed, nifty_trend=nt, regime=reg_name, regime_cap=reg_cap,
        rows=rows, eligible=elig, book=book, deployed=running, ban_list=sorted(ban),
        mismatches=feed.mismatches, warnings=feed.warnings)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("path")
    ap.add_argument("--header-rows", default="2")
    ap.add_argument("--strike-feed", default=None,
                    help="feed .xlsx for the `strike` sheet (default: same as path)")
    ap.add_argument("--prev-strike-tag", default=None,
                    help="prior reading tag under archive/strike_archive for the "
                         "adverse-position-reducing check")
    ap.add_argument("--ban", default=None,
                    help="comma-separated F&O ban symbols (overrides "
                         "config/ban_list.txt)")
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args()
    bl = [s for s in a.ban.split(",") if s.strip()] if a.ban is not None else None
    out = run(a.path, header_rows=int(a.header_rows),
              strike_feed=a.strike_feed, prev_strike_tag=a.prev_strike_tag,
              ban_list=bl)

    nt = out["nifty_trend"]
    print(f"# feed: {out['feed'].path}")
    print(f"# rows parsed: {len(out['feed'].records)}   "
          f"col-15 mismatches: {len(out['mismatches'])}   "
          f"warnings: {len(out['warnings'])}")
    for m in out["mismatches"][:30]:
        print(f"  MISMATCH r{m.row} {m.symbol}: expected={m.expected} actual={m.actual} [{m.note}]")
    for w in out["warnings"]:
        print(f"  WARN {w}")
    print(f"\n# NIFTY regime trend: series[{nt['n']}] "
          f"np_now={nt['np_now']} np_old={nt['np_old']} "
          f"score={nt['score']:+.3f} r2={nt['r2']:.3f}")
    print(f"# REGIME: {out['regime']}   book cap = Rs {out['regime_cap']:,}")
    if out["ban_list"]:
        bl = ", ".join(out["ban_list"])
        hit = [x["symbol"] for x in out["rows"] if x["elig_fut"] and x["banned"]]
        print(f"# F&O BAN (reduce-only, no new BUY): {bl}"
              + (f"   -> excluded from book: {', '.join(hit)}" if hit else ""))

    vetoed = [x for x in out["rows"] if x["elig_fut"] and x["opt_veto"]]
    if vetoed:
        print(f"\n# OPTIONS-VETOED (futures-eligible but options contrarian-bearish "
              f"and adverse position not reducing): {len(vetoed)}")
        for x in sorted(vetoed, key=lambda x: x["opt_score"]):
            print(f"  {x['symbol']:<14} score={x['score']:+.3f}  "
                  f"opt_score={x['opt_score']:+.3f}  reducing={x['opt_reducing']}")

    print(f"\n# ELIGIBLE LONGS (conviction = 1/3 netpos + 1/3 options + 1/3 price; "
          f"price_c=0 -> cap ~0.667): {len(out['eligible'])}")
    hdr = ("rank sym            sect         score  score_c  opt_s  opt_c   "
           "conv   lots  tag")
    print(hdr)
    for i, x in enumerate(out["eligible"], 1):
        tag = []
        if x["super_bull"]:
            tag.append("SUPER-BULL")
        if x["low_part"]:
            tag.append("low-participation")
        if x["opt_reducing"]:
            tag.append("opt-adverse-reducing")
        if not x["verified"]:
            tag.append("UNVERIFIED-col15")
        print(f"{i:>3}  {x['symbol']:<14} {str(x['sector'])[:11]:<11}  "
              f"{x['score']:+.3f}  {x['score_c']:.3f}  {x['opt_score']:+.3f}  "
              f"{x['opt_c']:.3f}  {x['conviction']:.3f}  {x['lots_raw']:>3}  "
              f"{','.join(tag)}")

    print(f"\n# BOOK after rank-and-fill to Rs {out['regime_cap']:,} "
          f"(1 lot = Rs {LOT_RUPEES:,})")
    print("sym            lots   score    R2     conv    Rs value        tag")
    tot = 0
    for x in out["book"]:
        tot += x["value"]
        tag = []
        if x["super_bull"]:
            tag.append("SUPER-BULL")
        if x["low_part"]:
            tag.append("low-participation")
        print(f"{x['symbol']:<14} {x['lots']:>4}   {x['score']:+.3f}  "
              f"{x['r2']:.3f}  {x['conviction']:.3f}  Rs {x['value']:>12,}  "
              f"{','.join(tag)}")
    print(f"{'TOTAL':<14} {'':>4}   {'':6} {'':6} {'':6} Rs {tot:>12,}  "
          f"({tot/out['regime_cap']*100:.1f}% of cap)")

    if a.json:
        slim = {k: v for k, v in out.items() if k not in ("feed",)}
        slim["mismatches"] = [vars(m) for m in out["mismatches"]]
        print("\n" + json.dumps(slim, default=str, indent=1))
