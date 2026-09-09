"""
strike_analysis.py -- retail option positioning by strike -> contrarian price band.

Source: the `strike` sheet (Excel pivot "Count of Net Qty."). Row labels are
"SYMBOL   STRIKE" (strike 0 = the FUT/cash leg). Columns:
    0 RowLabels
    1 FUT b   2 FUT s
    4 CALL b  5 CALL s          (col 3 = FUT total, 6 = CALL total)
    7 PUT  b  8 PUT  s          (col 9 = PUT total, 10 = grand total)
Values are COUNTS (how many retail long vs short at that strike), small ints.

Per strike:  netC = call_b - call_s ,  netP = put_b - put_s
  netC > 0  retail net LONG calls   -> contrarian: strike acts as a CEILING (decays)
  netC < 0  retail net SHORT calls  -> contrarian: strike likely BREAKS (shorts run over)
  netP > 0  retail net LONG puts    -> contrarian: strike acts as a FLOOR
  netP < 0  retail net SHORT puts   -> contrarian: strike likely BREAKS DOWN

The `pos` net position already accounts for total retail exposure; this overlay
is about the SHAPE / range only. Strikes are moneyness-weighted: in-the-money
strikes are more relevant than out-of-the-money (real delta / capital / market
impact vs cheap lottery tickets), so each strike's net count is multiplied by
`mny_weight()` -- full-or-bonus weight at/inside the money, exponential decay as
it goes OTM -- before walls, skews and the name-level lean are computed.

Band for a stock at spot S (LTP from `pos`):
  ceiling = lowest strike > S with netC >= MIN_NET      (nearest retail long-call wall)
  floor   = highest strike < S with netP >= MIN_NET     (nearest retail long-put wall)
  up-skew   if the nearest strike above S is net SHORT calls  (no wall / breakout fuel)
  down-skew if the nearest strike below S is net SHORT puts    (no floor / air below)

NOTE: strike positioning changes every reading. Snapshot every file into
archive/strike_archive/<reading>.csv and diff strike-by-strike over time.
"""
from __future__ import annotations
import argparse, csv, glob, os
import openpyxl

MIN_NET = 1          # weighted |net| >= this to treat a strike as a "wall". Tune.
ARCHIVE_DIR = "archive/strike_archive"

# Moneyness weighting: ITM strikes are weighted HIGHER because an ITM retail
# position carries intrinsic value -> a bigger P&L outflow for retail when the
# contrarian fade plays out -> a bigger contrarian signal. OTM = cheap, small
# outflow, decays to noise.
#   signed depth x  (positive = ITM for that option type):
#     call at K, spot S -> x = (S - K) / S
#     put  at K, spot S -> x = (K - S) / S
#   x >= 0 (ATM/ITM): w = 1 + ITM_BONUS * min(x, ITM_CAP)   -> 1.0 ATM .. 2.0 deep ITM
#   x <  0 (OTM):      w = exp(x / OTM_TAU)                  -> 5% OTM~0.43, 10%~0.19
ITM_BONUS = 2.0
ITM_CAP = 0.50
OTM_TAU = 0.06


def mny_weight(strike, spot, typ):
    if not spot or strike <= 0:
        return 1.0
    x = (spot - strike) / spot if typ == "C" else (strike - spot) / spot
    if x >= 0:
        return 1.0 + ITM_BONUS * min(x, ITM_CAP)
    import math
    return math.exp(x / OTM_TAU)

# Concentrated single-strike activity is more likely a BROKER CALL (many clients
# put on the same tip) than organic dispersed retail -> do not fade it as a
# retail wall; surface it separately. A strike is flagged when its gross option
# count is BOTH a large multiple of the name's median strike AND a large share
# of the whole chain, and it is not just the at-the-money strike (ATM
# concentration is normal).
BROKER_MULT = 3.0        # gross >= this * median non-zero strike gross
BROKER_SHARE = 0.30      # gross >= this * total option activity in the name
BROKER_MIN_GROSS = 8     # and gross >= this absolute floor (thin chains are not
                         #   "broker calls" just because one strike dominates)
ATM_PCT = 0.02           # strikes within +/- this of spot are ATM -> exempt

# A "pin" is only neutral if spot sits roughly BETWEEN retail's two walls. When
# spot is much closer to one of them, a short strangle at those strikes carries
# real directional delta (the near leg is ~ATM, the far leg is not) and is a
# disguised bet, not a range trade. Past this ratio, take the directional side
# the options read supports instead of pretending it is neutral.
PIN_SKEW_MAX = 2.0

# Minimum one-sidedness before the option book counts as directional at all:
# |bull - bear| / (bull + bear). Below this the chain is two-sided -> no trade.
DIR_MIN_CONVICTION = 0.20


def _n(x):
    return 0.0 if x in (None, "") else float(x)


def load_strike(path, sheet="strike"):
    """-> {symbol: {strike: {co_b,co_s,po_b,po_s,fut_b,fut_s}}}  (strike 0 kept)."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet]
    out = {}
    for r in ws.iter_rows(values_only=True):
        lab = r[0]
        if not lab or str(lab).strip() in ("Count of Net Qty.", "RowLabels"):
            continue
        parts = str(lab).split()
        if len(parts) < 2:
            continue
        try:
            strike = float(parts[-1])
        except ValueError:
            continue
        sym = " ".join(parts[:-1]).strip().upper()
        out.setdefault(sym, {})[strike] = dict(
            fut_b=_n(r[1]), fut_s=_n(r[2]),
            co_b=_n(r[4]), co_s=_n(r[5]), po_b=_n(r[7]), po_s=_n(r[8]))
    wb.close()
    return out


def load_ltp(path, sheet="pos"):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet]
    out = {}
    for row in list(ws.iter_rows(values_only=True))[2:]:
        if row and row[0] and isinstance(row[2], (int, float)):
            out[str(row[0]).strip().upper()] = float(row[2])
    wb.close()
    return out


def _median(xs):
    xs = sorted(xs)
    n = len(xs)
    if not n:
        return 0.0
    return xs[n // 2] if n % 2 else (xs[n // 2 - 1] + xs[n // 2]) / 2


def flag_broker_calls(strikes, spot, mult=BROKER_MULT, share=BROKER_SHARE,
                      min_gross=BROKER_MIN_GROSS, atm_pct=ATM_PCT):
    """-> {strike: {'gross':g, 'share':g/total}} for strikes that look like a
    broker call rather than organic retail: concentrated (large multiple of the
    name's median strike AND a large share of the whole chain), sizeable in
    absolute terms, and not just the at-the-money strike."""
    gross = {}
    for k, d in strikes.items():
        if k == 0:
            continue
        gross[k] = d["co_b"] + d["co_s"] + d["po_b"] + d["po_s"]
    total = sum(gross.values())
    if total <= 0:
        return {}
    med = _median([g for g in gross.values() if g > 0]) or 1.0
    flagged = {}
    for k, g in gross.items():
        if spot and abs(k - spot) <= atm_pct * spot:
            continue                       # ATM concentration is normal
        if g >= min_gross and g >= mult * med and g >= share * total:
            flagged[k] = {"gross": g, "share": g / total}
    return flagged


def band(symbol, strikes, spot, min_net=MIN_NET, exclude=None):
    exclude = exclude or set()
    rows = []          # (K, netC, netP, wnetC, wnetP)  weighted = raw * mny_weight
    for k in sorted(strikes):
        if k == 0:
            continue
        d = strikes[k]
        netc = d["co_b"] - d["co_s"]
        netp = d["po_b"] - d["po_s"]
        wc = netc * mny_weight(k, spot, "C")
        wp = netp * mny_weight(k, spot, "P")
        rows.append((k, netc, netp, wc, wp))
    usable = [t for t in rows if t[0] not in exclude]     # drop broker-call strikes
    above = [t for t in usable if t[0] > spot]
    below = [t for t in usable if t[0] < spot]
    # band WALLS from retail LONG options (their bought OTM options pin/decay)
    ceiling = next((k for k, _, _, wc, _ in above if wc >= min_net), None)
    floor = next((k for k, _, _, _, wp in reversed(below) if wp >= min_net), None)
    up_skew = bool(above) and above[0][3] <= -min_net       # nearest above = short calls
    down_skew = bool(below) and below[-1][4] <= -min_net    # nearest below = short puts

    # DIRECTIONAL lean -- FULLY CONTRARIAN, fade every retail option position:
    #   retail long  calls (bullish bet)     -> fade -> BEARISH  (-wNetC)
    #   retail short calls (capping bet)      -> fade -> BULLISH  (-wNetC)
    #   retail long  puts  (bearish bet)      -> fade -> BULLISH  (+wNetP)
    #   retail short puts  (floor / complacent)-> fade -> BEARISH (+wNetP)
    #   =>  retail_dir = Sum(wNetP) - Sum(wNetC)   (moneyness-weighted)
    #   retail_dir > 0  -> contrarian BULLISH (confirms / can upsize a long)
    #   retail_dir < 0  -> contrarian BEARISH (caution / veto a long)
    call_bias = sum(wc for *_, wc, _ in usable)
    put_bias = sum(wp for *_, wp in usable)
    retail_dir = put_bias - call_bias
    return dict(symbol=symbol, spot=spot, ceiling=ceiling, floor=floor,
               call_bias=call_bias, put_bias=put_bias, retail_dir=retail_dir,
               up_skew=up_skew, down_skew=down_skew, rows=rows,
               broker_calls=sorted(exclude))


def verdict(b):
    lo = f"{b['floor']:.0f}" if b["floor"] else ("open" if b["down_skew"] else "?")
    hi = f"{b['ceiling']:.0f}" if b["ceiling"] else ("open" if b["up_skew"] else "?")
    tags = []
    if b["floor"] and b["ceiling"]:
        tags.append(f"range-bound {lo}-{hi}")
    if b["down_skew"]:
        tags.append(f"floor fragile (retail short puts at {b['rows'][0][0]:.0f}+)")
    if b["up_skew"]:
        tags.append("upside open (retail short calls overhead)")
    if b["ceiling"] and not b["up_skew"]:
        tags.append(f"capped at {hi} (retail long calls)")
    if b["broker_calls"]:
        bc = ", ".join(f"{k:.0f}" for k in b["broker_calls"])
        tags.append(f"broker-call cluster at {bc} — omitted")
    # fully-contrarian weighted lean: retail_dir = Sum(wNetP) - Sum(wNetC).
    rd = b["retail_dir"]
    side = "contrarian BULLISH" if rd > 0 else "contrarian BEARISH" if rd < 0 else "neutral"
    tags.append(f"retail_dir {rd:+.1f}  (wNetC {b['call_bias']:+.1f} / wNetP {b['put_bias']:+.1f})"
                f" -> {side}")
    return f"{lo} - {hi}   [{'; '.join(tags) or 'no clear structure'}]"


def weighted_rows(symbol, strikes, spot, min_net=MIN_NET):
    """Per-strike rows in the operator's own `strike`-sheet pivot layout
    (RowLabels / b / s / Total, grouped FUT | CO | PO | GrandTotal), with the
    moneyness weighting applied to the option groups. FUT (strike 0) is weight
    1.0. Weighted columns are suffixed w; raw counts are kept alongside so the
    weighting is auditable."""
    bc = flag_broker_calls(strikes, spot)
    out = []
    for k in sorted(strikes):
        d = strikes[k]
        label = f"{symbol}   {k:g}"
        if k == 0:                                   # futures leg, weight 1.0
            fb, fs = d["fut_b"], d["fut_s"]
            out.append([label, symbol, 0, None, 1.0, 1.0,
                        fb, fs, fb - fs, round(fb - fs, 2),
                        0, 0, 0, 0.0, 0, 0, 0, 0.0,
                        round(fb - fs, 2), "FUT leg (weight 1.0)"])
            continue
        cw, pw = mny_weight(k, spot, "C"), mny_weight(k, spot, "P")
        netc, netp = d["co_b"] - d["co_s"], d["po_b"] - d["po_s"]
        wc, wp = netc * cw, netp * pw
        if k in bc:
            note = "BROKER-CALL cluster (omitted)"
        elif k > spot and wc >= min_net:
            note = "retail long calls -> ceiling"
        elif k > spot and wc <= -min_net:
            note = "retail short calls -> breakout up"
        elif k < spot and wp >= min_net:
            note = "retail long puts -> floor"
        elif k < spot and wp <= -min_net:
            note = "retail short puts -> break down"
        else:
            note = ""
        out.append([label, symbol, k,
                    round((k - spot) / spot * 100, 1) if spot else None,
                    round(cw, 3), round(pw, 3),
                    0, 0, 0, 0.0,                              # no FUT at a strike
                    d["co_b"], d["co_s"], netc, round(wc, 2),
                    d["po_b"], d["po_s"], netp, round(wp, 2),
                    round(wc + wp, 2), note])
    return out


WEIGHTED_HEADERS = ["RowLabels", "symbol", "strike", "vs_spot_%",
                    "call_wt", "put_wt",
                    "FUT b", "FUT s", "FUT net", "FUT netW",
                    "CO b", "CO s", "CO net", "CO netW",
                    "PO b", "PO s", "PO net", "PO netW",
                    "GrandTotalW", "note"]


import re as _re_mod


def _re_strikes(legs):
    """Strike numbers appearing in a legs string."""
    return _re_mod.findall(r"(?:Buy|Sell)\s+([\d.]+)\s+(?:CE|PE)", legs or "")


def suggest_strategy(symbol, strikes, spot, min_net=MIN_NET, netpos_score=None):
    """Two-step contrarian construction. Every output is a DEFINED STRUCTURE -
    never a single naked long option. If no second leg exists, there is no trade.

    STEP 1 - direction from the option book, in detail. Each retail position is
    scored on its moneyness-weighted size and then faded:
        retail LONG  calls -> a bullish bet that decays        -> BEARISH for us
        retail SHORT puts  -> "it won't fall", gets run over   -> BEARISH for us
        retail SHORT calls -> "it won't rise", gets run over   -> BULLISH for us
        retail LONG  puts  -> a bearish bet that decays        -> BULLISH for us
      bull_pressure = |short calls| + long puts
      bear_pressure = long calls + |short puts|
      conviction    = |bull - bear| / total

    STEP 2 - the structure, built on retail's own strikes:
      BEARISH  -> bear put spread (buy ATM PE, sell their wall below)
                  else bear call spread (sell their call wall, buy the wing)
      BULLISH  -> bull call spread (buy ATM CE, sell their wall above)
                  else bull put spread (sell their put wall, buy the wing)
      BALANCED -> they are net sellers of premium  -> long strangle (buy it back)
                  they are net buyers  of premium  -> iron condor (sell it to them)

    Mechanical output from positioning, not a market opinion; the operator
    places every order.
    """
    bc = flag_broker_calls(strikes, spot)
    b = band(symbol, strikes, spot, min_net, exclude=set(bc))
    ks = sorted(k for k in strikes if k > 0 and k not in bc)
    res = dict(symbol=symbol, spot=spot, floor=b["floor"], ceiling=b["ceiling"],
               netpos_score=(round(netpos_score, 4)
                             if netpos_score is not None else None),
               retail_dir=round(b["retail_dir"], 2),
               activity=round(abs(b["call_bias"]) + abs(b["put_bias"]), 2),
               bull_pressure=0.0, bear_pressure=0.0, verdict="", conviction=0.0,
               fav_strike=None, setup="", read="", strategy="none", legs="",
               alt="", caution="")
    if not ks:
        res["setup"] = "no chain"
        res["read"] = "no option chain"
        return res
    if res["activity"] < 3:
        res["setup"] = "thin"
        res["read"] = "chain too thin to fade"
        res["caution"] = "thin chain - ignore"
        return res

    wc = {k: (strikes[k]["co_b"] - strikes[k]["co_s"]) * mny_weight(k, spot, "C")
          for k in ks}
    wp = {k: (strikes[k]["po_b"] - strikes[k]["po_s"]) * mny_weight(k, spot, "P")
          for k in ks}
    above = [k for k in ks if k >= spot]
    below = [k for k in ks if k <= spot]
    atm_up = above[0] if above else None
    atm_dn = below[-1] if below else None

    def g(k):
        return f"{k:g}"

    def up_of(k):
        return next((x for x in above if x > k), None)

    def dn_of(k):
        return next((x for x in reversed(below) if x < k), None)

    # ---- STEP 1: direction ---------------------------------------------
    long_calls = sum(v for v in wc.values() if v > 0)
    short_calls = sum(-v for v in wc.values() if v < 0)
    long_puts = sum(v for v in wp.values() if v > 0)
    short_puts = sum(-v for v in wp.values() if v < 0)
    bull = short_calls + long_puts
    bear = long_calls + short_puts
    tot = bull + bear
    res["bull_pressure"] = round(bull, 2)
    res["bear_pressure"] = round(bear, 2)
    res["conviction"] = round(abs(bull - bear) / tot, 3) if tot > 0 else 0.0
    res["setup"] = (f"longC {long_calls:.1f} | shortC {short_calls:.1f} | "
                    f"longP {long_puts:.1f} | shortP {short_puts:.1f}")

    # walls retail bought (they overpaid, it decays -> we sell it)
    call_walls = sorted([k for k in above if wc[k] >= min_net],
                        key=lambda k: -wc[k])          # heaviest first
    put_walls = sorted([k for k in below if wp[k] >= min_net],
                       key=lambda k: -wp[k])
    call_wall = call_walls[0] if call_walls else None
    put_wall = put_walls[0] if put_walls else None
    # strikes retail sold (their pain point -> price is drawn there)
    call_pain = min([k for k in above if wc[k] <= -min_net],
                    key=lambda k: wc[k], default=None)
    put_pain = min([k for k in below if wp[k] <= -min_net],
                   key=lambda k: wp[k], default=None)

    # ---- balanced book -> a volatility structure, not a dead end --------
    if res["conviction"] < DIR_MIN_CONVICTION:
        res["verdict"] = "balanced"
        sold = short_calls + short_puts
        bought = long_calls + long_puts
        if sold > bought and call_pain and put_pain:
            res["read"] = (f"two-sided, but retail is a net SELLER of premium "
                           f"({sold:.1f} vs {bought:.1f}) -> buy back what they sold")
            res["strategy"] = "Long strangle at retail's own short strikes"
            res["legs"] = f"Buy {g(call_pain)} CE / Buy {g(put_pain)} PE"
            res["alt"] = (f"Tighter: long straddle at {g(atm_up)}"
                          if atm_up else "")
            res["fav_strike"] = call_pain
        elif bought >= sold and call_wall and put_wall:
            w_up, w_dn = up_of(call_wall), dn_of(put_wall)
            res["read"] = (f"two-sided, but retail is a net BUYER of premium "
                           f"({bought:.1f} vs {sold:.1f}) -> sell it to them")
            if w_up and w_dn:
                res["strategy"] = "Iron condor around retail's own walls"
                res["legs"] = (f"Sell {g(call_wall)} CE / Buy {g(w_up)} CE / "
                               f"Sell {g(put_wall)} PE / Buy {g(w_dn)} PE")
            else:
                res["strategy"] = "Short strangle at retail's own walls"
                res["legs"] = f"Sell {g(call_wall)} CE / Sell {g(put_wall)} PE"
                res["caution"] = "no wing strikes - undefined risk"
            res["fav_strike"] = call_wall
        else:
            res["read"] = (f"bull {bull:.1f} vs bear {bear:.1f} - two-sided and no "
                           f"usable pair of strikes")
        return res

    bullish = bull > bear
    res["verdict"] = "BULLISH" if bullish else "BEARISH"
    drivers = ([f"short calls {short_calls:.1f}"] if bullish and short_calls > 0 else
               [f"long calls {long_calls:.1f}"] if not bullish and long_calls > 0 else [])
    drivers += ([f"long puts {long_puts:.1f}"] if bullish and long_puts > 0 else
                [f"short puts {short_puts:.1f}"] if not bullish and short_puts > 0 else [])

    # ---- STEP 2: a defined structure on retail's strikes ----------------
    if bullish:
        # debit: buy ATM call, sell the wall above (theirs) or their pain strike
        sell_c = next((k for k in call_walls + ([call_pain] if call_pain else [])
                       if atm_up is not None and k > atm_up), None)
        if atm_up is not None and sell_c is not None:
            res["strategy"] = "Bull call spread (long call spread)"
            res["legs"] = f"Buy {g(atm_up)} CE / Sell {g(sell_c)} CE"
            res["fav_strike"] = sell_c
        elif put_wall is not None and dn_of(put_wall) is not None:
            res["strategy"] = "Bull put spread (credit)"
            res["legs"] = f"Sell {g(put_wall)} PE / Buy {g(dn_of(put_wall))} PE"
            res["fav_strike"] = put_wall
        # secondary
        if put_wall is not None and dn_of(put_wall) is not None and "Bull call" in res["strategy"]:
            res["alt"] = (f"Credit version: Sell {g(put_wall)} PE / "
                          f"Buy {g(dn_of(put_wall))} PE")
    else:
        sell_p = next((k for k in put_walls + ([put_pain] if put_pain else [])
                       if atm_dn is not None and k < atm_dn), None)
        if atm_dn is not None and sell_p is not None:
            res["strategy"] = "Bear put spread (debit)"
            res["legs"] = f"Buy {g(atm_dn)} PE / Sell {g(sell_p)} PE"
            res["fav_strike"] = sell_p
        elif call_wall is not None and up_of(call_wall) is not None:
            res["strategy"] = "Bear call spread (credit)"
            res["legs"] = f"Sell {g(call_wall)} CE / Buy {g(up_of(call_wall))} CE"
            res["fav_strike"] = call_wall
        if call_wall is not None and up_of(call_wall) is not None and "Bear put" in res["strategy"]:
            res["alt"] = (f"Credit version: Sell {g(call_wall)} CE / "
                          f"Buy {g(up_of(call_wall))} CE")

    res["read"] = (f"retail {' + '.join(drivers) or 'net one-sided'} -> fade -> "
                   f"{'up' if bullish else 'down'}"
                   + (f"; structure built to {g(res['fav_strike'])}"
                      if res["fav_strike"] is not None else ""))
    if not res["legs"]:
        res["strategy"] = "none"
        res["caution"] = "no pair of strikes to build a spread - no trade"
        return res

    ks_in = [float(x) for x in _re_strikes(res["legs"])]
    if len(ks_in) >= 2:
        width = abs(max(ks_in) - min(ks_in)) / spot * 100
        if width > 12:
            res["caution"] = f"spread is {width:.0f}% wide - low probability"

    if netpos_score is not None:
        nb = ("BULLISH" if netpos_score <= -0.15
              else "BEARISH" if netpos_score >= 0.15 else None)
        if nb and nb != res["verdict"]:
            res["caution"] = ((res["caution"] + "; ") if res["caution"] else "") + \
                f"net pos says {nb} ({netpos_score:+.2f}) - options disagree"
    if bc:
        res["caution"] = ((res["caution"] + "; ") if res["caution"] else "") + \
            "broker-call cluster omitted at " + ", ".join(f"{k:g}" for k in sorted(bc))
    return res


STRATEGY_HEADERS = ["symbol", "spot", "netpos_score", "opt_activity",
                    "bull_pressure", "bear_pressure", "verdict", "conviction",
                    "fav_strike", "activity breakdown", "contrarian read",
                    "strategy", "legs", "alternative", "caution"]


def snapshot(path, reading_tag):
    os.makedirs(ARCHIVE_DIR, exist_ok=True)
    data = load_strike(path)
    out = os.path.join(ARCHIVE_DIR, f"{reading_tag}.csv")
    with open(out, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["symbol", "strike", "co_b", "co_s", "po_b", "po_s",
                    "fut_b", "fut_s"])
        for sym in sorted(data):
            for k in sorted(data[sym]):
                d = data[sym][k]
                w.writerow([sym, k, d["co_b"], d["co_s"], d["po_b"], d["po_s"],
                            d["fut_b"], d["fut_s"]])
    return out


def _load_snapshot(tag):
    p = os.path.join(ARCHIVE_DIR, f"{tag}.csv")
    if not os.path.exists(p):
        return {}
    out = {}
    with open(p, newline="") as f:
        for row in csv.DictReader(f):
            out.setdefault(row["symbol"].upper(), {})[float(row["strike"])] = {
                k: float(row[k]) for k in
                ("co_b", "co_s", "po_b", "po_s", "fut_b", "fut_s")}
    return out


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("feed", help="2:30 feed .xlsx")
    ap.add_argument("symbols", nargs="*", help="symbols to analyse (default: a few)")
    ap.add_argument("--min-net", type=int, default=MIN_NET)
    ap.add_argument("--prev", help="prior reading tag under archive/strike_archive "
                                   "for strike-by-strike deltas")
    a = ap.parse_args()

    strikes = load_strike(a.feed)
    ltp = load_ltp(a.feed)
    prev = _load_snapshot(a.prev) if a.prev else {}
    syms = [s.upper() for s in a.symbols] or [
        "RELIANCE", "APOLLOHOSP", "MARICO", "SOLARINDS", "AUROPHARMA",
        "INDUSINDBK", "LICHSGFIN"]

    for sym in syms:
        if sym not in strikes or sym not in ltp:
            print(f"\n{sym}: not found in strike/pos"); continue
        bc = flag_broker_calls(strikes[sym], ltp[sym])
        b = band(sym, strikes[sym], ltp[sym], a.min_net, exclude=set(bc))
        print(f"\n===== {sym}   spot {b['spot']:.1f} =====")
        print(f"  contrarian band: {verdict(b)}")
        for k, info in bc.items():
            print(f"  ! broker-call? strike {k:.0f}: gross {info['gross']:.0f} "
                  f"({info['share']*100:.0f}% of the chain) — not faded as retail")
        print(f"  {'strike':>8} {'Cwt':>4} {'Pwt':>4} {'netC':>5} {'netP':>5} "
              f"{'wC':>6} {'wP':>6}  {'dC/dP':>7}  note")
        for k, nc, npp, wc, wp in b["rows"]:
            pk = prev.get(sym, {}).get(k)
            if pk:
                dc = nc - (pk["co_b"] - pk["co_s"])
                dp = npp - (pk["po_b"] - pk["po_s"])
                dd = f"{dc:+.0f}/{dp:+.0f}"
            else:
                dd = "new" if prev else ""
            cwt = mny_weight(k, b["spot"], "C")
            pwt = mny_weight(k, b["spot"], "P")
            if k in bc:
                note = "BROKER-CALL cluster (omitted)"
            elif k > b["spot"] and wc >= a.min_net: note = "long calls -> ceiling"
            elif k > b["spot"] and wc <= -a.min_net: note = "short calls -> breakout up"
            elif k < b["spot"] and wp >= a.min_net: note = "long puts -> floor"
            elif k < b["spot"] and wp <= -a.min_net: note = "short puts -> break down"
            else: note = ""
            print(f"  {k:>8.0f} {cwt:>4.2f} {pwt:>4.2f} {nc:>+5.0f} {npp:>+5.0f} "
                  f"{wc:>+6.1f} {wp:>+6.1f}  {dd:>7}  {note}")
