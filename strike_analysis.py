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
    """Flat per-strike rows with the moneyness weighting applied, for the
    `strikes` sheet. One row per symbol x strike."""
    bc = flag_broker_calls(strikes, spot)
    out = []
    for k in sorted(strikes):
        if k == 0:
            continue
        d = strikes[k]
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
        out.append([symbol, spot, k, round((k - spot) / spot * 100, 1) if spot else None,
                    round(cw, 3), round(pw, 3),
                    d["co_b"], d["co_s"], netc, round(wc, 2),
                    d["po_b"], d["po_s"], netp, round(wp, 2), note])
    return out


WEIGHTED_HEADERS = ["symbol", "spot", "strike", "vs_spot_%", "call_wt", "put_wt",
                    "callB", "callS", "netC", "wNetC",
                    "putB", "putS", "netP", "wNetP", "note"]


def suggest_strategy(symbol, strikes, spot, min_net=MIN_NET):
    """Deterministic contrarian option structure implied by the strike map.

    Every retail position is faded -- retail long calls/puts decay (walls),
    retail short calls/puts get run over (breaks). This is mechanical output
    from the positioning, NOT a market opinion, and the operator places any
    order themselves. Structures are defined-risk spreads by default.
    """
    bc = flag_broker_calls(strikes, spot)
    b = band(symbol, strikes, spot, min_net, exclude=set(bc))
    ks = sorted(k for k in strikes if k > 0 and k not in bc)
    res = dict(symbol=symbol, spot=spot, floor=b["floor"], ceiling=b["ceiling"],
               retail_dir=round(b["retail_dir"], 2),
               activity=round(abs(b["call_bias"]) + abs(b["put_bias"]), 2),
               read="", strategy="none", legs="", alt="", caution="")
    if not ks:
        res["read"] = "no option chain"
        return res
    act = res["activity"]
    if act < 3:
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
    # short-side walls must sit STRICTLY beyond the long leg, or the "spread"
    # collapses to a single strike (zero width, no trade).
    short_calls = [k for k in above
                   if wc[k] <= -min_net and atm_up is not None and k > atm_up]
    short_puts = [k for k in below
                  if wp[k] <= -min_net and atm_dn is not None and k < atm_dn]
    rd = b["retail_dir"]

    if rd > 0:                       # retail leaning bearish -> fade -> BULLISH
        res["read"] = "retail net bearish on options -> contrarian BULLISH"
        if short_calls:
            tgt = min(short_calls, key=lambda k: wc[k])   # biggest short-call cluster
            res["strategy"] = "Bull call spread - fade retail's short calls"
            res["legs"] = f"Buy {atm_up:g} CE / Sell {tgt:g} CE"
        elif b["ceiling"] and atm_up and b["ceiling"] > atm_up:
            res["strategy"] = "Bull call spread - capped at retail's long-call wall"
            res["legs"] = f"Buy {atm_up:g} CE / Sell {b['ceiling']:g} CE"
        elif atm_up:
            res["strategy"] = "Long call"
            res["legs"] = f"Buy {atm_up:g} CE"
        if b["floor"] and atm_dn and b["floor"] < atm_dn:
            res["alt"] = f"Bull put spread: Sell {atm_dn:g} PE / Buy {b['floor']:g} PE"
    elif rd < 0:                     # retail leaning bullish -> fade -> BEARISH
        res["read"] = "retail net bullish on options -> contrarian BEARISH"
        if short_puts:
            tgt = min(short_puts, key=lambda k: wp[k])    # biggest short-put cluster
            res["strategy"] = "Put debit spread - fade retail's short puts"
            res["legs"] = f"Buy {atm_dn:g} PE / Sell {tgt:g} PE"
        elif b["floor"] and atm_dn and b["floor"] < atm_dn:
            res["strategy"] = "Put debit spread - down to retail's long-put floor"
            res["legs"] = f"Buy {atm_dn:g} PE / Sell {b['floor']:g} PE"
        elif atm_dn:
            res["strategy"] = "Long put"
            res["legs"] = f"Buy {atm_dn:g} PE"
        if b["ceiling"] and atm_up and b["ceiling"] >= atm_up:
            nxt = next((k for k in above if k > b["ceiling"]), None)
            res["alt"] = (f"Bear call spread: Sell {b['ceiling']:g} CE"
                          + (f" / Buy {nxt:g} CE" if nxt else " (uncapped - add a long call)"))
    else:
        res["read"] = "options neutral"

    if not res["legs"]:
        res["strategy"] = "none"
        res["caution"] = "no usable strikes on the required side"
    if bc:
        res["caution"] = (res["caution"] + "; " if res["caution"] else "") + \
            "broker-call cluster omitted at " + ", ".join(f"{k:g}" for k in sorted(bc))
    return res


STRATEGY_HEADERS = ["symbol", "spot", "floor", "ceiling", "retail_dir",
                    "opt_activity", "contrarian read", "strategy", "legs",
                    "alternative", "caution"]


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
