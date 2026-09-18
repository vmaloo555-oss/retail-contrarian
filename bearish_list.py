"""
bearish_list.py -- contrarian-BEARISH candidates: net-position trend first,
max pain layered on top.

The book is LONG-ONLY, so this is the OPTION-STRUCTURE side only. It is the
mirror of the model's long gate, not a new rule:

    long  (model.py:210)   score <= -0.15  and |np_old| >= 10
                           and recent <=  0.15  and n >= 3
    bear  (here)           score >= +0.15  and |np_old| >= 10
                           and recent >= -0.15  and n >= 3

Reading: retail is BUILDING net length (score positive), not unwinding it, and
the build has not already started to reverse (recent). |np_old| >= 10 is the
same floor the model uses and is what kills the near-zero-denominator scores
that blew up to +32 and +12 on the ad-hoc ranking.

STAGE 1  net-position trend  -- the gate, and the primary sort.
         trend_strength = min(score, 3.0) * (0.5 + 0.5 * r2)
         r2 is the OLS fit of the series oldest->newest, so a steady build
         ranks above a jumpy one of the same magnitude.

STAGE 2  max pain           -- layered on, never promotes a name into the list.
         pain BELOW spot  -> confirms (a fall is what hurts most clients)
         pain ABOVE spot  -> contradicts, kept but flagged
         fall% to the pain strike is reported, never scored: a 30% gap is not
         three times better than a 10% gap, it is three times less likely.

STAGE 3  options read       -- reported alongside, for agreement or conflict.

    python bearish_list.py FEED.xlsx [--min-positions 25] [--top 30]
"""
from __future__ import annotations
import argparse

import model as M
import max_pain as MP
import strike_analysis as SA

FLAT = 0.15          # same band the model uses
MIN_NP_OLD = 10      # same floor as the model's long gate
MIN_N = 3
SCORE_CAP = 3.0      # beyond this, more "building" tells us nothing extra


def bear_gate(r):
    """Mirror of model.py's elig_fut, flipped to the building side."""
    return (r["score"] >= FLAT
            and abs(r["np_old"]) >= MIN_NP_OLD
            and r["recent"] >= -FLAT
            and r["n"] >= MIN_N)


def trend_strength(r):
    return min(r["score"], SCORE_CAP) * (0.5 + 0.5 * max(r["r2"], 0.0))


def build(feed, min_positions=25):
    out = M.run(feed, strike_feed=feed)
    mp = {x["symbol"]: x for x in MP.analyse(feed)}
    strikes = SA.load_strike(feed)
    ltp = SA.load_ltp(feed)

    rows = []
    for r in out["rows"]:
        sym = r["symbol"]
        if r.get("banned") or not M._is_tradable_symbol(sym):
            continue
        if not bear_gate(r):
            continue                      # STAGE 1 - trend is the gate

        m = mp.get(sym)
        spot = ltp.get(sym) or r.get("ltp")
        pain = fall = None
        pain_note = "no chain"
        tot = 0.0
        if m:
            tot = float(m["total_positions"] or 0)
            spot = float(m["spot"] or 0) or spot
            pain = float(m["PAIN_STRIKE"])
            fall = (spot - pain) / spot * 100 if spot else None
            if tot < min_positions:
                pain_note = f"thin chain ({tot:.0f})"
            elif pain < spot:
                pain_note = "confirms"       # STAGE 2
            elif pain > spot:
                pain_note = "CONTRADICTS - pain sits above spot"
            else:
                pain_note = "pain at spot"

        # STAGE 3 - what the option book says, for agreement / conflict
        opt = r.get("opt_score", 0.0)
        opt_note = ("options agree" if opt < -0.10 else
                    "options DISAGREE (contrarian-bullish)" if opt > 0.10 else
                    "options neutral")

        rows.append(dict(
            symbol=sym, strength=trend_strength(r), score=r["score"],
            r2=r["r2"], recent=r["recent"], np_now=r["np_now"],
            np_old=r["np_old"], n=r["n"], spot=spot, pain=pain, fall=fall,
            positions=tot, losers=float(m["losers_at_pain"]) if m else None,
            incl_fut=float(m.get("losers_incl_fut") or 0) if m else None,
            pct=float(m["pct_losing"]) if m else None,
            extra=float(m["extra_losers"]) if m else None,
            opt_score=opt, opt_note=opt_note, pain_note=pain_note,
            edge=(m or {}).get("edge", "")))

    rows.sort(key=lambda x: -x["strength"])
    return rows, out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("feed")
    ap.add_argument("--min-positions", type=int, default=25)
    ap.add_argument("--top", type=int, default=30)
    a = ap.parse_args()

    rows, out = build(a.feed, a.min_positions)
    conf = [r for r in rows if r["pain_note"] == "confirms"]
    print(f"NIFTY score {out['nifty_trend']['score']:+.3f} -> {out['regime']}")
    print(f"passed the net-pos bear gate: {len(rows)}   "
          f"of which max pain confirms: {len(conf)}")
    print()
    hdr = (f"{'sym':<12}{'stren':>6}{'score':>7}{'r2':>6}{'recent':>8}"
           f"{'np now/old':>12}{'spot':>10}{'PAIN':>10}{'fall%':>7}"
           f"{'lose':>6}{'of':>5}  pain / options")
    print(hdr)
    for r in rows[:a.top]:
        po = f"{r['np_now']:.0f}/{r['np_old']:.0f}"
        pain = f"{r['pain']:g}" if r["pain"] is not None else "-"
        fall = f"{r['fall']:.1f}" if r["fall"] is not None else "-"
        lose = f"{r['losers']:.0f}" if r["losers"] is not None else "-"
        tot = f"{r['positions']:.0f}" if r["positions"] else "-"
        print(f"{r['symbol']:<12}{r['strength']:>6.2f}{r['score']:>+7.2f}"
              f"{r['r2']:>6.2f}{r['recent']:>+8.2f}{po:>12}"
              f"{(r['spot'] or 0):>10.2f}{pain:>10}{fall:>7}"
              f"{lose:>6}{tot:>5}  {r['pain_note']} / {r['opt_note']}")


if __name__ == "__main__":
    main()
