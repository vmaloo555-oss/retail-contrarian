"""
max_pain.py -- for every stock, the ONE settlement strike that hurts retail most.

Two different questions get asked with the same words, so both are answered:

  PAIN_STRIKE  (headline)  the strike where the greatest NUMBER of retail option
                           positions expire at a loss. A position loses when:
                             long  call @k  ->  S <= k   (expires worthless)
                             short call @k  ->  S >  k   (exercised against)
                             long  put  @k  ->  S >= k   (expires worthless)
                             short put  @k  ->  S <  k   (exercised against)
                           losers(S) is summed over every strike; we take argmax.

  VALUE_STRIKE (cross-check) classic max-pain on NET positions: the strike that
                           minimises the intrinsic value retail's book is worth
                             V(S) = SUM netC[k]*max(0,S-k) + SUM netP[k]*max(0,k-S)
                           where netC = call_b - call_s, netP = put_b - put_s.
                           argmin V. Unlike open-interest max pain this nets buys
                           against sells, so V can fall away without bound when
                           retail is net short one side - `edge` flags that.

Candidate settlement prices are the strikes actually present in that symbol's
chain. Ties resolve to the strike nearest spot, so the answer is always ONE
definite number.

Both measures ignore premium paid - they are expiry-intrinsic, exactly as
standard max pain is. The `strike` sheet holds COUNTS of retail positions, so
losers/total are counts of positions, not lots or rupees.

    python max_pain.py FEED.xlsx [-o out.csv]
"""
from __future__ import annotations
import argparse, csv, sys

import strike_analysis as SA


def _cands(chain):
    return sorted(k for k in chain if k > 0)


def losers_at(chain, S):
    """How many retail option positions expire at a loss if settlement = S."""
    n = 0.0
    for k, v in chain.items():
        if k <= 0:
            continue
        if S <= k:
            n += v["co_b"]          # long call worthless
        else:
            n += v["co_s"]          # short call run over
        if S >= k:
            n += v["po_b"]          # long put worthless
        else:
            n += v["po_s"]          # short put run over
    return n


def value_at(chain, S):
    """Intrinsic value of retail's NET option book if settlement = S."""
    v = 0.0
    for k, d in chain.items():
        if k <= 0:
            continue
        netC = d["co_b"] - d["co_s"]
        netP = d["po_b"] - d["po_s"]
        if S > k:
            v += netC * (S - k)
        if S < k:
            v += netP * (k - S)
    return v


def fut_leg(chain):
    """(long, short) count on the futures/cash leg, stored at strike 0."""
    d = chain.get(0.0) or chain.get(0) or {}
    return d.get("fut_b", 0.0), d.get("fut_s", 0.0)


def fut_losers_at(chain, S, spot):
    """Futures positions losing at settlement S.

    NOTE a different baseline from the option count: options are judged on
    expiry intrinsic against their own strike, futures have no strike and we
    do not get entry prices, so they are judged against TODAY'S spot - i.e.
    "loses if it moves from here to S". Reported separately for that reason.
    """
    fb, fs = fut_leg(chain)
    if not spot or S == spot:
        return 0.0
    return fb if S < spot else fs


def total_positions(chain):
    return sum(d["co_b"] + d["co_s"] + d["po_b"] + d["po_s"]
               for k, d in chain.items() if k > 0)


def _pick(scored, spot, want_max):
    """argmax/argmin with ties broken by nearest-to-spot -> one definite strike."""
    if not scored:
        return None, None
    best = max(s for _, s in scored) if want_max else min(s for _, s in scored)
    tied = [k for k, s in scored if s == best]
    if spot:
        tied.sort(key=lambda k: abs(k - spot))
    return tied[0], best


COLS = ["symbol", "spot", "PAIN_STRIKE", "pct_from_spot", "losers_at_pain",
        "total_positions", "pct_losing", "losers_at_spot", "extra_losers",
        "fut_long", "fut_short", "fut_losers_at_pain", "losers_incl_fut",
        "VALUE_STRIKE", "net_book_value_at_value_strike", "edge", "n_strikes"]


def analyse(feed, strike_sheet="strike", pos_sheet="pos"):
    strikes = SA.load_strike(feed, strike_sheet)
    ltp = SA.load_ltp(feed, pos_sheet)
    rows = []
    for sym, chain in sorted(strikes.items()):
        ks = _cands(chain)
        if len(ks) < 2:
            continue                      # no chain worth pricing
        spot = ltp.get(sym)
        tot = total_positions(chain)
        if tot <= 0 or not spot:
            continue

        pain, n_lose = _pick([(k, losers_at(chain, k)) for k in ks], spot, True)
        vstrike, vval = _pick([(k, value_at(chain, k)) for k in ks], spot, False)

        at_spot = losers_at(chain, spot) if spot else None
        fb, fs = fut_leg(chain)
        fut_lose = fut_losers_at(chain, pain, spot)
        rows.append({
            "symbol": sym,
            "spot": spot,
            "PAIN_STRIKE": pain,
            "pct_from_spot": (round((pain - spot) / spot * 100, 2)
                              if spot else None),
            "losers_at_pain": round(n_lose, 1),
            "total_positions": round(tot, 1),
            "pct_losing": round(n_lose / tot * 100, 1) if tot else None,
            "losers_at_spot": round(at_spot, 1) if at_spot is not None else None,
            "extra_losers": (round(n_lose - at_spot, 1)
                             if at_spot is not None else None),
            "fut_long": round(fb, 1),
            "fut_short": round(fs, 1),
            "fut_losers_at_pain": round(fut_lose, 1),
            "losers_incl_fut": round(n_lose + fut_lose, 1),
            "VALUE_STRIKE": vstrike,
            "net_book_value_at_value_strike": round(vval, 1),
            # the minimum sitting on the first/last strike means the true answer
            # is "at or beyond this strike" - retail is net short that wing
            "edge": "at chain edge" if vstrike in (ks[0], ks[-1]) else "",
            "n_strikes": len(ks),
        })
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("feed")
    ap.add_argument("-o", "--out", default=None)
    ap.add_argument("--top", type=int, default=25)
    a = ap.parse_args()

    rows = analyse(a.feed)
    out = a.out or "max_pain.csv"
    with open(out, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=COLS)
        w.writeheader()
        w.writerows(rows)

    rows.sort(key=lambda r: -(r["losers_at_pain"] or 0))
    print(f"{len(rows)} symbols -> {out}")
    print(f"{'symbol':<12}{'spot':>10}{'PAIN':>10}{'from spot':>11}"
          f"{'losing':>8}{'of':>7}{'%':>7}{'+fut':>7}  edge")
    for r in rows[:a.top]:
        print(f"{r['symbol']:<12}{r['spot'] or 0:>10.2f}{r['PAIN_STRIKE']:>10g}"
              f"{(r['pct_from_spot'] or 0):>10.1f}%{r['losers_at_pain']:>8.0f}"
              f"{r['total_positions']:>7.0f}{r['pct_losing']:>6.0f}%"
              f"{r['losers_incl_fut']:>7.0f}  {r['edge']}")


if __name__ == "__main__":
    main()
