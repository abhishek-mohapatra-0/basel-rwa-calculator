"""Synthetic exposure book for two quarter-ends (Q2 and Q3 2025), amounts in INR crore."""
from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd

from .rules import RATING_BUCKET

RATINGS = list(RATING_BUCKET)
MIX = {"Sovereign": 0.06, "Bank": 0.10, "Corporate": 0.34, "Retail": 0.25, "Residential Mortgage": 0.25}


def _rating(cls, rng):
    if cls in ("Retail", "Residential Mortgage"):
        return "Unrated"
    p = {"Sovereign": [0.10, 0.05, 0.05, 0.25, 0.25, 0.05, 0.10, 0.10, 0.05, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
         "Bank":      [0, 0.03, 0.05, 0.07, 0.10, 0.12, 0.13, 0.15, 0.10, 0.08, 0.05, 0.03, 0.02, 0.01, 0.01, 0, 0, 0, 0, 0.05],
         "Corporate": [0, 0, 0.01, 0.02, 0.03, 0.05, 0.06, 0.08, 0.10, 0.12, 0.08, 0.07, 0.06, 0.04, 0.03, 0.02, 0.01, 0.01, 0, 0.21]}[cls]
    p = np.array(p) / sum(p)
    return rng.choice(RATINGS, p=p)


def _drawn(cls, rng):
    lo, hi = {"Sovereign": (200, 2500), "Bank": (50, 900), "Corporate": (10, 600),
              "Retail": (40, 600), "Residential Mortgage": (60, 900)}[cls]  # retail lines are pooled segments
    return round(float(np.exp(rng.uniform(np.log(lo), np.log(hi)))), 2)


def q2_book(n, rng):
    rows = []
    for i in range(n):
        cls = rng.choice(list(MIX), p=list(MIX.values()))
        drawn = _drawn(cls, rng)
        obt = rng.choice(["Other commitment", "Unconditionally cancellable", "Trade letter of credit",
                          "Guarantee / credit substitute", "No commitment"], p=[0.35, 0.15, 0.1, 0.1, 0.3]) \
            if cls in ("Corporate", "Bank") else ("Unconditionally cancellable" if cls == "Retail" else "No commitment")
        defaulted = "Y" if cls in ("Corporate", "Retail") and rng.random() < 0.025 else "N"
        rows.append({
            "exposure_id": f"EX{10000 + i}", "counterparty": f"CP{rng.integers(1000, 9999)}",
            "asset_class": cls, "rating": _rating(cls, rng),
            "drawn": drawn, "undrawn": round(drawn * rng.uniform(0.05, 0.6), 2) if obt != "No commitment" else 0.0,
            "off_balance_type": obt,
            "ltv": round(float(rng.uniform(0.35, 1.05)), 3) if cls == "Residential Mortgage" else None,
            "defaulted": defaulted,
            "specific_provision": round(drawn * rng.uniform(0.05, 0.45), 2) if defaulted == "Y" else 0.0,
        })
    return pd.DataFrame(rows)


def roll_forward(q2: pd.DataFrame, rng) -> pd.DataFrame:
    """Q3: some exposures mature, new ones are booked, balances move, a few ratings migrate."""
    q3 = q2.sample(frac=0.93, random_state=1).copy()  # 7% mature / repay
    q3["drawn"] = (q3["drawn"] * rng.normal(1.02, 0.06, len(q3))).round(2).clip(lower=0.1)
    # Corporate downgrades: a sector stress hits ~8% of rated corporates
    corp = q3[(q3.asset_class == "Corporate") & (q3.rating != "Unrated")].sample(frac=0.08, random_state=2).index
    for i in corp:
        order = list(RATING_BUCKET)[:-1]
        k = order.index(q3.at[i, "rating"])
        q3.at[i, "rating"] = order[min(k + int(rng.integers(1, 4)), len(order) - 1)]
    # Mortgage LTV drifts down as loans amortise
    m = q3.asset_class == "Residential Mortgage"
    q3.loc[m, "ltv"] = (q3.loc[m, "ltv"] - rng.uniform(0, 0.03, m.sum())).round(3)
    # A handful of new defaults in retail
    newdef = q3[(q3.asset_class == "Retail") & (q3.defaulted == "N")].sample(frac=0.02, random_state=3).index
    q3.loc[newdef, "defaulted"] = "Y"
    q3.loc[newdef, "specific_provision"] = (q3.loc[newdef, "drawn"] * 0.12).round(2)
    new = q2_book(int(len(q2) * 0.11), rng)
    new["exposure_id"] = [f"EX{20000 + i}" for i in range(len(new))]
    return pd.concat([q3, new]).reset_index(drop=True)


def main():
    rng = np.random.default_rng(2025)
    q2 = q2_book(1500, rng)
    q3 = roll_forward(q2, rng)
    out = Path("data"); out.mkdir(exist_ok=True)
    q2.to_csv(out / "exposures_2025Q2.csv", index=False)
    q3.to_csv(out / "exposures_2025Q3.csv", index=False)
    print(f"Q2: {len(q2)} exposures  Q3: {len(q3)} exposures")


if __name__ == "__main__":
    main()
