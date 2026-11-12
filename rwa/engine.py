"""Compute EAD, risk weight and RWA per exposure, and explain the quarter-on-quarter change."""
from __future__ import annotations

import pandas as pd

from .rules import CAPITAL_RATIO, CCF, risk_weight


def calculate(exp: pd.DataFrame) -> pd.DataFrame:
    df = exp.copy()
    df["ccf"] = df["off_balance_type"].map(CCF).fillna(0.0)
    df["ead"] = df["drawn"] + df["ccf"] * df["undrawn"]
    cover = (df["specific_provision"] / df["drawn"].where(df["drawn"] > 0)).fillna(0)
    df["risk_weight"] = [
        risk_weight(r.asset_class, r.rating, r.ltv, r.defaulted == "Y", c)
        for r, c in zip(df.itertuples(index=False), cover)
    ]
    df["rwa"] = df["ead"] * df["risk_weight"]
    df["capital_req"] = df["rwa"] * CAPITAL_RATIO
    return df


def summary(df: pd.DataFrame) -> pd.DataFrame:
    s = df.groupby("asset_class")[["ead", "rwa", "capital_req"]].sum()
    s["rw_density"] = s["rwa"] / s["ead"]
    s.loc["Total"] = s.sum(numeric_only=True)
    s.loc["Total", "rw_density"] = s.loc["Total", "rwa"] / s.loc["Total", "ead"]
    return s


def attribute(prev: pd.DataFrame, cur: pd.DataFrame) -> pd.DataFrame:
    """Split the RWA change into new business, maturities, EAD (volume) and risk-weight effects.

    For exposures in both quarters:
      volume effect      = (EAD_now - EAD_prev) x RW_prev
      risk-weight effect = EAD_now x (RW_now - RW_prev)
    which adds up exactly to RWA_now - RWA_prev.
    """
    m = cur.merge(prev[["exposure_id", "ead", "risk_weight", "rwa"]], on="exposure_id", how="outer",
                  suffixes=("", "_prev"), indicator=True)
    m["asset_class"] = m["asset_class"].fillna(m["exposure_id"].map(prev.set_index("exposure_id")["asset_class"]))
    both = m["_merge"] == "both"
    out = pd.DataFrame({
        "asset_class": m["asset_class"],
        "new_business": m["rwa"].where(m["_merge"] == "left_only", 0),
        "matured": -m["rwa_prev"].where(m["_merge"] == "right_only", 0),
        "volume_effect": ((m["ead"] - m["ead_prev"]) * m["risk_weight_prev"]).where(both, 0),
        "risk_weight_effect": (m["ead"] * (m["risk_weight"] - m["risk_weight_prev"])).where(both, 0),
    })
    g = out.groupby("asset_class").sum()
    g.insert(0, "rwa_prev", prev.groupby("asset_class")["rwa"].sum())
    g["rwa_now"] = cur.groupby("asset_class")["rwa"].sum()
    g = g.fillna(0)
    g.loc["Total"] = g.sum()
    g["check"] = (g["rwa_prev"] + g[["new_business", "matured", "volume_effect", "risk_weight_effect"]].sum(axis=1) - g["rwa_now"]).round(2)
    return g
