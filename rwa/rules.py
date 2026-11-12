"""Basel III standardised approach for credit risk, simplified.

References are to the Basel Framework, CRE20 (standardised approach: risk weights)
https://www.bis.org/basel_framework/chapter/CRE/20.htm
Simplifications are listed in the README; this is a learning model, not a regulatory engine.
"""

RATING_BUCKET = {
    **dict.fromkeys(["AAA", "AA+", "AA", "AA-"], "AAA to AA-"),
    **dict.fromkeys(["A+", "A", "A-"], "A+ to A-"),
    **dict.fromkeys(["BBB+", "BBB", "BBB-"], "BBB+ to BBB-"),
    **dict.fromkeys(["BB+", "BB", "BB-"], "BB+ to BB-"),
    **dict.fromkeys(["B+", "B", "B-"], "B+ to B-"),
    **dict.fromkeys(["CCC+", "CCC", "CC"], "Below B-"),
    "Unrated": "Unrated",
}
BUCKETS = ["AAA to AA-", "A+ to A-", "BBB+ to BBB-", "BB+ to BB-", "B+ to B-", "Below B-", "Unrated"]

# CRE20.7 sovereigns, CRE20.18 banks (ECRA; unrated taken as SCRA grade B = 75%), CRE20.42 corporates
RATED_RW = {
    "Sovereign": [0.00, 0.20, 0.50, 1.00, 1.00, 1.50, 1.00],
    "Bank":      [0.20, 0.30, 0.50, 1.00, 1.00, 1.50, 0.75],
    "Corporate": [0.20, 0.50, 0.75, 1.00, 1.50, 1.50, 1.00],
}
RETAIL_RW = 0.75  # CRE20.65 regulatory retail

# CRE20.82 residential real estate, repayment not materially dependent on the property's cash flows
LTV_BANDS = [(0, 0.20), (0.50, 0.25), (0.60, 0.30), (0.80, 0.40), (0.90, 0.50), (1.00, 0.70)]  # (lower bound, RW)

# CRE20.94 credit conversion factors for off-balance-sheet items
CCF = {
    "Unconditionally cancellable": 0.10,
    "Other commitment": 0.40,
    "Trade letter of credit": 0.20,
    "Guarantee / credit substitute": 1.00,
    "No commitment": 0.0,
}

CAPITAL_RATIO = 0.08  # minimum total capital (Pillar 1) as % of RWA


def risk_weight(asset_class: str, rating: str, ltv: float | None, defaulted: bool, provision_cover: float) -> float:
    if defaulted:  # CRE20.88: 150% if specific provisions < 20% of outstanding, else 100%
        return 1.50 if provision_cover < 0.20 else 1.00
    if asset_class == "Retail":
        return RETAIL_RW
    if asset_class == "Residential Mortgage":
        rw = LTV_BANDS[0][1]
        for lower, w in LTV_BANDS:
            if ltv is not None and ltv > lower:
                rw = w
        return rw
    return RATED_RW[asset_class][BUCKETS.index(RATING_BUCKET[rating])]
