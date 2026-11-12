import pandas as pd
import pytest

from rwa.engine import attribute, calculate
from rwa.rules import risk_weight


@pytest.mark.parametrize("cls,rating,expected", [
    ("Sovereign", "AA", 0.0), ("Sovereign", "BBB-", 0.5),
    ("Bank", "A", 0.3), ("Bank", "Unrated", 0.75),
    ("Corporate", "BBB+", 0.75), ("Corporate", "B", 1.5), ("Corporate", "Unrated", 1.0),
])
def test_rated_risk_weights(cls, rating, expected):
    assert risk_weight(cls, rating, None, False, 0) == expected


@pytest.mark.parametrize("ltv,expected", [(0.45, 0.20), (0.50, 0.20), (0.55, 0.25), (0.80, 0.30), (0.95, 0.50), (1.05, 0.70)])
def test_mortgage_ltv_bands_are_upper_inclusive(ltv, expected):
    assert risk_weight("Residential Mortgage", "Unrated", ltv, False, 0) == expected


def test_defaulted_depends_on_provision_cover():
    assert risk_weight("Corporate", "A", None, True, 0.10) == 1.50
    assert risk_weight("Corporate", "A", None, True, 0.25) == 1.00


def _book(rows):
    cols = ["exposure_id", "asset_class", "rating", "drawn", "undrawn", "off_balance_type", "ltv", "defaulted", "specific_provision"]
    return pd.DataFrame(rows, columns=cols)


def test_ead_uses_ccf_and_walk_reconciles():
    q2 = calculate(_book([
        ["E1", "Corporate", "A", 100, 50, "Other commitment", None, "N", 0],     # EAD 120, RW 50%
        ["E2", "Retail", "Unrated", 40, 0, "No commitment", None, "N", 0],        # matures
    ]))
    assert q2.loc[0, "ead"] == 120 and q2.loc[0, "rwa"] == 60
    q3 = calculate(_book([
        ["E1", "Corporate", "BBB", 110, 50, "Other commitment", None, "N", 0],   # grows and downgraded
        ["E3", "Bank", "AA", 200, 0, "No commitment", None, "N", 0],             # new
    ]))
    w = attribute(q2, q3).loc["Total"]
    assert w.new_business == 40 and w.matured == -30
    assert w.volume_effect == pytest.approx(10 * 0.5)
    assert w.risk_weight_effect == pytest.approx(130 * 0.25)
    assert w.check == 0
