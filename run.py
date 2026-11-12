"""Generate the exposure book, price it in Python, write the Excel model and charts."""
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

from rwa import generate_data
from rwa.engine import attribute, calculate, summary
from rwa.excel_model import build

generate_data.main()
q2 = pd.read_csv("data/exposures_2025Q2.csv")
q3 = pd.read_csv("data/exposures_2025Q3.csv")
p2, p3 = calculate(q2), calculate(q3)
s2, s3 = summary(p2), summary(p3)
walk = attribute(p2, p3)

out = Path("output"); out.mkdir(exist_ok=True)
s3.round(2).to_csv(out / "rwa_summary_2025Q3.csv")
walk.round(2).to_csv(out / "rwa_walk_Q2_Q3.csv")
print(walk.round(1).to_string())
print("Excel model:", build(q2, q3, out / "basel_rwa_calculator.xlsx", s2, s3))

# Waterfall for the README
t = walk.loc["Total"]
steps = [("Q2 RWA", t.rwa_prev), ("New business", t.new_business), ("Maturities", t.matured),
         ("Volume", t.volume_effect), ("Risk-weight\nchanges", t.risk_weight_effect)]
fig, ax = plt.subplots(figsize=(9, 4.2))
run = 0
for i, (lab, v) in enumerate(steps):
    base = 0 if i == 0 else run
    color = "#1F3864" if i == 0 else ("#C0392B" if v > 0 else "#27AE60")
    ax.bar(i, v, bottom=base, color=color)
    ax.text(i, base + v + (900 if v >= 0 else -900), f"{v/1000:+.1f}k" if i else f"{v/1000:.1f}k",
            ha="center", va="bottom" if v >= 0 else "top", fontsize=9)
    run = base + v
ax.bar(len(steps), t.rwa_now, color="#1F3864")
ax.text(len(steps), t.rwa_now + 900, f"{t.rwa_now/1000:.1f}k", ha="center", fontsize=9)
ax.set_xticks(range(len(steps) + 1), [s[0] for s in steps] + ["Q3 RWA"], fontsize=9)
ax.set_ylim(t.rwa_prev * 0.85, t.rwa_now * 1.12)
ax.set_title("RWA walk, Q2 to Q3 2025 (INR crore)", loc="left", fontweight="bold")
for s in ("top", "right"):
    ax.spines[s].set_visible(False)
fig.tight_layout(); fig.savefig("docs/images/rwa_walk.png", dpi=150); plt.close(fig)

fig, ax = plt.subplots(figsize=(9, 3.8))
d = s3.drop("Total")
ax.barh(d.index, d.ead / 1000, color="#BFC9D9", label="EAD")
ax.barh(d.index, d.rwa / 1000, color="#1F3864", label="RWA")
for i, (e, r) in enumerate(zip(d.ead, d.rwa)):
    ax.text(e / 1000 + 1, i, f"density {r/e:.0%}", va="center", fontsize=8)
ax.set_title("Q3 2025 EAD vs RWA by asset class (INR '000 crore)", loc="left", fontweight="bold")
ax.legend(frameon=False); [ax.spines[s].set_visible(False) for s in ("top", "right")]
fig.tight_layout(); fig.savefig("docs/images/rwa_density.png", dpi=150); plt.close(fig)
