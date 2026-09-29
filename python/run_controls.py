"""Reference implementation of the valuation controls in pandas.
Reads data/Raw_Positions.csv + data/Thresholds.csv, applies the same rules as the Excel/VBA/Access/Power BI
versions, and writes report files + a dashboard preview to outputs/.   Run from the repository root:
    python python/run_controls.py
"""
import json
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

REPORT_DATE = pd.Timestamp("2026-09-28")
raw = pd.read_csv("data/Raw_Positions.csv", dtype={"AssetClass": str})
thr = pd.read_csv("data/Thresholds.csv")

# 1. clean ---------------------------------------------------------------------------------------
d = raw.copy()
d["PositionID"] = d["PositionID"].astype(str).str.strip()
n_raw = len(d)
d = d.drop_duplicates("PositionID", keep="first").copy()
canon = {a.lower(): a for a in thr["AssetClass"]}
d["AssetClass"] = d["AssetClass"].astype(str).str.strip().map(lambda x: canon.get(x.lower(), x))
d["Desk"] = d["Desk"].str.strip()

# 2. test ----------------------------------------------------------------------------------------
d = d.merge(thr, on="AssetClass", how="left")
d["VariancePct"] = (d["BookPrice"] - d["IndependentPrice"]) / d["IndependentPrice"]
d["MVDiff"] = d["Quantity"] * (d["BookPrice"] - d["IndependentPrice"])

def status(r):
    if pd.isna(r.TolerancePct): return "UNMAPPED ASSET CLASS"
    if pd.isna(r.IndependentPrice): return "NO INDEPENDENT PRICE"
    return "BREACH" if abs(r.VariancePct) > r.TolerancePct else "OK"
d["Status"] = d.apply(status, axis=1)
d["DateFlag"] = (pd.to_datetime(d["ValuationDate"]) != REPORT_DATE).map({True: "STALE", False: "CURRENT"})

# 3. reports -------------------------------------------------------------------------------------
def summarise(by):
    t = d[d.Status.isin(["OK", "BREACH"])]
    g = t.groupby(by).agg(PositionsTested=("Status", "size"),
                          OK=("Status", lambda s: (s == "OK").sum()),
                          Breaches=("Status", lambda s: (s == "BREACH").sum()),
                          NetMVDiff=("MVDiff", "sum")).reset_index()
    g["NoIndepPrice"] = g[by].map(d[d.Status == "NO INDEPENDENT PRICE"].groupby(by).size()).fillna(0).astype(int)
    g["BreachRate"] = (g.Breaches / g.PositionsTested).round(4)
    g["NetMVDiff"] = g["NetMVDiff"].round(2)
    return g[[by, "PositionsTested", "OK", "Breaches", "NoIndepPrice", "BreachRate", "NetMVDiff"]]

exc = d[d.Status == "BREACH"].assign(AbsMVDiff=lambda x: x.MVDiff.abs()).sort_values("AbsMVDiff", ascending=False)
exc.drop(columns=["Status"]).round({"VariancePct": 6, "MVDiff": 2, "AbsMVDiff": 2}).to_csv("outputs/Exceptions.csv", index=False)
by_class, by_desk = summarise("AssetClass"), summarise("Desk")
by_class.to_csv("outputs/Summary_by_AssetClass.csv", index=False)
by_desk.to_csv("outputs/Summary_by_Desk.csv", index=False)
dq = pd.DataFrame([
    ("Raw rows loaded", n_raw), ("Duplicate rows removed", n_raw - len(d)), ("Unique positions", len(d)),
    ("Missing independent price", int((d.Status == "NO INDEPENDENT PRICE").sum())),
    ("Unmapped asset class", int((d.Status == "UNMAPPED ASSET CLASS").sum())),
    ("Stale valuation date", int((d.DateFlag == "STALE").sum()))], columns=["Check", "Count"])
dq.to_csv("outputs/Data_Quality.csv", index=False)
d.round({"VariancePct": 6, "MVDiff": 2}).to_csv("outputs/Checked_Positions.csv", index=False)

ok, br = int((d.Status == "OK").sum()), int((d.Status == "BREACH").sum())
print(dq.to_string(index=False)); print(f"OK {ok} | Breaches {br} | Breach rate {br/(ok+br):.1%}")
print(by_class.to_string(index=False))

# 4. verify against the independent expected results --------------------------------------------
exp = json.load(open("data/expected_results.json"))
assert (ok, br) == (exp["ok"], exp["breach"]), "status counts differ from expected_results.json"
assert int(dq.iloc[3, 1]) == exp["no_price"] and int(dq.iloc[4, 1]) == exp["unmapped"] and int(dq.iloc[5, 1]) == exp["stale"]
print("Verified against data/expected_results.json")

# 5. dashboard preview (matplotlib - NOT the Power BI file) ------------------------------------
NAVY, RED, AMBER, GREY = "#1F3864", "#C00000", "#ED9B00", "#8C8C8C"
fig = plt.figure(figsize=(14, 8.5), facecolor="white")
fig.suptitle("Valuation Control Overview - Python-generated preview (synthetic data)", fontsize=15, fontweight="bold", color=NAVY, x=0.02, ha="left")
kpis = [("Positions tested", f"{ok+br}"), ("Breaches", f"{br}"), ("Breach rate", f"{br/(ok+br):.1%}"),
        ("Net MV diff (breaches)", f"{d[d.Status=='BREACH'].MVDiff.sum():,.0f}")]
for i, (lab, val) in enumerate(kpis):
    ax = fig.add_axes([0.03 + i * 0.245, 0.80, 0.225, 0.11]); ax.axis("off")
    ax.add_patch(plt.Rectangle((0, 0), 1, 1, fc="#F2F5FA", ec="#C9D3E6", transform=ax.transAxes))
    ax.text(0.5, 0.62, val, ha="center", va="center", fontsize=22, fontweight="bold", color=RED if lab.startswith("Breach") else NAVY)
    ax.text(0.5, 0.2, lab, ha="center", va="center", fontsize=10, color=GREY)

ax1 = fig.add_axes([0.06, 0.45, 0.40, 0.28]); c = by_class.sort_values("Breaches")
ax1.barh(c.AssetClass, c.Breaches, color=RED); ax1.set_title("Breaches by asset class", loc="left", fontsize=11, color=NAVY)
for y, v in enumerate(c.Breaches): ax1.text(v + 0.15, y, str(v), va="center", fontsize=9)

ax2 = fig.add_axes([0.56, 0.45, 0.40, 0.28]); k = by_desk.sort_values("BreachRate", ascending=False)
ax2.bar(k.Desk, k.BreachRate * 100, color=NAVY); ax2.set_title("Breach rate by desk (%)", loc="left", fontsize=11, color=NAVY)
plt.setp(ax2.get_xticklabels(), rotation=35, ha="right", fontsize=8)

ax3 = fig.add_axes([0.06, 0.05, 0.36, 0.28]); s = d.Status.value_counts()
cols = {"OK": "#5B9BD5", "BREACH": RED, "NO INDEPENDENT PRICE": AMBER, "UNMAPPED ASSET CLASS": GREY}
w, _ = ax3.pie(s.values, colors=[cols[i] for i in s.index], startangle=90, wedgeprops=dict(width=0.4))
ax3.legend(w, [f"{i} ({v})" for i, v in s.items()], loc="center left", bbox_to_anchor=(1.0, 0.5), fontsize=8, frameon=False)
ax3.set_title("Status mix", loc="left", fontsize=11, color=NAVY)

ax4 = fig.add_axes([0.62, 0.05, 0.34, 0.28]); t = exc.head(8).iloc[::-1]
ax4.barh(t.PositionID + " " + t.AssetClass.str[:4], t.AbsMVDiff, color=RED)
ax4.set_title("Top exceptions by |MV diff|", loc="left", fontsize=11, color=NAVY); ax4.tick_params(labelsize=8)
for a in (ax1, ax2, ax4):
    a.spines[["top", "right"]].set_visible(False)
fig.savefig("outputs/dashboard_preview.png", dpi=140, bbox_inches="tight")
print("saved outputs/")
