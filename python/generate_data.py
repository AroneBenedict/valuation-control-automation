"""Generate SYNTHETIC position data with deliberate data-quality problems, and compute
the expected control results in pandas (used to cross-check Excel, VBA, Access and Power BI)."""
import random, json, datetime as dt
import pandas as pd

random.seed(42)
REPORT_DATE = dt.date(2026, 9, 28)
STALE_DATE = dt.date(2026, 9, 25)

CLASSES = {
    "Equity":          dict(prefix="EQ", n=80, desks=[("Zurich Equities", "CHF"), ("Singapore Equities", "SGD")], lo=20,   hi=300,  sigma=0.005,  tol=0.010,  qty=(100, 50000),      dp=2),
    "Corporate Bond":  dict(prefix="CB", n=70, desks=[("London Credit", "GBP"), ("New York Credit", "USD")],      lo=92,   hi=108,  sigma=0.0025, tol=0.005,  qty=(1000, 20000),     dp=3),
    "Government Bond": dict(prefix="GB", n=60, desks=[("London Rates", "GBP"), ("New York Rates", "USD")],        lo=96,   hi=104,  sigma=0.001,  tol=0.0025, qty=(1000, 20000),     dp=3),
    "FX Forward":      dict(prefix="FX", n=40, desks=[("Zurich FX", "CHF"), ("Singapore FX", "SGD")],            lo=0.6,  hi=1.6,  sigma=0.0015, tol=0.003,  qty=(100000, 5000000), dp=5),
}

rows = []
for ac, c in CLASSES.items():
    for i in range(1, c["n"] + 1):
        desk, ccy = random.choice(c["desks"])
        book = round(random.uniform(c["lo"], c["hi"]), c["dp"])
        eps = random.gauss(0, c["sigma"])
        ind = round(book / (1 + eps), c["dp"])
        rows.append(dict(AssetClass=ac, Instrument=f'{c["prefix"]}-{i:04d}', Desk=desk, Currency=ccy,
                         Quantity=random.randint(*c["qty"]), BookPrice=book, IndependentPrice=ind, ValuationDate=REPORT_DATE))
random.shuffle(rows)
for k, r in enumerate(rows, 1):
    r["PositionID"] = f"P{k:04d}"

idx = list(range(len(rows))); random.shuffle(idx)
take = lambda n: [idx.pop() for _ in range(n)]

for i in take(12):   # deliberate large pricing differences
    r = rows[i]; c = CLASSES[r["AssetClass"]]
    eps = random.choice([-1, 1]) * c["tol"] * random.uniform(1.3, 3.0)
    r["IndependentPrice"] = round(r["BookPrice"] / (1 + eps), c["dp"])
for i in take(6):  rows[i]["IndependentPrice"] = None     # missing independent price
for i in take(5):  rows[i]["ValuationDate"] = STALE_DATE  # stale valuation date
for i in take(2):  rows[i]["AssetClass"] = "Gov Bond"     # unmapped label
for i in take(16):                                        # messy text fixable by trim / case
    a = rows[i]["AssetClass"]
    rows[i]["AssetClass"] = random.choice([" " + a, a.upper(), a.lower() + " ", a + "  "])
for i in take(4):    # exact duplicate rows inserted later in the file
    pos = min(len(rows), i + random.randint(3, 20))
    rows.insert(pos, dict(rows[i]))

cols = ["PositionID", "ValuationDate", "Desk", "AssetClass", "Instrument", "Currency", "Quantity", "BookPrice", "IndependentPrice"]
raw = pd.DataFrame(rows)[cols]
raw.to_csv("data/Raw_Positions.csv", index=False, date_format="%Y-%m-%d")
thr = pd.DataFrame([(k, v["tol"]) for k, v in CLASSES.items()], columns=["AssetClass", "TolerancePct"])
thr.to_csv("data/Thresholds.csv", index=False)

def run_controls(raw, thr, report_date):
    d = raw.copy()
    d["PositionID"] = d["PositionID"].astype(str).str.strip()
    d = d.drop_duplicates("PositionID", keep="first").copy()
    canon = {a.lower(): a for a in thr["AssetClass"]}
    d["AssetClass"] = d["AssetClass"].astype(str).str.strip().map(lambda x: canon.get(x.lower(), x))
    d = d.merge(thr, on="AssetClass", how="left")
    d["VariancePct"] = (d["BookPrice"] - d["IndependentPrice"]) / d["IndependentPrice"]
    d["MVDiff"] = d["Quantity"] * (d["BookPrice"] - d["IndependentPrice"])
    def status(r):
        if pd.isna(r["TolerancePct"]): return "UNMAPPED ASSET CLASS"
        if pd.isna(r["IndependentPrice"]): return "NO INDEPENDENT PRICE"
        return "BREACH" if abs(r["VariancePct"]) > r["TolerancePct"] else "OK"
    d["Status"] = d.apply(status, axis=1)
    d["Stale"] = pd.to_datetime(d["ValuationDate"]).dt.date != report_date
    return d

if __name__ == "__main__":
    clean = run_controls(raw, thr, REPORT_DATE)
    clean[cols].to_csv("data/Clean_Positions.csv", index=False, date_format="%Y-%m-%d")
    exp = dict(raw_rows=len(raw), unique=len(clean), duplicates=len(raw) - len(clean),
               no_price=int((clean.Status == "NO INDEPENDENT PRICE").sum()),
               unmapped=int((clean.Status == "UNMAPPED ASSET CLASS").sum()),
               stale=int(clean.Stale.sum()), ok=int((clean.Status == "OK").sum()),
               breach=int((clean.Status == "BREACH").sum()))
    exp["checked"] = exp["ok"] + exp["breach"]
    exp["breach_rate"] = round(exp["breach"] / exp["checked"], 4)
    chk = clean[clean.Status.isin(["OK", "BREACH"])]
    exp["by_class"] = [dict(AssetClass=k, positions=len(g), breaches=int((g.Status == "BREACH").sum()), net_mv=round(float(g.MVDiff.sum()), 2))
                       for k, g in chk.groupby("AssetClass")]
    top = clean[clean.Status == "BREACH"].assign(a=lambda x: x.MVDiff.abs()).sort_values("a", ascending=False).head(5)
    exp["top5"] = [dict(PositionID=r.PositionID, AssetClass=r.AssetClass, VariancePct=round(float(r.VariancePct), 6), MVDiff=round(float(r.MVDiff), 2)) for r in top.itertuples()]
    json.dump(exp, open("data/expected_results.json", "w"), indent=2)
    print(json.dumps(exp, indent=2))
