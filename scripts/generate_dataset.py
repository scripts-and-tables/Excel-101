"""Generate the synthetic sales datasets used throughout the course.

UAE FMCG distributor context: sales of Food and HPC (home & personal care) brands
to Dubai supermarket chains, across their branches, by sales reps.

Outputs:
- ../docs/files/source/sales_data.csv  — order-line sales/returns export
- ../docs/files/source/reps.csv         — Sales Rep reference (rep covers customers)
- ../docs/files/source/brands.csv       — Brand reference (brand manager covers brands)

EVERYTHING IS SYNTHETIC. Customers are real Dubai supermarket *chain names* used only
as realistic labels; all codes, branches, reps, managers, brands, quantities and
values are invented from a fixed seed. The BRANDS ARE FICTIONAL — they belong to no
real company and are deliberately unrelated to any Transmed portfolio.

Deterministic — re-running produces the same files. A few deliberately "dirty" rows
are injected so the Module 1 cleanup exercises have something to clean.
"""

from __future__ import annotations

import csv
import random
from datetime import date, timedelta
from pathlib import Path

SEED = 20260101
ROW_COUNT = 2000
START_DATE = date(2024, 1, 1)
END_DATE = date(2025, 12, 31)
# Brands' base prices are per-unit; distributors sell by the CASE, so scale up to
# realistic wholesale line values (and therefore realistic AED totals & quotas).
CASE_FACTOR = 100

# Real Dubai supermarket chains (Food + HPC) — used only as labels. (name, customer code)
CUSTOMERS = [
    ("Carrefour", 1001),
    ("Lulu Hypermarket", 1002),
    ("Spinneys", 1003),
    ("Choithrams", 1004),
    ("Union Coop", 1005),
    ("Waitrose", 1006),
    ("Al Maya Supermarket", 1007),
    ("Nesto Hypermarket", 1008),
    ("West Zone Supermarket", 1009),
    ("Géant", 1010),
    ("Aswaaq", 1011),
    ("Viva Supermarket", 1012),
    ("Grandiose Supermarket", 1013),
    ("Day to Day", 1014),
    ("Safeer Market", 1015),
]

# Dubai areas a branch can sit in (also the rep territory dimension).
AREAS = [
    "Deira", "Bur Dubai", "Al Barsha", "Jumeirah", "Mirdif", "Dubai Marina",
    "Downtown", "Al Quoz", "Karama", "Jebel Ali", "Festival City",
    "International City", "Silicon Oasis", "JLT", "Discovery Gardens",
    "Motor City", "Al Qusais", "Satwa",
]

# Sales reps -> manager. Reps cover Dubai territories (areas).
SALES_MANAGERS = ["Khalid Rahman", "Sandra D'Souza", "Tariq Aziz"]
SALES_REPS = [
    "Rashid Al Marzooqi", "Anjali Menon", "Mohammed Saleh", "Grace Fernandes",
    "Vikram Nair", "Fatima Khan", "Joseph Mathew", "Ayesha Siddiqui",
    "Daniel Costa", "Priya Raj", "Omar Haddad", "Sunil Kumar",
    "Mariam Hassan", "Arjun Pillai", "Lina Aboud",
]

# Brand managers — people who each cover a set of brands ("category managers").
BRAND_MANAGERS = ["Imran Sheikh", "Rebecca Thomas", "Hassan Ali",
                  "Divya Krishnan", "Noor Abbas"]

# FICTIONAL brands by category (no real company / not Transmed). (brand, base unit price AED)
BRANDS = {
    "Food": [
        ("Crunchio", 4.5), ("Goldenfields", 9.0), ("FreshNest", 6.5),
        ("Cedarna", 12.0), ("Oasis Delights", 7.5), ("SunHarvest", 5.0),
        ("Bakehouse Co", 8.0), ("Zaytoona", 15.0), ("DeliMia", 10.5),
        ("Marhaba Gold", 6.0),
    ],
    "HPC": [
        ("Lumora", 18.0), ("Cleanova", 11.0), ("Mintleaf", 7.0),
        ("Auracare", 22.0), ("Silkene", 16.0), ("PureGlow", 25.0),
        ("FreshLine", 9.5), ("Caressa", 13.0), ("Sparklo", 8.5),
        ("Verdé", 19.0),
    ],
}


def random_date(rng: random.Random) -> date:
    delta = (END_DATE - START_DATE).days
    return START_DATE + timedelta(days=rng.randint(0, delta))


def build_branches(rng: random.Random) -> list[dict]:
    """For each customer, create 3-6 branches in distinct Dubai areas."""
    branches: list[dict] = []
    code = 50001
    for name, cust_code in CUSTOMERS:
        n = rng.randint(3, 6)
        areas = rng.sample(AREAS, n)
        for area in areas:
            branches.append({
                "Customer": name,
                "CustomerCode": cust_code,
                "Branch": f"{name} - {area}",
                "BranchCode": code,
                "Area": area,
            })
            code += 1
    return branches


def build_reps() -> tuple[list[dict], dict]:
    """Reps (SalesRep + Manager) + an area -> rep territory map.
    Quotas are assigned later, from each rep's actual sales (see main)."""
    reps = [{"SalesRep": name, "Manager": SALES_MANAGERS[i % len(SALES_MANAGERS)]}
            for i, name in enumerate(SALES_REPS)]
    # Assign every area to a rep (territory); reps may cover several areas.
    area_to_rep = {area: SALES_REPS[i % len(SALES_REPS)] for i, area in enumerate(AREAS)}
    return reps, area_to_rep


def build_brands() -> list[dict]:
    """Brand reference table: each brand -> its category and its brand manager."""
    brands: list[dict] = []
    for category, items in BRANDS.items():
        # two managers per category, plus Noor Abbas covering a premium slice
        if category == "Food":
            pool = ["Imran Sheikh", "Rebecca Thomas", "Noor Abbas"]
        else:
            pool = ["Hassan Ali", "Divya Krishnan", "Noor Abbas"]
        for j, (brand, _price) in enumerate(items):
            brands.append({
                "Brand": brand,
                "Category": category,
                "BrandManager": pool[j % len(pool)],
            })
    brands.sort(key=lambda b: b["Brand"])
    return brands


def main() -> None:
    rng = random.Random(SEED)
    branches = build_branches(rng)
    reps, area_to_rep = build_reps()
    brand_list = [(b, price, cat) for cat, items in BRANDS.items() for (b, price) in items]

    rows: list[dict] = []
    for i in range(1, ROW_COUNT + 1):
        br = rng.choice(branches)
        brand, base_price, category = rng.choice(brand_list)
        rep = area_to_rep[br["Area"]]
        # ~7% of lines are returns (negative qty & value)
        is_return = rng.random() < 0.07
        qty = rng.choices([1, 2, 3, 5, 8, 12, 20, 40, 60, 100],
                          weights=[20, 18, 15, 12, 10, 8, 7, 5, 3, 2])[0]
        # per-case price wobbles a little around the brand base (scaled to cases)
        unit_price = round(base_price * CASE_FACTOR * rng.uniform(0.9, 1.25), 2)
        value = round(unit_price * qty, 2)
        if is_return:
            qty = -qty
            value = -value
        rows.append({
            "OrderNumber": f"SO-{100000 + i}",
            "Date": random_date(rng).isoformat(),
            "InvoiceType": "Return" if is_return else "Sales",
            "CustomerCode": br["CustomerCode"],
            "Customer": br["Customer"],
            "BranchCode": br["BranchCode"],
            "Branch": br["Branch"],
            "Area": br["Area"],
            "SalesRep": rep,
            "Brand": brand,
            "Category": category,
            "SalesQuantity": qty,
            "SalesValue": value,
        })

    # ---- inject deliberate "dirty" rows for the Module 1 cleanup exercises ----
    # 1) leading/trailing spaces on a few Customer names
    for idx in (12, 47, 233, 871, 1402):
        if idx < len(rows):
            rows[idx]["Customer"] = "  " + rows[idx]["Customer"] + " "
    # 2) inconsistent casing on Area
    for idx in (88, 410, 905, 1567):
        if idx < len(rows):
            rows[idx]["Area"] = rows[idx]["Area"].upper()
    # 3) two exact duplicate order lines (fresh OrderNumber, otherwise identical)
    if len(rows) >= 50:
        dup1 = dict(rows[5]);  dup1["OrderNumber"] = "SO-199998"; rows.append(dup1)
        dup2 = dict(rows[42]); dup2["OrderNumber"] = "SO-199999"; rows.append(dup2)

    rows.sort(key=lambda r: r["Date"])

    # Data-driven annual quotas: base each rep's quota on their own net sales so
    # quota-attainment KPIs land in a believable ~80-125% range (rounded to 25k AED).
    rep_net: dict[str, float] = {}
    for r in rows:
        rep_net[r["SalesRep"]] = rep_net.get(r["SalesRep"], 0.0) + r["SalesValue"]
    for rep in reps:
        net = rep_net.get(rep["SalesRep"], 0.0)
        quota = net / rng.uniform(0.85, 1.25)
        rep["AnnualQuota"] = max(25_000, int(round(quota / 25_000)) * 25_000)
    reps.sort(key=lambda r: r["SalesRep"])

    out_path = Path(__file__).resolve().parent.parent / "docs" / "files" / "source" / "sales_data.csv"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = ["OrderNumber", "Date", "InvoiceType", "CustomerCode", "Customer",
                  "BranchCode", "Branch", "Area", "SalesRep", "Brand", "Category",
                  "SalesQuantity", "SalesValue"]
    with out_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    print(f"Wrote {len(rows)} rows to {out_path}")

    reps_path = out_path.parent / "reps.csv"
    rep_fields = ["SalesRep", "Manager", "AnnualQuota"]
    with reps_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=rep_fields)
        writer.writeheader()
        writer.writerows(reps)
    print(f"Wrote {len(reps)} reps to {reps_path}")

    brands_ref = build_brands()
    brands_path = out_path.parent / "brands.csv"
    with brands_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["Brand", "Category", "BrandManager"])
        writer.writeheader()
        writer.writerows(brands_ref)
    print(f"Wrote {len(brands_ref)} brands to {brands_path}")


if __name__ == "__main__":
    main()
