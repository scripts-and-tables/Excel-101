"""Generate a synthetic sales dataset used throughout the course.

Output: ../files/source/sales_data.csv

Deterministic — running it again produces the same file.
Includes a few deliberate "dirty" rows so Module 1 cleanup exercises have something to clean.
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

REGIONS = ["North", "South", "East", "West", "Central"]

SALES_REPS = {
    "North":   ["Anna Becker", "Liam O'Connor", "Priya Sharma"],
    "South":   ["Marco Rossi", "Yara Haddad", "Tom Nguyen"],
    "East":    ["Sofia Lindqvist", "Daniel Park", "Amaya Suzuki"],
    "West":    ["Carlos Mendes", "Hannah Cohen", "Jin Wei"],
    "Central": ["Olivia Brown", "Ravi Patel", "Eva Schmidt"],
}

CATEGORIES = {
    "Hardware":   ["Laptop Pro 14", "Laptop Pro 16", "Desktop Mini", "Monitor 27\"", "Monitor 32\"", "Docking Station"],
    "Accessories":["Wireless Mouse", "Mechanical Keyboard", "USB-C Hub", "Webcam HD", "Headset Pro"],
    "Software":   ["Office Suite", "Antivirus 1y", "VPN 1y", "Design Suite", "Project Tracker"],
    "Services":   ["Setup Service", "Onsite Training", "Premium Support", "Data Migration"],
}

PRICE_BY_PRODUCT = {
    "Laptop Pro 14": 1299, "Laptop Pro 16": 1899, "Desktop Mini": 899,
    "Monitor 27\"": 329,  "Monitor 32\"": 549,  "Docking Station": 189,
    "Wireless Mouse": 39, "Mechanical Keyboard": 119, "USB-C Hub": 59,
    "Webcam HD": 89, "Headset Pro": 149,
    "Office Suite": 199, "Antivirus 1y": 49, "VPN 1y": 69,
    "Design Suite": 599, "Project Tracker": 129,
    "Setup Service": 249, "Onsite Training": 1500,
    "Premium Support": 899, "Data Migration": 1200,
}

COST_RATIO = {
    "Hardware":    0.72,
    "Accessories": 0.55,
    "Software":    0.30,
    "Services":    0.40,
}

CUSTOMER_PREFIXES = ["Acme", "Globex", "Initech", "Umbrella", "Wayne", "Stark",
                     "Wonka", "Hooli", "Pied Piper", "Cyberdyne", "Soylent",
                     "Tyrell", "Vandelay", "Aperture", "Black Mesa", "Massive Dynamic",
                     "Oscorp", "Weyland", "Rekall", "Gringotts"]
CUSTOMER_SUFFIXES = ["Corp", "Industries", "LLC", "Group", "GmbH", "Holdings", "Partners", "Ltd"]

STATUSES_WEIGHTED = [("Closed Won", 0.78), ("Closed Won", 0.0),  # weight via duplicates
                     ("Refunded", 0.05), ("Pending", 0.10), ("Cancelled", 0.07)]


def weighted_choice(rng: random.Random, weighted: list[tuple[str, float]]) -> str:
    pool: list[str] = []
    for value, weight in weighted:
        pool.extend([value] * max(1, int(weight * 100)))
    return rng.choice(pool)


def random_date(rng: random.Random) -> date:
    delta = (END_DATE - START_DATE).days
    return START_DATE + timedelta(days=rng.randint(0, delta))


def make_customers(rng: random.Random, n: int = 80) -> list[str]:
    customers: set[str] = set()
    while len(customers) < n:
        customers.add(f"{rng.choice(CUSTOMER_PREFIXES)} {rng.choice(CUSTOMER_SUFFIXES)}")
    return sorted(customers)


def main() -> None:
    rng = random.Random(SEED)
    customers = make_customers(rng)

    rows: list[dict] = []
    for i in range(1, ROW_COUNT + 1):
        region = rng.choice(REGIONS)
        rep = rng.choice(SALES_REPS[region])
        category = rng.choice(list(CATEGORIES.keys()))
        product = rng.choice(CATEGORIES[category])
        unit_price = PRICE_BY_PRODUCT[product]
        quantity = rng.choices([1, 2, 3, 4, 5, 8, 10, 15, 25],
                               weights=[28, 22, 16, 10, 8, 6, 5, 3, 2])[0]
        # Discount: most orders 0, some 5/10/15/20%
        discount = rng.choices([0.00, 0.05, 0.10, 0.15, 0.20],
                               weights=[55, 18, 14, 8, 5])[0]
        cost = round(unit_price * COST_RATIO[category], 2)
        status = weighted_choice(rng, STATUSES_WEIGHTED)
        rows.append({
            "OrderID": f"SO-{10000 + i}",
            "OrderDate": random_date(rng).isoformat(),
            "Region": region,
            "SalesRep": rep,
            "Customer": rng.choice(customers),
            "Product": product,
            "Category": category,
            "UnitPrice": unit_price,
            "Quantity": quantity,
            "Discount": discount,
            "Cost": cost,
            "Status": status,
        })

    # Inject deliberate dirty data for the Module 1 cleanup exercise
    # 1) Trailing/leading spaces on a few customer names
    for idx in (12, 47, 233, 871, 1402):
        if idx < len(rows):
            rows[idx]["Customer"] = "  " + rows[idx]["Customer"] + " "
    # 2) Inconsistent casing on Region
    for idx in (88, 410, 905, 1567):
        if idx < len(rows):
            rows[idx]["Region"] = rows[idx]["Region"].upper()
    # 3) Two exact duplicate orders (different IDs but otherwise identical content)
    if len(rows) >= 2:
        dup_src = dict(rows[5])
        dup_src["OrderID"] = "SO-19998"
        rows.append(dup_src)
        dup_src2 = dict(rows[42])
        dup_src2["OrderID"] = "SO-19999"
        rows.append(dup_src2)

    # Sort by date for nicer presentation
    rows.sort(key=lambda r: r["OrderDate"])

    out_path = Path(__file__).resolve().parent.parent / "files" / "source" / "sales_data.csv"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = ["OrderID", "OrderDate", "Region", "SalesRep", "Customer",
                  "Product", "Category", "UnitPrice", "Quantity", "Discount",
                  "Cost", "Status"]
    with out_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    print(f"Wrote {len(rows)} rows to {out_path}")


if __name__ == "__main__":
    main()
