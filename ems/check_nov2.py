"""Replay the November 2, 2026 MyVenue pair with the event 15 assignments.

Mirrors EMS_Calculate: 10% food / 8% alcohol for nonprofits, the $200 commission
floor, 90% of tips cut off at the cent, and vendor commission by bucket.
A staff stand with tips and no headcount is reported and does not stop the rest.
"""

import csv
from collections import defaultdict
from decimal import Decimal, ROUND_FLOOR

PRICE = "/workspace/ems/samples/11-02-26_MBB_Price_Level_Categories.csv"
TIPS = "/workspace/ems/samples/11-02-26_MBB_Attendant_Tips.csv"


def money(text):
    text = (text or "").replace(",", "").replace("$", "").strip()
    return Decimal(text or "0")


def half_up(amount, rate):
    cents = (amount * Decimal(100) * rate)
    return (cents + Decimal("0.5")).to_integral_value(rounding=ROUND_FLOOR) / Decimal(100)


def floor_product(amount, rate):
    cents = amount * Decimal(100) * rate
    return cents.to_integral_value(rounding=ROUND_FLOOR) / Decimal(100)


def bucket(code):
    if code in ("313000", "FF-313076", "353008"):
        return "FOOD"
    if code in ("312115", "FF-312114"):
        return "NA"
    if code == "313082":
        return "BEER"
    if code in ("313004", "313096"):
        return "LIQUOR"
    return "OTHER"


def load():
    sales = defaultdict(lambda: defaultdict(lambda: [Decimal(0), Decimal(0)]))
    location = None
    for row in csv.reader(open(PRICE, newline="", encoding="utf-8-sig")):
        field = lambda index: row[index].strip() if index < len(row) else ""
        if field(1).lower().startswith("location:"):
            location = field(1)[9:].strip()
        elif location and field(2).endswith("Total:") and " - " in field(2)[:-6]:
            code = field(2)[:-6].split(" - ", 1)[0].strip()
            sales[location][bucket(code)][0] += money(field(15))
            sales[location][bucket(code)][1] += money(field(18))
    tips = {}
    for row in csv.reader(open(TIPS, newline="", encoding="utf-8-sig")):
        field = lambda index: row[index].strip() if index < len(row) else ""
        if field(1).endswith("Total:") and " - " in field(1):
            tips[field(1)[:-6].strip()] = money(field(12))
    return sales, tips


NPO = {"food": Decimal("0.10"), "alcohol": Decimal("0.08"), "withhold": Decimal("0.10"), "kind": "NPO"}
ASSIGNED = {
    "PIT 101 - A - POP & POUR (Q)": ("NPO", "ARVC", NPO),
    "PIT 101 - B -  NACHOMAMA (R)": ("NPO", "ARVC", NPO),
    "PIT 102 - THE PIT STOP (U)": ("NPO", "ARVC", NPO),
    "PIT 106 - LOUIE'S DAWG HOUSE (C)": ("NPO", "DCVA", NPO),
    "PIT 103 - SUB SUGAR SHIVERS (V)": ("SUB", "SUGAR", {"FOOD": Decimal("0.80"), "NA": Decimal("0.10"), "BEER": Decimal(0), "LIQUOR": Decimal(0), "withhold": Decimal("0.10"), "card": Decimal("0.03"), "kind": "SUB"}),
    "PIT 107 - SUB - LAGUNA BURGER (G)": ("SUB", "LAG", {"FOOD": Decimal("0.70"), "NA": Decimal("0.05"), "BEER": Decimal(0), "LIQUOR": Decimal(0), "withhold": Decimal("0.10"), "card": Decimal("0.03"), "kind": "SUB"}),
}


def main():
    sales, tips = load()
    groups = defaultdict(list)
    held = []
    for name, tip in tips.items():
        if name not in ASSIGNED:
            if tip:
                held.append((name, tip))
            continue
        role, payee, rate = ASSIGNED[name]
        nets = defaultdict(lambda: Decimal(0))
        gross = Decimal(0)
        for key, (g, n) in sales[name].items():
            nets[key] += n
            gross += g
        if role == "NPO":
            food = nets["FOOD"] + nets["NA"] + nets["OTHER"]
            alcohol = nets["BEER"] + nets["LIQUOR"]
            food_comm = half_up(food, rate["food"])
            alcohol_comm = half_up(alcohol, rate["alcohol"])
        else:
            food = nets["FOOD"] + nets["OTHER"]
            alcohol = nets["NA"] + nets["BEER"] + nets["LIQUOR"]
            food_comm = half_up(nets["FOOD"] + nets["OTHER"], rate["FOOD"])
            alcohol_comm = half_up(nets["NA"], rate["NA"]) + half_up(nets["BEER"], rate["BEER"]) + half_up(nets["LIQUOR"], rate["LIQUOR"])
        commission = food_comm + alcohol_comm
        topup = Decimal(0)
        if role == "NPO" and food + alcohol > 0 and commission < 200:
            topup = Decimal(200) - commission
        payable = floor_product(tip, Decimal(1) - rate["withhold"])
        groups[payee].append((name, food, alcohol, commission, payable, topup, gross, tip))
    print("STAFF OR OPEN TIPS LEFT UNPAID (no headcount, event still calculates)")
    for name, tip in held:
        print(f"  {tip:8}  {name}")
    print("held tips", sum(tip for _, tip in held))
    for payee, rows in groups.items():
        kind = ASSIGNED[rows[0][0]][1]
        rate = ASSIGNED[rows[0][0]][2]
        commission = sum(row[3] for row in rows)
        gratuity = sum(row[4] for row in rows)
        topup = sum(row[5] for row in rows)
        net = sum(row[1] + row[2] for row in rows)
        card = half_up(net, rate["card"]) if rate["kind"] == "SUB" else Decimal(0)
        due = commission + gratuity + topup - card
        print(f"\n{payee}  net {net}  comm {commission}  top-up {topup}  grat {gratuity}  card {card}  due {due}")
        for row in rows:
            print(f"  {row[0][:42]:42} food {row[1]} alc {row[2]} comm {row[3]} grat {row[4]} top {row[5]}")


if __name__ == "__main__":
    main()
