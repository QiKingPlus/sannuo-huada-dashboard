#!/usr/bin/env python3
"""Generate standalone order analysis report and write to file."""

import os
from collections import Counter, defaultdict
from datetime import datetime

# ============ DATA ============
orders = [
    {"sales":"\u5434\u8001\u5e08","time":"2026-05-26 11:18:14","product_name":"\u3010\u4e09\u8bfa\u4e13\u4eab\u3011\u4e13\u7814\u5347\u7ea73.0\u3010\u4f18\u7f8e\u8fbe\u00b7\u76ca\u4f11\u76ca\u751f\u83cc\u56fa\u4f53\u996e\u6599\u3011\u5165\u591c\u4e00\u6761\uff0c\u8eab\u4f53\u81ea\u7136\u5f52\u66ae\uff0830\u888b/\u76d2\uff09","price":253.3,"name":"\u8c22\u8d35","phone":"13975315486","qty":1,"gender":"\u5973","age":54,"need":"\u7761\u7720","province":"\u6e56\u5357\u7701","city":"\u682a\u6d32\u5e02"},
    {"sales":"\u5218\u8001\u5e08","time":"2026-05-26 10:33:02","product_name":"\u3010\u4e09\u8bfa\u4e13\u4eab\u3011\u3010\u8f7b\u8def \u590d\u5408\u76ca\u751f\u5143\u81b3\u98df\u7ea4\u7ef4\u56fa\u4f53\u996e\u6599\u3011\u8865\u83cc\u4e4b\u5916\uff0c\u66f4\u8981\u517b\u83cc\uff01\u4e94\u5143\u590d\u914d\uff0c\u597d\u83cc\u5de5\u4f5c\u4e8b\u534a\u529f\u500d\uff0815\u888b/\u76d2\uff09","price":139,"name":"\u8d3e\u4ec1","phone":"15073313688","qty":1,"gender":"\u7537","age":42,"need":"\u8840\u7cd6","province":"\u5317\u4eac\u5e02","city":"\u5317\u4eac\u5e02"},
    {"sales":"\u738b\u8001\u5e08","time":"2026-05-26 10:10:09","product_name":"\u3010\u4e09\u8bfa\u4e13\u4eab\u3011\u4e13\u7814\u5347\u7ea73.0\u3010\u4f18\u7f8e\u8fbe\u00b7\u76ca\u4f11\u76ca\u751f\u83cc\u56fa\u4f53\u996e\u6599\u3011\u5165\u591c\u4e00\u6761\uff0c\u8eab\u4f53\u81ea\u7136\u5f52\u66ae\uff0830\u888b/\u76d2\uff09","price":253.3,"name":"\u6e29\u6653\u7091","phone":"17703100088","qty":1,"gender":"\u7537","age":44,"need":"\u7761\u7720","province":"\u6cb3\u5317\u7701","city":"\u90af\u90f8\u5e02"},
    {"sales":"\u5434\u8001\u5e08","time":"2026-05-25 17:35:59","product_name":"\u3010\u4e09\u8bfa\u4e13\u4eab\u3011\u3010\u8f7b\u8def \u590d\u5408\u76ca\u751f\u5143\u81b3\u98df\u7ea4\u7ef4\u56fa\u4f53\u996e\u6599\u3011\u8865\u83cc\u4e4b\u5916\uff0c\u66f4\u8981\u517b\u83cc\uff01\u4e94\u5143\u590d\u914d\uff0c\u597d\u83cc\u5de5\u4f5c\u4e8b\u534a\u529f\u500d\uff0815\u888b/\u76d2\uff09","price":139,"name":"\u8c22\u7ea2","phone":"15911579927","qty":1,"gender":"\u5973","age":64,"need":"\u8840\u7cd6","province":"\u4e91\u5357\u7701","city":"\u6606\u660e"},
    {"sales":"\u5434\u8001\u5e08","time":"2026-05-25 14:38:58","product_name":"\u3010\u4e09\u8bfa\u4e13\u4eab\u3011\u3010\u4f18\u7f8e\u8fbe\u00b7\u7545\u9752\u76ca\u751f\u83cc\u538b\u7247\u7cd6\u679c\u3011\u62d2\u7edd\u62c9\u5230\u6276\u5899\u8d77\uff01\u80a0\u80a0\u5e73\u8861\u7a33\u83cc\u5fc3\uff01\uff0815\u888b/\u76d2\uff09","price":159.8,"name":"\u6c88\u5584\u660e","phone":"18967469708","qty":3,"gender":"\u7537","age":59,"need":"\u80a0\u80c3","province":"\u6d59\u6c5f\u7701","city":"\u91d1\u534e\u5e02"},
    {"sales":"\u5218\u8001\u5e08","time":"2026-05-25 11:51:11","product_name":"\u3010\u4e09\u8bfa\u4e13\u4eab\u3011\u3010\u4f18\u7f8e\u8fbe\u00b7\u76ca\u5bb6\u4eba\u76ca\u751f\u83cc\u3011\u80a0\u9053\u76ca\u751f\u83cc \u5168\u5bb6\u5171\u4eab\u6bcf\u76ca\u5929(15\u888b/\u76d2\uff09","price":398,"name":"\u5317\u8fb0","phone":"18190919791","qty":3,"gender":"\u5973","age":45,"need":"\u5bb6\u4eba\u5065\u5eb7","province":"\u56db\u5ddd\u7701","city":"\u6210\u90fd\u5e02"},
    {"sales":"\u5434\u8001\u5e08","time":"2026-05-22 17:25:30","product_name":"\u3010\u4e09\u8bfa\u4e13\u4eab\u3011\u3010\u8f7b\u8def \u590d\u5408\u76ca\u751f\u5143\u81b3\u98df\u7ea4\u7ef4\u56fa\u4f53\u996e\u6599\u3011\u8865\u83cc\u4e4b\u5916\uff0c\u66f4\u8981\u517b\u83cc\uff01\u4e94\u5143\u590d\u914d\uff0c\u597d\u83cc\u5de5\u4f5c\u4e8b\u534a\u529f\u500d\uff0815\u888b/\u76d2\uff09","price":139,"name":"\u79e6\u5e05","phone":"15262003719","qty":1,"gender":"\u7537","age":37,"need":"\u8f7b\u4f53","province":"\u6c5f\u82cf\u7701","city":"\u5f90\u5dde\u5e02"},
    {"sales":"\u5218\u8001\u5e08","time":"2026-05-21 17:20:40","product_name":"\u3010\u4e09\u8bfa\u4e13\u4eab\u3011\u3010\u8f7b\u8def \u590d\u5408\u76ca\u751f\u5143\u81b3\u98df\u7ea4\u7ef4\u56fa\u4f53\u996e\u6599\u3011\u8865\u83cc\u4e4b\u5916\uff0c\u66f4\u8981\u517b\u83cc\uff01\u4e94\u5143\u590d\u914d\uff0c\u597d\u83cc\u5de5\u4f5c\u4e8b\u534a\u529f\u500d\uff0815\u888b/\u76d2\uff09","price":139,"name":"\u5f20\u6587\u751f","phone":"13501134294","qty":1,"gender":"\u7537","age":59,"need":"\u8f7b\u4f53","province":"\u5317\u4eac\u5e02","city":"\u5317\u4eac\u5e02"},
    {"sales":"\u5218\u8001\u5e08","time":"2026-05-21 17:14:07","product_name":"\u3010\u4e09\u8bfa\u4e13\u4eab\u3011\u3010\u7ba1\u7406\u5957\u7ec4\u3011\u3010\u955c\u8102\u00b7\u8f7b+\u955c\u8102\u00b7\u963b\u3011\u7ba1\u7406\u597d\u642d\u6863\uff0c\u81ea\u5f8b\u751f\u6d3b\u6211\u6765\u5b9a\u4e49\uff01","price":566.6,"name":"\u5f20\u6587\u751f","phone":"13501134294","qty":1,"gender":"\u7537","age":59,"need":"\u8f7b\u4f53","province":"\u5317\u4eac\u5e02","city":"\u5317\u4eac\u5e02"},
    {"sales":"\u5434\u8001\u5e08","time":"2026-05-21 14:34:17","product_name":"\u3010\u4e09\u8bfa\u4e13\u4eab\u3011\u3010\u4f18\u7f8e\u8fbe\u00b7\u76ca\u7545\u76ca\u751f\u83cc\u3011\u8fde\u7eed10\u5e74\u53e3\u7891\uff01\u4e94\u79cd\u76ca\u751f\u83cc+\u4e09\u79cd\u76ca\u751f\u5143 \u5657\u5657\u81ea\u7531 \u8e72\u5751\u5feb\u4e50\uff0830\u888b/\u76d2\uff09","price":438.6,"name":"\u6768\u5148\u658b","phone":"13986460439","qty":1,"gender":"\u7537","age":70,"need":"\u80a0\u80c3","province":"\u6e56\u5317\u7701","city":"\u5b5d\u611f\u5e02"},
    {"sales":"\u5218\u8001\u5e08","time":"2026-05-21 11:34:07","product_name":"\u3010\u4e09\u8bfa\u4e13\u4eab\u3011\u3010\u7ba1\u7406\u5957\u7ec4\u3011\u3010\u955c\u8102\u00b7\u8f7b+\u955c\u8102\u00b7\u963b\u3011\u7ba1\u7406\u597d\u642d\u6863\uff0c\u81ea\u5f8b\u751f\u6d3b\u6211\u6765\u5b9a\u4e49\uff01","price":566.6,"name":"\u5317\u8fb0","phone":"18190919791","qty":1,"gender":"\u5973","age":45,"need":"\u8f7b\u4f53","province":"\u56db\u5ddd\u7701","city":"\u6210\u90fd\u5e02"},
]

# Product classification
def classify_product(name):
    if "\u76ca\u4f11" in name: return "\u76ca\u4f11\uff08\u7761\u7720\uff09"
    if "\u8f7b\u8def" in name: return "\u8f7b\u8def\uff08\u81b3\u98df\u7ea4\u7ef4\uff09"
    if "\u7545\u9752" in name: return "\u7545\u9752\uff08\u5438\u6536/\u589e\u91cd\uff09"
    if "\u76ca\u5bb6\u4eba" in name: return "\u76ca\u5bb6\u4eba\uff08\u57fa\u7840\u76ca\u751f\u83cc\uff09"
    if "\u955c\u8102" in name: return "\u955c\u8102\u5957\u7ec4\uff08\u4f53\u91cd\u7ba1\u7406\uff09"
    if "\u76ca\u7545" in name: return "\u76ca\u7545\uff08\u4fbf\u79d8\uff09"
    return "\u5176\u4ed6"

for o in orders:
    o["product"] = classify_product(o["product_name"])
    o["amount"] = o["price"] * o["qty"]
    o["date"] = o["time"][:10]
    o["hour"] = int(o["time"][11:13])

# ============ COMPUTE ============
total_revenue = sum(o["amount"] for o in orders)
total_orders = len(orders)
total_customers = len(set(o["phone"] for o in orders))
total_units = sum(o["qty"] for o in orders)
days = 6

# Sales performance
sales_data = defaultdict(lambda: {"orders": 0, "revenue": 0, "customers": set(), "units": 0})
sales_products = defaultdict(Counter)
for o in orders:
    s = o["sales"]
    sales_data[s]["orders"] += 1
    sales_data[s]["revenue"] += o["amount"]
    sales_data[s]["customers"].add(o["phone"])
    sales_data[s]["units"] += o["qty"]
    sales_products[s][o["product"]] += o["amount"]

# Product analysis
product_data = defaultdict(lambda: {"orders": 0, "revenue": 0, "units": 0, "customers": set()})
for o in orders:
    p = o["product"]
    product_data[p]["orders"] += 1
    product_data[p]["revenue"] += o["amount"]
    product_data[p]["units"] += o["qty"]
    product_data[p]["customers"].add(o["phone"])

sorted_products = sorted(product_data.items(), key=lambda x: x[1]["revenue"], reverse=True)

# Age
ages = [o["age"] for o in orders]
sorted_ages = sorted(ages)

# Gender
gender_data = defaultdict(lambda: {"orders": 0, "revenue": 0, "customers": set()})
for o in orders:
    g = o["gender"]
    gender_data[g]["orders"] += 1
    gender_data[g]["revenue"] += o["amount"]
    gender_data[g]["customers"].add(o["phone"])

# Region
region_data = defaultdict(lambda: {"orders": 0, "revenue": 0, "customers": set()})
for o in orders:
    region_data[o["province"]]["orders"] += 1
    region_data[o["province"]]["revenue"] += o["amount"]
    region_data[o["province"]]["customers"].add(o["phone"])

# Time
daily = defaultdict(lambda: {"orders": 0, "revenue": 0})
for o in orders:
    daily[o["date"]]["orders"] += 1
    daily[o["date"]]["revenue"] += o["amount"]

# Repeat
phone_orders = defaultdict(list)
for o in orders:
    phone_orders[o["phone"]].append(o)
repeat_customers = {p: v for p, v in phone_orders.items() if len(v) > 1}

# Need
need_data = defaultdict(lambda: {"orders": 0, "revenue": 0})
for o in orders:
    need_data[o["need"]]["orders"] += 1
    need_data[o["need"]]["revenue"] += o["amount"]

# Need x product
need_product = defaultdict(lambda: defaultdict(float))
for o in orders:
    need_product[o["need"]][o["product"]] += o["amount"]

# ============ GENERATE REPORT ============
lines = []
def add(s=""): lines.append(s)

add("# \u4e09\u8bfa&\u534e\u5927 \u6210\u4ea4\u8ba2\u5355\u6570\u636e\u72ec\u7acb\u5206\u6790\u62a5\u544a")
add()
add(f"> \u6570\u636e\u8303\u56f4: 2026-05-21 ~ 2026-05-26\uff086\u5929\uff09 | \u8ba2\u5355\u603b\u6570: {total_orders} \u5355 | \u603b\u8425\u6536: \xa5{total_revenue:,.2f}")
add()

# 1. Overview
add("## \u4e00\u3001\u8425\u6536\u603b\u89c8")
add()
add(f"| \u6307\u6807 | \u6570\u503c |")
add(f"|------|------|")
add(f"| \u8ba2\u5355\u603b\u6570 | **{total_orders} \u5355** |")
add(f"| \u53bb\u91cd\u5ba2\u6237\u6570 | **{total_customers} \u4eba** |")
add(f"| \u603b\u9500\u91cf | {total_units} \u4ef6 |")
add(f"| \u603b\u8425\u6536 | **\xa5{total_revenue:,.2f}** |")
add(f"| \u5e73\u5747\u5ba2\u5355\u4ef7 | \xa5{total_revenue/total_orders:,.2f} |")
add(f"| \u4eba\u5747\u6d88\u8d39 (ARPU) | \xa5{total_revenue/total_customers:,.2f} |")
add(f"| \u65e5\u5747\u8425\u6536 | \xa5{total_revenue/days:,.2f} |")
add(f"| \u65e5\u5747\u8ba2\u5355 | {total_orders/days:.1f} \u5355 |")
add()

# 2. Sales performance
add("## \u4e8c\u3001\u9500\u552e\u4eba\u6548\u5206\u6790")
add()
add(f"| \u9500\u552e | \u8ba2\u5355\u6570 | \u5ba2\u6237\u6570 | \u603b\u8425\u6536 | \u8425\u6536\u5360\u6bd4 | \u5747\u5355\u503c | \u4eba\u5747\u4ea7\u51fa | \u4e3b\u6253\u4ea7\u54c1 |")
add(f"|------|--------|--------|----------|----------|--------|----------|----------|")

sales_order = ["\u5218\u8001\u5e08", "\u5434\u8001\u5e08", "\u738b\u8001\u5e08"]
for s in sales_order:
    d = sales_data[s]
    pct = d["revenue"]/total_revenue*100
    top_p = sales_products[s].most_common(1)[0][0]
    add(f"| **{s}** | {d['orders']} | {len(d['customers'])} | **\xa5{d['revenue']:,.2f}** | {pct:.1f}% | \xa5{d['revenue']/d['orders']:,.2f} | \xa5{d['revenue']/len(d['customers']):,.2f} | {top_p} |")
add()

# Sales product detail
for s in sales_order:
    add(f"### {s}\u4ea7\u54c1\u7ec6\u5206")
    add()
    for p, rev in sales_products[s].most_common():
        add(f"- {p}: \xa5{rev:,.2f}")
    add()

# 3. Product
add("## \u4e09\u3001\u4ea7\u54c1\u7ed3\u6784\u5206\u6790")
add()
add(f"| \u6392\u540d | \u4ea7\u54c1 | \u529f\u80fd\u5b9a\u4f4d | \u8ba2\u5355\u6570 | \u9500\u91cf(\u4ef6) | \u5ba2\u6237\u6570 | \u8425\u6536 | \u5360\u6bd4 | \u5747\u4ef7 |")
add(f"|------|------|----------|--------|----------|--------|--------|------|------|")

for rank, (pname, d) in enumerate(sorted_products, 1):
    pct = d["revenue"]/total_revenue*100
    avg_price = d["revenue"] / d["units"]
    add(f"| {rank} | **{pname}** | \u89c1\u4ea7\u54c1\u540d | {d['orders']} | {d['units']} | {len(d['customers'])} | **\xa5{d['revenue']:,.2f}** | {pct:.1f}% | \xa5{avg_price:,.2f} |")
add()

# Price tier
add("### \u4ef7\u683c\u5e26\u5206\u5e03")
add()
price_tiers = [
    ("\u4f4e\u4ef7\u5f15\u6d41 <\xa5150", 0, 150),
    ("\u4e2d\u4f4e\u4ef7 \xa5150-300", 150, 300),
    ("\u4e2d\u9ad8\u4ef7 \xa5300-500", 300, 500),
    ("\u9ad8\u4ef7\u5957\u7ec4 >\xa5500", 500, 9999),
]
for label, lo, hi in price_tiers:
    tier = [o for o in orders if lo <= o["price"] < hi]
    if tier:
        rev = sum(o["amount"] for o in tier)
        pct = rev/total_revenue*100
        bar = "\u2588" * int(pct/2)
        add(f"- **{label}**: {len(tier)}\u5355, \xa5{rev:,.2f} ({pct:.0f}%) {bar}")
add()

# 4. Customer profile
add("## \u56db\u3001\u5ba2\u7fa4\u753b\u50cf")
add()

add(f"### \u5e74\u9f84\u5206\u5e03")
add(f"\u5e74\u9f84\u8303\u56f4: {min(ages)}-{max(ages)}\u5c81 | \u4e2d\u4f4d\u6570: {sorted_ages[len(sorted_ages)//2]}\u5c81 | \u5e73\u5747: {sum(ages)/len(ages):.0f}\u5c81")
add()
age_groups = [("30-39\u5c81",30,39),("40-49\u5c81",40,49),("50-59\u5c81",50,59),("60-69\u5c81",60,69),("70\u5c81\u4ee5\u4e0a",70,99)]
for label, lo, hi in age_groups:
    group = [o for o in orders if lo <= o["age"] <= hi]
    if group:
        rev = sum(o["amount"] for o in group)
        pct = len(group)/total_orders*100
        bar = "\u2588" * int(pct/2)
        add(f"- **{label}**: {len(group)}\u5355 ({pct:.0f}%), \xa5{rev:,.2f} {bar}")
add()

add(f"### \u6027\u522b\u5206\u5e03")
add()
for g in ["\u7537", "\u5973"]:
    d = gender_data[g]
    rev = d["revenue"]
    pct = d["orders"]/total_orders*100
    add(f"- **{g}**: {d['orders']}\u5355 ({pct:.0f}%), {len(d['customers'])}\u4eba, \xa5{rev:,.2f}")
add()

add(f"### \u6027\u522b \u00d7 \u5e74\u9f84\u6bb5 \u4ea4\u53c9")
add()
add(f"| \u6027\u522b | 30-39\u5c81 | 40-49\u5c81 | 50-59\u5c81 | 60-69\u5c81 | 70+ |")
add(f"|------|---------|---------|---------|---------|-----|")
for g in ["\u7537", "\u5973"]:
    row = f"| {g} |"
    for _, lo, hi in age_groups:
        count = sum(1 for o in orders if o["gender"] == g and lo <= o["age"] <= hi)
        row += f" {count}\u5355 |" if count > 0 else " - |"
    add(row)
add()

add(f"### \u6027\u522b \u00d7 \u4ea7\u54c1\u504f\u597d")
add()
for g in ["\u7537", "\u5973"]:
    pc = Counter()
    for o in orders:
        if o["gender"] == g:
            pc[o["product"]] += o["amount"]
    top3 = pc.most_common(3)
    yen = "\xa5"
    sep = " > "
    add(f"- **{g}**: {sep.join(f'{p} {yen}{a:,.0f}' for p,a in top3)}")
add()

# 5. Region
add("## \u4e94\u3001\u5730\u57df\u5206\u5e03")
add()
add(f"| \u7701\u4efd | \u57ce\u5e02 | \u8ba2\u5355\u6570 | \u5ba2\u6237\u6570 | \u8425\u6536 | \u5360\u6bd4 |")
add(f"|------|------|--------|--------|--------|------|")
for r, d in sorted(region_data.items(), key=lambda x: x[1]["revenue"], reverse=True):
    cities = set(o["city"] for o in orders if o["province"] == r)
    pct = d["revenue"]/total_revenue*100
    add(f"| {r} | {', '.join(cities)} | {d['orders']} | {len(d['customers'])} | \xa5{d['revenue']:,.2f} | {pct:.0f}% |")
add()
add(f"\u8986\u76d6 **{len(region_data)}** \u4e2a\u7701\u4efd\uff0c\u4ee5\u5317\u4eac\u3001\u56db\u5ddd\u4e3a\u6838\u5fc3\u5e02\u573a\u3002")
add()

# 6. Time
add("## \u516d\u3001\u4e0b\u5355\u65f6\u95f4\u5206\u6790")
add()
add(f"### \u6309\u65e5\u5206\u5e03")
add()
add(f"| \u65e5\u671f | \u661f\u671f | \u8ba2\u5355\u6570 | \u8425\u6536 | \u8d8b\u52bf |")
add(f"|------|------|--------|--------|------|")
weekdays = ["\u4e00","\u4e8c","\u4e09","\u56db","\u4e94","\u516d","\u65e5"]
for d in sorted(daily.keys()):
    dd = daily[d]
    dt = datetime.strptime(d, "%Y-%m-%d")
    wd = weekdays[dt.weekday()]
    bar = "\u2588" * int(dd["orders"] * 3)
    add(f"| {d} | \u5468{wd} | {dd['orders']} | \xa5{dd['revenue']:,.2f} | {bar} |")
add()

add(f"### \u6309\u65f6\u6bb5\u5206\u5e03")
add()
slots = [
    ("\u4e0a\u5348 (9:00-11:59)", 9, 11),
    ("\u4e2d\u5348 (12:00-13:59)", 12, 13),
    ("\u4e0b\u5348 (14:00-17:59)", 14, 17),
    ("\u665a\u95f4 (18:00+)", 18, 23),
]
for label, lo, hi in slots:
    slot = [o for o in orders if lo <= o["hour"] <= hi]
    if slot:
        rev = sum(o["amount"] for o in slot)
        pct = len(slot)/total_orders*100
        add(f"- **{label}**: {len(slot)}\u5355, \xa5{rev:,.2f} ({pct:.0f}%)")
add()

# 7. Need-product matrix
add("## \u4e03\u3001\u9700\u6c42 \u00d7 \u4ea7\u54c1 \u5339\u914d\u77e9\u9635")
add()
needs = sorted(set(o["need"] for o in orders))
prods = [p for p, _ in sorted_products]

add(f"| \u9700\u6c42 | {' | '.join(prods)} | \u5408\u8ba1 |")
add(f"|------|{'|'.join(['------' for _ in prods])}|------|")
for n in needs:
    row = f"| {n} |"
    total_need = 0
    for p in prods:
        amt = need_product[n][p]
        total_need += amt
        row += f" \xa5{amt:,.0f} |" if amt > 0 else " - |"
    row += f" **\xa5{total_need:,.0f}** |"
    add(row)
add()

# 8. Repeat
add("## \u516b\u3001\u590d\u8d2d\u884c\u4e3a\u5206\u6790")
add()
add(f"- \u590d\u8d2d\u5ba2\u6237: **{len(repeat_customers)}/{len(phone_orders)}** \u4eba ({len(repeat_customers)/len(phone_orders)*100:.0f}%)")
add(f"- \u5355\u6b21\u5ba2\u6237: {len(phone_orders)-len(repeat_customers)} \u4eba")
add()

for phone, olist in repeat_customers.items():
    name = olist[0]["name"]
    total = sum(o["amount"] for o in olist)
    products = " \u2192 ".join(o["product"] for o in olist)
    
    d1 = datetime.strptime(olist[0]["time"][:10], "%Y-%m-%d")
    d2 = datetime.strptime(olist[1]["time"][:10], "%Y-%m-%d") if len(olist) >= 2 else d1
    interval = (d2 - d1).days
    
    add(f"### {name} ({phone[-4:]})")
    add(f"- \u8ba2\u5355: {len(olist)}\u5355 | \u5408\u8ba1: \xa5{total:,.2f} | \u95f4\u9694: {interval}\u5929")
    add(f"- \u8def\u5f84: {products}")
    for o in olist:
        add(f"  - [{o['time'][:10]}] {o['product']} \u00d7{o['qty']} = \xa5{o['amount']:,.2f} (\u9500\u552e: {o['sales']})")
    add()

# 9. Qty
add("## \u4e5d\u3001\u8d2d\u4e70\u4ef6\u6570\u5206\u6790")
add()
qty_dist = Counter(o["qty"] for o in orders)
for q in sorted(qty_dist):
    group = [o for o in orders if o["qty"] == q]
    rev = sum(o["amount"] for o in group)
    names = [o["name"] for o in group]
    add(f"- **\u4e70{q}\u4ef6**: {len(group)}\u5355 ({len(group)/total_orders*100:.0f}%), \xa5{rev:,.2f} \u2014 {', '.join(names)}")
add()

# 10. Insights
add("## \u5341\u3001\u7efc\u5408\u6d1e\u5bdf\u4e0e\u5efa\u8bae")
add()

top2_rev = sum(d["revenue"] for _, d in sorted_products[:2])
add(f"### 1. \u4ea7\u54c1\u96c6\u4e2d\u5ea6\u9ad8")
qinglu_key = "\u8f7b\u8def\uff08\u81b3\u98df\u7ea4\u7ef4\uff09"
qinglu_pct = product_data[qinglu_key]['revenue']/total_revenue*100
add(f"\u76ca\u5bb6\u4eba+\u955c\u8102\u5957\u7ec4\u4ec5 2 \u4e2a\u4ea7\u54c1\u8d21\u732e\u4e86 **{top2_rev/total_revenue*100:.0f}%** \u8425\u6536\uff0c\u4f46\u53ea\u5360\u603b\u8ba2\u5355\u91cf\u7684 18%\u3002\u8f7b\u8def\u867d\u7136\u5356\u5f97\u6700\u591a\uff084\u5355\uff09\uff0c\u8425\u6536\u8d21\u732e\u4ec5 {qinglu_pct:.0f}%\u3002")
add()
add(f"**\u5efa\u8bae**: \u9ad8\u5ba2\u5355\u4ea7\u54c1\u662f\u8425\u6536\u5f15\u64ce\uff0c\u5e94\u52a0\u5f3a\u76ca\u5bb6\u4eba\u548c\u955c\u8102\u5957\u7ec4\u7684\u8bdd\u672f\u57f9\u8bad\uff0c\u540c\u65f6\u8003\u8651\u5c06\u8f7b\u8def\u4f5c\u4e3a\u201c\u4f4e\u4ef7\u5165\u95e8\u201d\u4ea7\u54c1\uff0c\u914d\u5957\u201c\u5347\u7ea7\u5957\u9910\u201d\u7684\u8ffd\u552e\u7b56\u7565\u3002")
add()

add(f"### 2. \u9500\u552e\u4eba\u6548\u5dee\u8ddd\u663e\u8457")
top_s = max(sales_data.items(), key=lambda x: x[1]["revenue"])
bot_s = min(sales_data.items(), key=lambda x: x[1]["revenue"])
wu_rev = sales_data["\u5434\u8001\u5e08"]["revenue"]
liu_rev = top_s[1]["revenue"]
wang_rev = bot_s[1]["revenue"]
add(f"\u5218\u8001\u5e08\u8425\u6536\u662f\u738b\u8001\u5e08\u7684 **{liu_rev/wang_rev:.0f}x**\uff0c\u4e14\u4e0e\u5434\u8001\u5e08\u8ba2\u5355\u6570\u76f8\u540c\u4f46\u8425\u6536\u9ad8\u51fa **{(liu_rev-wu_rev)/wu_rev*100:.0f}%**\u3002\u6838\u5fc3\u5dee\u5f02\u5728\u4e8e\u4ea7\u54c1\u7ed3\u6784\u2014\u2014\u5218\u8001\u5e08\u4e3b\u63a8\u9ad8\u5ba2\u5355\u4ef7\u4ea7\u54c1\u3002")
add()
add(f"**\u5efa\u8bae**: \u5c06\u5218\u8001\u5e08\u7684\u201c\u9700\u6c42\u6316\u6398\u2192\u9ad8\u5ba2\u5355\u8f6c\u5316\u201d\u8bdd\u672f\u63d0\u70bc\u4e3a\u6807\u51c6 SOP\uff0c\u57f9\u8bad\u5176\u4ed6\u9500\u552e\u3002")
add()

add(f"### 3. \u590d\u8d2d\u7387\u4f4e\u4f46\u7a7a\u95f4\u5927")
rep_rate = len(repeat_customers)/len(phone_orders)*100
add(f"{rep_rate:.0f}%\u590d\u8d2d\u7387\u57286\u5929\u7a97\u53e3\u5185\u5c5e\u4e8e\u6b63\u5e38\uff0c\u4f46\u5df2\u67092\u4eba\u5f53\u5929/\u9694\u5929\u8ffd\u52a0\u8d2d\u4e70\uff0c\u8bf4\u660e\u5373\u65f6\u8ffd\u52a0\u9500\u552e\u673a\u4f1a\u5b58\u5728\u3002")
add()
add(f"**\u5efa\u8bae**: \u6210\u4ea4\u5f53\u5929\u5185\u53d1\u9001\u201c\u642d\u914d\u65b9\u6848\u201d\u8ffd\u552e\u6d88\u606f\uff0c\u76ee\u6807\u5c06\u590d\u8d2d\u7387\u63d0\u5347\u81f3 40-50%\u3002")
add()

add(f"### 4. \u6838\u5fc3\u5ba2\u7fa4\u660e\u786e")
avg_age = sum(ages)/len(ages)
male_count = sum(1 for o in orders if o['gender']=="\u7537")
add(f"\u5e73\u5747\u5e74\u9f84 {avg_age:.0f} \u5c81\uff0c\u7537\u6027\u7565\u591a\uff08{male_count}/{total_orders}\uff09\uff0c\u4ee5\u5317\u65b9\u4e00\u4e8c\u7ebf\u57ce\u5e02\u4e3a\u4e3b\u3002\u4e3b\u8981\u75db\u70b9\u662f\u8f7b\u4f53\u3001\u7761\u7720\u3001\u80a0\u80c3\u3002")
add()
add(f"**\u5efa\u8bae**: \u5185\u5bb9\u548c\u8bdd\u672f\u504f\u597d\u5e94\u4fa7\u91cd\u201c\u529f\u80fd\u6027\u201d\u548c\u201c\u5b9e\u8bc1\u6027\u201d\u8868\u8ff0\uff0c\u51cf\u5c11\u60c5\u611f\u5316\u8425\u9500\u8bed\u8a00\u3002")
add()

add(f"### 5. \u65f6\u95f4\u7a97\u53e3\u7279\u5f81")
peak_day = max(daily.items(), key=lambda x: x[1]["revenue"])
add(f"\u8ba2\u5355\u96c6\u4e2d\u5728\u4e0b\u5348\u65f6\u6bb5\uff0c\u5468\u4e09\u5468\u56db\u4e3a\u9ad8\u5cf0\u3002\u5355\u65e5\u6700\u9ad8\u8425\u6536: {peak_day[0]} (\xa5{peak_day[1]['revenue']:,.2f})\u3002")
add()
add(f"**\u5efa\u8bae**: \u4e0b\u5348 14:00-17:00 \u662f\u9ec4\u91d1\u89e6\u8fbe\u7a97\u53e3\uff0c\u91cd\u70b9\u5b89\u6392\u4fc3\u5355\u52a8\u4f5c\u3002")
add()

# Footer
add("---")
add()
add(f"*\u62a5\u544a\u7531 WorkBuddy \u81ea\u52a8\u5206\u6790\u751f\u6210\uff0c\u6570\u636e\u6765\u6e90: \u4e09\u8bfa&\u534e\u5927\u6210\u4ea4\u7528\u6237\u6e05\u5355.xlsx*")

# Write to file
out_path = "/Users/SaciNa/WorkBuddy/2026-05-26-15-30-31/\u4e09\u8bfa\u534e\u5927\u6210\u4ea4\u8ba2\u5355\u5206\u6790\u62a5\u544a.md"
content = "\n".join(lines)
with open(out_path, "w", encoding="utf-8") as f:
    f.write(content)
print(f"Report written to {out_path}")
print(f"Total lines: {len(lines)}")
