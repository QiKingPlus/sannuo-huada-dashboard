#!/usr/bin/env python3
"""Cross-reference chat data with order data and enrich the report."""

# ============ ORDER DATA (from Excel) ============
orders = [
    {"row": 2, "sales": "WU", "time": "2026-05-26 11:18", "product": "yixiu", "price": 253.3, "name": "XieGui", "phone": "13975315486", "qty": 1, "gender": "F", "age": 54, "need": "sleep", "region": "Hunan-Zhuzhou"},
    {"row": 3, "sales": "LIU", "time": "2026-05-26 10:33", "product": "qinglu", "price": 139, "name": "JiaRen", "phone": "15073313688", "qty": 1, "gender": "M", "age": 42, "need": "glucose", "region": "Beijing"},
    {"row": 4, "sales": "WANG", "time": "2026-05-26 10:10", "product": "yixiu", "price": 253.3, "name": "WenXiaowei", "phone": "17703100088", "qty": 1, "gender": "M", "age": 44, "need": "sleep", "region": "Hebei-Handan"},
    {"row": 5, "sales": "WU", "time": "2026-05-25 17:35", "product": "qinglu", "price": 139, "name": "XieHong", "phone": "15911579927", "qty": 1, "gender": "F", "age": 64, "need": "glucose", "region": "Yunnan-Kunming"},
    {"row": 6, "sales": "WU", "time": "2026-05-25 14:38", "product": "changqing", "price": 159.8, "name": "ShenShanming", "phone": "18967469708", "qty": 3, "gender": "M", "age": 59, "need": "gut", "region": "Zhejiang-Jinhua"},
    {"row": 7, "sales": "LIU", "time": "2026-05-25 11:51", "product": "yijiaren", "price": 398, "name": "BeiChen", "phone": "18190919791", "qty": 3, "gender": "F", "age": 45, "need": "family", "region": "Sichuan-Chengdu"},
    {"row": 8, "sales": "WU", "time": "2026-05-22 17:25", "product": "qinglu", "price": 139, "name": "QinShuai", "phone": "15262003719", "qty": 1, "gender": "M", "age": 37, "need": "weight", "region": "Jiangsu-Xuzhou"},
    {"row": 9, "sales": "LIU", "time": "2026-05-21 17:20", "product": "qinglu", "price": 139, "name": "ZhangWensheng", "phone": "13501134294", "qty": 1, "gender": "M", "age": 59, "need": "weight", "region": "Beijing"},
    {"row": 10, "sales": "LIU", "time": "2026-05-21 17:14", "product": "jingzhi_taozu", "price": 566.6, "name": "ZhangWensheng", "phone": "13501134294", "qty": 1, "gender": "M", "age": 59, "need": "weight", "region": "Beijing"},
    {"row": 11, "sales": "WU", "time": "2026-05-21 14:34", "product": "yichang", "price": 438.6, "name": "YangXianzhai", "phone": "13986460439", "qty": 1, "gender": "M", "age": 70, "need": "gut", "region": "Hubei-Xiaogan"},
    {"row": 12, "sales": "LIU", "time": "2026-05-21 11:34", "product": "jingzhi_taozu", "price": 566.6, "name": "BeiChen", "phone": "18190919791", "qty": 1, "gender": "F", "age": 45, "need": "weight", "region": "Sichuan-Chengdu"},
]

# Product name mapping
product_cn = {
    "yixiu": "\u76ca\u4f11\uff08\u7761\u7720\uff09",
    "qinglu": "\u8f7b\u8def\uff08\u81b3\u98df\u7ea4\u7ef4\uff09",
    "changqing": "\u7545\u9752\uff08\u5438\u6536/\u589e\u91cd\uff09",
    "yijiaren": "\u76ca\u5bb6\u4eba\uff08\u57fa\u7840\u76ca\u751f\u83cc\uff09",
    "jingzhi_taozu": "\u955c\u8102\u00b7\u8f7b+\u955c\u8102\u00b7\u963b\u5957\u7ec4",
    "yichang": "\u76ca\u7545\uff08\u4fbf\u79d8\uff09",
}

sales_cn = {"WU": "\u5434\u8001\u5e08", "LIU": "\u5218\u8001\u5e08", "WANG": "\u738b\u8001\u5e08"}
gender_cn = {"M": "\u7537", "F": "\u5973"}
need_cn = {
    "sleep": "\u7761\u7720", "glucose": "\u8840\u7cd6", "gut": "\u80a0\u80c3",
    "weight": "\u8f7b\u4f53", "family": "\u5bb6\u4eba\u5065\u5eb7"
}

# ============ CHAT-TO-ORDER MAPPING ============
chat_users = [
    {"cid": 1, "advisor": "\u8d75\u8001\u5e08", "customer": "\u5317\u8fb0", "phone": "18190919791", "msgs": 124, "a_msgs": 65, "c_msgs": 59},
    {"cid": 2, "advisor": "\u8d75\u8001\u5e08", "customer": "\u5eb7\u5a01\u70df\u91521350", "phone": "13501134294", "msgs": 65, "a_msgs": 32, "c_msgs": 33},
    {"cid": 3, "advisor": "\u59e3\u59e3", "customer": "yxz", "phone": "13986460439", "msgs": 40, "a_msgs": 23, "c_msgs": 17},
    {"cid": 4, "advisor": "\u59e3\u59e3", "customer": "\u963f\u660e", "phone": "18967469708", "msgs": 99, "a_msgs": 52, "c_msgs": 47},
    {"cid": 5, "advisor": "\u8bfa\u8bfa", "customer": "\u6cb3\u5317\u521b\u683c\u7279", "phone": "17703100088", "msgs": 105, "a_msgs": 56, "c_msgs": 49},
    {"cid": 6, "advisor": "\u59e3\u59e3", "customer": "1.2.3.", "phone": "15262003719", "msgs": 44, "a_msgs": 23, "c_msgs": 21},
    {"cid": 7, "advisor": "\u59e3\u59e3", "customer": "\u6307\u7eb9", "phone": "13975315486", "msgs": 110, "a_msgs": 49, "c_msgs": 61},
    {"cid": 8, "advisor": "\u8d75\u8001\u5e08", "customer": "\u4f3c\u66fe\u76f8\u8bc6", "phone": "15073313688", "msgs": 98, "a_msgs": 55, "c_msgs": 43},
]

# Cross-reference
matched = []
unmatched = []
user_data = {}  # phone -> {chat, orders}

for o in orders:
    phone = o["phone"]
    found = False
    for c in chat_users:
        if c["phone"] == phone:
            found = True
            if phone not in user_data:
                user_data[phone] = {"chat": c, "orders": [], "total": 0}
            user_data[phone]["orders"].append(o)
            user_data[phone]["total"] += o["price"] * o["qty"]
            matched.append(o)
            break
    if not found:
        unmatched.append(o)

total_revenue = sum(o["price"] * o["qty"] for o in orders)

# Print results
print("=" * 70)
print("Cross-reference: Chat Records vs Order Data")
print("=" * 70)
print(f"Total orders: {len(orders)}")
print(f"Matched to chats: {len(matched)}")
print(f"Unmatched: {len(unmatched)}")
print(f"Total revenue: {total_revenue:,.1f}")

print(f"\n{'=' * 70}")
print("Per-User Match Details")
print("=" * 70)

for phone, data in sorted(user_data.items(), key=lambda x: x[1]["total"], reverse=True):
    c = data["chat"]
    olist = data["orders"]
    t = data["total"]
    o0 = olist[0]
    
    print(f"\n-- {c['customer']} -> {o0['name']} --")
    print(f"Phone: {phone} | Age: {o0['age']} | Gender: {gender_cn[o0['gender']]} | Region: {o0['region']}")
    print(f"Chat advisor: {c['advisor']} | Order sales: {sales_cn[o0['sales']]}")
    print(f"Chat msgs: {c['msgs']} (A:{c['a_msgs']}/C:{c['c_msgs']})")
    print(f"Orders:")
    for o in olist:
        print(f"  [{o['time']}] {product_cn[o['product']]} x{o['qty']} | {o['price']} | Need: {need_cn[o['need']]}")
    print(f"Total spend: {t:,.1f}")

if unmatched:
    print(f"\n{'=' * 70}")
    print("Unmatched Orders (no chat record)")
    print("=" * 70)
    for o in unmatched:
        print(f"  {o['name']} ({o['phone']}) - {product_cn[o['product']]} x{o['qty']} | {o['price']} | {sales_cn[o['sales']]}")

# === METRICS ===
print(f"\n{'=' * 70}")
print("Revenue Metrics")
print("=" * 70)

# By sales
sr = {}
so = {}
for o in orders:
    s = o["sales"]
    sr[s] = sr.get(s, 0) + o["price"] * o["qty"]
    so[s] = so.get(s, 0) + 1

print("\nSales performance:")
for s, rev in sorted(sr.items(), key=lambda x: x[1], reverse=True):
    print(f"  {sales_cn[s]}: {rev:,.1f} ({so[s]} orders, AOV {rev/so[s]:,.1f})")

# By product
pr = {}
pq = {}
for o in orders:
    p = o["product"]
    pr[p] = pr.get(p, 0) + o["price"] * o["qty"]
    pq[p] = pq.get(p, 0) + o["qty"]

print("\nProduct revenue:")
for p, rev in sorted(pr.items(), key=lambda x: x[1], reverse=True):
    print(f"  {product_cn[p]}: {rev:,.1f} ({pq[p]} units)")

# Key ratios
arpu = total_revenue / len(chat_users)
aov = total_revenue / len(orders)
repeats = sum(1 for d in user_data.values() if len(d["orders"]) > 1)

print(f"\nKey ratios:")
print(f"  ARPU (per chat user): {arpu:,.1f}")
print(f"  AOV (per order): {aov:,.1f}")
print(f"  Repeat rate: {repeats}/{len(user_data)} users")

# Demographics
ages = [o["age"] for o in orders]
genders = [o["gender"] for o in orders]
print(f"\nDemographics:")
print(f"  Age: {min(ages)}-{max(ages)}, avg {sum(ages)/len(ages):.0f}")
print(f"  Gender: M={genders.count('M')}/{len(genders)}, F={genders.count('F')}/{len(genders)}")

# Regions
regions = {}
for o in orders:
    r = o["region"]
    regions[r] = regions.get(r, 0) + 1
print(f"  Regions: {', '.join(f'{k}({v})' for k,v in sorted(regions.items(), key=lambda x: x[1], reverse=True))}")

# Advisor mapping
print(f"\nAdvisor mapping (chat nickname -> order name):")
am = {}
for phone, data in user_data.items():
    ca = data["chat"]["advisor"]
    os = sales_cn[data["orders"][0]["sales"]]
    key = f"{ca} -> {os}"
    if key not in am:
        am[key] = []
    am[key].append(data["chat"]["customer"])
for k, v in am.items():
    print(f"  {k}: {', '.join(v)}")

print("\nDone!")
