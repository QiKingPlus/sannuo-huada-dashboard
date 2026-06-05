#!/usr/bin/env python3
"""三诺&华大成交订单独立分析"""

import re
from collections import Counter, defaultdict
from datetime import datetime

# ============ DATA ============
orders = [
    {"row":2, "sales":"吴老师", "time":"2026-05-26 11:18:14", "product_name":"【三诺专享】专研升级3.0【优美达·益休益生菌固体饮料】入夜一条，身体自然归暮（30袋/盒）", "price":253.3, "name":"谢贵", "phone":"13975315486", "qty":1, "gender":"女", "age":54, "need":"睡眠", "province":"湖南省", "city":"株洲市"},
    {"row":3, "sales":"刘老师", "time":"2026-05-26 10:33:02", "product_name":"【三诺专享】【轻路 复合益生元膳食纤维固体饮料】补菌之外，更要养菌！五元复配，好菌工作事半功倍（15袋/盒）", "price":139, "name":"贾仁", "phone":"15073313688", "qty":1, "gender":"男", "age":42, "need":"血糖", "province":"北京市", "city":"北京市"},
    {"row":4, "sales":"王老师", "time":"2026-05-26 10:10:09", "product_name":"【三诺专享】专研升级3.0【优美达·益休益生菌固体饮料】入夜一条，身体自然归暮（30袋/盒）", "price":253.3, "name":"温晓炜", "phone":"17703100088", "qty":1, "gender":"男", "age":44, "need":"睡眠", "province":"河北省", "city":"邯郸市"},
    {"row":5, "sales":"吴老师", "time":"2026-05-25 17:35:59", "product_name":"【三诺专享】【轻路 复合益生元膳食纤维固体饮料】补菌之外，更要养菌！五元复配，好菌工作事半功倍（15袋/盒）", "price":139, "name":"谢红", "phone":"15911579927", "qty":1, "gender":"女", "age":64, "need":"血糖", "province":"云南省", "city":"昆明"},
    {"row":6, "sales":"吴老师", "time":"2026-05-25 14:38:58", "product_name":"【三诺专享】【优美达·畅青益生菌压片糖果】拒绝拉到扶墙起！肠肠平衡稳菌心！（15袋/盒）", "price":159.8, "name":"沈善明", "phone":"18967469708", "qty":3, "gender":"男", "age":59, "need":"肠胃", "province":"浙江省", "city":"金华市"},
    {"row":7, "sales":"刘老师", "time":"2026-05-25 11:51:11", "product_name":"【三诺专享】【优美达·益家人益生菌】肠道益生菌 全家共享每益天(15袋/盒）", "price":398, "name":"北辰", "phone":"18190919791", "qty":3, "gender":"女", "age":45, "need":"家人健康", "province":"四川省", "city":"成都市"},
    {"row":8, "sales":"吴老师", "time":"2026-05-22 17:25:30", "product_name":"【三诺专享】【轻路 复合益生元膳食纤维固体饮料】补菌之外，更要养菌！五元复配，好菌工作事半功倍（15袋/盒）", "price":139, "name":"秦帅", "phone":"15262003719", "qty":1, "gender":"男", "age":37, "need":"轻体", "province":"江苏省", "city":"徐州市"},
    {"row":9, "sales":"刘老师", "time":"2026-05-21 17:20:40", "product_name":"【三诺专享】【轻路 复合益生元膳食纤维固体饮料】补菌之外，更要养菌！五元复配，好菌工作事半功倍（15袋/盒）", "price":139, "name":"张文生", "phone":"13501134294", "qty":1, "gender":"男", "age":59, "need":"轻体", "province":"北京市", "city":"北京市"},
    {"row":10,"sales":"刘老师", "time":"2026-05-21 17:14:07", "product_name":"【三诺专享】【管理套组】【镜脂·轻+镜脂·阻】管理好搭档，自律生活我来定义！", "price":566.6, "name":"张文生", "phone":"13501134294", "qty":1, "gender":"男", "age":59, "need":"轻体", "province":"北京市", "city":"北京市"},
    {"row":11,"sales":"吴老师", "time":"2026-05-21 14:34:17", "product_name":"【三诺专享】【优美达·益畅益生菌】连续10年口碑！五种益生菌+三种益生元 噗噗自由 蹲坑快乐（30袋/盒）", "price":438.6, "name":"杨先斋", "phone":"13986460439", "qty":1, "gender":"男", "age":70, "need":"肠胃", "province":"湖北省", "city":"孝感市"},
    {"row":12,"sales":"刘老师", "time":"2026-05-21 11:34:07", "product_name":"【三诺专享】【管理套组】【镜脂·轻+镜脂·阻】管理好搭档，自律生活我来定义！", "price":566.6, "name":"北辰", "phone":"18190919791", "qty":1, "gender":"女", "age":45, "need":"轻体", "province":"四川省", "city":"成都市"},
]

# Product classification
def classify_product(name):
    if "益休" in name: return "益休（睡眠）"
    if "轻路" in name: return "轻路（膳食纤维）"
    if "畅青" in name: return "畅青（吸收/增重）"
    if "益家人" in name: return "益家人（基础益生菌）"
    if "镜脂" in name: return "镜脂套组（体重管理）"
    if "益畅" in name: return "益畅（便秘）"
    return "其他"

for o in orders:
    o["product"] = classify_product(o["product_name"])
    o["amount"] = o["price"] * o["qty"]
    o["date"] = o["time"][:10]
    o["hour"] = int(o["time"][11:13])

# ============ KEY METRICS ============
total_revenue = sum(o["amount"] for o in orders)
total_orders = len(orders)
total_customers = len(set(o["phone"] for o in orders))
total_units = sum(o["qty"] for o in orders)

print("=" * 70)
print("三诺&华大 成交订单数据分析报告")
print(f"数据范围: 2026-05-21 ~ 2026-05-26（6天）")
print("=" * 70)

# 1. 总览
print("\n" + "=" * 70)
print("一、营收总览")
print("=" * 70)
print(f"  订单总数: {total_orders} 单")
print(f"  客户数: {total_customers} 人（去重手机号）")
print(f"  总销量: {total_units} 件")
print(f"  总营收: {total_revenue:,.2f}")
print(f"  平均客单价: {total_revenue/total_orders:,.2f}")
print(f"  人均消费: {total_revenue/total_customers:,.2f}")
print(f"  日均营收: {total_revenue/6:,.2f}")
print(f"  日均订单: {total_orders/6:.1f}")

# 2. 销售人效
print("\n" + "=" * 70)
print("二、销售人效分析")
print("=" * 70)

sales_data = defaultdict(lambda: {"orders": 0, "revenue": 0, "customers": set(), "units": 0})
for o in orders:
    s = o["sales"]
    sales_data[s]["orders"] += 1
    sales_data[s]["revenue"] += o["amount"]
    sales_data[s]["customers"].add(o["phone"])
    sales_data[s]["units"] += o["qty"]

# Product mix per sales
sales_products = defaultdict(Counter)
for o in orders:
    sales_products[o["sales"]][o["product"]] += o["amount"]

print(f"{'销售':<8} {'订单':<5} {'客户':<5} {'营收':>12} {'占比':>7} {'均单':>8} {'人均':>8} {'主打产品'}")
print("-" * 70)
for s in ["刘老师", "吴老师", "王老师"]:
    d = sales_data[s]
    pct = d["revenue"]/total_revenue*100
    top_product = sales_products[s].most_common(1)[0][0] if sales_products[s] else "-"
    print(f"{s:<8} {d['orders']:<5} {len(d['customers']):<5} {d['revenue']:>10,.2f} {pct:>6.1f}% {d['revenue']/d['orders']:>7,.2f} {d['revenue']/len(d['customers']):>7,.2f} {top_product}")

# 3. 产品分析
print("\n" + "=" * 70)
print("三、产品结构分析")
print("=" * 70)

product_data = defaultdict(lambda: {"orders": 0, "revenue": 0, "units": 0, "customers": set(), "prices": []})
for o in orders:
    p = o["product"]
    product_data[p]["orders"] += 1
    product_data[p]["revenue"] += o["amount"]
    product_data[p]["units"] += o["qty"]
    product_data[p]["customers"].add(o["phone"])
    product_data[p]["prices"].append(o["price"])

# Sort by revenue desc
sorted_products = sorted(product_data.items(), key=lambda x: x[1]["revenue"], reverse=True)

print(f"{'产品':<22} {'订单':>4} {'销量':>4} {'客户':>4} {'营收':>10} {'占比':>6} {'均价':>8} {'单品营收贡献'}")
print("-" * 70)
for pname, d in sorted_products:
    pct = d["revenue"]/total_revenue*100
    avg_price = sum(d["prices"])/len(d["prices"])
    per_order = d["revenue"]/d["orders"]
    print(f"{pname:<22} {d['orders']:>4} {d['units']:>4} {len(d['customers']):>4} {d['revenue']:>8,.2f} {pct:>5.1f}% {avg_price:>7,.2f} {per_order:>8,.2f}")

# Price tier analysis
print("\n价格带分析:")
price_tiers = [
    ("<150 (低价引流)", 0, 150),
    ("150-300 (中低价)", 150, 300),
    ("300-500 (中高价)", 300, 500),
    (">500 (高价套组)", 500, 9999),
]
for label, lo, hi in price_tiers:
    tier_orders = [o for o in orders if lo <= o["price"] < hi]
    if tier_orders:
        rev = sum(o["amount"] for o in tier_orders)
        print(f"  {label}: {len(tier_orders)}单, {rev:,.2f} ({rev/total_revenue*100:.1f}%)")

# 4. 客群画像
print("\n" + "=" * 70)
print("四、客群画像")
print("=" * 70)

# Age
ages = [o["age"] for o in orders]
age_groups = [
    ("30-39岁", 30, 39),
    ("40-49岁", 40, 49),
    ("50-59岁", 50, 59),
    ("60-69岁", 60, 69),
    ("70岁以上", 70, 99),
]
print(f"\n年龄分布: {min(ages)}-{max(ages)}岁, 中位数 {sorted(ages)[len(ages)//2]}岁, 平均 {sum(ages)/len(ages):.0f}岁")
for label, lo, hi in age_groups:
    group = [o for o in orders if lo <= o["age"] <= hi]
    if group:
        rev = sum(o["amount"] for o in group)
        print(f"  {label}: {len(group)}单 ({len(group)/total_orders*100:.0f}%), 营收 {rev:,.2f}")

# Gender
print(f"\n性别分布:")
for g in ["男", "女"]:
    group = [o for o in orders if o["gender"] == g]
    rev = sum(o["amount"] for o in group)
    customers = len(set(o["phone"] for o in group))
    print(f"  {g}: {len(group)}单 ({len(group)/total_orders*100:.0f}%), {customers}人, 营收 {rev:,.2f}")

# Gender x Product
print(f"\n性别 × 产品偏好:")
for g in ["男", "女"]:
    g_orders = [o for o in orders if o["gender"] == g]
    pc = Counter()
    for o in g_orders:
        pc[o["product"]] += o["amount"]
    top3 = pc.most_common(3)
    print(f"  {g}: {', '.join(f'{p}({a:,.0f})' for p,a in top3)}")

# Age x Need
print(f"\n年龄段 × 核心需求:")
for label, lo, hi in age_groups:
    group = [o for o in orders if lo <= o["age"] <= hi]
    if group:
        nc = Counter(o["need"] for o in group)
        print(f"  {label}: {', '.join(f'{n}({c})' for n,c in nc.most_common())}")

# 5. 地域
print("\n" + "=" * 70)
print("五、地域分布")
print("=" * 70)

region_data = defaultdict(lambda: {"orders": 0, "revenue": 0, "customers": set()})
for o in orders:
    region = o["province"]
    region_data[region]["orders"] += 1
    region_data[region]["revenue"] += o["amount"]
    region_data[region]["customers"].add(o["phone"])

print(f"{'省份':<8} {'订单':<5} {'客户':<5} {'营收':>10} {'占比':>6} {'人均':>8}")
print("-" * 50)
for r, d in sorted(region_data.items(), key=lambda x: x[1]["revenue"], reverse=True):
    pct = d["revenue"]/total_revenue*100
    print(f"{r:<8} {d['orders']:<5} {len(d['customers']):<5} {d['revenue']:>8,.2f} {pct:>5.1f}% {d['revenue']/len(d['customers']):>7,.2f}")

# 6. 时间
print("\n" + "=" * 70)
print("六、下单时间分析")
print("=" * 70)

# Daily
daily = defaultdict(lambda: {"orders": 0, "revenue": 0})
for o in orders:
    daily[o["date"]]["orders"] += 1
    daily[o["date"]]["revenue"] += o["amount"]

print("\n按日分布:")
for d in sorted(daily.keys()):
    dd = daily[d]
    bar = "█" * int(dd["orders"] * 3)
    print(f"  {d}  {dd['orders']}单  {dd['revenue']:>8,.2f}  {bar}")

# Hourly
hourly = defaultdict(lambda: {"orders": 0, "revenue": 0})
for o in orders:
    hourly[o["hour"]]["orders"] += 1
    hourly[o["hour"]]["revenue"] += o["amount"]

print(f"\n按时段分布:")
time_slots = [("上午 (9-12)", 9, 12), ("下午 (13-17)", 13, 17), ("晚间 (18-23)", 18, 23)]
for label, lo, hi in time_slots:
    slot_orders = [o for o in orders if lo <= o["hour"] <= hi]
    if slot_orders:
        rev = sum(o["amount"] for o in slot_orders)
        print(f"  {label}: {len(slot_orders)}单, {rev:,.2f} ({rev/total_revenue*100:.0f}%)")

# 7. 需求-产品匹配
print("\n" + "=" * 70)
print("七、需求 × 产品 匹配矩阵")
print("=" * 70)

need_product = defaultdict(lambda: defaultdict(lambda: {"orders": 0, "revenue": 0}))
needs_list = sorted(set(o["need"] for o in orders))
products_list = sorted(set(o["product"] for o in orders))

# Matrix header
header = f"{'需求':<10}"
for p in products_list:
    short = p[:4]
    header += f" {short:>6}"
print(header)
print("-" * (10 + 7 * len(products_list)))

for n in needs_list:
    row = f"{n:<10}"
    for p in products_list:
        d = need_product[n][p]
        if d["orders"] > 0:
            row += f" {d['revenue']:>5,.0f}"
        else:
            row += f" {'-':>6}"
    print(row)

# Fill the actual data
for o in orders:
    need_product[o["need"]][o["product"]]["orders"] += 1
    need_product[o["need"]][o["product"]]["revenue"] += o["amount"]

print("\n实际数据（重新填充）:")
header = f"{'需求\\产品':<12}"
for p in products_list:
    header += f" {p[:6]:>8}"
print(header)
print("-" * (12 + 9 * len(products_list)))

for n in needs_list:
    row = f"{n:<12}"
    for p in products_list:
        d = need_product[n][p]
        if d["orders"] > 0:
            row += f" {d['revenue']:>7,.0f}"
        else:
            row += f" {'-':>8}"
    print(row)

# 8. 复购行为
print("\n" + "=" * 70)
print("八、复购行为分析")
print("=" * 70)

phone_orders = defaultdict(list)
for o in orders:
    phone_orders[o["phone"]].append(o)

repeat = {p: v for p, v in phone_orders.items() if len(v) > 1}
single = {p: v for p, v in phone_orders.items() if len(v) == 1}

print(f"  总客户: {len(phone_orders)} 人")
print(f"  复购客户: {len(repeat)} 人 ({len(repeat)/len(phone_orders)*100:.0f}%)")
print(f"  单次客户: {len(single)} 人")

for phone, olist in repeat.items():
    name = olist[0]["name"]
    total = sum(o["amount"] for o in olist)
    products = [o["product"] for o in olist]
    interval_days = (datetime.strptime(olist[1]["time"][:10], "%Y-%m-%d") - datetime.strptime(olist[0]["time"][:10], "%Y-%m-%d")).days if len(olist) >= 2 else 0
    
    print(f"\n  {name} ({phone[-4:]}):")
    print(f"    订单: {len(olist)}单, 合计 {total:,.2f}")
    print(f"    产品: {' → '.join(products)}")
    print(f"    间隔: {interval_days}天")
    for o in olist:
        print(f"    [{o['time'][:10]}] {o['product']} x{o['qty']} = {o['amount']:,.2f}")

# 9. 件数分布
print("\n" + "=" * 70)
print("九、购买件数分析")
print("=" * 70)

qty_dist = Counter(o["qty"] for o in orders)
for q, c in sorted(qty_dist.items()):
    rev = sum(o["amount"] for o in orders if o["qty"] == q)
    print(f"  买{q}件: {c}单, 营收 {rev:,.2f}")

# Multi-unit orders detail
multi_unit = [o for o in orders if o["qty"] > 1]
if multi_unit:
    print("\n  多件订单明细:")
    for o in multi_unit:
        print(f"    {o['name']} | {o['product']} x{o['qty']} | {o['amount']:,.2f} | {o['sales']}")

# 10. 综合洞察
print("\n" + "=" * 70)
print("十、综合洞察与建议")
print("=" * 70)

# Revenue concentration
top2_rev = sum(d["revenue"] for _, d in sorted_products[:2])
print(f"\n  1. 产品集中度: 前2产品贡献 {top2_rev/total_revenue*100:.0f}% 营收（益家人+镜脂套组）")
print(f"     → 高客单价产品是营收引擎，应加强高客单产品的话术培训和推送策略")

top1_sales = max(sales_data.items(), key=lambda x: x[1]["revenue"])
print(f"\n  2. 销售人效差距: {top1_sales[0]}营收是其他销售的 {top1_sales[1]['revenue']/min(d['revenue'] for d in sales_data.values()):.0f} 倍")
print(f"     → 需要将高绩效销售的话术和方法论标准化，赋能全团队")

repeat_rate = len(repeat)/len(phone_orders)*100
print(f"\n  3. 复购率: {repeat_rate:.0f}%，处于偏低水平")
print(f"     → 6天窗口太短，但已有2人当天/隔天复购，说明即时追加销售机会存在")

print(f"\n  4. 价格天花板: 最高客单 ¥566.60（镜脂套组），最低 ¥139.00（轻路），跨幅 4.1倍")
print(f"     → 价格带覆盖完整，但中高价带(300-500)仅1单，存在提价空间")

avg_age = sum(ages)/len(ages)
print(f"\n  5. 核心客群: 平均{avg_age:.0f}岁，男女比例 {sum(1 for o in orders if o['gender']=='男')}/{sum(1 for o in orders if o['gender']=='女')}")
print(f"     → 中老年男性为主，产品沟通应侧重功能性、实证性表述")

print("\nDone!")
