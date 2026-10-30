#!/usr/bin/env python3
"""三诺&华大私域成交数据可视化看板"""

import openpyxl
import json
import math
import os
from datetime import datetime
from collections import defaultdict, Counter

# ─── 1. 数据加载 ───────────────────────────────────────────
wb = openpyxl.load_workbook(
    '/Users/SaciNa/Documents/臻选&芥子/臻选私域/三诺&华大成交用户清单.xlsx',
    data_only=True
)
ws = wb.active

# ─── 1.0 表头驱动的列映射 ───────────────────────────────────
# 源表列顺序曾变动（新增过「收件人」列），硬编码 row[n] 会整体错位，
# 导致年龄图取到「件数」、需求图取到「年龄」、省份图取到「需求」。
# 改为按表头名称定位列，源表再插列也不会错位。
HEADER_ALIASES = {
    'sales':    ['销售', '销售顾问', '顾问'],
    'time_raw': ['下单时间', '下单日期', '时间'],
    'product':  ['产品名称', '商品名称', '产品'],
    'amount':   ['产品价格', '金额', '实付金额'],
    'wechat':   ['微信', '微信昵称', '微信名'],
    'recipient': ['收件人', '收货人', '姓名'],
    'phone':    ['手机', '手机号', '联系电话', '电话'],
    'qty':      ['件数', '数量'],
    'gender':   ['性别'],
    'age':      ['年龄'],
    'demand':   ['需求', '需求类型', '健康需求'],
    'province': ['省份', '省'],
    'city':     ['城市', '市'],
}
# 兜底：旧版 12 列布局（无「收件人」）
LEGACY_IDX = {
    'sales': 0, 'time_raw': 1, 'product': 2, 'amount': 3, 'wechat': 4,
    'phone': 5, 'qty': 6, 'gender': 7, 'age': 8, 'demand': 9,
    'province': 10, 'city': 11,
}

header = [str(c).strip() if c is not None else '' for c in
          next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
COL = {}
for key, aliases in HEADER_ALIASES.items():
    for a in aliases:
        if a in header:
            COL[key] = header.index(a)
            break
    else:
        COL[key] = LEGACY_IDX.get(key)


def cell(row, key):
    i = COL.get(key)
    if i is None or i >= len(row):
        return None
    return row[i]


def clean_str(v):
    """去首尾空白（含全角空格）与内部连续空白"""
    if v is None:
        return ''
    return ' '.join(str(v).replace('　', ' ').split())


PROV_KEYWORDS = ('北京', '天津', '上海', '重庆', '河北', '山西', '辽宁', '吉林',
                 '黑龙江', '江苏', '浙江', '安徽', '福建', '江西', '山东', '河南',
                 '湖北', '湖南', '广东', '广西', '海南', '四川', '贵州', '云南',
                 '西藏', '陕西', '甘肃', '青海', '宁夏', '新疆', '内蒙古',
                 '台湾', '香港', '澳门')


def looks_like_region(s):
    return any(p in s for p in PROV_KEYWORDS)


def parse_age(v):
    """年龄必须是 1~120 的整数，否则视为缺失"""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        n = int(v)
    else:
        s = str(v).strip()
        if not s or not s.lstrip('-').isdigit():
            return None
        n = int(s)
    return n if 1 <= n <= 120 else None


def infer_demand(demand):
    """需求列被误填为省市时返回空串，交给后处理按产品主流需求推断"""
    d = clean_str(demand)
    if d == '胃肠':          # 与「肠胃」为同一需求口径，统一
        d = '肠胃'
    return '' if looks_like_region(d) else d


rows = []
for row in ws.iter_rows(min_row=2, values_only=True):
    # 跳过完全空行（产品名和金额都为空），允许顾问为空的行
    if cell(row, 'product') is None and cell(row, 'amount') is None:
        continue

    product = clean_str(cell(row, 'product'))
    amount = cell(row, 'amount')
    try:
        amount = float(amount) if amount not in (None, '') else 0.0
    except (TypeError, ValueError):
        amount = 0.0

    gender = clean_str(cell(row, 'gender'))
    if gender not in ('男', '女'):
        gender = ''

    wechat = clean_str(cell(row, 'wechat'))
    recipient = clean_str(cell(row, 'recipient'))

    rows.append({
        'sales': clean_str(cell(row, 'sales')) or '未知',
        'time_raw': cell(row, 'time_raw'),
        'product': product,
        'amount': amount,
        'name': recipient or wechat,          # 展示名：收件人姓名优先（与 10/07 基线一致），缺失则用微信昵称
        'wechat': wechat,
        'recipient': recipient,
        'phone': clean_str(cell(row, 'phone')),
        'qty': str(cell(row, 'qty')) if cell(row, 'qty') else '',
        'gender': gender,
        'age': parse_age(cell(row, 'age')),
        'demand': infer_demand(cell(row, 'demand')),
        'province': clean_str(cell(row, 'province')),
        'city': clean_str(cell(row, 'city')),
    })

# Parse time
for r in rows:
    raw = r['time_raw']
    if isinstance(raw, datetime):
        r['date'] = raw.date()
        r['datetime'] = raw
    else:
        # Try "下单时间：YYYY-MM-DD HH:MM:SS"
        s = str(raw).replace('下单时间：', '').strip()
        try:
            dt = datetime.strptime(s, '%Y-%m-%d %H:%M:%S')
            r['date'] = dt.date()
            r['datetime'] = dt
        except:
            r['date'] = 'unknown'

# ─── 1.5 客户基盘数据 ─────────────────────────────────────
# 源文件位于 iCloud Drive，离线/未下载时仅为占位文件（openpyxl 报
# "Need authenticator"）。加载失败时回退到上次成功缓存，保证看板核心
# 成交指标仍可正常生成；文件恢复后下次成功运行会自动刷新缓存。
BASE_XLSX = '/Users/SaciNa/Documents/臻选&芥子/臻选私域/客户统计-按员工查看.xlsx'
BASE_CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'base_data_cache.json')


def load_base_data(path):
    wb2 = openpyxl.load_workbook(path, data_only=True)
    ws2 = wb2.active
    bd = {}
    for row in ws2.iter_rows(min_row=6, max_row=10, values_only=True):
        name_raw = str(row[0]).strip()
        bd[name_raw] = {
            'base': int(row[1]) if row[1] else 0,
            'new': int(row[2]) if row[2] else 0,
            'net': int(row[3]) if row[3] else 0,
            'churn': int(row[4]) if row[4] else 0,
        }
    return bd


base_data = {}
try:
    base_data = load_base_data(BASE_XLSX)
    with open(BASE_CACHE, 'w', encoding='utf-8') as f:
        json.dump(base_data, f, ensure_ascii=False)
except Exception as e:
    print(f"[WARN] 客户基盘文件加载失败（{e}），尝试使用本地缓存")
    if os.path.exists(BASE_CACHE):
        with open(BASE_CACHE, 'r', encoding='utf-8') as f:
            base_data = json.load(f)
        print("[INFO] 已使用上次成功缓存的客户基盘数据")
    else:
        print("[WARN] 无本地缓存，客户基盘数据将为空")
total_base = sum(v['base'] for v in base_data.values())
total_churn_all = sum(v['churn'] for v in base_data.values())

# ─── 2. 产品分类简化 ───────────────────────────────────────
def simplify_product(name):
    name = str(name)
    if '镜脂·轻+镜脂·阻' in name or '管理套组' in name:
        return '镜脂套组（轻+阻）'
    # 合并订单：轻路+镜脂阻
    if ('轻路' in name and ('镜脂' in name or '镜脂阻' in name)):
        return '轻路+镜脂阻（合并）'
    if '镜脂·轻' in name:
        return '镜脂·轻（单品）'
    if '镜脂·阻' in name or '镜脂*阻' in name:
        return '镜脂·阻（单品）'
    if '益衡' in name or '轻路' not in name and '益衡' in name:
        return '益衡+轻路套组'
    if '轻路' in name:
        return '轻路'
    if '益净' in name:
        return '益净'
    if '益休' in name:
        return '益休'
    if '益畅' in name:
        return '益畅'
    if '畅青' in name:
        return '畅青'
    if '益家人' in name:
        return '益家人'
    return '其他'

for r in rows:
    r['category'] = simplify_product(r['product'])

# 需求列异常行（误填省市）：按同类产品的需求众数推断
_prod_demand = defaultdict(Counter)
for r in rows:
    if r['demand']:
        _prod_demand[r['category']][r['demand']] += 1
for r in rows:
    if not r['demand']:
        cands = _prod_demand.get(r['category'])
        r['demand'] = cands.most_common(1)[0][0] if cands else '未标注'

# ─── 3. 统计计算 ───────────────────────────────────────────
total_orders = len(rows)
total_revenue = sum(r['amount'] for r in rows)
avg_order = total_revenue / total_orders

# Unique customers (by phone)
phone_names = defaultdict(set)
for r in rows:
    phone_names[r['phone']].add(r['name'])
unique_phones = len(phone_names)

# 复购客户
repeat_phones = [p for p, names in phone_names.items() if len([r for r in rows if r['phone'] == p]) > 1]
repeat_customers = len(repeat_phones)

# 顾问业绩
counselor_stats = defaultdict(lambda: {'orders': 0, 'revenue': 0, 'customers': set()})
for r in rows:
    c = r['sales']
    counselor_stats[c]['orders'] += 1
    counselor_stats[c]['revenue'] += r['amount']
    counselor_stats[c]['customers'].add(r['phone'])

# 日趋势
daily = defaultdict(lambda: {'orders': 0, 'revenue': 0})
for r in rows:
    d = str(r['date'])
    daily[d]['orders'] += 1
    daily[d]['revenue'] += r['amount']

# 月度趋势
monthly = defaultdict(lambda: {'orders': 0, 'revenue': 0, 'customers': set()})
for r in rows:
    d = r['date']
    if d == 'unknown':
        continue
    m = str(d)[:7]  # YYYY-MM
    monthly[m]['orders'] += 1
    monthly[m]['revenue'] += r['amount']
    monthly[m]['customers'].add(r['phone'])

# 产品结构
product_stats = defaultdict(lambda: {'orders': 0, 'revenue': 0})
for r in rows:
    cat = r['category']
    product_stats[cat]['orders'] += 1
    product_stats[cat]['revenue'] += r['amount']

# 需求分布
demand_stats = Counter(r['demand'] for r in rows if r['demand'])

# 性别分布
gender_stats = Counter(r['gender'] for r in rows if r['gender'])

# 年龄分布
ages = [r['age'] for r in rows if r['age'] is not None]
age_known = len(ages)
age_missing = len(rows) - age_known
age_groups = {'<30': 0, '30-39': 0, '40-49': 0, '50-59': 0, '60+': 0}
for a in ages:
    if a < 30: age_groups['<30'] += 1
    elif a < 40: age_groups['30-39'] += 1
    elif a < 50: age_groups['40-49'] += 1
    elif a < 60: age_groups['50-59'] += 1
    else: age_groups['60+'] += 1

# 性别×年龄交叉分析
gender_age = defaultdict(lambda: {'<30': 0, '30-39': 0, '40-49': 0, '50-59': 0, '60+': 0})
for r in rows:
    g = r['gender']
    a = r['age']
    if g and a is not None:
        if a < 30: gender_age[g]['<30'] += 1
        elif a < 40: gender_age[g]['30-39'] += 1
        elif a < 50: gender_age[g]['40-49'] += 1
        elif a < 60: gender_age[g]['50-59'] += 1
        else: gender_age[g]['60+'] += 1

# 地域分布 Top — 省份+城市明细
province_raw = Counter(r['province'] for r in rows if r['province'])
# 省→城市→单数
province_cities = defaultdict(lambda: Counter())
for r in rows:
    p = r['province']
    c = r['city']
    if p and c:
        province_cities[p][c] += 1

province_stats = province_raw

# 时段分布
hour_stats = Counter()
for r in rows:
    if isinstance(r.get('datetime'), datetime):
        hour_stats[r['datetime'].hour] += 1

# 高价值Top
top_customers = defaultdict(lambda: {'total': 0, 'orders': 0, 'products': set(), 'sales': set()})
for r in rows:
    p = r['phone']
    top_customers[p]['total'] += r['amount']
    top_customers[p]['orders'] += 1
    top_customers[p]['products'].add(r['category'])
    top_customers[p]['sales'].add(r['sales'])
    if 'name' not in top_customers[p]:
        top_customers[p]['name'] = r['name']
        top_customers[p]['demand'] = r['demand'] or ''
        top_customers[p]['province'] = r['province'] or ''
        top_customers[p]['city'] = r['city'] or ''

top_sorted = sorted(top_customers.items(), key=lambda x: x[1]['total'], reverse=True)

# Quarter-hour trend
qtr = defaultdict(int)
for r in rows:
    if isinstance(r.get('datetime'), datetime):
        h = r['datetime'].hour
        m = r['datetime'].minute
        q = h * 4 + m // 15
        qtr[q] += 1

# ─── 4. 输出 HTML ──────────────────────────────────────────
daily_sorted = sorted(daily.items())
counselor_sorted = sorted(counselor_stats.items(), key=lambda x: x[1]['revenue'], reverse=True)
product_sorted = sorted(product_stats.items(), key=lambda x: x[1]['revenue'], reverse=True)
demand_sorted = sorted(demand_stats.items(), key=lambda x: x[1], reverse=True)
province_top = province_stats.most_common(10)

# 日期标签
date_labels_json = json.dumps([d[0] for d in daily_sorted])
date_revenue_json = json.dumps([round(d[1]['revenue'], 2) for d in daily_sorted])
date_orders_json = json.dumps([d[1]['orders'] for d in daily_sorted])

# 月度标签与数据
monthly_sorted = sorted(monthly.items())
monthly_labels = [m[0] for m in monthly_sorted]
monthly_labels_json = json.dumps(monthly_labels)
monthly_revenue_json = json.dumps([round(m[1]['revenue'], 2) for m in monthly_sorted])
monthly_orders_json = json.dumps([m[1]['orders'] for m in monthly_sorted])
monthly_customers_json = json.dumps([len(m[1]['customers']) for m in monthly_sorted])

# 月度明细表
monthly_table_rows = ""
for m, s in monthly_sorted:
    avg = s['revenue'] / s['orders'] if s['orders'] else 0
    monthly_table_rows += f"""
    <tr>
        <td>{m}</td>
        <td>{s['orders']}</td>
        <td>¥{s['revenue']:,.0f}</td>
        <td>{len(s['customers'])}</td>
        <td>¥{avg:,.0f}</td>
    </tr>"""

# 产品数据
prod_labels = [p[0] for p in product_sorted]
prod_values = [round(p[1]['revenue'], 2) for p in product_sorted]
prod_colors = ['#FF6B6B','#4ECDC4','#45B7D1','#96CEB4','#FFEAA7','#DFE6E9','#FDCB6E','#E17055','#00B894']

# 产品数量数据（按订单数排序）
product_qty_sorted = sorted(product_stats.items(), key=lambda x: x[1]['orders'], reverse=True)
prod_qty_labels = [p[0] for p in product_qty_sorted]
prod_qty_values = [p[1]['orders'] for p in product_qty_sorted]
total_qty = sum(prod_qty_values)

# 需求数据
demand_labels = [d[0] for d in demand_sorted]
demand_values = [d[1] for d in demand_sorted]

# 顾问数据
csl_names = [c[0] for c in counselor_sorted]
csl_revenue = [round(c[1]['revenue'], 2) for c in counselor_sorted]
csl_orders = [c[1]['orders'] for c in counselor_sorted]
csl_customers = [len(c[1]['customers']) for c in counselor_sorted]

# 顾问人效数据（客户基盘/转化率/客均产出/流失率）
csl_base = [base_data.get(n, {}).get('base', 0) for n in csl_names]
csl_churn = [base_data.get(n, {}).get('churn', 0) for n in csl_names]
csl_conv_rate = [round(len(counselor_stats[n]['customers']) / b * 100, 1) if b else 0 for n, b in zip(csl_names, csl_base)]
csl_rev_per_cust = [round(counselor_stats[n]['revenue'] / b, 1) if b else 0 for n, b in zip(csl_names, csl_base)]
csl_churn_rate = [round(c / (b + c) * 100, 1) if (b + c) else 0 for n, b, c in zip(csl_names, csl_base, csl_churn)]
overall_conv = round(unique_phones / total_base * 100, 1) if total_base else 0
overall_rev_per_cust = round(total_revenue / total_base, 1) if total_base else 0
overall_churn_rate = round(total_churn_all / (total_base + total_churn_all) * 100, 1) if (total_base + total_churn_all) else 0

# 性别
g_labels = [g[0] for g in gender_stats.most_common()]
g_values = [g[1] for g in gender_stats.most_common()]

# 年龄
ag_labels = list(age_groups.keys())
ag_values = list(age_groups.values())

# 省份 — 纯省名，确保完整显示
prov_top_raw = province_stats.most_common(10)
prov_labels = [p[0] for p in prov_top_raw]
prov_values = [p[1] for p in prov_top_raw]

# 城市分布
city_stats = Counter(r['city'] for r in rows if r['city'])
city_top = city_stats.most_common(10)
city_labels = [c[0] for c in city_top]
city_values = [c[1] for c in city_top]

# 时段
all_hours = list(range(24))
hour_labels = [f"{h}:00" for h in all_hours]
hour_values = [hour_stats.get(h, 0) for h in all_hours]

# Top 客户
top_cust_rows = ""
for i, (phone, data) in enumerate(top_sorted[:10], 1):
    name = data.get('name', phone)
    top_cust_rows += f"""
    <tr>
        <td>{i}</td>
        <td>{name}</td>
        <td>{'、'.join(sorted(data.get('sales', [])))}</td>
        <td>{phone[:3]}****{phone[-4:]}</td>
        <td>¥{data['total']:,.0f}</td>
        <td>{data['orders']}单</td>
        <td>{data.get('demand', '')}</td>
        <td>{data.get('province', '')} {data.get('city', '')}</td>
        <td>{'、'.join(list(data['products'])[:3])}</td>
    </tr>"""

latest_date = max(r['date'] for r in rows if r.get('date') != 'unknown')
latest_date_str = str(latest_date).replace('2026-0', '').replace('2026-', '')
# 最新动态摘要
latest_orders = sorted([r for r in rows if r.get('date') != 'unknown'], key=lambda r: r['date'], reverse=True)[:3]
latest_summary = []
for r in latest_orders:
    latest_summary.append(f"{r['date']} {r['sales']} {r['category']} ¥{r['amount']:,.0f}（{r.get('name','') or '未知客户'}）")

html = f"""<!DOCTYPE html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>三诺&华大私域成交数据看板（截至{latest_date_str}）</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC', 'Microsoft YaHei', sans-serif; background: #f0f2f5; color: #333; padding: 20px; }}
.header {{ background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: #fff; padding: 30px 40px; border-radius: 16px; margin-bottom: 24px; }}
.header h1 {{ font-size: 24px; font-weight: 600; margin-bottom: 8px; }}
.header p {{ opacity: 0.85; font-size: 14px; }}
.kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; margin-bottom: 24px; }}
.kpi-card {{ background: #fff; border-radius: 12px; padding: 20px 24px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }}
.kpi-label {{ font-size: 13px; color: #888; margin-bottom: 6px; }}
.kpi-value {{ font-size: 28px; font-weight: 700; }}
.kpi-sub {{ font-size: 12px; color: #999; margin-top: 4px; }}
.kpi-up {{ color: #27ae60; }}
.kpi-down {{ color: #e74c3c; }}
.dashboard {{ display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 16px; }}
@media (max-width: 960px) {{ .dashboard {{ grid-template-columns: 1fr; }} }}
.card {{ background: #fff; border-radius: 12px; padding: 20px 24px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }}
.card-title {{ font-size: 16px; font-weight: 600; margin-bottom: 16px; color: #333; display: flex; align-items: center; gap: 8px; }}
.chart-wrap {{ position: relative; width: 100%; }}
.chart-wrap canvas {{ max-height: 280px; }}
.chart-note {{ margin-top: 10px; font-size: 12px; color: #888; line-height: 1.5; }}
.full-width {{ grid-column: 1 / -1; }}
table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
th {{ background: #f8f9fa; padding: 10px 12px; text-align: left; font-weight: 600; color: #555; border-bottom: 2px solid #e9ecef; }}
td {{ padding: 10px 12px; border-bottom: 1px solid #f0f0f0; }}
tr:hover td {{ background: #fafbfc; }}
.highlight {{ background: #FFF3CD; font-weight: 600; padding: 10px 16px; border-radius: 8px; border-left: 4px solid #ffc107; margin-bottom: 16px; font-size: 13px; }}
.footer {{ text-align: center; padding: 20px; color: #aaa; font-size: 12px; }}
.badge {{ display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 500; }}
.badge-green {{ background: #d4edda; color: #155724; }}
.badge-red {{ background: #f8d7da; color: #721c24; }}
.badge-blue {{ background: #d1ecf1; color: #0c5460; }}
</style>
</head>
<body>

<div class="header">
    <h1>📊 三诺&华大 私域成交数据看板</h1>
    <p>数据截至 {max(r['date'] for r in rows if r.get('date') != 'unknown')} | 数据源：三诺&华大成交用户清单 | 自动生成</p>
</div>

<div class="kpi-grid">
    <div class="kpi-card">
        <div class="kpi-label">💰 累计营收</div>
        <div class="kpi-value">¥{total_revenue:,.0f}</div>
        <div class="kpi-sub">较上期 ¥6,999 → <span class="kpi-up">+¥{total_revenue - 6999.70:,.0f}</span></div>
    </div>
    <div class="kpi-card">
        <div class="kpi-label">📦 累计订单</div>
        <div class="kpi-value">{total_orders}</div>
        <div class="kpi-sub">较上期 17单 → <span class="kpi-up">+{total_orders - 17}单</span></div>
    </div>
    <div class="kpi-card">
        <div class="kpi-label">👤 成交客户</div>
        <div class="kpi-value">{unique_phones}</div>
        <div class="kpi-sub">复购客户 {repeat_customers}人</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-label">📊 平均客单</div>
        <div class="kpi-value">¥{avg_order:,.0f}</div>
        <div class="kpi-sub">较上期 ¥411 → <span class="kpi-up">+¥{avg_order - 411.75:,.0f}</span></div>
    </div>
    <div class="kpi-card">
        <div class="kpi-label">🔁 复购率</div>
        <div class="kpi-value">{repeat_customers/unique_phones*100:.0f}%</div>
        <div class="kpi-sub">{repeat_customers}/{unique_phones} 客户复购</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-label">🛒 顾问数量</div>
        <div class="kpi-value">{len(counselor_stats)}</div>
        <div class="kpi-sub">最新成交: {csl_names[0] if csl_names else '-'}</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-label">👥 客户基盘</div>
        <div class="kpi-value">{total_base:,}</div>
        <div class="kpi-sub">{len(counselor_stats)}位顾问服务中</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-label">🎯 整体转化率</div>
        <div class="kpi-value">{overall_conv}%</div>
        <div class="kpi-sub">{total_orders}单 / {total_base:,}客户</div>
    </div>
</div>

<div class="highlight">
    ⚠️ <strong>最新动态</strong>：<br>{'<br>'.join(f'• {s}' for s in latest_summary)}
</div>

<div class="dashboard">
    <div class="card">
        <div class="card-title">📈 日成交趋势</div>
        <div class="chart-wrap"><canvas id="dailyChart"></canvas></div>
    </div>
    <div class="card">
        <div class="card-title">🏆 顾问业绩对比</div>
        <div class="chart-wrap"><canvas id="counselorChart"></canvas></div>
    </div>
</div>

<div class="dashboard">
    <div class="card full-width">
        <div class="card-title">📅 月度成交趋势</div>
        <div class="chart-wrap" style="height: 320px;"><canvas id="monthlyChart"></canvas></div>
    </div>
</div>

<div class="dashboard">
    <div class="card full-width">
        <div class="card-title">📅 月度数据明细</div>
        <table>
            <thead><tr><th>月份</th><th>订单数</th><th>营收</th><th>成交客户</th><th>月均客单</th></tr></thead>
            <tbody>{monthly_table_rows}</tbody>
        </table>
    </div>
</div>

<div class="dashboard">
    <div class="card">
        <div class="card-title">🎯 顾问转化率对比</div>
        <div class="chart-wrap"><canvas id="conversionChart"></canvas></div>
    </div>
    <div class="card">
        <div class="card-title">📊 顾问客均产出（¥/客户）</div>
        <div class="chart-wrap"><canvas id="revPerCustChart"></canvas></div>
    </div>
</div>

<div class="dashboard">
    <div class="card">
        <div class="card-title">🍩 产品营收结构</div>
        <div class="chart-wrap"><canvas id="productChart"></canvas></div>
    </div>
    <div class="card">
        <div class="card-title">📦 产品销量结构</div>
        <div class="chart-wrap"><canvas id="productQtyChart"></canvas></div>
    </div>
</div>

<div class="dashboard">
    <div class="card">
        <div class="card-title">🎯 客户需求分布</div>
        <div class="chart-wrap"><canvas id="demandChart"></canvas></div>
    </div>
    <div class="card">
        <div class="card-title">👥 性别×年龄分布（堆叠）</div>
        <div class="chart-wrap"><canvas id="genderAgeChart"></canvas></div>
        <div class="chart-note">基于 {age_known} 单有年龄记录的数据统计（{age_missing} 单未填年龄，不计入分组）</div>
    </div>
</div>

<div class="dashboard">
    <div class="card">
        <div class="card-title">👫 性别占比</div>
        <div class="chart-wrap"><canvas id="genderChart"></canvas></div>
    </div>
    <div class="card">
        <div class="card-title">⏰ 下单时段分布</div>
        <div class="chart-wrap"><canvas id="hourChart"></canvas></div>
    </div>
</div>

<div class="dashboard">
    <div class="card">
        <div class="card-title">📍 省份分布 Top10</div>
        <div class="chart-wrap"><canvas id="provinceChart"></canvas></div>
    </div>
    <div class="card">
        <div class="card-title">🏙️ 城市分布 Top10</div>
        <div class="chart-wrap"><canvas id="cityChart"></canvas></div>
    </div>
</div>

<div class="dashboard">
    <div class="card full-width">
        <div class="card-title">⭐ 高价值客户 Top10</div>
        <table>
            <thead><tr><th>#</th><th>姓名</th><th>服务顾问</th><th>手机</th><th>累计消费</th><th>订单</th><th>需求</th><th>所在地</th><th>产品</th></tr></thead>
            <tbody>{top_cust_rows}</tbody>
        </table>
    </div>
</div>

<div class="dashboard">
    <div class="card full-width">
        <div class="card-title">📋 顾问业绩明细</div>
        <table>
            <thead><tr><th>顾问</th><th>订单数</th><th>营收</th><th>客单价</th><th>成交客户</th><th>客户基盘</th><th>转化率</th><th>客均产出</th><th>流失率</th><th>营收占比</th></tr></thead>
            <tbody>
"""
for i, cname in enumerate(csl_names):
    s = counselor_stats[cname]
    share = s['revenue'] / total_revenue * 100
    asp = s['revenue'] / s['orders']
    c_count = len(s['customers'])
    b = csl_base[i]
    html += f"""
            <tr>
                <td><strong>{cname}</strong></td>
                <td>{s['orders']}</td>
                <td>¥{s['revenue']:,.0f}</td>
                <td>¥{asp:,.0f}</td>
                <td>{c_count}</td>
                <td>{b:,}</td>
                <td><strong>{csl_conv_rate[i]}%</strong></td>
                <td>¥{csl_rev_per_cust[i]}</td>
                <td>{csl_churn_rate[i]}%</td>
                <td>{share:.1f}%</td>
            </tr>"""

html += f"""
            <tr style="background:#f8f9fa; font-weight:600;">
                <td>合计/均值</td>
                <td>{total_orders}</td>
                <td>¥{total_revenue:,.0f}</td>
                <td>¥{avg_order:,.0f}</td>
                <td>{unique_phones}</td>
                <td>{total_base:,}</td>
                <td><strong>{overall_conv}%</strong></td>
                <td>¥{overall_rev_per_cust}</td>
                <td>{overall_churn_rate}%</td>
                <td>100%</td>
            </tr>
            </tbody>
        </table>
    </div>
</div>

<div class="footer">
    三诺&华大私域成交数据看板 | 数据截至 {latest_date} | 自动生成于 {datetime.now().strftime('%Y-%m-%d %H:%M')}
</div>

<script>
const COLORS = ['#667eea','#764ba2','#f093fb','#4facfe','#43e97b','#fa709a','#fee140','#30cfd0','#a8edea','#fed6e3'];

// 日成交趋势
new Chart(document.getElementById('dailyChart'), {{
    type: 'bar',
    data: {{
        labels: {date_labels_json},
        datasets: [
            {{
                label: '营收 (¥)',
                data: {date_revenue_json},
                backgroundColor: '#667eea',
                borderRadius: 6,
                yAxisID: 'y',
                order: 1
            }},
            {{
                label: '订单数',
                data: {date_orders_json},
                type: 'line',
                borderColor: '#f093fb',
                backgroundColor: 'transparent',
                borderWidth: 2,
                pointRadius: 4,
                pointBackgroundColor: '#f093fb',
                tension: 0.3,
                yAxisID: 'y1',
                order: 0
            }}
        ]
    }},
    options: {{
        responsive: true,
        maintainAspectRatio: false,
        interaction: {{ mode: 'index', intersect: false }},
        plugins: {{ legend: {{ position: 'top', labels: {{ usePointStyle: true }} }} }},
        scales: {{
            y: {{ position: 'left', title: {{ display: true, text: '营收 (¥)' }}, grid: {{ color: '#f0f0f0' }} }},
            y1: {{ position: 'right', title: {{ display: true, text: '订单数' }}, grid: {{ display: false }}, min: 0 }}
        }}
    }}
}});

// 顾问业绩
new Chart(document.getElementById('counselorChart'), {{
    type: 'bar',
    data: {{
        labels: {json.dumps(csl_names)},
        datasets: [
            {{ label: '营收 (¥)', data: {json.dumps(csl_revenue)}, backgroundColor: COLORS.slice(0, {len(csl_names)}), borderRadius: 6 }}
        ]
    }},
    options: {{
        responsive: true, maintainAspectRatio: false,
        plugins: {{ legend: {{ display: false }} }},
        scales: {{ y: {{ grid: {{ color: '#f0f0f0' }} }} }}
    }}
}});

// 顾问转化率对比（真实数据：成交客户/客户基盘）
new Chart(document.getElementById('conversionChart'), {{
    type: 'bar',
    data: {{
        labels: {json.dumps(csl_names)},
        datasets: [{{
            label: '转化率 (%)',
            data: {json.dumps(csl_conv_rate)},
            backgroundColor: COLORS.slice(0, {len(csl_names)}),
            borderRadius: 6,
        }}]
    }},
    options: {{
        responsive: true, maintainAspectRatio: false,
        plugins: {{
            legend: {{ display: false }},
            tooltip: {{ callbacks: {{ label: function(ctx) {{ const ci = ctx.dataIndex; return ctx.raw.toFixed(1) + '% (' + {json.dumps(csl_customers)}[ci] + '成交客户/' + {json.dumps(csl_base)}[ci] + '客户基盘)'; }} }} }}
        }},
        scales: {{
            y: {{ beginAtZero: true, grid: {{ color: '#f0f0f0' }}, title: {{ display: true, text: '转化率 (%)' }} }}
        }}
    }}
}});

// 客均产出（真实数据：营收/客户基盘）
new Chart(document.getElementById('revPerCustChart'), {{
    type: 'bar',
    data: {{
        labels: {json.dumps(csl_names)},
        datasets: [{{
            label: '客均产出 (¥)',
            data: {json.dumps(csl_rev_per_cust)},
            backgroundColor: COLORS.slice(0, {len(csl_names)}),
            borderRadius: 6,
        }}]
    }},
    options: {{
        responsive: true, maintainAspectRatio: false,
        plugins: {{ legend: {{ display: false }} }},
        scales: {{ y: {{ beginAtZero: true, grid: {{ color: '#f0f0f0' }}, title: {{ display: true, text: '营收/客户 (¥)' }} }} }}
    }}
}});

// 产品营收结构
new Chart(document.getElementById('productChart'), {{
    type: 'doughnut',
    data: {{
        labels: {json.dumps(prod_labels)},
        datasets: [{{ data: {json.dumps(prod_values)}, backgroundColor: {json.dumps(prod_colors[:len(prod_labels)])}, borderWidth: 2, borderColor: '#fff' }}]
    }},
    options: {{
        responsive: true, maintainAspectRatio: false,
        plugins: {{
            legend: {{ position: 'right', labels: {{ padding: 12, usePointStyle: true, font: {{ size: 11 }} }} }},
            tooltip: {{ callbacks: {{ label: function(ctx) {{ return ctx.label + ': ¥' + ctx.raw.toLocaleString() + ' (' + (ctx.raw / {total_revenue} * 100).toFixed(1) + '%)'; }} }} }}
        }}
    }}
}});

// 产品销量结构
new Chart(document.getElementById('productQtyChart'), {{
    type: 'doughnut',
    data: {{
        labels: {json.dumps(prod_qty_labels)},
        datasets: [{{ data: {json.dumps(prod_qty_values)}, backgroundColor: {json.dumps(prod_colors[:len(prod_qty_labels)])}, borderWidth: 2, borderColor: '#fff' }}]
    }},
    options: {{
        responsive: true, maintainAspectRatio: false,
        plugins: {{
            legend: {{ position: 'right', labels: {{ padding: 12, usePointStyle: true, font: {{ size: 11 }} }} }},
            tooltip: {{ callbacks: {{ label: function(ctx) {{ return ctx.label + ': ' + ctx.raw + '单 (' + (ctx.raw / {total_qty} * 100).toFixed(1) + '%)'; }} }} }}
        }}
    }}
}});

// 需求分布
new Chart(document.getElementById('demandChart'), {{
    type: 'polarArea',
    data: {{
        labels: {json.dumps(demand_labels)},
        datasets: [{{ data: {json.dumps(demand_values)}, backgroundColor: ['#667eea','#764ba2','#f093fb','#4facfe','#ff6b6b','#2ecc71','#f39c12'], borderWidth: 2, borderColor: '#fff' }}]
    }},
    options: {{
        responsive: true, maintainAspectRatio: false,
        plugins: {{ legend: {{ position: 'bottom', labels: {{ padding: 10, usePointStyle: true }} }} }}
    }}
}});

// 性别占比
new Chart(document.getElementById('genderChart'), {{
    type: 'doughnut',
    data: {{
        labels: {json.dumps(g_labels)},
        datasets: [{{ data: {json.dumps(g_values)}, backgroundColor: ['#4facfe','#f093fb'], borderWidth: 3, borderColor: '#fff' }}]
    }},
    options: {{
        responsive: true, maintainAspectRatio: false,
        plugins: {{ legend: {{ position: 'bottom' }} }}
    }}
}});

// 性别×年龄堆叠柱状图（含总体）
const gaMale = {json.dumps([gender_age.get('男', {}).get(k, 0) for k in ag_labels])};
const gaFemale = {json.dumps([gender_age.get('女', {}).get(k, 0) for k in ag_labels])};
const gaTotal = gaMale.map((v,i) => v + gaFemale[i]);
new Chart(document.getElementById('genderAgeChart'), {{
    type: 'bar',
    data: {{
        labels: {json.dumps(ag_labels)},
        datasets: [
            {{ label: '👨 男', data: gaMale, backgroundColor: '#4facfe' }},
            {{ label: '👩 女', data: gaFemale, backgroundColor: '#f093fb' }}
        ]
    }},
    options: {{
        responsive: true, maintainAspectRatio: false,
        plugins: {{
            legend: {{ position: 'top', labels: {{ usePointStyle: true }} }},
            tooltip: {{
                callbacks: {{
                    footer: function(items) {{
                        const idx = items[0].dataIndex;
                        return '合计: ' + gaTotal[idx] + '人';
                    }}
                }}
            }}
        }},
        scales: {{
            x: {{ stacked: true, grid: {{ display: false }} }},
            y: {{ stacked: true, beginAtZero: true, grid: {{ color: '#f0f0f0' }}, title: {{ display: true, text: '人数' }} }}
        }}
    }},
    plugins: [{{
        id: 'totalLabels',
        afterDatasetsDraw(chart) {{
            const ctx = chart.ctx;
            const meta = chart.getDatasetMeta(0);
            ctx.font = 'bold 13px -apple-system, sans-serif';
            ctx.fillStyle = '#333';
            ctx.textAlign = 'center';
            meta.data.forEach((bar, i) => {{
                const t = gaTotal[i];
                if (t > 0) {{
                    ctx.fillText(t + '人', bar.x, bar.y - 8);
                }}
            }});
        }}
    }}]
}});

// 时段
new Chart(document.getElementById('hourChart'), {{
    type: 'line',
    data: {{
        labels: {json.dumps(hour_labels)},
        datasets: [{{ label: '下单数', data: {json.dumps(hour_values)}, borderColor: '#667eea', backgroundColor: 'rgba(102,126,234,0.1)', fill: true, tension: 0.4, pointRadius: 3 }}]
    }},
    options: {{
        responsive: true, maintainAspectRatio: false,
        plugins: {{ legend: {{ display: false }} }},
        scales: {{ y: {{ beginAtZero: true, grid: {{ color: '#f0f0f0' }} }} }}
    }}
}});

// 省份
new Chart(document.getElementById('provinceChart'), {{
    type: 'bar',
    data: {{
        labels: {json.dumps(prov_labels)},
        datasets: [{{ label: '订单数', data: {json.dumps(prov_values)}, backgroundColor: '#764ba2', borderRadius: 4 }}]
    }},
    options: {{
        indexAxis: 'y',
        responsive: true, maintainAspectRatio: false,
        layout: {{ padding: {{ left: 10 }} }},
        plugins: {{ legend: {{ display: false }} }},
        scales: {{
            x: {{ beginAtZero: true, grid: {{ color: '#f0f0f0' }} }},
            y: {{
                ticks: {{ font: {{ size: 12 }}, autoSkip: false }}
            }}
        }}
    }}
}});

// 城市
new Chart(document.getElementById('cityChart'), {{
    type: 'bar',
    data: {{
        labels: {json.dumps(city_labels)},
        datasets: [{{ label: '订单数', data: {json.dumps(city_values)}, backgroundColor: '#4facfe', borderRadius: 4 }}]
    }},
    options: {{
        indexAxis: 'y',
        responsive: true, maintainAspectRatio: false,
        layout: {{ padding: {{ left: 10 }} }},
        plugins: {{ legend: {{ display: false }} }},
        scales: {{
            x: {{ beginAtZero: true, grid: {{ color: '#f0f0f0' }} }},
            y: {{
                ticks: {{ font: {{ size: 12 }}, autoSkip: false }}
            }}
        }}
    }}
}});

// 月度成交趋势
new Chart(document.getElementById('monthlyChart'), {{
    type: 'bar',
    data: {{
        labels: {monthly_labels_json},
        datasets: [
            {{
                label: '营收 (¥)',
                data: {monthly_revenue_json},
                backgroundColor: '#667eea',
                borderRadius: 6,
                yAxisID: 'y',
                order: 1
            }},
            {{
                label: '订单数',
                data: {monthly_orders_json},
                type: 'line',
                borderColor: '#f093fb',
                backgroundColor: 'transparent',
                borderWidth: 2,
                pointRadius: 5,
                pointBackgroundColor: '#f093fb',
                tension: 0.3,
                yAxisID: 'y1',
                order: 0
            }}
        ]
    }},
    options: {{
        responsive: true,
        maintainAspectRatio: false,
        interaction: {{ mode: 'index', intersect: false }},
        plugins: {{
            legend: {{ position: 'top', labels: {{ usePointStyle: true }} }},
            tooltip: {{
                callbacks: {{
                    afterBody: function(items) {{
                        const i = items[0].dataIndex;
                        return '成交客户: ' + {monthly_customers_json}[i] + '人';
                    }}
                }}
            }}
        }},
        scales: {{
            y: {{ position: 'left', title: {{ display: true, text: '营收 (¥)' }}, grid: {{ color: '#f0f0f0' }} }},
            y1: {{ position: 'right', title: {{ display: true, text: '订单数' }}, grid: {{ display: false }}, min: 0 }}
        }}
    }}
}});
</script>
</body>
</html>
"""

output_path = '/Users/SaciNa/WorkBuddy/2026-05-26-15-30-31/三诺华大数据看板.html'
with open(output_path, 'w', encoding='utf-8') as f:
    f.write(html)

print(f"✅ 看板已生成: {output_path}")
print(f"\n═══ 数据摘要 ═══")
print(f"订单总数: {total_orders}")
print(f"总营收: ¥{total_revenue:,.2f}")
print(f"成交客户: {unique_phones}人 (复购{repeat_customers}人)")
print(f"平均客单: ¥{avg_order:,.0f}")
print(f"顾问数: {len(counselor_stats)}人")
print()
print("── 顾问业绩 ──")
for cname in csl_names:
    s = counselor_stats[cname]
    print(f"  {cname}: {s['orders']}单 ¥{s['revenue']:,.0f} ({s['revenue']/total_revenue*100:.1f}%)")
print()
print("── 产品营收 Top5 ──")
for i, (cat, s) in enumerate(product_sorted[:5], 1):
    print(f"  {i}. {cat}: {s['orders']}单 ¥{s['revenue']:,.0f}")
print()
print("── 日成交 ──")
for d in daily_sorted:
    print(f"  {d[0]}: {d[1]['orders']}单 ¥{d[1]['revenue']:,.0f}")
print()
print("── 月度成交 ──")
for m, s in monthly_sorted:
    avg = s['revenue'] / s['orders'] if s['orders'] else 0
    print(f"  {m}: {s['orders']}单 ¥{s['revenue']:,.0f} 客户{len(s['customers'])}人 客单¥{avg:,.0f}")
print()
print("── 需求分布 ──")
for d, c in demand_sorted:
    print(f"  {d}: {c}人")
print()
print("── 高价值客户 Top5 ──")
for i, (phone, data) in enumerate(top_sorted[:5], 1):
    print(f"  {i}. {data['name']}: ¥{data['total']:,.0f} ({data['orders']}单)")
