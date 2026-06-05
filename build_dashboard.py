#!/usr/bin/env python3
"""三诺&华大私域成交数据可视化看板"""

import openpyxl
import json
import math
from datetime import datetime
from collections import defaultdict, Counter

# ─── 1. 数据加载 ───────────────────────────────────────────
wb = openpyxl.load_workbook(
    '/Users/SaciNa/Documents/臻选&芥子/臻选私域/三诺&华大成交用户清单.xlsx',
    data_only=True
)
ws = wb.active

rows = []
for row in ws.iter_rows(min_row=2, values_only=True):
    if row[0] is None:
        continue
    rows.append({
        'sales': row[0],
        'time_raw': row[1],
        'product': row[2],
        'amount': float(row[3]) if row[3] else 0,
        'name': row[4],
        'phone': str(row[5]) if row[5] else '',
        'qty': str(row[6]) if row[6] else '',
        'gender': row[7],
        'age': int(row[8]) if row[8] else None,
        'demand': row[9],
        'province': row[10],
        'city': row[11],
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

# ─── 2. 产品分类简化 ───────────────────────────────────────
def simplify_product(name):
    name = str(name)
    if '镜脂·轻+镜脂·阻' in name or '管理套组' in name:
        return '镜脂套组（轻+阻）'
    if '镜脂·轻' in name:
        return '镜脂·轻（单品）'
    if '镜脂·阻' in name:
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
age_groups = {'<30': 0, '30-39': 0, '40-49': 0, '50-59': 0, '60+': 0}
for a in ages:
    if a < 30: age_groups['<30'] += 1
    elif a < 40: age_groups['30-39'] += 1
    elif a < 50: age_groups['40-49'] += 1
    elif a < 60: age_groups['50-59'] += 1
    else: age_groups['60+'] += 1

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
top_customers = defaultdict(lambda: {'total': 0, 'orders': 0, 'products': set()})
for r in rows:
    p = r['phone']
    top_customers[p]['total'] += r['amount']
    top_customers[p]['orders'] += 1
    top_customers[p]['products'].add(r['category'])
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

# 产品数据
prod_labels = [p[0] for p in product_sorted]
prod_values = [round(p[1]['revenue'], 2) for p in product_sorted]
prod_colors = ['#FF6B6B','#4ECDC4','#45B7D1','#96CEB4','#FFEAA7','#DFE6E9','#FDCB6E','#E17055','#00B894']

# 需求数据
demand_labels = [d[0] for d in demand_sorted]
demand_values = [d[1] for d in demand_sorted]

# 顾问数据
csl_names = [c[0] for c in counselor_sorted]
csl_revenue = [round(c[1]['revenue'], 2) for c in counselor_sorted]
csl_orders = [c[1]['orders'] for c in counselor_sorted]
csl_customers = [len(c[1]['customers']) for c in counselor_sorted]

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
        <td>{phone[:3]}****{phone[-4:]}</td>
        <td>¥{data['total']:,.0f}</td>
        <td>{data['orders']}单</td>
        <td>{data.get('demand', '')}</td>
        <td>{data.get('province', '')} {data.get('city', '')}</td>
        <td>{'、'.join(list(data['products'])[:3])}</td>
    </tr>"""

html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>三诺&华大私域成交数据看板（截至6/4）</title>
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
    <p>数据截至 2026年6月4日 | 数据源：三诺&华大成交用户清单 | 自动生成</p>
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
        <div class="kpi-sub">新增朱老师</div>
    </div>
</div>

<div class="highlight">
    ⚠️ <strong>新增动态</strong>：6/3-6/4 密集成交，新增大单 糖友专享套组 ¥2,697（黎老师）、镜脂·轻 ¥1,519.80（赵老师）、益净复购 ¥1,009.80（赵老师/彭南其）
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
    <div class="card">
        <div class="card-title">🍩 产品营收结构</div>
        <div class="chart-wrap"><canvas id="productChart"></canvas></div>
    </div>
    <div class="card">
        <div class="card-title">🎯 客户需求分布</div>
        <div class="chart-wrap"><canvas id="demandChart"></canvas></div>
    </div>
</div>

<div class="dashboard">
    <div class="card">
        <div class="card-title">👥 客户画像</div>
        <div class="dashboard" style="grid-template-columns: 1fr 1fr; gap: 12px;">
            <div>
                <div style="font-size:12px;color:#888;margin-bottom:8px;">性别分布</div>
                <div class="chart-wrap"><canvas id="genderChart"></canvas></div>
            </div>
            <div>
                <div style="font-size:12px;color:#888;margin-bottom:8px;">年龄分布</div>
                <div class="chart-wrap"><canvas id="ageChart"></canvas></div>
            </div>
        </div>
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
            <thead><tr><th>#</th><th>姓名</th><th>手机</th><th>累计消费</th><th>订单</th><th>需求</th><th>所在地</th><th>产品</th></tr></thead>
            <tbody>{top_cust_rows}</tbody>
        </table>
    </div>
</div>

<div class="dashboard">
    <div class="card full-width">
        <div class="card-title">📋 顾问业绩明细</div>
        <table>
            <thead><tr><th>顾问</th><th>订单数</th><th>营收</th><th>客单价</th><th>客户数</th><th>占比</th><th>人均产出</th></tr></thead>
            <tbody>
"""

for cname in csl_names:
    s = counselor_stats[cname]
    share = s['revenue'] / total_revenue * 100
    asp = s['revenue'] / s['orders']
    c_count = len(s['customers'])
    html += f"""
            <tr>
                <td><strong>{cname}</strong></td>
                <td>{s['orders']}</td>
                <td>¥{s['revenue']:,.0f}</td>
                <td>¥{asp:,.0f}</td>
                <td>{c_count}</td>
                <td>{share:.1f}%</td>
                <td>¥{s['revenue']/c_count:,.0f}/人</td>
            </tr>"""

html += f"""
            </tbody>
        </table>
    </div>
</div>

<div class="footer">
    三诺&华大私域成交数据看板 | 数据截至 2026年6月4日 | 自动生成于 {datetime.now().strftime('%Y-%m-%d %H:%M')}
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

// 性别
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

// 年龄
new Chart(document.getElementById('ageChart'), {{
    type: 'bar',
    data: {{
        labels: {json.dumps(ag_labels)},
        datasets: [{{ label: '人数', data: {json.dumps(ag_values)}, backgroundColor: '#667eea', borderRadius: 4 }}]
    }},
    options: {{
        responsive: true, maintainAspectRatio: false,
        plugins: {{ legend: {{ display: false }} }},
        scales: {{ y: {{ beginAtZero: true, grid: {{ color: '#f0f0f0' }} }} }}
    }}
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
print("── 需求分布 ──")
for d, c in demand_sorted:
    print(f"  {d}: {c}人")
print()
print("── 高价值客户 Top5 ──")
for i, (phone, data) in enumerate(top_sorted[:5], 1):
    print(f"  {i}. {data['name']}: ¥{data['total']:,.0f} ({data['orders']}单)")
