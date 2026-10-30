# -*- coding: utf-8 -*-
"""三诺华大 月度经营报表 v3
组织口径变更（2026-09 起）：
  · 黎老师 9月起停止服务 → 9月数据不计入团队与顾问表（5-8月历史贡献保留）
  · 王老师账号 9月起由赵老师一并接待 → 9月数据并入赵老师
输出：3 张核心表 + 2 张客单价表，共 5 个 sheet
"""
import openpyxl, json, shutil, os
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
from collections import defaultdict

BASE = '/Users/SaciNa/WorkBuddy/2026-05-26-15-30-31'
XLSX = f'{BASE}/三诺华大月度转化率与单客价值_2026年5-9月.xlsx'
ORDER_SRC = '/Users/SaciNa/Documents/臻选&芥子/臻选私域/三诺&华大成交用户清单.xlsx'
GROWTH = '/Users/SaciNa/WorkBuddy/2026-09-17-11-36-56/outputs/wbz_customer_growth_raw.json'

MONTHS = [5, 6, 7, 8, 9]
NAME_MAP = {"赵老师": "赵老师", "三诺健康顾问丨朱国银": "朱老师", "三诺健康管家 | 诺诺": "王老师",
            "三诺健康管家｜吴老师": "吴老师", "三诺健康管家 | 黎老师": "黎老师"}
ADV_ORDER = ['赵老师', '朱老师', '王老师', '吴老师', '黎老师']

# ── 备份 ──
os.makedirs(f'{BASE}/outputs', exist_ok=True)
shutil.copy(XLSX, f'{BASE}/outputs/三诺华大月度转化率与单客价值_备份_{datetime.now():%Y%m%d_%H%M}.xlsx')

# ══════ 数据源 1：客户增长（微伴）══════
growth = json.load(open(GROWTH, encoding='utf-8'))
G = defaultdict(lambda: defaultdict(dict))       # G[month][顾问] = {'add':,'lost':}
for r in growth:
    sh = NAME_MAP.get(r['staff_name'])
    p = r.get('_period', '')
    if sh and p.startswith('2026-'):
        G[int(p[5:7])][sh] = {'add': r.get('increase_count', 0), 'lost': r.get('decrease_count', 0)}

# ══════ 数据源 2：成交清单 ══════
wb_src = openpyxl.load_workbook(ORDER_SRC, data_only=True)
ws_src = wb_src.active
# 按表头名定位列（源表曾插入「微信」列，硬编码 row[n] 会整体错位）
_hdr = [str(c).strip() if c is not None else '' for c in
        next(ws_src.iter_rows(min_row=1, max_row=1, values_only=True))]
IDX_PHONE = _hdr.index('手机') if '手机' in _hdr else 5
orders = []
for row in ws_src.iter_rows(min_row=2, values_only=True):
    if row[2] is None and row[3] is None:
        continue
    raw = row[1]
    if isinstance(raw, datetime):
        d = raw.date()
    else:
        s = str(raw).replace('下单时间：', '').strip()
        d = None
        for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%d'):
            try:
                d = datetime.strptime(s, fmt).date()
                break
            except Exception:
                pass
    orders.append({'sales': str(row[0] or ''), 'date': d, 'product': str(row[2] or ''),
                   'amt': float(row[3] or 0), 'phone': str(row[IDX_PHONE] or '')})
for r in orders:
    r['m'] = r['date'].month if (r['date'] and r['date'].year == 2026) else None

first_month = {}
for r in sorted([x for x in orders if x['m']], key=lambda x: x['date']):
    if r['phone'] and r['phone'] not in first_month:
        first_month[r['phone']] = r['m']

def msub(m):
    return [r for r in orders if r['m'] == m]

def is_grp(p):
    return ('套组' in p) or ('+' in p) or ('合并' in p)

# ══════ 组织口径：9月起归属调整 ══════
# 9月起：王 → 赵；黎 停服
def eff_adv(adv, m):
    """该顾问当月数据归属到谁（None = 不计入）"""
    if m >= 9:
        if adv == '王老师':
            return '赵老师'
        if adv == '黎老师':
            return None
    return adv

# ── 月度团队（9月：赵+朱+吴+王，黎剔除）──
monthly = []
cum = 0
for m in MONTHS:
    add = lost = 0
    for adv in ADV_ORDER:
        if eff_adv(adv, m) is None:
            continue
        add += G[m][adv]['add']
        lost += G[m][adv]['lost']
    cum += add - lost
    sub = msub(m)
    rev = sum(r['amt'] for r in sub)
    n = len(sub)
    cust = len(set(r['phone'] for r in sub))
    grp = [r for r in sub if is_grp(r['product'])]
    single = [r for r in sub if not is_grp(r['product'])]
    hi = [r for r in sub if r['amt'] >= 1000]
    amts = sorted(r['amt'] for r in sub)
    newcust = len(set(r['phone'] for r in sub
                     if first_month.get(r['phone']) == m))
    monthly.append(dict(
        m=m, add=add, lost=lost, net=add - lost, cum=cum,
        n=n, cust=cust, rev=rev,
        aov=rev / n if n else 0, aov_c=rev / cust if cust else 0,
        arpu=rev / cum if cum else 0,
        conv_new=newcust / add if add else 0,
        conv_base=cust / cum if cum else 0,
        newcust=newcust, repcust=cust - newcust,
        grp_n=len(grp), grp_rev=sum(r['amt'] for r in grp),
        grp_aov=(sum(r['amt'] for r in grp) / len(grp)) if grp else 0,
        sg_n=len(single), sg_rev=sum(r['amt'] for r in single),
        sg_aov=(sum(r['amt'] for r in single) / len(single)) if single else 0,
        hi_n=len(hi), hi_rev=sum(r['amt'] for r in hi),
        median=amts[len(amts) // 2] if amts else 0,
        maxv=amts[-1] if amts else 0))

# ── 顾问维度（9月起王并入赵）──
adv_rows = []
SPEC = [
    ('赵老师', '赵老师（含王老师号·9月起）', True),
    ('王老师', '王老师（5-8月，9月起并入赵）', False),
    ('朱老师', '朱老师', True),
    ('吴老师', '吴老师', True),
    ('黎老师', '黎老师（5-8月，9月起停服）', False),
]
for key, label, take9 in SPEC:
    a58 = sum(G[m][key]['add'] for m in (5, 6, 7, 8))
    l58 = sum(G[m][key]['lost'] for m in (5, 6, 7, 8))
    if take9:
        a9 = sum(G[9][x]['add'] for x in (['赵老师', '王老师'] if key == '赵老师' else [key]))
        l9 = sum(G[9][x]['lost'] for x in (['赵老师', '王老师'] if key == '赵老师' else [key]))
    else:
        a9 = l9 = 0
    # 成交：王的9月单归赵
    def owner(s):
        return '赵老师' if (s == '王老师' and take9 is False) else s
    sub = [r for r in orders if r['m'] and r['sales'] == key
           and not (key == '王老师' and r['m'] >= 9)]
    if key == '赵老师':
        sub += [r for r in orders if r['m'] == 9 and r['sales'] == '王老师']
    rev = sum(r['amt'] for r in sub)
    cn = len(sub)
    cc = len(set(r['phone'] for r in sub))
    nc = len(set(r['phone'] for r in sub if first_month.get(r['phone']) is not None))
    tot_add = a58 + a9
    tot_lost = l58 + l9
    adv_rows.append(dict(key=key, label=label, a58=a58, l58=l58, a9=a9, l9=l9,
                         add=tot_add, lost=tot_lost, net=tot_add - tot_lost,
                         cn=cn, cc=cc, rev=rev,
                         aov=rev / cn if cn else 0,
                         conv=cc / tot_add if tot_add else 0,
                         arpu=rev / (tot_add - tot_lost) if (tot_add - tot_lost) else 0))

ALL_CC = len(set(r['phone'] for r in orders if r['phone']))
TOT = dict(add=sum(x['add'] for x in adv_rows), lost=sum(x['lost'] for x in adv_rows),
           cn=sum(x['cn'] for x in adv_rows), cc=sum(x['cc'] for x in adv_rows),
           rev=sum(x['rev'] for x in adv_rows))

# ── 顾问×月度 ──
adv_n, adv_rev = defaultdict(dict), defaultdict(dict)
for key in ADV_ORDER:
    for m in MONTHS:
        sub = [r for r in orders if r['m'] == m and r['sales'] == key]
        if key == '王老师' and m >= 9:
            sub = []
        if key == '赵老师' and m >= 9:
            sub += [r for r in orders if r['m'] == m and r['sales'] == '王老师']
        adv_n[key][m] = len(sub)
        adv_rev[key][m] = sum(r['amt'] for r in sub)

# ══════ 样式 ══════
F_TITLE = Font(size=12, bold=True, color='FF1F3864')
F_SUB = Font(size=9, color='FF595959')
F_SEC = Font(size=11, bold=True, color='FF1F3864')
F_HDR = Font(size=10, bold=True, color='FFFFFFFF')
F_BODY = Font(size=10)
F_KEY = Font(size=10, bold=True, color='FF1F3864')
F_NOTE = Font(size=9, color='FFC00000')
FILL_HDR = PatternFill('solid', fgColor='FF2E5C8A')
FILL_TOTAL = PatternFill('solid', fgColor='FFEFF3F9')
THIN = Side(style='thin', color='FFBFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
C_CTR = Alignment(horizontal='center', vertical='center')
C_RIGHT = Alignment(horizontal='right', vertical='center')
NUM2, NUM0, PCT = '#,##0.00', '#,##0', '0.00%'

wb = openpyxl.load_workbook(XLSX)
for sn in ['月度总表', '月度转化率', '顾问明细', '客单价月度跟进', '顾问×月度客单价']:
    if sn in wb.sheetnames:
        del wb[sn]

def hdr_row(ws, r, cols):
    for i, h in enumerate(cols, 1):
        c = ws.cell(r, i, h)
        c.font = F_HDR; c.fill = FILL_HDR; c.alignment = C_CTR; c.border = BORDER

def note(ws, r, text, red=False):
    ws.cell(r, 1, text).font = F_NOTE if red else F_SUB

# ══════════ Sheet 1：月度总表 ══════════
ws = wb.create_sheet('月度总表')
ws.sheet_properties.tabColor = '2E5C8A'
for i, w in enumerate([10, 12, 12, 12, 16, 11, 11, 13, 12, 13], 1):
    ws.column_dimensions[get_column_letter(i)].width = w
r = 1
ws.cell(r, 1, '三诺×华大私域 · 月度经营总表（2026年5-9月）').font = F_TITLE; r += 1
note(ws, r, '口径：客户基数 = 自5月项目起点的累计净增客户（新增−流失）；'
            '单客价值 = 当月营收 ÷ 累计客户基数　|　生成时间：'
            f'{datetime.now():%Y-%m-%d %H:%M}'); r += 1
note(ws, r, '★ 组织变更（2026-09起）：黎老师停止服务（9月新增12/流失9已剔除团队口径）；'
            '王老师账号由赵老师一并接待（9月新增199/流失45并入赵）。', red=True); r += 2
hdr_row(ws, r, ['月份', '新增客户', '流失客户', '当月净增', '累计客户基数\n(5月起净增)',
                '成交单量', '成交客户', '营收(元)', '客单价(元)', '单客价值(元)'])
r += 1
for x in monthly:
    ws.cell(r, 1, f"{x['m']:02d}月")
    ws.cell(r, 2, x['add']); ws.cell(r, 3, x['lost']); ws.cell(r, 4, x['net'])
    ws.cell(r, 5, x['cum']); ws.cell(r, 6, x['n']); ws.cell(r, 7, x['cust'])
    ws.cell(r, 8, round(x['rev'], 2)); ws.cell(r, 9, round(x['aov'], 2))
    ws.cell(r, 10, round(x['arpu'], 2))
    for col in range(1, 11):
        c = ws.cell(r, col); c.border = BORDER
        c.font = F_KEY if col == 10 else F_BODY
        c.alignment = C_CTR if col == 1 else C_RIGHT
        if col == 8: c.number_format = NUM2
        elif col in (9, 10): c.number_format = NUM0
    r += 1
_all_rev = sum(x['rev'] for x in monthly); _all_n = sum(x['n'] for x in monthly)
TA = sum(x['add'] for x in monthly); TL = sum(x['lost'] for x in monthly)
ws.cell(r, 1, '合计'); ws.cell(r, 2, TA); ws.cell(r, 3, TL); ws.cell(r, 4, TA - TL)
ws.cell(r, 5, monthly[-1]['cum']); ws.cell(r, 6, _all_n); ws.cell(r, 7, ALL_CC)
ws.cell(r, 8, round(_all_rev, 2)); ws.cell(r, 9, round(_all_rev / _all_n, 2))
ws.cell(r, 10, round(_all_rev / monthly[-1]['cum'], 2))
for col in range(1, 11):
    c = ws.cell(r, col); c.border = BORDER; c.font = F_KEY; c.fill = FILL_TOTAL
    c.alignment = C_CTR if col == 1 else C_RIGHT
    if col == 8: c.number_format = NUM2
    elif col in (9, 10): c.number_format = NUM0
r += 2
for t in [
    '注：① 9月为1-17日数据，非整月；',
    '　　② 9月起黎老师停止服务，其9月新增12/流失9不计入团队；王老师账号由赵老师一并接待，人数不变（团队由5人变4个号/3位接待人）；',
    '　　③ 5-8月为原5人架构，数据未作调整，可直接与前月对比；',
    '　　④ 按整段区间直取5-9月净增为6,373，按月度加总为6,368，差5为区间边界重复计算。',
]:
    note(ws, r, t); r += 1

# ══════════ Sheet 2：月度转化率 ══════════
ws = wb.create_sheet('月度转化率')
ws.sheet_properties.tabColor = '5B8DC8'
for i, w in enumerate([10, 13, 15, 16, 13, 14, 15, 12, 16], 1):
    ws.column_dimensions[get_column_letter(i)].width = w
r = 1
ws.cell(r, 1, '月度转化率（2026年5-9月）').font = F_TITLE; r += 1
note(ws, r, '★ 9月起组织变更：黎老师停服、王老师号并入赵老师，9月分母已按4个号口径调整。', red=True); r += 2
hdr_row(ws, r, ['月份', '新增客户(A)', '当月首单成交人数(B)', '新客转化率\nB/A (目标30%)',
                '成交客户总数', '累计客户基数(C)', '全基盘转化率', '老客成交人数', '老客转化率(估)\n目标10%'])
r += 1
for x in monthly:
    rep = x['cust'] - x['newcust']
    rep_base = x['cum'] - x['add']
    ws.cell(r, 1, f"{x['m']:02d}月")
    ws.cell(r, 2, x['add']); ws.cell(r, 3, x['newcust'])
    ws.cell(r, 4, x['conv_new']); ws.cell(r, 5, x['cust'])
    ws.cell(r, 6, x['cum']); ws.cell(r, 7, x['conv_base'])
    ws.cell(r, 8, rep)
    ws.cell(r, 9, rep / rep_base if rep_base else None)
    for col in range(1, 10):
        c = ws.cell(r, col); c.border = BORDER; c.font = F_BODY
        c.alignment = C_CTR if col == 1 else C_RIGHT
        if col in (4, 7, 9): c.number_format = PCT
    r += 1
TA2 = sum(x['add'] for x in monthly)
ws.cell(r, 1, '合计'); ws.cell(r, 2, TA2)
ws.cell(r, 3, sum(x['newcust'] for x in monthly))
ws.cell(r, 4, sum(x['newcust'] for x in monthly) / TA2)
ws.cell(r, 5, ALL_CC); ws.cell(r, 6, monthly[-1]['cum'])
ws.cell(r, 7, ALL_CC / monthly[-1]['cum'])
for col in range(1, 10):
    c = ws.cell(r, col); c.border = BORDER; c.font = F_KEY; c.fill = FILL_TOTAL
    c.alignment = C_CTR if col == 1 else C_RIGHT
    if col in (4, 7): c.number_format = PCT
r += 2
for t in ['注：① 新客转化率分母 = 当月新增客户数；② 新客口径为「当月产生首单的客户」，非严格同月加微队列（缺客户级加微时间）；',
          '　　③ 老客分母 = 累计基数 − 当月新增，为推算值；④ 项目总目标为月转化率(下单) ≥ 10%；',
          '　　⑤ 9月新增客户数 821→809（剔除黎老师已停服号的12人），流失 183→174。']:
    note(ws, r, t, red=(t.startswith('　　⑤'))); r += 1

# ══════════ Sheet 3：顾问明细 ══════════
ws = wb.create_sheet('顾问明细')
ws.sheet_properties.tabColor = '7BA7D7'
for i, w in enumerate([24, 11, 11, 11, 11, 11, 11, 11, 9, 10, 12, 11, 11, 11], 1):
    ws.column_dimensions[get_column_letter(i)].width = w
r = 1
ws.cell(r, 1, '按顾问 · 5-9月累计（客户基数 = 该顾问5月起累计净增）').font = F_TITLE; r += 1
note(ws, r, '★ 9月起：王老师账号并入赵老师（9月数据已合并）；黎老师停服（仅计5-8月贡献）。', red=True); r += 2
hdr_row(ws, r, ['顾问', '5-8月新增', '5-8月流失', '9月新增', '9月流失', '累计新增', '累计流失',
                '净增(基数)', '成交单', '成交客户', '营收(元)', '客单价(元)', '新客转化率', '单客价值(元)'])
r += 1
for x in adv_rows:
    hi = x['key'] in ('赵老师', '黎老师')
    ws.cell(r, 1, x['label'])
    ws.cell(r, 2, x['a58']); ws.cell(r, 3, x['l58'])
    ws.cell(r, 4, x['a9'] if x['a9'] else ('并入赵' if x['key'] == '王老师' else ('停服' if x['key'] == '黎老师' else 0)))
    ws.cell(r, 5, x['l9'] if x['l9'] else '')
    ws.cell(r, 6, x['add']); ws.cell(r, 7, x['lost']); ws.cell(r, 8, x['net'])
    ws.cell(r, 9, x['cn']); ws.cell(r, 10, x['cc'])
    ws.cell(r, 11, round(x['rev'], 2)); ws.cell(r, 12, round(x['aov'], 2))
    ws.cell(r, 13, x['conv']); ws.cell(r, 14, round(x['arpu'], 2))
    for col in range(1, 15):
        c = ws.cell(r, col); c.border = BORDER
        c.font = F_KEY if hi else F_BODY
        c.alignment = Alignment(horizontal='left', vertical='center') if col == 1 else C_RIGHT
        if col == 11: c.number_format = NUM2
        elif col in (12, 14): c.number_format = NUM0
        elif col == 13: c.number_format = PCT
    r += 1
ws.cell(r, 1, '合计')
for col, v in [(2, sum(x['a58'] for x in adv_rows)), (3, sum(x['l58'] for x in adv_rows)),
               (4, sum(x['a9'] for x in adv_rows)), (5, sum(x['l9'] for x in adv_rows)),
               (6, TOT['add']), (7, TOT['lost']), (8, TOT['add'] - TOT['lost']),
               (9, TOT['cn']), (10, ALL_CC), (11, round(TOT['rev'], 2)),
               (12, round(TOT['rev'] / TOT['cn'], 2)), (13, ALL_CC / TOT['add']),
               (14, round(TOT['rev'] / (TOT['add'] - TOT['lost']), 2))]:
    ws.cell(r, col, v)
for col in range(1, 15):
    c = ws.cell(r, col); c.border = BORDER; c.font = F_KEY; c.fill = FILL_TOTAL
    c.alignment = Alignment(horizontal='left', vertical='center') if col == 1 else C_RIGHT
    if col == 11: c.number_format = NUM2
    elif col in (12, 14): c.number_format = NUM0
    elif col == 13: c.number_format = PCT
r += 2
for t in ['注：① 「赵老师（含王老师号·9月起）」的9月数据 = 赵号222新增/43流失 + 王号199新增/45流失；王老师行仅含5-8月；',
          '　　② 王老师9月成交为 0 单，故并入后赵老师营收与单量不变；',
          '　　③ 黎老师5-8月客户盘仍计入项目累计基数（889净增），9月起不再产生新增；',
          '　　④ 单客价值 = 该顾问累计营收 ÷ 其累计净增基数；',
          f'　　⑤ 合计成交客户 {ALL_CC} 人为全局去重；各顾问相加为 {TOT["cc"]} 人，差 {TOT["cc"]-ALL_CC} 人为跨顾问重复成交客户。']:
    note(ws, r, t); r += 1

# ══════════ Sheet 4：客单价月度跟进 ══════════
ws = wb.create_sheet('客单价月度跟进')
ws.sheet_properties.tabColor = '2E5C8A'
for i, w in enumerate([10, 11, 11, 13, 13, 14, 12, 12, 12, 12], 1):
    ws.column_dimensions[get_column_letter(i)].width = w
r = 1
ws.cell(r, 1, '客单价月度跟进（2026年5-9月）').font = F_TITLE; r += 1
note(ws, r, '口径：客单价(按单) = 当月营收 ÷ 成交单量；客单价(按客户) = 当月营收 ÷ 成交客户数。'
            '数据源：三诺&华大成交用户清单.xlsx（清单口径，低于结算净额约10~13%）'); r += 1
note(ws, r, f'生成时间：{datetime.now():%Y-%m-%d %H:%M}　|　9月为1-17日数据，非整月　|　'
            '★ 9月起黎老师停服、王老师号并入赵老师（团队合计不变）', red=True); r += 2
ws.cell(r, 1, '一、月度趋势').font = F_SEC; r += 1
hdr_row(ws, r, ['月份', '成交单量', '成交客户', '营收(元)', '客单价(按单)', '客单价(按客户)',
                '环比', '单笔中位数', '单笔最高']); r += 1
for idx, x in enumerate(monthly):
    prev = monthly[idx - 1]['aov'] if idx > 0 else None
    ws.cell(r, 1, f"{x['m']:02d}月")
    ws.cell(r, 2, x['n']); ws.cell(r, 3, x['cust']); ws.cell(r, 4, round(x['rev'], 2))
    ws.cell(r, 5, round(x['aov'], 2)); ws.cell(r, 6, round(x['aov_c'], 2))
    ws.cell(r, 7, (x['aov'] / prev - 1) if prev else None)
    ws.cell(r, 8, round(x['median'], 2)); ws.cell(r, 9, round(x['maxv'], 2))
    for col in range(1, 10):
        c = ws.cell(r, col); c.border = BORDER
        c.font = F_KEY if col == 5 else F_BODY
        c.alignment = C_CTR if col == 1 else C_RIGHT
        if col == 4: c.number_format = NUM2
        elif col in (5, 6, 8, 9): c.number_format = NUM0
        elif col == 7: c.number_format = '+0.0%;-0.0%;0.0%'
    r += 1
for m in (10, 11, 12):
    ws.cell(r, 1, f'{m:02d}月')
    ws.cell(r, 5, f'=IFERROR(D{r}/B{r},"")')
    ws.cell(r, 6, f'=IFERROR(D{r}/C{r},"")')
    ws.cell(r, 7, f'=IFERROR(E{r}/E{r-1},"")')
    for col in range(1, 10):
        c = ws.cell(r, col); c.border = BORDER; c.font = Font(size=10, color='FFA6A6A6')
        c.alignment = C_CTR if col == 1 else C_RIGHT
        if col == 4: c.number_format = NUM2
        elif col in (5, 6, 8, 9): c.number_format = NUM0
        elif col == 7: c.number_format = '+0.0%;-0.0%;0.0%'
    r += 1
ws.cell(r, 1, '合计'); ws.cell(r, 2, _all_n); ws.cell(r, 3, ALL_CC)
ws.cell(r, 4, round(_all_rev, 2)); ws.cell(r, 5, round(_all_rev / _all_n, 2))
ws.cell(r, 6, round(_all_rev / ALL_CC, 2))
for col in range(1, 10):
    c = ws.cell(r, col); c.border = BORDER; c.font = F_KEY; c.fill = FILL_TOTAL
    c.alignment = C_CTR if col == 1 else C_RIGHT
    if col == 4: c.number_format = NUM2
    elif col in (5, 6, 8, 9): c.number_format = NUM0
r += 3
ws.cell(r, 1, '二、客单价驱动拆解（套组 vs 单品 vs 高价单）').font = F_SEC; r += 1
hdr_row(ws, r, ['月份', '套组单量', '套组客单价', '单品单量', '单品客单价',
                '高价位(≥1000)单量', '高价位营收', '高价位营收占比', '套组营收占比']); r += 1
for x in monthly:
    ws.cell(r, 1, f"{x['m']:02d}月")
    ws.cell(r, 2, x['grp_n']); ws.cell(r, 3, round(x['grp_aov'], 2) if x['grp_aov'] else None)
    ws.cell(r, 4, x['sg_n']); ws.cell(r, 5, round(x['sg_aov'], 2) if x['sg_aov'] else None)
    ws.cell(r, 6, x['hi_n']); ws.cell(r, 7, round(x['hi_rev'], 2))
    ws.cell(r, 8, x['hi_rev'] / x['rev'] if x['rev'] else None)
    ws.cell(r, 9, x['grp_rev'] / x['rev'] if x['rev'] else None)
    for col in range(1, 10):
        c = ws.cell(r, col); c.border = BORDER; c.font = F_BODY
        c.alignment = C_CTR if col == 1 else C_RIGHT
        if col in (3, 5, 7): c.number_format = NUM0
        elif col in (8, 9): c.number_format = PCT
    r += 1
r += 2
ws.cell(r, 1, '三、客单价分档结构（单位：单量 / 占当月单量比例）').font = F_SEC; r += 1
hdr_row(ws, r, ['月份', '<300', '300~600', '600~1000', '1000~2000', '≥2000', '当月单量']); r += 1
BANDS = [(0, 300), (300, 600), (600, 1000), (1000, 2000), (2000, 10 ** 9)]
for x in monthly:
    sub = msub(x['m']); tot = len(sub)
    ws.cell(r, 1, f"{x['m']:02d}月")
    for i, (lo, hi) in enumerate(BANDS, 2):
        cnt = len([q for q in sub if lo <= q['amt'] < hi])
        ws.cell(r, i, f'{cnt}单 / {cnt / tot * 100:.0f}%' if tot else '0单')
    ws.cell(r, 7, tot)
    for col in range(1, 8):
        c = ws.cell(r, col); c.border = BORDER; c.font = F_BODY; c.alignment = C_CTR
    r += 1
r += 2
ws.cell(r, 1, '四、月度续填指引').font = F_SEC; r += 1
for t in ['① 每月更新：把当月「成交单量 / 成交客户 / 营收」填入「一、月度趋势」对应行，客单价、环比自动计算（10-12月行已预留公式）。',
          '② 营收与单量来源：三诺&华大成交用户清单.xlsx 当月成交记录汇总（清单口径）。',
          '③ 客单价(按单) 反映订单结构，客单价(按客户) 反映客户价值，两者一起看。',
          '④ 环比为正说明单笔金额在放大，需同时对照「成交单量」——若单量下滑而客单价上升，是「少单变贵」，不等于增长。',
          '⑤ 「二、驱动拆解」看客单价上涨是套组拉动还是单品涨价；套组客单价通常是单品的 3~6 倍。',
          '⑥ 「三、分档结构」看高价单占比是否在扩大；≥1000 元订单是客单价的主要推手。',
          '⑦ 9月起顾问归属变更（王→赵、黎停服）不影响团队合计，但影响「顾问×月度客单价」的横向对比。']:
    note(ws, r, t); r += 1
ws.freeze_panes = 'A11'

# ══════════ Sheet 5：顾问×月度客单价 ══════════
ws2 = wb.create_sheet('顾问×月度客单价')
ws2.sheet_properties.tabColor = '5B8DC8'
for i, w in enumerate([26, 13, 13, 13, 13, 13, 13], 1):
    ws2.column_dimensions[get_column_letter(i)].width = w
r = 1
ws2.cell(r, 1, '顾问 × 月度客单价（2026年5-9月）').font = F_TITLE; r += 1
note(ws2, r, '客单价(元) = 该顾问当月营收 ÷ 当月单量。★ 9月起：王老师号并入赵老师（9月数据已合并），黎老师停服（9月留空）。',
     red=True); r += 2

def adv_block(title, kind, fmt):
    global r
    ws2.cell(r, 1, title).font = F_SEC; r += 1
    ws2.cell(r, 1, '顾问')
    for i, m in enumerate(MONTHS, 2):
        ws2.cell(r, i, f'{m}月')
    ws2.cell(r, 7, '累计')
    for col in range(1, 8):
        c = ws2.cell(r, col); c.font = F_HDR; c.fill = FILL_HDR
        c.alignment = C_CTR; c.border = BORDER
    r += 1
    for key, label, _ in SPEC:
        label_s = label.split('（')[0]
        ws2.cell(r, 1, label if len(label) < 20 else label)
        for i, m in enumerate(MONTHS, 2):
            n = adv_n[key][m]; rv = adv_rev[key][m]
            if kind == 'aov':
                ws2.cell(r, i, round(rv / n, 2) if n else None)
            elif kind == 'n':
                ws2.cell(r, i, n if not (key in ('王老师', '黎老师') and m >= 9) else None)
            else:
                ws2.cell(r, i, round(rv, 2) if not (key in ('王老师', '黎老师') and m >= 9) else None)
        tn = sum(adv_n[key][m] for m in MONTHS); tr = sum(adv_rev[key][m] for m in MONTHS)
        ws2.cell(r, 7, round(tr / tn, 2) if kind == 'aov' and tn else
                 (tn if kind == 'n' else (round(tr, 2) if kind == 'rev' else None)))
        for col in range(1, 8):
            c = ws2.cell(r, col); c.border = BORDER; c.font = F_BODY
            c.alignment = Alignment(horizontal='left', vertical='center') if col == 1 else C_RIGHT
            c.number_format = fmt
        r += 1
    ws2.cell(r, 1, '合计（4个号）')
    for i, m in enumerate(MONTHS, 2):
        x = next(y for y in monthly if y['m'] == m)
        ws2.cell(r, i, round(x['aov'], 2) if kind == 'aov' else (x['n'] if kind == 'n' else round(x['rev'], 2)))
    alln = sum(x['n'] for x in monthly); allr = sum(x['rev'] for x in monthly)
    ws2.cell(r, 7, round(allr / alln, 2) if kind == 'aov' else (alln if kind == 'n' else round(allr, 2)))
    for col in range(1, 8):
        c = ws2.cell(r, col); c.border = BORDER; c.font = F_KEY; c.fill = FILL_TOTAL
        c.alignment = Alignment(horizontal='left', vertical='center') if col == 1 else C_RIGHT
        c.number_format = fmt
    r += 3

adv_block('一、客单价矩阵（元/单）', 'aov', NUM0)
adv_block('二、成交单量矩阵（单）', 'n', NUM0)
adv_block('三、营收矩阵（元）', 'rev', NUM2)
note(ws2, r, '注：① 空白表示当月无成交或已停服/并入（王老师9月起并入赵老师，黎老师9月起停服）；'
             '② 「赵老师（含王老师号·9月起）」的9月客单价 = (赵¥14,243 + 王¥0) ÷ 11单；'
             '③ 客单价受单笔大单影响明显，单量少的月份波动大，建议结合「二、成交单量矩阵」一起看。')

wb.save(XLSX)
print('✅ 已生成 5 sheet:', wb.sheetnames)
print('\n月度（团队·4个号口径）:')
for x in monthly:
    print(f"  {x['m']}月 新增{x['add']:>5} 流失{x['lost']:>4} 净增{x['net']:>5} 基数{x['cum']:>5} "
          f"{x['n']:>3}单 ¥{x['rev']:>10,.2f} 客单¥{x['aov']:>7,.0f} 单客价值¥{x['arpu']:>6.2f} "
          f"新客转化{x['conv_new']*100:>5.2f}% 全基盘{x['conv_base']*100:>5.2f}%")
print('\n顾问:')
for x in adv_rows:
    print(f"  {x['label']:<26} 新增{x['add']:>5} 流失{x['lost']:>4} 净增{x['net']:>5} "
          f"{x['cn']:>3}单 ¥{x['rev']:>10,.2f} 客单¥{x['aov']:>7,.0f} 单客价值¥{x['arpu']:>6.2f}")
print(f"  {'合计':<26} 新增{TOT['add']:>5} 流失{TOT['lost']:>4} 净增{TOT['add']-TOT['lost']:>5} "
      f"{TOT['cn']:>3}单 ¥{TOT['rev']:>10,.2f}")
