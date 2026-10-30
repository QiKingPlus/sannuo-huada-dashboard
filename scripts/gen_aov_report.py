# -*- coding: utf-8 -*-
"""客单价月度跟进报表生成：在现有 xlsx 中新增 2 个 sheet"""
import openpyxl, json, shutil, os
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
from collections import defaultdict, OrderedDict

BASE = '/Users/SaciNa/WorkBuddy/2026-05-26-15-30-31'
XLSX = f'{BASE}/三诺华大月度转化率与单客价值_2026年5-9月.xlsx'
SRC = '/Users/SaciNa/Documents/臻选&芥子/臻选私域/三诺&华大成交用户清单.xlsx'
GROWTH = '/Users/SaciNa/WorkBuddy/2026-09-17-11-36-56/outputs/wbz_customer_growth_raw.json'
NAME_MAP = {"赵老师": "赵老师", "三诺健康顾问丨朱国银": "朱老师", "三诺健康管家 | 诺诺": "王老师",
            "三诺健康管家｜吴老师": "吴老师", "三诺健康管家 | 黎老师": "黎老师"}
SHORT = ["赵老师", "朱老师", "黎老师", "吴老师", "王老师"]
MONTHS = [5, 6, 7, 8, 9]

# ── 备份 ──
os.makedirs(f'{BASE}/outputs', exist_ok=True)
shutil.copy(XLSX, f'{BASE}/outputs/三诺华大月度转化率与单客价值_备份_{datetime.now():%Y%m%d_%H%M}.xlsx')

# ── 读取成交清单 ──
wb_src = openpyxl.load_workbook(SRC, data_only=True)
ws_src = wb_src.active
# 按表头名定位列（源表曾插入「微信」列，硬编码 row[n] 会整体错位）
_hdr = [str(c).strip() if c is not None else '' for c in
        next(ws_src.iter_rows(min_row=1, max_row=1, values_only=True))]
IDX_PHONE = _hdr.index('手机') if '手机' in _hdr else 5
rows = []
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
    rows.append({'sales': str(row[0] or ''), 'date': d, 'product': str(row[2] or ''),
                 'amt': float(row[3] or 0), 'phone': str(row[IDX_PHONE] or '')})

def msub(m):
    return [r for r in rows if r['date'] and r['date'].month == m and r['date'].year == 2026]

def is_grp(p):
    return ('套组' in p) or ('+' in p) or ('合并' in p)

# ── 月度指标 ──
monthly = []
for m in MONTHS:
    sub = msub(m)
    rev = sum(r['amt'] for r in sub)
    n = len(sub)
    cust = len(set(r['phone'] for r in sub))
    grp = [r for r in sub if is_grp(r['product'])]
    single = [r for r in sub if not is_grp(r['product'])]
    hi = [r for r in sub if r['amt'] >= 1000]
    amts = sorted(r['amt'] for r in sub)
    monthly.append(dict(
        m=m, n=n, cust=cust, rev=rev,
        aov=rev / n if n else 0, aov_c=rev / cust if cust else 0,
        grp_n=len(grp), grp_rev=sum(r['amt'] for r in grp),
        grp_aov=(sum(r['amt'] for r in grp) / len(grp)) if grp else 0,
        sg_n=len(single), sg_rev=sum(r['amt'] for r in single),
        sg_aov=(sum(r['amt'] for r in single) / len(single)) if single else 0,
        hi_n=len(hi), hi_rev=sum(r['amt'] for r in hi),
        median=amts[len(amts) // 2] if amts else 0,
        maxv=amts[-1] if amts else 0))

# ── 分档 ──
BANDS = [(0, 300, '<300'), (300, 600, '300~600'), (600, 1000, '600~1000'),
         (1000, 2000, '1000~2000'), (2000, 10 ** 9, '≥2000')]
band_data = {}
for m in MONTHS:
    sub = msub(m)
    band_data[m] = [(lab, len([r for r in sub if lo <= r['amt'] < hi]).__int__()) for lo, hi, lab in BANDS]

# ── 顾问 × 月度 ──
adv_aov, adv_n, adv_rev = {}, {}, {}
for s in SHORT:
    adv_aov[s], adv_n[s], adv_rev[s] = {}, {}, {}
    for m in MONTHS:
        sub = [r for r in msub(m) if r['sales'] == s]
        rv = sum(r['amt'] for r in sub)
        adv_n[s][m] = len(sub)
        adv_rev[s][m] = rv
        adv_aov[s][m] = rv / len(sub) if sub else 0

# ── 样式 ──
F_TITLE = Font(size=12, bold=True, color='FF1F3864')
F_SUB = Font(size=9, color='FF595959')
F_SEC = Font(size=11, bold=True, color='FF1F3864')
F_HDR = Font(size=10, bold=True, color='FFFFFFFF')
F_BODY = Font(size=10)
F_KEY = Font(size=10, bold=True, color='FF1F3864')
F_COM = Font(size=9, color='FF7F7F7F')
FILL_HDR = PatternFill('solid', fgColor='FF2E5C8A')
FILL_SEC = PatternFill('solid', fgColor='FFDCE6F1')
FILL_TOTAL = PatternFill('solid', fgColor='FFEFF3F9')
FILL_KEY = PatternFill('solid', fgColor='FFF2F7FC')
THIN = Side(style='thin', color='FFBFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
C_CTR = Alignment(horizontal='center', vertical='center')
C_LEFT = Alignment(horizontal='left', vertical='center')
C_RIGHT = Alignment(horizontal='right', vertical='center')
NUM2 = '#,##0.00'
NUM0 = '#,##0'
PCT = '0.0%'

wb = openpyxl.load_workbook(XLSX)
for sn in ['客单价月度跟进', '顾问×月度客单价']:
    if sn in wb.sheetnames:
        del wb[sn]

# ══════════════ Sheet: 客单价月度跟进 ══════════════
ws = wb.create_sheet('客单价月度跟进')
ws.sheet_properties.tabColor = '2E5C8A'
WIDTHS = [10, 11, 11, 13, 13, 14, 12, 12, 12, 12]
for i, w in enumerate(WIDTHS, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

r = 1
ws.cell(r, 1, '客单价月度跟进（2026年5-9月）').font = F_TITLE
r += 1
ws.cell(r, 1, '口径：客单价(按单) = 当月营收 ÷ 成交单量；客单价(按客户) = 当月营收 ÷ 成交客户数。'
              '数据源：三诺&华大成交用户清单.xlsx（清单口径，低于结算净额约10~13%）').font = F_SUB
r += 1
ws.cell(r, 1, f'生成时间：{datetime.now():%Y-%m-%d %H:%M}　|　9月为1-17日数据，非整月').font = F_SUB
r += 2

# ── 一、月度趋势 ──
ws.cell(r, 1, '一、月度趋势').font = F_SEC
r += 1
hdr = ['月份', '成交单量', '成交客户', '营收(元)', '客单价(按单)', '客单价(按客户)', '环比', '单笔中位数', '单笔最高']
for i, h in enumerate(hdr, 1):
    c = ws.cell(r, i, h)
    c.font = F_HDR; c.fill = FILL_HDR; c.alignment = C_CTR; c.border = BORDER
hdr_row = r
r += 1
first_row = r
for idx, x in enumerate(monthly):
    prev = monthly[idx - 1]['aov'] if idx > 0 else None
    ws.cell(r, 1, f"{x['m']:02d}月")
    ws.cell(r, 2, x['n'])
    ws.cell(r, 3, x['cust'])
    ws.cell(r, 4, round(x['rev'], 2))
    ws.cell(r, 5, round(x['aov'], 2))
    ws.cell(r, 6, round(x['aov_c'], 2))
    ws.cell(r, 7, (x['aov'] / prev - 1) if prev else None)
    ws.cell(r, 8, round(x['median'], 2))
    ws.cell(r, 9, round(x['maxv'], 2))
    for col in range(1, 10):
        c = ws.cell(r, col)
        c.border = BORDER
        c.font = F_KEY if col == 5 else F_BODY
        c.alignment = C_CTR if col == 1 else C_RIGHT
        if col == 4: c.number_format = NUM2
        elif col in (5, 6, 8, 9): c.number_format = NUM0
        elif col == 7: c.number_format = '+0.0%;-0.0%;0.0%'
    r += 1
# 空行（续填）10-12月
for m in (10, 11, 12):
    ws.cell(r, 1, f'{m:02d}月')
    ws.cell(r, 5, f'=IFERROR(D{r}/B{r},"")')
    ws.cell(r, 6, f'=IFERROR(D{r}/C{r},"")')
    ws.cell(r, 7, f'=IFERROR(E{r}/E{r-1},"")')
    for col in range(1, 10):
        c = ws.cell(r, col)
        c.border = BORDER; c.font = F_COM
        c.alignment = C_CTR if col == 1 else C_RIGHT
        if col == 4: c.number_format = NUM2
        elif col in (5, 6, 8, 9): c.number_format = NUM0
        elif col == 7: c.number_format = '+0.0%;-0.0%;0.0%'
    r += 1
# 合计
_all_rev = sum(x['rev'] for x in monthly)
_all_n = sum(x['n'] for x in monthly)
_all_cust = len(set(r['phone'] for r in rows if r['phone']))
ws.cell(r, 1, '合计')
ws.cell(r, 2, _all_n)
ws.cell(r, 3, _all_cust)
ws.cell(r, 4, round(_all_rev, 2))
ws.cell(r, 5, round(_all_rev / _all_n, 2))
ws.cell(r, 6, round(_all_rev / _all_cust, 2))
for col in range(1, 10):
    c = ws.cell(r, col)
    c.border = BORDER; c.font = F_KEY; c.fill = FILL_TOTAL
    c.alignment = C_CTR if col == 1 else C_RIGHT
    if col == 4: c.number_format = NUM2
    elif col in (5, 6, 8, 9): c.number_format = NUM0
r += 3

# ── 二、驱动拆解 ──
ws.cell(r, 1, '二、客单价驱动拆解（套组 vs 单品 vs 高价单）').font = F_SEC
r += 1
hdr2 = ['月份', '套组单量', '套组客单价', '单品单量', '单品客单价',
        '高价位(≥1000)单量', '高价位营收', '高价位营收占比', '套组营收占比']
for i, h in enumerate(hdr2, 1):
    c = ws.cell(r, i, h)
    c.font = F_HDR; c.fill = FILL_HDR; c.alignment = C_CTR; c.border = BORDER
r += 1
for x in monthly:
    ws.cell(r, 1, f"{x['m']:02d}月")
    ws.cell(r, 2, x['grp_n'])
    ws.cell(r, 3, round(x['grp_aov'], 2) if x['grp_aov'] else None)
    ws.cell(r, 4, x['sg_n'])
    ws.cell(r, 5, round(x['sg_aov'], 2) if x['sg_aov'] else None)
    ws.cell(r, 6, x['hi_n'])
    ws.cell(r, 7, round(x['hi_rev'], 2))
    ws.cell(r, 8, x['hi_rev'] / x['rev'] if x['rev'] else None)
    ws.cell(r, 9, x['grp_rev'] / x['rev'] if x['rev'] else None)
    for col in range(1, 10):
        c = ws.cell(r, col)
        c.border = BORDER; c.font = F_BODY
        c.alignment = C_CTR if col == 1 else C_RIGHT
        if col in (3, 5, 7): c.number_format = NUM0
        elif col in (8, 9): c.number_format = PCT
    r += 1
r += 2

# ── 三、分档 ──
ws.cell(r, 1, '三、客单价分档结构（单位：单量 / 占当月单量比例）').font = F_SEC
r += 1
hdr3 = ['月份', '<300', '300~600', '600~1000', '1000~2000', '≥2000', '当月单量']
for i, h in enumerate(hdr3, 1):
    c = ws.cell(r, i, h)
    c.font = F_HDR; c.fill = FILL_HDR; c.alignment = C_CTR; c.border = BORDER
r += 1
for x in monthly:
    tot = x['n']
    ws.cell(r, 1, f"{x['m']:02d}月")
    for i, (lab, cnt) in enumerate(band_data[x['m']], 2):
        ws.cell(r, i, f'{cnt}单 / {cnt / tot * 100:.0f}%' if tot else '0单')
    ws.cell(r, 7, tot)
    for col in range(1, 8):
        c = ws.cell(r, col)
        c.border = BORDER; c.font = F_BODY
        c.alignment = C_CTR if col == 1 else C_CTR
    r += 1
r += 2

# ── 四、续填指引 ──
ws.cell(r, 1, '四、月度续填指引').font = F_SEC
r += 1
notes = [
    '① 每月更新：把当月「成交单量 / 成交客户 / 营收」填入「一、月度趋势」对应行，客单价、环比自动计算（10-12月行已预留公式）。',
    '② 营收与单量来源：三诺&华大成交用户清单.xlsx 当月成交记录汇总（清单口径）。',
    '③ 客单价(按单) = 营收 ÷ 单量；客单价(按客户) = 营收 ÷ 成交客户数。前者反映订单结构，后者反映客户价值，两者一起看。',
    '④ 环比为正说明单笔金额在放大，需同时对照「成交单量」——若单量下滑而客单价上升，是「少单变贵」，不等于增长。',
    '⑤ 「二、驱动拆解」用来看客单价上涨是套组拉动还是单品涨价；套组客单价通常是单品的 3~6 倍。',
    '⑥ 「三、分档结构」看高价单占比是否在扩大；≥1000 元订单是客单价的主要推手。',
]
for n in notes:
    ws.cell(r, 1, n).font = F_SUB
    r += 1

ws.freeze_panes = f'A{hdr_row + 1}'

# ══════════════ Sheet: 顾问×月度客单价 ══════════════
ws2 = wb.create_sheet('顾问×月度客单价')
ws2.sheet_properties.tabColor = '5B8DC8'
for i, w in enumerate([10, 13, 13, 13, 13, 13, 13], 1):
    ws2.column_dimensions[get_column_letter(i)].width = w
r = 1
ws2.cell(r, 1, '顾问 × 月度客单价（2026年5-9月）').font = F_TITLE
r += 1
ws2.cell(r, 1, '客单价(元) = 该顾问当月营收 ÷ 当月单量。三个矩阵配套看：客单价看结构，单量看基数，营收看结果。').font = F_SUB
r += 2

def adv_block(title, kind, data_func, fmt):
    """kind: 'aov' | 'n' | 'rev'"""
    global r
    ws2.cell(r, 1, title).font = F_SEC
    r += 1
    ws2.cell(r, 1, '顾问')
    for i, m in enumerate(MONTHS, 2):
        ws2.cell(r, i, f'{m}月')
    ws2.cell(r, 7, '累计')
    for col in range(1, 8):
        c = ws2.cell(r, col)
        c.font = F_HDR; c.fill = FILL_HDR; c.alignment = C_CTR; c.border = BORDER
    r += 1

    def agg(s):
        tot_n = sum(adv_n[s][m] for m in MONTHS)
        tot_r = sum(adv_rev[s][m] for m in MONTHS)
        if kind == 'aov':
            return round(tot_r / tot_n, 2) if tot_n else None
        if kind == 'n':
            return tot_n
        return round(tot_r, 2)

    for s in SHORT:
        ws2.cell(r, 1, s)
        for i, m in enumerate(MONTHS, 2):
            v = data_func(s, m)
            ws2.cell(r, i, round(v, 2) if isinstance(v, float) else v)
        ws2.cell(r, 7, agg(s))
        for col in range(1, 8):
            c = ws2.cell(r, col)
            c.border = BORDER; c.font = F_BODY
            c.alignment = C_CTR if col == 1 else C_RIGHT
            c.number_format = fmt
        r += 1
    # 合计行
    alln = sum(x['n'] for x in monthly)
    allr = sum(x['rev'] for x in monthly)
    ws2.cell(r, 1, '合计')
    for i, m in enumerate(MONTHS, 2):
        x = next(y for y in monthly if y['m'] == m)
        ws2.cell(r, i, round(x['aov'], 2) if kind == 'aov' else (x['n'] if kind == 'n' else round(x['rev'], 2)))
    ws2.cell(r, 7, round(allr / alln, 2) if kind == 'aov' else (alln if kind == 'n' else round(allr, 2)))
    for col in range(1, 8):
        c = ws2.cell(r, col)
        c.border = BORDER; c.font = F_KEY; c.fill = FILL_TOTAL
        c.alignment = C_CTR if col == 1 else C_RIGHT
        c.number_format = fmt
    r += 3

adv_block('一、客单价矩阵（元/单）', 'aov', lambda s, m: adv_aov[s][m] if adv_n[s][m] else None, NUM0)
adv_block('二、成交单量矩阵（单）', 'n', lambda s, m: adv_n[s][m], NUM0)
adv_block('三、营收矩阵（元）', 'rev', lambda s, m: round(adv_rev[s][m], 2), NUM2)

ws2.cell(r, 1, '注：① 空白表示该顾问当月无成交单（5月黎老师，9月黎老师、王老师）；'
               '② 客单价受单笔大单影响明显，单量少的月份波动大，建议结合「二、成交单量矩阵」一起看。').font = F_SUB

wb.save(XLSX)
print('✅ 已生成，sheets:', wb.sheetnames)
print('月度客单价:', [f"{x['m']}月 ¥{x['aov']:.0f}" for x in monthly])
