#!/usr/bin/env python3
"""生成三诺&华大本周销售排名海报 — 适合发朋友圈/微信"""

from PIL import Image, ImageDraw, ImageFont
import os

W = 1080
H = 1520

img = Image.new("RGB", (W, H), "#0F172A")
draw = ImageDraw.Draw(img)

# ── 字体 ──
# macOS 上优先用 Hiragino Sans GB（支持中文），PingFang 可能不存在
font_candidates = [
    "/System/Library/Fonts/Hiragino Sans GB.ttc",
    "/System/Library/Fonts/STHeiti Medium.ttc",
    "/Library/Fonts/Arial Unicode.ttf",
]
font_path = None
for fp in font_candidates:
    if os.path.exists(fp):
        font_path = fp
        break

if font_path:
    f_title = ImageFont.truetype(font_path, 56)
    f_sub = ImageFont.truetype(font_path, 28)
    f_label = ImageFont.truetype(font_path, 20)
    f_name = ImageFont.truetype(font_path, 32)
    f_num = ImageFont.truetype(font_path, 36)
    f_big = ImageFont.truetype(font_path, 44)
    f_small = ImageFont.truetype(font_path, 24)
    f_medal = ImageFont.truetype(font_path, 40)
else:
    f_title = ImageFont.load_default()
    f_sub = f_title; f_label = f_title; f_name = f_title; f_num = f_title
    f_big = f_title; f_small = f_title; f_medal = f_title

# ── 顶部渐变背景 ──
for y in range(0, 380):
    r = int(15 + y * 0.15)
    g = int(23 + y * 0.15)
    b = int(42 + y * 0.15)
    draw.rectangle([(0, y), (W, y+1)], fill=(r, g, b))

# ── 顶部标题区 ──
draw.text((540, 80), "三诺×华大", fill="#38BDF8", font=f_title, anchor="mt")
draw.text((540, 155), "本周销售战报", fill="#FFFFFF", font=f_title, anchor="mt")

draw.text((540, 230), "2026.06.02 — 06.05", fill="#94A3B8", font=f_sub, anchor="mt")

# ── KPI 横条 ──
kpi_y = 300
kpi_items = [
    ("¥7,364", "本周营收"),
    ("10单", "成交订单"),
    ("7位", "成交客户"),
    ("¥736", "平均客单"),
]
kpi_w = W // 4
for i, (val, label) in enumerate(kpi_items):
    cx = kpi_w * i + kpi_w // 2
    draw.text((cx, kpi_y), val, fill="#38BDF8", font=f_big, anchor="mt")
    draw.text((cx, kpi_y + 52), label, fill="#94A3B8", font=f_label, anchor="mt")

# 分隔线
draw.rectangle([(60, 400), (W-60, 401)], fill="#1E293B")

# ── 排名列表 ──
rankings = [
    {"rank": 1, "medal": "1", "medal_fg": "#FFD700", "name": "赵老师", "orders": "5单", "revenue": "¥3,972", "cust": "4位客户", "bar_pct": 1.00},
    {"rank": 2, "medal": "2", "medal_fg": "#C0C0C0", "name": "黎老师", "orders": "2单", "revenue": "¥2,836", "cust": "2位客户", "bar_pct": 0.71},
    {"rank": 3, "medal": "3", "medal_fg": "#CD7F32", "name": "王老师", "orders": "2单", "revenue": "¥417",   "cust": "2位客户", "bar_pct": 0.10},
    {"rank": 4, "medal": "4", "medal_fg": "#94A3B8", "name": "朱老师", "orders": "1单", "revenue": "¥139",   "cust": "1位客户", "bar_pct": 0.035},
    {"rank": 5, "medal": "5", "medal_fg": "#64748B", "name": "吴老师", "orders": "0单", "revenue": "¥0",      "cust": "本周暂未开单", "bar_pct": 0},
]

start_y = 430
row_h = 130

for i, r in enumerate(rankings):
    y = start_y + i * row_h

    # 行背景
    bg_color = "#1E293B" if i % 2 == 0 else "#0F172A"
    draw.rectangle([(60, y), (W-60, y + row_h - 12)], fill=bg_color, outline=None)

    # 排名/奖牌 — 用彩色圆底
    medal_r = 24
    medal_cx = 110
    medal_cy = y + row_h//2 - 5
    draw.ellipse(
        [(medal_cx - medal_r, medal_cy - medal_r), (medal_cx + medal_r, medal_cy + medal_r)],
        fill=r["medal_fg"]
    )
    draw.text((medal_cx, medal_cy), r["medal"], fill="#0F172A", font=f_name, anchor="mm")

    # 名字
    name_color = "#F8FAFC" if r["revenue"] != "¥0" else "#64748B"
    draw.text((180, y + row_h//2 - 18), r["name"], fill=name_color, font=f_name, anchor="lm")

    # 排名进度条
    bar_x = 300
    bar_w = 320
    bar_h = 14
    bar_y = y + row_h//2 + 6
    draw.rectangle([(bar_x, bar_y), (bar_x + bar_w, bar_y + bar_h)], fill="#334155")
    if r["bar_pct"] > 0:
        fill_w = int(bar_w * r["bar_pct"])
        # 渐变条
        bar_color = "#38BDF8" if r["rank"] == 1 else ("#06B6D4" if r["rank"] == 2 else "#64748B")
        draw.rectangle([(bar_x, bar_y), (bar_x + fill_w, bar_y + bar_h)], fill=bar_color)

        # 占比文字
        pct_text = f"{r['revenue']}" if r["revenue"] != "¥0" else ""
        if pct_text:
            draw.text((bar_x + fill_w + 12, bar_y + bar_h//2), pct_text, fill=bar_color, font=f_small, anchor="lm")

    # 订单数
    order_color = "#E2E8F0" if r["orders"] != "0单" else "#64748B"
    draw.text((680, y + row_h//2 - 18), r["orders"], fill=order_color, font=f_num, anchor="lm")

    # 客户数
    cust_color = "#94A3B8" if r["cust"] != "本周暂未开单" else "#EF4444"
    draw.text((680, y + row_h//2 + 18), r["cust"], fill=cust_color, font=f_label, anchor="lm")

    # 标签
    draw.text((810, y + row_h//2 - 18), "订单", fill="#475569", font=f_label, anchor="lm")
    draw.text((810, y + row_h//2 + 18), "客户", fill="#475569", font=f_label, anchor="lm")

# ── 底部 ──
bottom_y = start_y + len(rankings) * row_h + 30

# 警示 + 数据源
draw.text((540, bottom_y + 20), "! 吴老师连续挂零，建议本周主动跟进", fill="#F97316", font=f_sub, anchor="mt")
draw.text((540, bottom_y + 70), "数据来源：三诺&华大成交用户清单 ｜ 实时看板：qikingplus.github.io/sannuo-huada-dashboard", fill="#475569", font=f_label, anchor="mt")

draw.text((540, bottom_y + 120), "三诺×华大 · 私域运营 · 数据驱动增长", fill="#1E293B", font=f_sub, anchor="mt")

# ── 保存 ──
out = "/Users/SaciNa/Desktop/本周销售排名_0605.png"
img.save(out, "PNG")
print(f"✅ 海报已保存: {out}")
print(f"   尺寸: {W}x{H}")
