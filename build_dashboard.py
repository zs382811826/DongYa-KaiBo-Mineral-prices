#!/usr/bin/env python3
"""矿产行情看板生成脚本 —— 读取 Excel 生成 HTML 看板"""
import openpyxl
import json
import os
import math
from datetime import datetime, timedelta
from collections import defaultdict

EXCEL = r"D:\矿产行情看板\my_website\矿产行情记录表.xlsx"
OUTPUT = r"D:\矿产行情看板\my_website\index.html"

CATEGORIES = [
    ("稀土行情", ["氧化钕", "氧化镨钕", "氧化铽", "氧化镝", "氧化钇", "独居石"]),
    ("锆业行情", ["锆精矿(国内)", "锆精矿(越南)", "高级硅酸锆"]),
    ("钛业行情", ["钛精矿", "金红石"]),
    ("小金属行情", ["锡锭", "氧化钽", "氧化铌"]),
]

DISPLAY_NAMES = {
    "氧化钕": "氧化钕(Nd2O3/TREO：99.0%-99.9%)",
	"氧化镨钕": "氧化镨钕（Pr6O11+Nd2O3/TREO≥99%、Nd2O3/TREO≥75%）",
	"氧化铽": "氧化铽（Tb4O7/TREO：99.99%）",
	"氧化镝": "氧化镝（Dy2O3/TREO：99.5%-99.9%）",
	"氧化钇": "氧化钇（Y2O3/TREO：99.99%－99.999%）",
	"独居石": "独居石(REO=54,Pr+Nd=22.5,Tb=0.11,Dy=0.58)",
    "高级硅酸锆": "广东高级硅酸锆((Zr+Hf)O2≥64.5%,D50=1.0μm)",
    "金红石": "金红石型钛白粉（氯化法）",
    "锆精矿(国内)": "锆精矿(国内ZrO2>65%,TiO2<0.15%,Fe2O3<0.1%)",
    "锆精矿(越南)": "锆精矿(越南ZrO2>65%,TiO2<0.15%,Fe2O3<0.1%)",
    "氧化钽": "氧化钽(Ta2O5≥99% 工业级)",
    "氧化铌": "氧化铌(Nb2O5≥99.5% 冶金级)",
    "钛精矿": "钛精矿(TiO2>49-50%,Fe2O3≤8%)",
}

def get_display_name(name):
    return DISPLAY_NAMES.get(name, name)

def load_data(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Sheet1"]
    headers = [c.value for c in ws[2]]
    rows = []
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row):
        vals = [c.value for c in row]
        if vals[0] is None or not isinstance(vals[0], datetime):
            continue
        rows.append(vals)
    return headers, rows

def fmt_wan(val):
    """Format value in wan yuan"""
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return "-"
    v = val / 10000
    if v >= 10000:
        return f"{v/10000:.2f}亿"
    if v >= 1:
        return f"{v:.2f}万"
    return f"{v:.4f}万"

def fmt_pct(val):
    if val is None:
        return "-"
    sign = "+" if val > 0 else ""
    return f"{sign}{val:.2f}%"

def pct_class(val):
    if val is None:
        return ""
    return "up" if val > 0 else ("down" if val < 0 else "flat")

def compute_changes(headers, rows):
    products = headers[1:]
    if not rows:
        return products, [], [], [], {}, []

    latest = rows[-1]
    prev = rows[-2] if len(rows) >= 2 else None

    # Weekly: latest vs previous row
    weekly = []
    for i, name in enumerate(products, 1):
        if prev and prev[i] is not None and latest[i] is not None and prev[i] != 0:
            w = (latest[i] - prev[i]) / prev[i] * 100
        else:
            w = None
        weekly.append(w)

    # Monthly: average by month
    month_groups = defaultdict(list)
    for r in rows:
        d = r[0]
        key = (d.year, d.month)
        for i, name in enumerate(products, 1):
            if r[i] is not None:
                month_groups[(key, name)].append(r[i])

    month_avgs = {}
    for (key, name), vals in month_groups.items():
        month_avgs[(key, name)] = sum(vals) / len(vals)

    sorted_months = sorted(set(k[0] for k in month_groups.keys()))
    cur_month = sorted_months[-1] if sorted_months else None
    prev_month = sorted_months[-2] if len(sorted_months) >= 2 else None

    monthly = []
    for name in products:
        if cur_month and prev_month and (cur_month, name) in month_avgs and (prev_month, name) in month_avgs:
            cur_avg = month_avgs[(cur_month, name)]
            prev_avg = month_avgs[(prev_month, name)]
            m = (cur_avg - prev_avg) / prev_avg * 100 if prev_avg != 0 else None
        else:
            m = None
        monthly.append(m)

    # Yearly: latest vs earliest this year
    this_year = latest[0].year
    year_rows = [r for r in rows if r[0].year == this_year]
    earliest = year_rows[0] if year_rows else latest

    yearly = []
    for i, name in enumerate(products, 1):
        if earliest[i] is not None and latest[i] is not None and earliest[i] != 0:
            y = (latest[i] - earliest[i]) / earliest[i] * 100
        else:
            y = None
        yearly.append(y)

    # Trend data for charts
    trend_data = {}
    for i, name in enumerate(products, 1):
        points = []
        for r in rows:
            if r[i] is not None:
                points.append({"date": r[0].strftime("%Y-%m-%d"), "price": r[i]})
        trend_data[name] = points

    return products, weekly, monthly, yearly, trend_data, rows
def build_html(products, latest_prices, weekly, monthly, yearly, trend_data, rows):
    date_range = f"{rows[0][0].strftime('%Y-%m-%d')} ~ {rows[-1][0].strftime('%Y-%m-%d')}"
    latest_date = rows[-1][0].strftime("%Y-%m-%d")
    total = len(products)
    valid_weekly = [w for w in weekly if w is not None]
    valid_monthly = [m for m in monthly if m is not None]
    valid_yearly = [y for y in yearly if y is not None]
    avg_w = sum(valid_weekly) / max(1, len(valid_weekly))
    avg_m = sum(valid_monthly) / max(1, len(valid_monthly))
    avg_y = sum(valid_yearly) / max(1, len(valid_yearly))
    rise_count = sum(1 for w in valid_weekly if w > 0)
    fall_count = sum(1 for w in valid_weekly if w < 0)
    flat_count = sum(1 for w in valid_weekly if w == 0)

    idx_map = {name: i for i, name in enumerate(products)}
    weekly_map = {name: weekly[i] for i, name in enumerate(products)}

    rows_html = ""
    for cat_name, cat_products in CATEGORIES:
        rows_html += f'<tr class="cat-header" data-category="{cat_name}"><td colspan="5">{cat_name}</td></tr>'
        for name in cat_products:
            if name not in idx_map:
                continue
            i = idx_map[name]
            display = get_display_name(name)
            lp = latest_prices[i]
            w = weekly[i]
            m = monthly[i]
            y = yearly[i]
            rows_html += f"""<tr class="product-row" data-category="{cat_name}" data-name="{display.lower()}">
            <td class="product-name">{display}</td>
            <td class="price">{fmt_wan(lp)}</td>
            <td class="{pct_class(w)}">{fmt_pct(w)}</td>
            <td class="{pct_class(m)}">{fmt_pct(m)}</td>
            <td class="{pct_class(y)}">{fmt_pct(y)}</td>
        </tr>"""

    options = ""
    category_options = '<option value="all">全部行业</option>'
    for cat_name, cat_products in CATEGORIES:
        category_options += f'<option value="{cat_name}">{cat_name}</option>'
        options += f'<optgroup label="{cat_name}">'
        for name in cat_products:
            if name not in idx_map:
                continue
            display = get_display_name(name)
            options += f'<option value="{name}">{display}</option>'
        options += '</optgroup>'

    category_cards = ""
    for cat_index, (cat_name, cat_products) in enumerate(CATEGORIES):
        indices = [idx_map[name] for name in cat_products if name in idx_map]
        cat_changes = [weekly[i] for i in indices if weekly[i] is not None]
        cat_avg = sum(cat_changes) / max(1, len(cat_changes))
        cat_rise = sum(1 for value in cat_changes if value > 0)
        cat_fall = sum(1 for value in cat_changes if value < 0)
        leader = max(indices, key=lambda i: abs(weekly[i] or 0)) if indices else None
        leader_text = (
            f"{products[leader]} {fmt_pct(weekly[leader])}"
            if leader is not None and weekly[leader] is not None
            else "暂无变动"
        )
        category_cards += f"""
        <article class="sector-card sector-{cat_index + 1}">
            <div class="sector-head">
                <div>
                    <div class="sector-name">{cat_name}</div>
                    <div class="sector-count">{len(indices)} 个品种</div>
                </div>
                <div class="sector-change {pct_class(cat_avg)}">{fmt_pct(cat_avg)}</div>
            </div>
            <div class="sector-breadth">
                <span class="up">上涨 {cat_rise}</span>
                <span class="down">下跌 {cat_fall}</span>
            </div>
            <div class="sector-leader">波动关注：{leader_text}</div>
        </article>"""

    ranked = [(i, value) for i, value in enumerate(weekly) if value is not None]
    gainers = sorted([item for item in ranked if item[1] > 0], key=lambda item: item[1], reverse=True)[:5]
    losers = sorted([item for item in ranked if item[1] < 0], key=lambda item: item[1])[:5]
    max_abs_change = max([abs(value) for _, value in ranked] or [1])

    def build_rank(items, empty_text):
        if not items:
            return f'<div class="rank-empty">{empty_text}</div>'
        result = ""
        for i, value in items:
            width = max(4, abs(value) / max_abs_change * 100) if value != 0 else 4
            result += f"""
            <div class="rank-item">
                <div class="rank-line">
                    <span class="rank-name">{products[i]}</span>
                    <span class="{pct_class(value)}">{fmt_pct(value)}</span>
                </div>
                <div class="rank-track"><span class="{pct_class(value)}" style="width:{width:.1f}%"></span></div>
            </div>"""
        return result

    gainers_html = build_rank(gainers, "本期暂无上涨品种")
    losers_html = build_rank(losers, "本期暂无下跌品种")
    focus_product = products[max(ranked, key=lambda item: abs(item[1]))[0]] if ranked else products[0]
    market_status = "偏强运行" if avg_w > 0 else ("偏弱运行" if avg_w < 0 else "横盘整理")

    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>矿产行情看板</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
html, body {{ width:100%; min-width:0; overflow-x:hidden; }}
body {{ font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif; background:#f2f5f7; color:#24313d; }}
button, input, select {{ font:inherit; }}
.topbar {{ background:#17232d; color:#fff; border-bottom:3px solid #c69a45; }}
.topbar-inner {{ width:100%; max-width:1440px; margin:auto; padding:20px 28px; display:flex; justify-content:space-between; align-items:center; gap:24px; }}
.brand-kicker {{ color:#c9d2d9; font-size:11px; margin-bottom:4px; }}
.brand h1 {{ font-size:23px; font-weight:700; }}
.market-state {{ display:flex; align-items:center; gap:12px; text-align:right; }}
.state-dot {{ width:9px; height:9px; border-radius:50%; background:#c69a45; box-shadow:0 0 0 4px rgba(198,154,69,.15); }}
.state-label {{ font-size:12px; color:#aebbc5; }}
.state-value {{ margin-top:2px; font-size:15px; font-weight:700; }}
.meta-strip {{ background:#fff; border-bottom:1px solid #dfe5e8; }}
.meta-inner {{ max-width:1440px; margin:auto; padding:10px 28px; display:flex; flex-wrap:wrap; gap:8px 24px; font-size:12px; color:#65737e; }}
.meta-inner strong {{ color:#33424e; font-weight:600; }}
.container {{ width:100%; max-width:1440px; min-width:0; margin:0 auto; padding:24px 28px 30px; }}
.section-head {{ display:flex; align-items:flex-end; justify-content:space-between; gap:18px; margin:0 0 12px; }}
.section-head h2 {{ font-size:16px; color:#263744; }}
.section-head p {{ font-size:12px; color:#7b8790; }}
.kpi-grid {{ display:grid; grid-template-columns:repeat(5,minmax(0,1fr)); background:#fff; border:1px solid #dfe5e8; border-radius:6px; margin-bottom:24px; }}
.kpi {{ min-width:0; padding:18px 20px; border-right:1px solid #e6ebee; }}
.kpi:last-child {{ border-right:0; }}
.kpi-label {{ font-size:12px; color:#75828c; margin-bottom:7px; }}
.kpi-value {{ font-size:25px; line-height:1.1; font-weight:750; font-variant-numeric:tabular-nums; }}
.kpi-note {{ margin-top:7px; color:#8a969f; font-size:11px; }}
.up {{ color:#c74444 !important; font-weight:650; }}
.down {{ color:#208060 !important; font-weight:650; }}
.flat {{ color:#7c8790 !important; font-weight:650; }}
.sector-grid {{ display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:14px; margin-bottom:24px; }}
.sector-card {{ position:relative; overflow:hidden; background:#fff; border:1px solid #dfe5e8; border-radius:6px; padding:17px 18px 15px; }}
.sector-card::before {{ content:""; position:absolute; inset:0 auto 0 0; width:4px; background:#407ea5; }}
.sector-2::before {{ background:#b0833c; }}
.sector-3::before {{ background:#657b67; }}
.sector-4::before {{ background:#785f8b; }}
.sector-head {{ display:flex; justify-content:space-between; align-items:flex-start; gap:12px; }}
.sector-name {{ font-size:14px; font-weight:700; color:#2d3b46; }}
.sector-count {{ margin-top:4px; font-size:11px; color:#89949c; }}
.sector-change {{ font-size:20px; font-variant-numeric:tabular-nums; }}
.sector-breadth {{ display:flex; gap:16px; margin-top:18px; font-size:12px; }}
.sector-leader {{ margin-top:10px; padding-top:10px; border-top:1px solid #edf0f2; color:#687781; font-size:11px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
.analytics-grid {{ display:grid; grid-template-columns:minmax(0,2fr) minmax(290px,1fr); gap:16px; margin-bottom:24px; }}
.panel {{ background:#fff; border:1px solid #dfe5e8; border-radius:6px; }}
.panel-head {{ min-height:58px; padding:13px 16px; border-bottom:1px solid #e7ebee; display:flex; justify-content:space-between; align-items:center; gap:14px; }}
.panel-title {{ font-size:14px; font-weight:700; color:#2d3b46; }}
.panel-subtitle {{ margin-top:3px; font-size:11px; color:#87929a; }}
.chart-controls {{ display:flex; align-items:center; justify-content:flex-end; gap:10px; min-width:0; }}
select, .search-input {{ height:34px; border:1px solid #ccd5da; border-radius:4px; color:#33424d; background:#fff; outline:none; }}
select {{ max-width:260px; padding:0 30px 0 10px; }}
select:focus, .search-input:focus {{ border-color:#5286a8; box-shadow:0 0 0 2px rgba(82,134,168,.12); }}
.range-group {{ display:flex; border:1px solid #ccd5da; border-radius:4px; overflow:hidden; flex-shrink:0; }}
.range-btn {{ height:32px; padding:0 10px; border:0; border-right:1px solid #d9e0e4; color:#687680; background:#fff; cursor:pointer; font-size:12px; }}
.range-btn:last-child {{ border-right:0; }}
.range-btn.active {{ color:#fff; background:#3b647d; }}
.chart-wrapper {{ position:relative; height:360px; padding:14px 16px 16px; }}
.ranking-body {{ display:grid; grid-template-columns:1fr 1fr; }}
.rank-column {{ min-width:0; padding:14px 15px 16px; }}
.rank-column + .rank-column {{ border-left:1px solid #edf0f2; }}
.rank-heading {{ display:flex; align-items:center; justify-content:space-between; margin-bottom:13px; font-size:12px; font-weight:700; color:#52616c; }}
.rank-empty {{ padding:28px 4px; color:#919ca4; font-size:12px; text-align:center; }}
.rank-item {{ margin-bottom:14px; }}
.rank-item:last-child {{ margin-bottom:0; }}
.rank-line {{ display:flex; justify-content:space-between; gap:10px; margin-bottom:6px; font-size:12px; }}
.rank-name {{ overflow:hidden; text-overflow:ellipsis; white-space:nowrap; color:#3c4b56; }}
.rank-track {{ height:3px; background:#edf1f3; overflow:hidden; }}
.rank-track span {{ display:block; height:100%; background:#7c8790; }}
.rank-track span.up {{ background:#c74444; }}
.rank-track span.down {{ background:#208060; }}
.table-panel {{ background:#fff; border:1px solid #dfe5e8; border-radius:6px; overflow:hidden; }}
.table-tools {{ display:flex; gap:8px; }}
.search-input {{ width:180px; padding:0 10px; }}
.table-wrap {{ width:100%; max-width:100%; overflow-x:auto; -webkit-overflow-scrolling:touch; }}
table {{ width:100%; min-width:760px; border-collapse:collapse; }}
th {{ background:#f2f5f6; padding:11px 14px; text-align:left; font-size:11px; color:#62717c; font-weight:700; border-bottom:1px solid #dfe5e8; }}
td {{ padding:11px 14px; font-size:12px; border-top:1px solid #edf0f2; }}
tbody tr.product-row:hover td {{ background:#f7fafb; }}
.price {{ font-weight:700; color:#293945; font-variant-numeric:tabular-nums; }}
.cat-header td {{ background:#e9eff2; font-weight:700; font-size:12px; color:#385568; padding:8px 14px; border-top:1px solid #d8e0e5; }}
.product-name {{ padding-left:24px; max-width:500px; }}
.footer {{ text-align:center; padding:20px; font-size:11px; color:#8a969f; }}
[hidden] {{ display:none !important; }}
@media (max-width:1050px) {{
    .kpi-grid {{ grid-template-columns:repeat(3,minmax(0,1fr)); }}
    .kpi {{ border-bottom:1px solid #e6ebee; }}
    .kpi:nth-child(3) {{ border-right:0; }}
    .kpi:nth-child(n+4) {{ border-bottom:0; }}
    .sector-grid {{ grid-template-columns:repeat(2,minmax(0,1fr)); }}
    .analytics-grid {{ grid-template-columns:1fr; }}
}}
@media (max-width:640px) {{
    .topbar-inner {{ display:block; padding:16px; }}
    .brand h1 {{ font-size:18px; white-space:nowrap; }}
    .market-state {{ margin-top:11px; justify-content:flex-start; text-align:left; }}
    .state-label {{ display:block; }}
    .meta-inner {{ padding:9px 16px; gap:5px 14px; }}
    .container {{ padding:16px 12px 24px; }}
    .section-head {{ align-items:flex-start; }}
    .section-head p {{ display:none; }}
    .kpi-grid {{ grid-template-columns:repeat(2,minmax(0,1fr)); }}
    .kpi {{ padding:14px 13px; border-right:1px solid #e6ebee; border-bottom:1px solid #e6ebee; }}
    .kpi:nth-child(even) {{ border-right:0; }}
    .kpi:nth-child(3) {{ border-right:1px solid #e6ebee; }}
    .kpi:nth-child(n+4) {{ border-bottom:1px solid #e6ebee; }}
    .kpi:last-child {{ grid-column:1 / -1; border-bottom:0; }}
    .kpi-value {{ font-size:21px; }}
    .sector-grid {{ grid-template-columns:1fr; gap:9px; }}
    .analytics-grid {{ gap:12px; }}
    .panel-head {{ align-items:flex-start; flex-direction:column; }}
    .chart-controls {{ width:100%; flex-direction:column; align-items:stretch; }}
    .chart-controls select {{ width:100%; max-width:none; min-width:0; }}
    .range-group {{ width:100%; }}
    .range-btn {{ flex:1; padding:0 8px; }}
    .chart-wrapper {{ height:290px; padding:10px; }}
    .ranking-body {{ grid-template-columns:1fr; }}
    .rank-column + .rank-column {{ border-left:0; border-top:1px solid #edf0f2; }}
    .table-panel .panel-head {{ flex-direction:column; }}
    .table-tools {{ width:100%; }}
    .table-tools select, .search-input {{ width:50%; min-width:0; }}
}}
@media print {{
    body {{ background:#fff; }}
    .topbar {{ background:#fff; color:#17232d; border-bottom:2px solid #17232d; }}
    .market-state, .chart-controls, .table-tools {{ display:none; }}
    .container {{ max-width:none; padding:16px 0; }}
    .panel, .sector-card, .kpi-grid, .table-panel {{ break-inside:avoid; box-shadow:none; }}
}}
</style>
</head>
<body>
<header class="topbar">
    <div class="topbar-inner">
        <div class="brand">
            <div class="brand-kicker">MINERAL MARKET INTELLIGENCE</div>
            <h1>东亚 &amp; 凯博矿产行情看板</h1>
        </div>
        <div class="market-state">
            <span class="state-dot"></span>
            <div>
                <div class="state-label">本周市场状态</div>
                <div class="state-value">{market_status}</div>
            </div>
        </div>
    </div>
</header>
<div class="meta-strip">
    <div class="meta-inner">
        <span>最新数据 <strong>{latest_date}</strong></span>
        <span>统计周期 <strong>{date_range}</strong></span>
        <span>覆盖品种 <strong>{total} 个</strong></span>
        <span>生成时间 <strong>{datetime.now().strftime("%Y-%m-%d %H:%M")}</strong></span>
    </div>
</div>
<div class="container">
    <section class="kpi-grid" aria-label="市场核心指标">
        <div class="kpi"><div class="kpi-label">周度平均涨跌</div><div class="kpi-value {pct_class(avg_w)}">{fmt_pct(avg_w)}</div><div class="kpi-note">市场短期方向</div></div>
        <div class="kpi"><div class="kpi-label">上涨品种</div><div class="kpi-value up">{rise_count}</div><div class="kpi-note">占有效报价 {rise_count / max(1, len(valid_weekly)) * 100:.0f}%</div></div>
        <div class="kpi"><div class="kpi-label">下跌品种</div><div class="kpi-value down">{fall_count}</div><div class="kpi-note">持平 {flat_count} 个</div></div>
        <div class="kpi"><div class="kpi-label">月度平均涨跌</div><div class="kpi-value {pct_class(avg_m)}">{fmt_pct(avg_m)}</div><div class="kpi-note">月均价环比</div></div>
        <div class="kpi"><div class="kpi-label">年度平均涨跌</div><div class="kpi-value {pct_class(avg_y)}">{fmt_pct(avg_y)}</div><div class="kpi-note">年初至今</div></div>
    </section>

    <div class="section-head">
        <div><h2>行业行情概览</h2></div>
        <p>周度平均涨跌与行业内部市场宽度</p>
    </div>
    <section class="sector-grid">{category_cards}</section>

    <section class="analytics-grid">
        <div class="panel">
            <div class="panel-head">
                <div>
                    <div class="panel-title">重点品种价格趋势</div>
                    <div class="panel-subtitle" id="chartProductName">{get_display_name(focus_product)}</div>
                </div>
                <div class="chart-controls">
                    <select id="productSelect" aria-label="选择品种">{options}</select>
                    <div class="range-group" aria-label="时间范围">
                        <button class="range-btn" type="button" data-range="30">30天</button>
                        <button class="range-btn" type="button" data-range="90">90天</button>
                        <button class="range-btn active" type="button" data-range="0">全部</button>
                    </div>
                </div>
            </div>
            <div class="chart-wrapper"><canvas id="trendChart"></canvas></div>
        </div>
        <div class="panel">
            <div class="panel-head">
                <div>
                    <div class="panel-title">周度涨跌排行</div>
                    <div class="panel-subtitle">按最新一期涨跌幅排序</div>
                </div>
            </div>
            <div class="ranking-body">
                <div class="rank-column">
                    <div class="rank-heading"><span>领涨品种</span><span class="up">涨幅</span></div>
                    {gainers_html}
                </div>
                <div class="rank-column">
                    <div class="rank-heading"><span>领跌品种</span><span class="down">跌幅</span></div>
                    {losers_html}
                </div>
            </div>
        </div>
    </section>

    <section class="table-panel">
        <div class="panel-head">
            <div>
                <div class="panel-title">完整行情明细</div>
                <div class="panel-subtitle">最新报价及周度、月度、年度变化</div>
            </div>
            <div class="table-tools">
                <select id="categoryFilter" aria-label="筛选行业">{category_options}</select>
                <input class="search-input" id="productSearch" type="search" placeholder="搜索品种" aria-label="搜索品种">
            </div>
        </div>
        <div class="table-wrap">
            <table>
                <thead><tr><th>品种</th><th>最新报价（万元）</th><th>周度涨跌</th><th>月度涨跌</th><th>年度涨跌</th></tr></thead>
                <tbody>{rows_html}</tbody>
            </table>
        </div>
    </section>
    <div class="footer">数据来源：瑞道金属网、上海有色网、生意社</div>
</div>
<script>
var trendData = {json.dumps(trend_data, ensure_ascii=False)};
var weeklyChanges = {json.dumps(weekly_map, ensure_ascii=False)};
var chart = null;
var activeRange = 0;

function fmtWan(v) {{
    var w = v / 10000;
    if (w >= 10000) return (w/10000).toFixed(2) + '\u4ebf';
    if (w >= 1) return w.toFixed(2) + '\u4e07';
    return w.toFixed(4) + '\u4e07';
}}

function drawChart(name) {{
    if (chart) chart.destroy();
    var points = trendData[name] || [];
    if (activeRange && points.length) {{
        var lastDate = new Date(points[points.length - 1].date);
        var cutoff = new Date(lastDate);
        cutoff.setDate(cutoff.getDate() - activeRange);
        points = points.filter(function(p) {{ return new Date(p.date) >= cutoff; }});
    }}
    var labels = points.map(function(p) {{ return p.date; }});
    var prices = points.map(function(p) {{ return p.price; }});
    var isDown = (weeklyChanges[name] || 0) < 0;
    var lineColor = isDown ? '#208060' : '#c74444';
    var fillColor = isDown ? 'rgba(32,128,96,0.08)' : 'rgba(199,68,68,0.08)';
    var ctx = document.getElementById('trendChart').getContext('2d');
    chart = new Chart(ctx, {{
        type: 'line',
        data: {{
            labels: labels,
            datasets: [{{
                label: name + ' (\u4e07\u5143)',
                data: prices,
                borderColor: lineColor,
                backgroundColor: fillColor,
                fill: true,
                tension: 0.25,
                borderWidth: 2,
                pointRadius: 3,
                pointHoverRadius: 6,
                pointBackgroundColor: lineColor
            }}]
        }},
        options: {{
            responsive: true,
            maintainAspectRatio: false,
            plugins: {{
                legend: {{ display: false }},
                tooltip: {{
                    backgroundColor: '#17232d',
                    padding: 10,
                    callbacks: {{
                        label: function(ctx) {{ return name + ': ' + fmtWan(ctx.parsed.y); }}
                    }}
                }}
            }},
            scales: {{
                x: {{
                    grid: {{ display: false }},
                    ticks: {{ color: '#7d8992', maxRotation: 0, autoSkip: true, maxTicksLimit: 7 }}
                }},
                y: {{
                    ticks: {{ color: '#7d8992', callback: function(v) {{ return fmtWan(v); }} }},
                    grid: {{ color: '#edf0f2' }}
                }}
            }}
        }}
    }});
    document.getElementById('chartProductName').textContent =
        document.getElementById('productSelect').selectedOptions[0].textContent;
}}

document.getElementById('productSelect').addEventListener('change', function(e) {{ drawChart(e.target.value); }});
document.querySelectorAll('.range-btn').forEach(function(button) {{
    button.addEventListener('click', function() {{
        document.querySelectorAll('.range-btn').forEach(function(item) {{ item.classList.remove('active'); }});
        button.classList.add('active');
        activeRange = Number(button.dataset.range);
        drawChart(document.getElementById('productSelect').value);
    }});
}});

function applyTableFilters() {{
    var category = document.getElementById('categoryFilter').value;
    var keyword = document.getElementById('productSearch').value.trim().toLowerCase();
    var rows = Array.from(document.querySelectorAll('.product-row'));
    rows.forEach(function(row) {{
        var categoryMatch = category === 'all' || row.dataset.category === category;
        var keywordMatch = !keyword || row.dataset.name.indexOf(keyword) !== -1;
        row.hidden = !(categoryMatch && keywordMatch);
    }});
    document.querySelectorAll('.cat-header').forEach(function(header) {{
        var visibleChild = rows.some(function(row) {{
            return row.dataset.category === header.dataset.category && !row.hidden;
        }});
        header.hidden = !visibleChild;
    }});
}}

document.getElementById('categoryFilter').addEventListener('change', applyTableFilters);
document.getElementById('productSearch').addEventListener('input', applyTableFilters);
document.getElementById('productSelect').value = '{focus_product}';
drawChart('{focus_product}');
</script>
</body>
</html>"""
    return html

def main():
    print("读取 Excel...")
    headers, rows = load_data(EXCEL)
    print(f"共 {len(rows)} 条记录, {len(headers)-1} 个品种")

    print("计算涨跌幅...")
    products, weekly, monthly, yearly, trend_data, rows2 = compute_changes(headers, rows)

    latest_prices = [rows[-1][i] if rows[-1][i] is not None else None for i in range(1, len(headers))]

    print("生成看板...")
    html = build_html(products, latest_prices, weekly, monthly, yearly, trend_data, rows2)

    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
    with open(OUTPUT, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"完成！看板已生成: {OUTPUT}")

if __name__ == "__main__":
    main()
