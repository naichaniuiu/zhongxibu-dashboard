#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
完全对齐原模板弹窗下钻风格的数据看板生成器
下钻逻辑：部门行点击 -> 弹窗销售员列表 -> 点击销售员名 -> 弹窗切换客户明细（带返回按钮）
"""
import json, sys, os
from datetime import datetime, timedelta

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

with open(os.path.join(BASE_DIR, 'dashboard_data.json'), encoding='utf-8') as f:
    data = json.load(f)
with open(os.path.join(BASE_DIR, 'customer_detail.json'), encoding='utf-8') as f:
    cust_detail = json.load(f)  # {dept: {seller: [{customer, perf, collect, total_debt, d30, d30_90, d90_180, d180, cycle, max_days, orders}]}}

# 数据截止日期为生成日的前一天，统计基日=数据截止日，生成于=实际生成日期
raw_data_date = data.get('data_date', data['today'])
data_date_obj = datetime.strptime(raw_data_date, '%Y-%m-%d')
stat_date_obj = data_date_obj - timedelta(days=1)
gen_date_obj = datetime.now()

data_date = data_date_obj.strftime('%Y-%m-%d')
stat_date = stat_date_obj.strftime('%Y-%m-%d')
gen_date = gen_date_obj.strftime('%Y-%m-%d')

dept_list = data['dept_data']
kpi = data['kpi']
debt_kpi = data['debt_kpi']
cycle_kpi = data['cycle_kpi']
total = {
    'data_date': data_date,
    'stat_date': stat_date,
    'gen_date': gen_date,
    'v26': kpi['total_perf26'],
    'v25': kpi['total_perf25'] if kpi['total_perf25'] is not None else None,
    'target': kpi['total_target'],
    'completion': kpi['total_completion'],
    'yoy': kpi['total_yoy'] if kpi['total_yoy'] is not None else None,
    'sales': kpi['total_active_sellers'],
    'total_debt': debt_kpi['total'],
    'overdue': debt_kpi.get('overdue', debt_kpi['d30_90'] + debt_kpi['d90_180'] + debt_kpi['d180']),
    'd30': debt_kpi['d30'],
    'd30_90': debt_kpi['d30_90'],
    'd90_180': debt_kpi['d90_180'],
    'd180': debt_kpi['d180'],
    'collect': sum(d['collect'] for d in dept_list),
    'avg_cycle': cycle_kpi['avg_cycle'],
}

# 直接使用已有的销售员数据（dashboard_data中已有）
sales_detail = data['sales_detail_data']
sales_cycle_raw = data['sales_cycle_detail']

# salesCycleData 中 debt_amt/rec_amt 是元（原始数据），保留
sales_cycle = {}
for dept_name, sellers in sales_cycle_raw.items():
    sales_cycle[dept_name] = [
        {'name': s['name'], 'debt_amt': s.get('debt_amt', 0), 'rec_amt': s.get('rec_amt', 0), 'cycle': round(s.get('cycle', 0), 1)}
        for s in sellers
    ]

# 构建 deptCustPerfData: {dept: [{name:seller, customers:[{name:cust, perf, collect, orders}]}]}
# 格式用于业绩tab下点击销售员后展示客户
dept_cust_perf = {}
dept_cust_debt = {}
cust_cycle_data = {}

for dept_name, sellers_map in cust_detail.items():
    dept_cust_perf[dept_name] = []
    dept_cust_debt[dept_name] = []
    cust_cycle_data[dept_name] = {}
    
    for seller_name, custs in sellers_map.items():
        # 业绩客户列表
        perf_custs = [{'name': c['customer'],
                       'perf': round(c.get('perf', 0), 2),
                       'collect': round(c.get('collect', 0), 2),
                       'orders': c.get('orders', 0)} for c in custs]
        dept_cust_perf[dept_name].append({'name': seller_name, 'customers': perf_custs})
        
        # 欠款客户列表
        debt_custs = [{'name': c['customer'],
                       'total_debt': round(c.get('total_debt', 0), 2),
                       'd30': round(c.get('d30', 0), 2),
                       'd30_90': round(c.get('d30_90', 0), 2),
                       'd90_180': round(c.get('d90_180', 0), 2),
                       'd180': round(c.get('d180', 0), 2),
                       'max_days': c.get('max_days', 0)} for c in custs if c.get('total_debt', 0) > 0]
        dept_cust_debt[dept_name].append({'name': seller_name, 'customers': debt_custs})
        
        # 回款周期客户列表（debt_amt, rec_amt用元）
        cycle_custs = [{'name': c['customer'],
                        'debt_amt': round(c.get('total_debt', 0) * 10000, 2),
                        'rec_amt': round(c.get('collect', 0) * 10000, 2),
                        'cycle': round(c.get('cycle', 0), 1)} for c in custs]
        cust_cycle_data[dept_name][seller_name] = cycle_custs

# 生成高风险客户（90天以上欠款）
risky_customers = []
for dept_name, sellers_map in cust_detail.items():
    for seller_name, custs in sellers_map.items():
        for c in custs:
            d90 = c.get('d90_180', 0) + c.get('d180', 0)
            if d90 > 0:
                risky_customers.append({
                    'customer': c['customer'],
                    'dept': dept_name,
                    'seller': seller_name,
                    'd90_180': round(c.get('d90_180', 0), 2),
                    'd180': round(c.get('d180', 0), 2),
                    'total_90plus': round(d90, 2),
                    'max_days': c.get('max_days', 0),
                })
risky_customers.sort(key=lambda x: x['total_90plus'], reverse=True)
risky_top15 = risky_customers[:15]

# ========== 生成部门表格行 ==========
def get_dept_status(completion):
    if completion >= 60:
        return 'badge-good', '较好'
    elif completion >= 30:
        return 'badge-warning', '待关注'
    else:
        return 'badge-down', '严重下滑'

def fmt_yoy(v):
    if v is None:
        return '<span class="trend-neutral">-</span>'
    if v > 0:
        return f'<span class="trend-up">▲ +{v:.1f}%</span>'
    elif v < -50:
        return f'<span class="trend-down">▼ {v:.1f}%</span>'
    else:
        return f'<span class="trend-down">▼ {v:.1f}%</span>'

def get_debt_status(d):
    total = d['total_debt']
    risky = d['d90_180'] + d['d180']
    if risky > 50 or (total > 0 and risky/total > 0.3):
        return 'badge-down', '高风险'
    elif risky > 10 or total > 100:
        return 'badge-warning', '关注'
    else:
        return 'badge-good', '较好'

def get_cycle_status(cycle):
    if cycle > 90:
        return 'badge-down', '需关注', 'negative'
    elif cycle > 60:
        return 'badge-warning', '一般', 'warning'
    else:
        return 'badge-good', '良好', 'highlight'

# 业绩表格行
perf_rows = ''
for d in dept_list:
    badge_cls, badge_text = get_dept_status(d['completion'])
    yoy_html = fmt_yoy(d['yoy'])
    v25_cell = f'{d["v25"]:.2f}' if d['v25'] is not None else '-'
    perf_rows += f'''<tr onclick="showSalesDetail('{d["dept"]}','perf')" style="cursor:pointer;" title="点击查看销售员业绩明细">
        <td>🏢 {d["dept"]}</td>
        <td class="highlight">{d["v26"]:.2f}</td>
        <td>{d["target"]:.1f}</td>
        <td>{d["completion"]:.1f}%</td>
        <td>{v25_cell}</td>
        <td>{yoy_html}</td>
        <td>{d["sales"]}</td>
        <td><span class="status-badge {badge_cls}">{badge_text}</span></td>
    </tr>'''

# 汇总行同比显示
yoy_total_str = f"▼ {total['yoy']:.1f}%" if total['yoy'] is not None else '-'
v25_total_str = f"{total['v25']:.2f}" if total['v25'] is not None else '-'

# 欠款表格行
debt_sorted = sorted(dept_list, key=lambda x: x['total_debt'], reverse=True)
debt_rows = ''
for d in debt_sorted:
    badge_cls, badge_text = get_debt_status(d)
    risky = d['d90_180'] + d['d180']
    d90_cls = 'warning' if risky < 50 else 'negative'
    d180_cls = 'negative' if d['d180'] > 0 else ''
    debt_rows += f'''<tr onclick="showSalesDetail('{d["dept"]}','debt')" style="cursor:pointer;">
        <td>🏢 {d["dept"]}</td>
        <td>{d["d30"]:.2f}</td>
        <td>{d["d30_90"]:.2f}</td>
        <td class="warning">{d["d90_180"]:.2f}</td>
        <td class="negative">{d["d180"]:.2f}</td>
        <td class="negative">{d["total_debt"]:.2f}</td>
        <td><span class="status-badge {badge_cls}">{badge_text}</span></td>
    </tr>'''

# 回款周期表格行
cycle_sorted = sorted(dept_list, key=lambda x: x['cycle'], reverse=True)
cycle_rows = ''
for d in cycle_sorted:
    badge_cls, badge_text, val_cls = get_cycle_status(d['cycle'])
    cycle_rows += f'''<tr onclick="showSalesDetail('{d["dept"]}','cycle')" style="cursor:pointer;">
        <td>🏢 {d["dept"]}</td>
        <td>{d["total_debt"]:.2f}</td>
        <td>{d["collect"]:.2f}</td>
        <td class="{val_cls}">{d["cycle"]:.1f}</td>
        <td><span class="status-badge {badge_cls}">{badge_text}</span></td>
    </tr>'''

# 高风险客户行
risky_rows = ''
for i, r in enumerate(risky_top15):
    rank = i + 1
    total_90 = r['total_90plus']
    if total_90 > 50:
        risk_text = '极高风险'
        risk_cls = 'badge-down'
    elif total_90 > 20:
        risk_text = '高风险'
        risk_cls = 'badge-down'
    else:
        risk_text = '关注'
        risk_cls = 'badge-warning'
    risky_rows += f'<tr><td>{rank}</td><td style="text-align:left;">{r["customer"]}（{r["dept"]}）</td><td class="negative">{total_90:.2f}</td><td><span class="status-badge {risk_cls}">{risk_text}</span></td></tr>'

# 超90天部门列表
over90_depts = [d for d in dept_list if d['cycle'] > 90]

# 回款周期图表数据
chart_cycle_depts = [d['dept'] for d in cycle_sorted]
chart_cycle_vals = [d['cycle'] for d in cycle_sorted]
chart_cycle_colors = ['#ff4757' if v > 90 else '#ffa502' if v > 60 else '#00ff88' for v in chart_cycle_vals]

# 业绩图表数据
perf_depts = [d['dept'] for d in dept_list]
perf_v26 = [d['v26'] for d in dept_list]
perf_v25 = [d['v25'] if d['v25'] is not None else 0 for d in dept_list]
perf_pie = [(d['dept'], d['v26']) for d in sorted(dept_list, key=lambda x: x['v26'], reverse=True) if d['v26'] > 0]

# 欠款图表数据
debt_pie_labels = ['30天内', '30-90天', '90-180天', '180天以上']
debt_pie_vals = [total['d30'], total['d30_90'], total['d90_180'], total['d180']]
debt_bar_depts = [d['dept'] for d in debt_sorted[:8]]
debt_bar_vals = [d['total_debt'] for d in debt_sorted[:8]]

# JSON序列化数据
sales_detail_json = json.dumps(sales_detail, ensure_ascii=False)
sales_cycle_json = json.dumps(sales_cycle, ensure_ascii=False)
dept_cust_perf_json = json.dumps(dept_cust_perf, ensure_ascii=False)
dept_cust_debt_json = json.dumps(dept_cust_debt, ensure_ascii=False)
cust_cycle_json = json.dumps(cust_cycle_data, ensure_ascii=False)
dept_list_json = json.dumps([{
    'dept': d['dept'], 'v26': d['v26'], 'v25': d['v25'] if d['v25'] is not None else 0, 'yoy': d['yoy'] if d['yoy'] is not None else 0,
    'target': d['target'], 'completion': d['completion'], 'sales': d['sales'],
    'd30': d['d30'], 'd30_90': d['d30_90'], 'd90_180': d['d90_180'], 'd180': d['d180'],
    'total_debt': d['total_debt'], 'collect': d['collect'], 'cycle': d['cycle']
} for d in dept_list], ensure_ascii=False)

html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>中西部大区 26财年Q1 数据看板</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC', 'Microsoft YaHei', sans-serif; background: #0a0e27; color: #fff; line-height: 1.6; min-height: 100vh; }}
        .container {{ max-width: 1400px; margin: 0 auto; padding: 30px 20px; }}
        .header {{ text-align: center; margin-bottom: 40px; }}
        .header h1 {{ font-size: 2.2em; background: linear-gradient(135deg, #00d4ff 0%, #7b2ff7 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text; margin-bottom: 10px; }}
        .header p {{ color: #8892b0; font-size: 1.1em; }}

        .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-bottom: 30px; }}
        .kpi-card {{ background: linear-gradient(135deg, rgba(255,255,255,0.1) 0%, rgba(255,255,255,0.05) 100%); border-radius: 16px; padding: 24px; text-align: center; border: 1px solid rgba(255,255,255,0.1); transition: transform 0.3s, box-shadow 0.3s; }}
        .kpi-card:hover {{ transform: translateY(-5px); box-shadow: 0 20px 40px rgba(0,212,255,0.15); }}
        .kpi-icon {{ font-size: 2.5em; margin-bottom: 10px; }}
        .kpi-card h3 {{ color: #8892b0; font-size: 0.9em; font-weight: 500; margin-bottom: 8px; }}
        .kpi-card .value {{ font-size: 2.2em; font-weight: 700; color: #00d4ff; }}
        .kpi-card .sub {{ font-size: 0.85em; color: #8892b0; margin-top: 5px; }}

        .highlight {{ color: #00ff88 !important; }}
        .negative {{ color: #ff4757 !important; }}
        .warning {{ color: #ffa502 !important; }}

        .tab-nav {{ display: flex; gap: 10px; margin-bottom: 25px; padding: 8px; background: rgba(255,255,255,0.03); border-radius: 12px; border: 1px solid rgba(255,255,255,0.08); overflow-x: auto; scrollbar-width: none; }}
        .tab-nav::-webkit-scrollbar {{ display: none; }}
        .tab-btn {{ display: flex; align-items: center; gap: 8px; padding: 12px 24px; background: transparent; border: none; border-radius: 8px; color: #8892b0; font-size: 0.95em; font-weight: 500; cursor: pointer; transition: all 0.3s; white-space: nowrap; }}
        .tab-btn:hover {{ background: rgba(255,255,255,0.05); color: #ccd6f6; }}
        .tab-btn.active {{ background: linear-gradient(135deg, #00d4ff 0%, #7b2ff7 100%); color: #fff; box-shadow: 0 4px 15px rgba(0,212,255,0.3); }}

        .tab-content {{ display: none; }}
        .tab-content.active {{ display: block; animation: fadeIn 0.3s ease; }}
        @keyframes fadeIn {{ from {{ opacity: 0; transform: translateY(10px); }} to {{ opacity: 1; transform: translateY(0); }} }}

        .module {{ background: rgba(255,255,255,0.05); border-radius: 20px; padding: 30px; border: 1px solid rgba(255,255,255,0.1); margin-bottom: 25px; }}
        .module-title {{ font-size: 1.3em; margin-bottom: 10px; color: #00d4ff; display: flex; align-items: center; gap: 10px; }}
        .module-desc {{ color: #8892b0; margin-bottom: 20px; font-size: 0.9em; }}

        .charts-section {{ display: grid; grid-template-columns: 1fr 1fr; gap: 25px; margin-bottom: 30px; }}
        .chart-box {{ background: rgba(255,255,255,0.03); border-radius: 16px; padding: 20px; border: 1px solid rgba(255,255,255,0.08); }}
        .chart-box h3 {{ color: #00d4ff; margin-bottom: 15px; font-size: 1.1em; text-align: center; }}
        .chart-container {{ position: relative; height: 320px; }}

        .dept-table {{ width: 100%; border-collapse: collapse; font-size: 0.9em; }}
        .dept-table thead th {{ background: rgba(0,212,255,0.1); color: #00d4ff; padding: 12px 8px; text-align: center; font-weight: 500; border-bottom: 1px solid rgba(0,212,255,0.2); white-space: nowrap; }}
        .dept-table tbody tr {{ border-bottom: 1px solid rgba(255,255,255,0.05); transition: background 0.2s; cursor: pointer; }}
        .dept-table tbody tr:hover {{ background: rgba(255,255,255,0.05); }}
        .dept-table tbody td {{ padding: 11px 8px; text-align: center; color: #ccd6f6; }}
        .dept-table tbody td:first-child {{ text-align: left; padding-left: 12px; }}
        .dept-table tfoot tr {{ background: rgba(0,212,255,0.08); border-top: 2px solid rgba(0,212,255,0.3); }}
        .dept-table tfoot td {{ padding: 12px 8px; text-align: center; color: #00d4ff; font-weight: 600; }}
        .dept-table tfoot td:first-child {{ text-align: left; padding-left: 12px; }}

        .trend-down {{ color: #ff4757; font-weight: 600; }}
        .trend-up {{ color: #00ff88; font-weight: 600; }}
        .trend-neutral {{ color: #ffa502; }}
        .status-badge {{ padding: 3px 10px; border-radius: 12px; font-size: 0.8em; font-weight: 600; }}
        .badge-down {{ background: rgba(255,71,87,0.2); color: #ff4757; border: 1px solid rgba(255,71,87,0.3); }}
        .badge-warning {{ background: rgba(255,165,2,0.2); color: #ffa502; border: 1px solid rgba(255,165,2,0.3); }}
        .badge-new {{ background: rgba(0,255,136,0.2); color: #00ff88; border: 1px solid rgba(0,255,136,0.3); }}
        .badge-good {{ background: rgba(0,212,255,0.2); color: #00d4ff; border: 1px solid rgba(0,212,255,0.3); }}

        /* 下钻弹窗 */
        .modal-overlay {{ display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.7); z-index: 1000; }}
        .modal-overlay.active {{ display: flex; align-items: center; justify-content: center; }}
        .modal {{ background: #1a2040; border-radius: 20px; padding: 30px; max-width: 900px; width: 90%; max-height: 80vh; overflow-y: auto; border: 1px solid rgba(0,212,255,0.3); }}
        .modal-header {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; }}
        .modal-title {{ color: #00d4ff; font-size: 1.2em; font-weight: 600; }}
        .modal-close {{ background: rgba(255,255,255,0.1); border: none; color: #fff; width: 30px; height: 30px; border-radius: 50%; cursor: pointer; font-size: 1.2em; display: flex; align-items: center; justify-content: center; }}
        .modal-close:hover {{ background: rgba(255,71,87,0.3); }}

        @media (max-width: 768px) {{ .charts-section {{ grid-template-columns: 1fr; }} .kpi-grid {{ grid-template-columns: repeat(2, 1fr); }} .dept-table {{ font-size: 0.8em; }} }}
    </style>
</head>
<body>
<div class="container">
    <div class="header">
        <h1>中西部大区 26财年Q1 数据看板</h1>
        <p>数据截止：{total['data_date']} &nbsp;|&nbsp; 统计基日：{total['stat_date']} &nbsp;|&nbsp; 生成于：{total['gen_date']}</p>
    </div>

    <!-- 核心KPI -->
    <div class="kpi-grid">
        <div class="kpi-card">
            <div class="kpi-icon">💰</div>
            <h3>26Q1 实际业绩</h3>
            <div class="value">{total['v26']:.2f}<span style="font-size:0.5em;">万</span></div>
            <div class="sub">目标：{total['target']:.1f}万</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-icon">🎯</div>
            <h3>Q1目标完成率</h3>
            <div class="value negative">{total['completion']:.1f}%</div>
            <div class="sub">距目标还差 {total['target']-total['v26']:.2f}万</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-icon">📉</div>
            <h3>同比25Q1</h3>
            {'<div class="value negative">'+str(total['yoy'])+'%</div><div class="sub">25Q1：'+str(total['v25'])+'万</div>' if total['yoy'] is not None else '<div class="value" style="color:#8892b0;">-</div><div class="sub">25Q1数据缺失</div>'}
        </div>
        <div class="kpi-card">
            <div class="kpi-icon">👥</div>
            <h3>在职销售员总数</h3>
            <div class="value" style="color:#ffa502;">{total['sales']}</div>
            <div class="sub">人</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-icon">⚠️</div>
            <h3>逾期欠款总额</h3>
            <div class="value negative">{total['overdue']:.2f}<span style="font-size:0.5em;">万</span></div>
            <div class="sub">总欠款：{total['total_debt']:.2f}万 | 90天以上：{total['d90_180']+total['d180']:.2f}万</div>
        </div>
    </div>

    <!-- Tab导航 -->
    <div class="tab-nav">
        <button class="tab-btn active" onclick="switchTab('performance', this)"><span class="tab-icon">📊</span>业绩分析</button>
        <button class="tab-btn" onclick="switchTab('debt', this)"><span class="tab-icon">💳</span>欠款分析</button>
        <button class="tab-btn" onclick="switchTab('collection', this)"><span class="tab-icon">⏱️</span>平均回款周期分析</button>
    </div>

    <!-- ===== 业绩分析 Tab ===== -->
    <div id="tab-performance" class="tab-content active">
        <div class="module">
            <h2 class="module-title">📊 26财年Q1 部门业绩总览</h2>
            <div class="charts-section">
                <div class="chart-box">
                    <h3>25Q1 vs 26Q1 部门业绩对比（万元）</h3>
                    <div class="chart-container"><canvas id="perfBarChart"></canvas></div>
                </div>
                <div class="chart-box">
                    <h3>26Q1 部门业绩占比分布</h3>
                    <div class="chart-container"><canvas id="perfPieChart"></canvas></div>
                </div>
            </div>
            <p class="module-desc" style="color:#00d4ff;font-weight:600;">🔽 点击任意部门行 → 查看该部门销售员明细（可继续点击销售员名查看客户明细）</p>
            <table class="dept-table">
                <thead>
                    <tr>
                        <th>部门</th>
                        <th>26Q1业绩(万)</th>
                        <th>Q1目标(万)</th>
                        <th>目标完成率</th>
                        <th>25Q1同期(万)</th>
                        <th>同比%</th>
                        <th>销售员数</th>
                        <th>状态</th>
                    </tr>
                </thead>
                <tbody>
                    {perf_rows}
                </tbody>
                <tfoot>
                    <tr>
                        <td>合计</td><td>{total['v26']:.2f}</td><td>{total['target']:.1f}</td><td>{total['completion']:.1f}%</td>
                        <td>{v25_total_str}</td><td class="trend-down">{yoy_total_str}</td><td>{total['sales']}</td><td></td>
                    </tr>
                </tfoot>
            </table>
        </div>
    </div>

    <!-- ===== 欠款分析 Tab ===== -->
    <div id="tab-debt" class="tab-content">
        <div class="module">
            <h2 class="module-title">💳 欠款分析总览</h2>

            <!-- 欠款KPI -->
            <div class="kpi-grid" style="margin-bottom:25px;">
                <div class="kpi-card">
                    <div class="kpi-icon">📋</div>
                    <h3>欠款总额</h3>
                    <div class="value negative">{total['total_debt']:.2f}<span style="font-size:0.5em;">万</span></div>
                    <div class="sub">全大区欠款合计</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-icon">🟢</div>
                    <h3>30天内</h3>
                    <div class="value highlight">{total['d30']:.2f}<span style="font-size:0.5em;">万</span></div>
                    <div class="sub">占比 {total['d30']/total['total_debt']*100:.1f}%</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-icon">🟡</div>
                    <h3>30-90天</h3>
                    <div class="value warning">{total['d30_90']:.2f}<span style="font-size:0.5em;">万</span></div>
                    <div class="sub">占比 {total['d30_90']/total['total_debt']*100:.1f}%</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-icon">🔴</div>
                    <h3>90天以上</h3>
                    <div class="value negative">{total['d90_180']+total['d180']:.2f}<span style="font-size:0.5em;">万</span></div>
                    <div class="sub">90-180天：{total['d90_180']:.2f}万 | 180天+：{total['d180']:.2f}万</div>
                </div>
            </div>

            <!-- 图表 -->
            <div class="charts-section">
                <div class="chart-box">
                    <h3>逾期欠款金额占比分布</h3>
                    <div class="chart-container"><canvas id="debtPieChart"></canvas></div>
                </div>
                <div class="chart-box">
                    <h3>各部门欠款总额排名（万元）</h3>
                    <div class="chart-container"><canvas id="debtBarChart"></canvas></div>
                </div>
            </div>

            <!-- 欠款明细表 -->
            <h3 style="color:#00d4ff;margin-bottom:15px;font-size:1.1em;">各部门分账龄欠款明细 <span style="color:#8892b0;font-size:0.8em;font-weight:normal;">（点击部门查看销售员明细 → 点击销售员查看客户明细）</span></h3>
            <table class="dept-table">
                <thead>
                    <tr>
                        <th>部门（点击查看明细）</th>
                        <th>30天内(万)</th>
                        <th>30-90天(万)</th>
                        <th>90-180天(万)</th>
                        <th>180天以上(万)</th>
                        <th>合计(万)</th>
                        <th>状态</th>
                    </tr>
                </thead>
                <tbody>
                    {debt_rows}
                </tbody>
                <tfoot>
                    <tr><td>合计</td><td>{total['d30']:.2f}</td><td>{total['d30_90']:.2f}</td><td>{total['d90_180']:.2f}</td><td>{total['d180']:.2f}</td><td>{total['total_debt']:.2f}</td><td></td></tr>
                </tfoot>
            </table>

            <!-- 高风险客户 -->
            <h3 style="color:#ff4757;margin:25px 0 15px;font-size:1.1em;">🚨 高风险客户（90天以上欠款，前15名）</h3>
            <table class="dept-table">
                <thead>
                    <tr><th>排名</th><th>客户名称（所属部门）</th><th>欠款金额(万)</th><th>风险等级</th></tr>
                </thead>
                <tbody>
                    {risky_rows}
                </tbody>
            </table>
        </div>
    </div>

    <!-- ===== 平均回款周期分析 Tab ===== -->
    <div id="tab-collection" class="tab-content">
        <div class="module">
            <h2 class="module-title">⏱️ 平均回款周期分析</h2>
            <div class="kpi-grid" style="margin-bottom:25px;">
                <div class="kpi-card">
                    <div class="kpi-icon">📊</div>
                    <h3>全大区平均回款周期</h3>
                    <div class="value warning">{total['avg_cycle']:.1f}</div>
                    <div class="sub">天（欠款+回款加权综合）</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-icon">🔴</div>
                    <h3>最长部门回款周期</h3>
                    <div class="value negative">{cycle_sorted[0]['cycle']:.1f}</div>
                    <div class="sub">{cycle_sorted[0]['dept']}</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-icon">🟢</div>
                    <h3>最短部门回款周期</h3>
                    <div class="value highlight">{cycle_sorted[-1]['cycle']:.1f}</div>
                    <div class="sub">{cycle_sorted[-1]['dept']}</div>
                </div>
                <div class="kpi-card" onclick="showOver90Depts()" style="cursor:pointer;" title="点击查看回款周期大于90天的部门">
                    <div class="kpi-icon">⚠️</div>
                    <h3>超90天部门数</h3>
                    <div class="value negative">{len(over90_depts)}</div>
                    <div class="sub">回款周期大于90天的部门（点击查看）</div>
                </div>
            </div>

            <!-- 图例 -->
            <div style="background:rgba(255, 255, 255, 0.03);border-radius:12px;padding:15px 20px;margin-bottom:20px;border:1px solid rgba(255, 255, 255, 0.08);">
                <div style="display:flex;align-items:center;gap:8px;margin-bottom:10px;">
                    <span style="font-weight:600;color:#00d4ff;">部门平均回款周期（天）</span>
                    <span style="color:#8892b0;font-size:0.9em;">| 回款周期计算 =（欠款加权天数 + 回款加权天数）÷（欠款总额 + 认款协同金额）</span>
                </div>
                <div style="display:flex;gap:20px;flex-wrap:wrap;">
                    <div style="display:flex;align-items:center;gap:6px;"><span style="width:10px;height:10px;border-radius:50%;background:#00ff88;display:inline-block;"></span><span style="color:#ccd6f6;font-size:0.9em;">≤60天（良好）</span></div>
                    <div style="display:flex;align-items:center;gap:6px;"><span style="width:10px;height:10px;border-radius:50%;background:#ffa502;display:inline-block;"></span><span style="color:#ccd6f6;font-size:0.9em;">61-90天（一般）</span></div>
                    <div style="display:flex;align-items:center;gap:6px;"><span style="width:10px;height:10px;border-radius:50%;background:#ff4757;display:inline-block;"></span><span style="color:#ccd6f6;font-size:0.9em;">&gt;90天（需关注）</span></div>
                </div>
            </div>

            <div class="chart-box" style="margin-bottom:25px;">
                <h3>📊 部门平均回款周期排名</h3>
                <div class="chart-container" style="height:380px;"><canvas id="cycleChart"></canvas></div>
            </div>

            <!-- 回款周期表格 -->
            <table class="dept-table">
                <thead>
                    <tr>
                        <th>部门（点击查看明细）</th>
                        <th>欠款总额(万)</th>
                        <th>回款金额(万)</th>
                        <th>回款周期(天)</th>
                        <th>状态</th>
                    </tr>
                </thead>
                <tbody>
                    {cycle_rows}
                </tbody>
                <tfoot>
                    <tr><td>合计</td><td>{total['total_debt']:.2f}</td><td>{total['collect']:.2f}</td><td>{total['avg_cycle']:.1f}</td><td></td></tr>
                </tfoot>
            </table>
        </div>
    </div>
</div>

<!-- 弹窗 -->
<div class="modal-overlay" id="modalOverlay" onclick="if(event.target===this)closeModal()">
    <div class="modal">
        <div class="modal-header">
            <div class="modal-title" id="modalTitle">部门详情</div>
            <button class="modal-close" onclick="closeModal()">✕</button>
        </div>
        <div id="modalContent"></div>
    </div>
</div>

<!-- 销售员/客户明细弹窗 -->
<div class="modal-overlay" id="salesModal" onclick="if(event.target===this)closeSalesModal()" style="z-index:1001;">
    <div style="background:#1a2040;border-radius:20px;padding:30px;max-width:1100px;width:95%;max-height:85vh;overflow-y:auto;border:1px solid rgba(0,212,255,0.3);">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px;">
            <div class="modal-title" id="salesModalTitle">销售员明细</div>
            <button class="modal-close" onclick="closeSalesModal()">✕</button>
        </div>
        <div id="salesTableContainer"></div>
    </div>
</div>

<script>
// ===== 数据 =====
const deptData = {dept_list_json};

// ===== Tab切换 =====
let currentTabType = 'perf';
function switchTab(id, btn) {{
    document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.getElementById('tab-' + id).classList.add('active');
    btn.classList.add('active');
    if (id === 'performance') currentTabType = 'perf';
    else if (id === 'debt') currentTabType = 'debt';
    else if (id === 'collection') currentTabType = 'cycle';
}}
function getCurrentTab() {{ return currentTabType; }}

// ===== 销售员明细数据（万元）=====
const salesDetailData = {sales_detail_json};

// ===== 销售员回款周期数据（元）=====
const salesCycleData = {sales_cycle_json};

// ===== 客户业绩数据 =====
// deptCustPerfData: {{dept: [{{name:seller, customers:[{{name, perf, collect, orders}}]}}]}}
const deptCustPerfData = {dept_cust_perf_json};

// ===== 客户欠款数据 =====
// deptCustDebtData: {{dept: [{{name:seller, customers:[{{name, total_debt, d30, d30_90, d90_180, d180, max_days}}]}}]}}
const deptCustDebtData = {dept_cust_debt_json};

// ===== 客户回款周期数据（元）=====
// custCycleData: {{dept: {{seller: [{{name, debt_amt, rec_amt, cycle}}]}}}}
const custCycleData = {cust_cycle_json};

// ===== 部门弹窗 =====
function showDeptDetail(dept) {{
    const d = deptData.find(x => x.dept === dept);
    if (!d) return;
    document.getElementById('modalTitle').textContent = dept + ' — 部门概览';
    const risky = d.d90_180 + d.d180;
    document.getElementById('modalContent').innerHTML = `
        <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:20px;">
            <div style="background:rgba(0,212,255,0.1);border-radius:12px;padding:14px;text-align:center;">
                <div style="color:#8892b0;font-size:0.8em;margin-bottom:4px;">26Q1业绩</div>
                <div style="color:#00ff88;font-size:1.5em;font-weight:700;">${{d.v26}}万</div>
            </div>
            <div style="background:rgba(0,212,255,0.1);border-radius:12px;padding:14px;text-align:center;">
                <div style="color:#8892b0;font-size:0.8em;margin-bottom:4px;">完成率</div>
                <div style="color:#ffa502;font-size:1.5em;font-weight:700;">${{d.completion}}%</div>
            </div>
            <div style="background:rgba(255,71,87,0.1);border-radius:12px;padding:14px;text-align:center;">
                <div style="color:#8892b0;font-size:0.8em;margin-bottom:4px;">同比25Q1</div>
                <div style="color:#ff4757;font-size:1.5em;font-weight:700;">${{d.yoy}}%</div>
            </div>
            <div style="background:rgba(0,212,255,0.1);border-radius:12px;padding:14px;text-align:center;">
                <div style="color:#8892b0;font-size:0.8em;margin-bottom:4px;">欠款总额</div>
                <div style="color:#ff4757;font-size:1.5em;font-weight:700;">${{d.total_debt}}万</div>
            </div>
        </div>
        <div style="text-align:center;margin-top:10px;">
            <button onclick="showSalesDetail('${{dept}}', getCurrentTab())" style="background:linear-gradient(135deg,#00d4ff,#7b2ff7);color:#fff;border:none;border-radius:8px;padding:10px 28px;font-size:1em;cursor:pointer;">📋 查看销售员明细</button>
        </div>`;
    document.getElementById('modalOverlay').classList.add('active');
}}
function closeModal() {{
    document.getElementById('modalOverlay').classList.remove('active');
}}

// ===== 销售员明细弹窗 =====
function showSalesDetail(dept, type) {{
    if (type === 'perf') renderSalesPerf(dept);
    else if (type === 'debt') renderSalesDebt(dept);
    else if (type === 'cycle') renderSalesCycle(dept);
    else renderSalesPerf(dept);
}}

function renderSalesPerf(dept) {{
    const list = [...(salesDetailData[dept] || [])];
    list.sort((a, b) => b.perf - a.perf);
    const rows = list.map(s => {{
        const color = s.perf < 0 ? '#ff4757' : s.perf < 5 ? '#ffa502' : '#00ff88';
        const status = s.perf < 0 ? '🔴 无业绩' : s.perf < 5 ? '🟡 待提升' : '🟢 正常';
        // 查找客户数据
        const custEntry = deptCustPerfData[dept] && deptCustPerfData[dept].find(x => x.name === s.name);
        const hasCust = custEntry && custEntry.customers && custEntry.customers.length > 0;
        const nameStyle = hasCust ? 'color:#00d4ff;cursor:pointer;text-decoration:underline;' : 'color:#ccd6f6;';
        const nameClick = hasCust ? `onclick="renderCustPerf('${{dept}}','${{s.name}}')"` : '';
        return `<tr style="border-bottom:1px solid rgba(255,255,255,0.05);">
            <td style="padding:7px 5px;${{nameStyle}}" ${{nameClick}}>${{s.name}}${{hasCust ? ' 📂' : ''}}</td>
            <td style="padding:7px 5px;color:${{color}};text-align:right;font-weight:600;">${{s.perf.toFixed(2)}}</td>
            <td style="padding:7px 5px;color:#00ff88;text-align:right;">${{s.collect.toFixed(2)}}</td>
            <td style="padding:7px 5px;color:#ccd6f6;text-align:center;">${{status}}</td>
        </tr>`;
    }}).join('');
    const tp = list.reduce((s, v) => s + v.perf, 0).toFixed(2);
    const tc = list.reduce((s, v) => s + v.collect, 0).toFixed(2);
    document.getElementById('salesModalTitle').innerHTML = `📊 ${{dept}} - 销售员业绩明细 <span style="font-size:0.7em;color:#8892b0;">（点击蓝色销售员名查看客户明细）</span>`;
    document.getElementById('salesTableContainer').innerHTML = `<table style="width:100%;border-collapse:collapse;font-size:0.85em;">
        <thead><tr style="background:rgba(0,212,255,0.12);">
            <th style="padding:8px 6px;color:#00d4ff;">销售员</th>
            <th style="padding:8px 6px;color:#00d4ff;text-align:right;">26Q1业绩(万)</th>
            <th style="padding:8px 6px;color:#00d4ff;text-align:right;">回款(万)</th>
            <th style="padding:8px 6px;color:#00d4ff;text-align:center;">状态</th>
        </tr></thead>
        <tbody>${{rows}}
        <tr style="background:rgba(0,212,255,0.08);font-weight:600;">
            <td style="padding:8px 6px;color:#00d4ff;">合计（${{list.length}}人）</td>
            <td style="padding:8px 6px;color:#00ff88;text-align:right;">${{tp}}</td>
            <td style="padding:8px 6px;color:#00ff88;text-align:right;">${{tc}}</td>
            <td></td>
        </tr></tbody></table>`;
    document.getElementById('salesModal').classList.add('active');
}}

function renderSalesDebt(dept) {{
    const list = [...(salesDetailData[dept] || [])];
    list.sort((a, b) => b.total_debt - a.total_debt);
    const rows = list.map(s => {{
        const dc = s.total_debt > 50 ? '#ff4757' : s.total_debt > 20 ? '#ffa502' : '#00ff88';
        const status = s.total_debt > 50 ? '🔴 高风险' : s.total_debt > 20 ? '🟡 关注' : '🟢 较好';
        const custEntry = deptCustDebtData[dept] && deptCustDebtData[dept].find(x => x.name === s.name);
        const hasCust = custEntry && custEntry.customers && custEntry.customers.length > 0;
        const nameStyle = hasCust ? 'color:#00d4ff;cursor:pointer;text-decoration:underline;' : 'color:#ccd6f6;';
        const nameClick = hasCust ? `onclick="renderCustDebt('${{dept}}','${{s.name}}')"` : '';
        return `<tr style="border-bottom:1px solid rgba(255,255,255,0.05);">
            <td style="padding:7px 5px;${{nameStyle}}" ${{nameClick}}>${{s.name}}${{hasCust ? ' 📂' : ''}}</td>
            <td style="padding:7px 5px;color:${{dc}};text-align:right;font-weight:600;">${{s.total_debt.toFixed(2)}}</td>
            <td style="padding:7px 5px;color:#00ff88;text-align:right;">${{s.d30.toFixed(2)}}</td>
            <td style="padding:7px 5px;color:#ffa502;text-align:right;">${{s.d30_90.toFixed(2)}}</td>
            <td style="padding:7px 5px;color:${{s.d90_180>0?'#ff4757':'#8892b0'}};text-align:right;">${{s.d90_180.toFixed(2)}}</td>
            <td style="padding:7px 5px;color:${{s.d180>0?'#ff4757':'#8892b0'}};text-align:right;">${{s.d180.toFixed(2)}}</td>
            <td style="padding:7px 5px;color:#ccd6f6;text-align:center;">${{status}}</td>
        </tr>`;
    }}).join('');
    const td = list.reduce((s, v) => s + v.total_debt, 0).toFixed(2);
    document.getElementById('salesModalTitle').innerHTML = `💰 ${{dept}} - 销售员欠款明细 <span style="font-size:0.7em;color:#8892b0;">（点击蓝色销售员名查看客户明细）</span>`;
    document.getElementById('salesTableContainer').innerHTML = `<table style="width:100%;border-collapse:collapse;font-size:0.82em;">
        <thead><tr style="background:rgba(0,212,255,0.12);">
            <th style="padding:8px 6px;color:#00d4ff;">销售员</th>
            <th style="padding:8px 6px;color:#00d4ff;text-align:right;">合计欠款(万)</th>
            <th style="padding:8px 6px;color:#00d4ff;text-align:right;">30天内</th>
            <th style="padding:8px 6px;color:#ffa502;text-align:right;">30-90天</th>
            <th style="padding:8px 6px;color:#ff4757;text-align:right;">90-180天</th>
            <th style="padding:8px 6px;color:#ff4757;text-align:right;">180天以上</th>
            <th style="padding:8px 6px;color:#00d4ff;text-align:center;">状态</th>
        </tr></thead>
        <tbody>${{rows}}
        <tr style="background:rgba(0,212,255,0.08);font-weight:600;">
            <td style="padding:8px 6px;color:#00d4ff;">合计（${{list.length}}人）</td>
            <td style="padding:8px 6px;color:#ff4757;text-align:right;">${{td}}</td>
            <td colspan="5"></td>
        </tr></tbody></table>`;
    document.getElementById('salesModal').classList.add('active');
}}

function renderSalesCycle(dept) {{
    const list = [...(salesCycleData[dept] || [])];
    list.sort((a, b) => b.cycle - a.cycle);
    const rows = list.map(s => {{
        const cycle = s.cycle;
        const cycleStr = cycle > 0 ? cycle.toFixed(1) : '-';
        const cc = cycle <= 0 ? '#8892b0' : cycle > 90 ? '#ff4757' : cycle > 60 ? '#ffa502' : '#00ff88';
        const status = cycle <= 0 ? '⚪ 无数据' : cycle > 90 ? '🔴 需关注' : cycle > 60 ? '🟡 偏高' : '🟢 正常';
        const hasCust = custCycleData[dept] && custCycleData[dept][s.name] && custCycleData[dept][s.name].length > 0;
        const nameStyle = hasCust ? 'color:#00d4ff;cursor:pointer;text-decoration:underline;' : 'color:#ccd6f6;';
        const nameClick = hasCust ? `onclick="renderCustCycle('${{dept}}','${{s.name}}')"` : '';
        return `<tr style="border-bottom:1px solid rgba(255,255,255,0.05);">
            <td style="padding:7px 5px;${{nameStyle}}" ${{nameClick}}>${{s.name}}${{hasCust ? ' 📂' : ''}}</td>
            <td style="padding:7px 5px;color:#00ff88;text-align:right;">${{(s.rec_amt/10000).toFixed(2)}}</td>
            <td style="padding:7px 5px;color:#ffa502;text-align:right;">${{(s.debt_amt/10000).toFixed(2)}}</td>
            <td style="padding:7px 5px;color:${{cc}};text-align:right;font-weight:600;">${{cycleStr}}</td>
            <td style="padding:7px 5px;color:#ccd6f6;text-align:center;">${{status}}</td>
        </tr>`;
    }}).join('');
    const tc = list.reduce((s, v) => s + v.rec_amt, 0);
    const td = list.reduce((s, v) => s + v.debt_amt, 0);
    document.getElementById('salesModalTitle').innerHTML = `⏱️ ${{dept}} - 销售员回款周期明细 <span style="font-size:0.7em;color:#8892b0;">（点击蓝色销售员名查看客户明细）</span>`;
    document.getElementById('salesTableContainer').innerHTML = `<table style="width:100%;border-collapse:collapse;font-size:0.85em;">
        <thead><tr style="background:rgba(0,212,255,0.12);">
            <th style="padding:8px 6px;color:#00d4ff;">销售员</th>
            <th style="padding:8px 6px;color:#00d4ff;text-align:right;">认款金额(万)</th>
            <th style="padding:8px 6px;color:#00d4ff;text-align:right;">欠款金额(万)</th>
            <th style="padding:8px 6px;color:#00d4ff;text-align:right;">回款周期(天)</th>
            <th style="padding:8px 6px;color:#00d4ff;text-align:center;">状态</th>
        </tr></thead>
        <tbody>${{rows}}
        <tr style="background:rgba(0,212,255,0.08);font-weight:600;">
            <td style="padding:8px 6px;color:#00d4ff;">合计（${{list.length}}人）</td>
            <td style="padding:8px 6px;color:#00ff88;text-align:right;">${{(tc/10000).toFixed(2)}}</td>
            <td style="padding:8px 6px;color:#ffa502;text-align:right;">${{(td/10000).toFixed(2)}}</td>
            <td colspan="2"></td>
        </tr></tbody></table>`;
    document.getElementById('salesModal').classList.add('active');
}}

function closeSalesModal() {{
    document.getElementById('salesModal').classList.remove('active');
    window._salesDebtBack = null;
    window._salesPerfBack = null;
    window._salesCycleBack = null;
}}

// ===== 客户回款周期明细（第二级下钻）=====
function renderCustCycle(dept, salesName) {{
    window._salesCycleBack = document.getElementById('salesTableContainer').innerHTML;
    window._salesCycleTitle = document.getElementById('salesModalTitle').innerHTML;

    const custList = custCycleData[dept] && custCycleData[dept][salesName] ? custCycleData[dept][salesName] : [];
    if (!custList.length) return;

    const sorted = [...custList].sort((a, b) => b.cycle - a.cycle);
    const rows = sorted.map(c => {{
        const cycle = c.cycle;
        const cycleStr = cycle > 0 ? cycle.toFixed(1) : '-';
        const cc = cycle <= 0 ? '#8892b0' : cycle > 90 ? '#ff4757' : cycle > 60 ? '#ffa502' : '#00ff88';
        const status = cycle <= 0 ? '⚪ 无数据' : cycle > 90 ? '🔴 需关注' : cycle > 60 ? '🟡 偏高' : '🟢 正常';
        return `<tr style="border-bottom:1px solid rgba(255,255,255,0.05);">
            <td style="padding:7px 5px;color:#ccd6f6;max-width:300px;word-break:break-all;">${{c.name}}</td>
            <td style="padding:7px 5px;color:#00ff88;text-align:right;">${{(c.rec_amt/10000).toFixed(2)}}</td>
            <td style="padding:7px 5px;color:#ffa502;text-align:right;">${{(c.debt_amt/10000).toFixed(2)}}</td>
            <td style="padding:7px 5px;color:${{cc}};text-align:right;font-weight:600;">${{cycleStr}}</td>
            <td style="padding:7px 5px;color:#ccd6f6;text-align:center;">${{status}}</td>
        </tr>`;
    }}).join('');

    const tc = sorted.reduce((s, c) => s + c.rec_amt, 0);
    const td = sorted.reduce((s, c) => s + c.debt_amt, 0);

    document.getElementById('salesModalTitle').innerHTML = `⏱️ ${{salesName}} - 客户回款周期明细 <span style="font-size:0.75em;color:#8892b0;">（${{dept}}）</span> <button onclick="backToSalesCycle()" style="margin-left:12px;background:rgba(0,212,255,0.2);border:1px solid rgba(0,212,255,0.4);color:#00d4ff;border-radius:6px;padding:3px 12px;font-size:0.85em;cursor:pointer;">← 返回销售员</button>`;
    document.getElementById('salesTableContainer').innerHTML = `<table style="width:100%;border-collapse:collapse;font-size:0.8em;">
        <thead><tr style="background:rgba(0,212,255,0.12);">
            <th style="padding:8px 6px;color:#00d4ff;">客户名称</th>
            <th style="padding:8px 6px;color:#00d4ff;text-align:right;">认款金额(万)</th>
            <th style="padding:8px 6px;color:#00d4ff;text-align:right;">欠款金额(万)</th>
            <th style="padding:8px 6px;color:#00d4ff;text-align:right;">回款周期(天)</th>
            <th style="padding:8px 6px;color:#00d4ff;text-align:center;">状态</th>
        </tr></thead>
        <tbody>${{rows}}
        <tr style="background:rgba(0,212,255,0.08);font-weight:600;">
            <td style="padding:8px 6px;color:#00d4ff;">合计（${{sorted.length}}个客户）</td>
            <td style="padding:8px 6px;color:#00ff88;text-align:right;">${{(tc/10000).toFixed(2)}}</td>
            <td style="padding:8px 6px;color:#ffa502;text-align:right;">${{(td/10000).toFixed(2)}}</td>
            <td colspan="2"></td>
        </tr></tbody></table>`;
}}

function backToSalesCycle() {{
    if (window._salesCycleBack) {{
        document.getElementById('salesTableContainer').innerHTML = window._salesCycleBack;
        document.getElementById('salesModalTitle').innerHTML = window._salesCycleTitle;
        window._salesCycleBack = null;
    }}
}}

// ===== 客户欠款明细（第二级下钻）=====
function renderCustDebt(dept, salesName) {{
    window._salesDebtBack = document.getElementById('salesTableContainer').innerHTML;
    window._salesDebtTitle = document.getElementById('salesModalTitle').innerHTML;

    const custEntry = deptCustDebtData[dept] && deptCustDebtData[dept].find(x => x.name === salesName);
    const custList = custEntry ? custEntry.customers : [];
    if (!custList.length) return;

    const sorted = [...custList].sort((a, b) => b.total_debt - a.total_debt);
    const rows = sorted.map(c => {{
        const dc = c.total_debt > 50 ? '#ff4757' : c.total_debt > 20 ? '#ffa502' : '#00ff88';
        const risky = c.d90_180 + c.d180;
        const status = risky > 20 ? '🔴 高风险' : risky > 5 ? '🟡 关注' : '🟢 较好';
        const daysText = c.max_days > 0 ? `最长${{c.max_days}}天` : '-';
        return `<tr style="border-bottom:1px solid rgba(255,255,255,0.05);">
            <td style="padding:7px 5px;color:#ccd6f6;max-width:280px;word-break:break-all;">${{c.name}}</td>
            <td style="padding:7px 5px;color:${{dc}};text-align:right;font-weight:600;">${{c.total_debt.toFixed(2)}}</td>
            <td style="padding:7px 5px;color:#00ff88;text-align:right;">${{c.d30.toFixed(2)}}</td>
            <td style="padding:7px 5px;color:#ffa502;text-align:right;">${{c.d30_90.toFixed(2)}}</td>
            <td style="padding:7px 5px;color:${{c.d90_180>0?'#ff4757':'#8892b0'}};text-align:right;">${{c.d90_180.toFixed(2)}}</td>
            <td style="padding:7px 5px;color:${{c.d180>0?'#ff4757':'#8892b0'}};text-align:right;">${{c.d180.toFixed(2)}}</td>
            <td style="padding:7px 5px;color:#8892b0;text-align:center;font-size:0.9em;">${{daysText}}</td>
            <td style="padding:7px 5px;color:#ccd6f6;text-align:center;">${{status}}</td>
        </tr>`;
    }}).join('');

    const td = sorted.reduce((s, c) => s + c.total_debt, 0).toFixed(2);

    document.getElementById('salesModalTitle').innerHTML = `💰 ${{salesName}} - 客户欠款明细 <span style="font-size:0.75em;color:#8892b0;">（${{dept}}）</span> <button onclick="backToSalesDebt()" style="margin-left:12px;background:rgba(0,212,255,0.2);border:1px solid rgba(0,212,255,0.4);color:#00d4ff;border-radius:6px;padding:3px 12px;font-size:0.85em;cursor:pointer;">← 返回销售员</button>`;
    document.getElementById('salesTableContainer').innerHTML = `<table style="width:100%;border-collapse:collapse;font-size:0.78em;">
        <thead><tr style="background:rgba(0,212,255,0.12);">
            <th style="padding:8px 6px;color:#00d4ff;">客户名称</th>
            <th style="padding:8px 6px;color:#ff4757;text-align:right;">合计欠款(万)</th>
            <th style="padding:8px 6px;color:#00d4ff;text-align:right;">30天内</th>
            <th style="padding:8px 6px;color:#ffa502;text-align:right;">30-90天</th>
            <th style="padding:8px 6px;color:#ff4757;text-align:right;">90-180天</th>
            <th style="padding:8px 6px;color:#ff4757;text-align:right;">180天以上</th>
            <th style="padding:8px 6px;color:#8892b0;text-align:center;">最长天数</th>
            <th style="padding:8px 6px;color:#00d4ff;text-align:center;">状态</th>
        </tr></thead>
        <tbody>${{rows}}
        <tr style="background:rgba(0,212,255,0.08);font-weight:600;">
            <td style="padding:8px 6px;color:#00d4ff;">合计（${{sorted.length}}个客户）</td>
            <td style="padding:8px 6px;color:#ff4757;text-align:right;">${{td}}</td>
            <td colspan="6"></td>
        </tr></tbody></table>`;
}}

function backToSalesDebt() {{
    if (window._salesDebtBack) {{
        document.getElementById('salesTableContainer').innerHTML = window._salesDebtBack;
        document.getElementById('salesModalTitle').innerHTML = window._salesDebtTitle;
        window._salesDebtBack = null;
    }}
}}

// ===== 客户业绩明细（第二级下钻）=====
function renderCustPerf(dept, salesName) {{
    window._salesPerfBack = document.getElementById('salesTableContainer').innerHTML;
    window._salesPerfTitle = document.getElementById('salesModalTitle').innerHTML;

    const custEntry = deptCustPerfData[dept] && deptCustPerfData[dept].find(x => x.name === salesName);
    const custList = custEntry ? custEntry.customers : [];
    if (!custList.length) return;

    const sorted = [...custList].sort((a, b) => b.perf - a.perf);
    const rows = sorted.map(c => {{
        const color = c.perf <= 0 ? '#8892b0' : c.perf < 2 ? '#ffa502' : '#00ff88';
        const status = c.perf <= 0 ? '⚪ 无业绩' : c.perf < 5 ? '🟡 较低' : '🟢 正常';
        return `<tr style="border-bottom:1px solid rgba(255,255,255,0.05);">
            <td style="padding:7px 5px;color:#ccd6f6;max-width:300px;word-break:break-all;">${{c.name}}</td>
            <td style="padding:7px 5px;color:${{color}};text-align:right;font-weight:600;">${{c.perf.toFixed(2)}}</td>
            <td style="padding:7px 5px;color:#00ff88;text-align:right;">${{c.collect.toFixed(2)}}</td>
            <td style="padding:7px 5px;color:#8892b0;text-align:center;">${{c.orders}}</td>
            <td style="padding:7px 5px;color:#ccd6f6;text-align:center;">${{status}}</td>
        </tr>`;
    }}).join('');

    const tp = sorted.reduce((s, c) => s + c.perf, 0).toFixed(2);
    const tc = sorted.reduce((s, c) => s + c.collect, 0).toFixed(2);

    document.getElementById('salesModalTitle').innerHTML = `📊 ${{salesName}} - 客户业绩明细 <span style="font-size:0.75em;color:#8892b0;">（${{dept}}）</span> <button onclick="backToSalesPerf()" style="margin-left:12px;background:rgba(0,212,255,0.2);border:1px solid rgba(0,212,255,0.4);color:#00d4ff;border-radius:6px;padding:3px 12px;font-size:0.85em;cursor:pointer;">← 返回销售员</button>`;
    document.getElementById('salesTableContainer').innerHTML = `<table style="width:100%;border-collapse:collapse;font-size:0.82em;">
        <thead><tr style="background:rgba(0,212,255,0.12);">
            <th style="padding:8px 6px;color:#00d4ff;">客户名称</th>
            <th style="padding:8px 6px;color:#00d4ff;text-align:right;">26Q1业绩(万)</th>
            <th style="padding:8px 6px;color:#00d4ff;text-align:right;">回款(万)</th>
            <th style="padding:8px 6px;color:#00d4ff;text-align:center;">订单数</th>
            <th style="padding:8px 6px;color:#00d4ff;text-align:center;">状态</th>
        </tr></thead>
        <tbody>${{rows}}
        <tr style="background:rgba(0,212,255,0.08);font-weight:600;">
            <td style="padding:8px 6px;color:#00d4ff;">合计（${{sorted.length}}个客户）</td>
            <td style="padding:8px 6px;color:#00ff88;text-align:right;">${{tp}}</td>
            <td style="padding:8px 6px;color:#00ff88;text-align:right;">${{tc}}</td>
            <td colspan="2"></td>
        </tr></tbody></table>`;
}}

function backToSalesPerf() {{
    if (window._salesPerfBack) {{
        document.getElementById('salesTableContainer').innerHTML = window._salesPerfBack;
        document.getElementById('salesModalTitle').innerHTML = window._salesPerfTitle;
        window._salesPerfBack = null;
    }}
}}

// ===== 超90天部门 =====
function showOver90Depts() {{
    const over90 = deptData.filter(d => d.cycle > 90);
    if (!over90.length) {{
        alert('当前没有回款周期超过90天的部门');
        return;
    }}
    const rows = over90.map(d => `<tr style="border-bottom:1px solid rgba(255,255,255,0.05);cursor:pointer;" onclick="showSalesDetail('${{d.dept}}','cycle')">
        <td style="padding:8px 6px;color:#ccd6f6;">🏢 ${{d.dept}}</td>
        <td style="padding:8px 6px;color:#ff4757;text-align:right;font-weight:600;">${{d.cycle.toFixed(1)}}天</td>
        <td style="padding:8px 6px;color:#ffa502;text-align:right;">${{d.total_debt.toFixed(2)}}万</td>
        <td style="padding:8px 6px;text-align:center;"><span class="status-badge badge-down">需关注</span></td>
    </tr>`).join('');
    document.getElementById('salesModalTitle').textContent = `⚠️ 回款周期超90天部门（共${{over90.length}}个）`;
    document.getElementById('salesTableContainer').innerHTML = `<table style="width:100%;border-collapse:collapse;font-size:0.85em;">
        <thead><tr style="background:rgba(0,212,255,0.12);">
            <th style="padding:8px 6px;color:#00d4ff;">部门</th>
            <th style="padding:8px 6px;color:#ff4757;text-align:right;">回款周期</th>
            <th style="padding:8px 6px;color:#ffa502;text-align:right;">欠款总额</th>
            <th style="padding:8px 6px;color:#00d4ff;text-align:center;">状态</th>
        </tr></thead><tbody>${{rows}}</tbody></table>
        <p style="color:#8892b0;font-size:0.85em;margin-top:10px;text-align:center;">点击部门行可查看该部门销售员明细</p>`;
    document.getElementById('salesModal').classList.add('active');
}}

// ===== 图表 =====
Chart.register(ChartDataLabels);
const DEPT_COLORS = ['#00d4ff','#7b2ff7','#00ff88','#ffa502','#ff4757','#ff6b6b','#48dbfb','#ff9f43','#1dd1a1'];

// 业绩柱图
new Chart(document.getElementById('perfBarChart'), {{
    type: 'bar',
    plugins: [ChartDataLabels],
    data: {{
        labels: {json.dumps(perf_depts, ensure_ascii=False)},
        datasets: [
            {{ label: '25Q1', data: {perf_v25}, backgroundColor: 'rgba(123,47,247,0.5)', borderColor: '#7b2ff7', borderWidth: 1 }},
            {{ label: '26Q1', data: {perf_v26}, backgroundColor: 'rgba(0,212,255,0.7)', borderColor: '#00d4ff', borderWidth: 1 }}
        ]
    }},
    options: {{
        responsive: true, maintainAspectRatio: false,
        plugins: {{ legend: {{ labels: {{ color: '#ccd6f6' }} }}, datalabels: {{ display: false }} }},
        scales: {{
            x: {{ ticks: {{ color: '#8892b0', maxRotation: 45 }}, grid: {{ color: 'rgba(255,255,255,0.05)' }} }},
            y: {{ ticks: {{ color: '#8892b0' }}, grid: {{ color: 'rgba(255,255,255,0.08)' }} }}
        }}
    }}
}});

// 业绩饼图
new Chart(document.getElementById('perfPieChart'), {{
    type: 'doughnut',
    plugins: [ChartDataLabels],
    data: {{
        labels: {json.dumps([x[0] for x in perf_pie], ensure_ascii=False)},
        datasets: [{{ data: {[x[1] for x in perf_pie]}, backgroundColor: DEPT_COLORS, borderWidth: 1, borderColor: '#0a0e27' }}]
    }},
    options: {{
        responsive: true, maintainAspectRatio: false,
        plugins: {{
            legend: {{ position: 'right', labels: {{ color: '#ccd6f6', font: {{ size: 11 }} }} }},
            datalabels: {{
                color: '#fff', font: {{ size: 10 }},
                formatter: (v, ctx) => {{
                    const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
                    return (v/total*100).toFixed(1) + '%';
                }}
            }}
        }}
    }}
}});

// 欠款饼图
new Chart(document.getElementById('debtPieChart'), {{
    type: 'doughnut',
    plugins: [ChartDataLabels],
    data: {{
        labels: {json.dumps(debt_pie_labels, ensure_ascii=False)},
        datasets: [{{ data: {debt_pie_vals}, backgroundColor: ['#00ff88','#ffa502','#ff6b6b','#ff4757'], borderWidth: 1, borderColor: '#0a0e27' }}]
    }},
    options: {{
        responsive: true, maintainAspectRatio: false,
        plugins: {{
            legend: {{ position: 'bottom', labels: {{ color: '#ccd6f6' }} }},
            datalabels: {{
                color: '#fff', font: {{ size: 11, weight: 'bold' }},
                formatter: (v, ctx) => {{
                    const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
                    return (v/total*100).toFixed(1) + '%';
                }}
            }}
        }}
    }}
}});

// 欠款柱图
new Chart(document.getElementById('debtBarChart'), {{
    type: 'bar',
    plugins: [ChartDataLabels],
    data: {{
        labels: {json.dumps(debt_bar_depts, ensure_ascii=False)},
        datasets: [{{ data: {debt_bar_vals}, backgroundColor: DEPT_COLORS, borderWidth: 1, borderColor: '#0a0e27' }}]
    }},
    options: {{
        indexAxis: 'y',
        responsive: true, maintainAspectRatio: false,
        plugins: {{ legend: {{ display: false }}, datalabels: {{ color: '#fff', font: {{ size: 10 }}, anchor: 'end', align: 'end', formatter: v => v.toFixed(1) }} }},
        scales: {{
            x: {{ ticks: {{ color: '#8892b0' }}, grid: {{ color: 'rgba(255,255,255,0.08)' }} }},
            y: {{ ticks: {{ color: '#8892b0', font: {{ size: 11 }} }}, grid: {{ display: false }} }}
        }}
    }}
}});

// 回款周期横向柱图
new Chart(document.getElementById('cycleChart'), {{
    type: 'bar',
    plugins: [ChartDataLabels],
    data: {{
        labels: {json.dumps(chart_cycle_depts, ensure_ascii=False)},
        datasets: [{{
            data: {chart_cycle_vals},
            backgroundColor: {json.dumps(chart_cycle_colors)},
            borderWidth: 1, borderColor: '#0a0e27'
        }}]
    }},
    options: {{
        indexAxis: 'y',
        responsive: true, maintainAspectRatio: false,
        plugins: {{
            legend: {{ display: false }},
            datalabels: {{ color: '#fff', font: {{ size: 11 }}, anchor: 'end', align: 'end', formatter: v => v.toFixed(1) + '天' }},
            annotation: {{ annotations: {{ line1: {{ type: 'line', xMin: 90, xMax: 90, borderColor: '#ff4757', borderWidth: 2, borderDash: [5, 5] }} }} }}
        }},
        scales: {{
            x: {{ ticks: {{ color: '#8892b0' }}, grid: {{ color: 'rgba(255,255,255,0.08)' }}, max: Math.max(...{chart_cycle_vals}) * 1.15 }},
            y: {{ ticks: {{ color: '#8892b0' }}, grid: {{ display: false }} }}
        }}
    }}
}});
</script>
</body>
</html>'''

outpath = os.path.join(BASE_DIR, 'index.html')
with open(outpath, 'w', encoding='utf-8') as f:
    f.write(html)
print(f"Done: {outpath}")
print(f"Size: {len(html.encode('utf-8'))/1024:.1f} KB")
