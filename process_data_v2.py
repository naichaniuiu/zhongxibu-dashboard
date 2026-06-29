# -*- coding: utf-8 -*-
"""
从Excel提取数据并计算所有看板指标 - 修正版
"""
import sys
import io
import os
import json
from collections import defaultdict
from datetime import datetime

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from openpyxl import load_workbook

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(os.path.expanduser('~'), 'Downloads', '业绩 欠款看板.xlsx')

wb = load_workbook(EXCEL_PATH, read_only=True, data_only=True)

def read_sheet_data(wb, idx):
    ws = wb.worksheets[idx]
    rows = list(ws.rows)
    if not rows:
        return [], []
    headers = [str(cell.value).strip() if cell.value is not None else '' for cell in rows[0]]
    data = []
    for row in rows[1:]:
        vals = [cell.value for cell in row]
        if all(v is None for v in vals):
            continue
        data.append(dict(zip(headers, vals)))
    return headers, data

def safe_float(v, default=0.0):
    try:
        if v is None:
            return default
        return float(v)
    except:
        return default

def get_dept_norm(row):
    """获取归一化部门名（中西部大区）"""
    d1 = str(row.get('一级部门', '') or '').strip()
    if d1 != '中西部大区':
        return None
    d3 = str(row.get('三级部门', '') or '').strip()
    if d3 and d3 != 'None':
        return normalize_dept(d3)
    d2 = str(row.get('二级部门', '') or '').strip()
    if d2 and d2 != 'None':
        return normalize_dept(d2)
    return '其他'

def normalize_dept(dept):
    if not dept or dept == 'None':
        return '其他'
    if '基建' in dept or '制造' in dept:
        return '武汉基建制造行业组'
    if '金融' in dept:
        return '武汉金融行业组'
    if '能源' in dept or '交通' in dept:
        return '武汉能源交通行业组'
    if '成都' in dept:
        return '成都站'
    if '长沙' in dept:
        return '长沙站'
    if '西安' in dept:
        return '西安站'
    if '郑州' in dept:
        return '郑州站'
    if '重庆' in dept:
        return '重庆站'
    return '其他'

ALL_DEPTS = ['武汉基建制造行业组', '武汉金融行业组', '成都站', '长沙站', '武汉能源交通行业组', '西安站', '郑州站', '重庆站', '其他']

# ============================
# Sheet 4: 目标数据（单位已是万元）
# ============================
_, target_rows = read_sheet_data(wb, 4)
dept_targets = {}
for r in target_rows:
    dept = str(r.get('区域') or '').strip()
    q1 = r.get('26财年Q1')
    if dept and q1 and isinstance(q1, (int, float)):
        dept_targets[dept] = round(float(q1), 1)  # 已经是万元

# ============================
# Sheet 0: 25财年Q1业绩
# ============================
_, perf25_rows = read_sheet_data(wb, 0)
dept_perf25 = defaultdict(float)
for r in perf25_rows:
    dept = get_dept_norm(r)
    if dept is None:
        continue
    dept_perf25[dept] += safe_float(r.get('业绩总金额', 0))

for k in list(dept_perf25.keys()):
    dept_perf25[k] = round(dept_perf25[k] / 10000, 2)

# ============================
# Sheet 1: 26财年Q1业绩
# ============================
_, perf26_rows = read_sheet_data(wb, 1)
dept_perf26 = defaultdict(float)
dept_active_sellers = defaultdict(set)
sales_perf26 = defaultdict(lambda: defaultdict(float))
sales_collect26 = defaultdict(lambda: defaultdict(float))
# 建立业绩单号映射
order_dept_map = {}
order_seller_map = {}

for r in perf26_rows:
    dept = get_dept_norm(r)
    if dept is None:
        continue
    dept_perf26[dept] += safe_float(r.get('业绩总金额', 0))
    seller = str(r.get('销售员名称', '') or '').strip()
    status = str(r.get('销售员状态', '') or '').strip()
    if seller:
        sales_perf26[dept][seller] += safe_float(r.get('业绩总金额', 0))
        sales_collect26[dept][seller] += safe_float(r.get('回款金额', 0))
        if status == '在职' and seller not in ('李国栋', '白雨'):
            dept_active_sellers[dept].add(seller)
    order_no = str(r.get('业绩单号', '') or '').strip()
    if order_no:
        order_dept_map[order_no] = dept
        if seller:
            order_seller_map[order_no] = seller

for k in list(dept_perf26.keys()):
    dept_perf26[k] = round(dept_perf26[k] / 10000, 2)

# 从25Q1也建映射
for r in perf25_rows:
    order_no = str(r.get('业绩单号', '') or '').strip()
    if order_no and order_no not in order_dept_map:
        dept = get_dept_norm(r)
        seller = str(r.get('销售员名称', '') or '').strip()
        if dept:
            order_dept_map[order_no] = dept
            if seller:
                order_seller_map[order_no] = seller

# ============================
# 在职销售员统计：从26财年Q1业绩 + 欠款数据两个sheet收集，剔除支振岗（李国栋、白雨）
# ============================
# 先读取欠款数据（供在职销售员统计和欠款处理共用）
_, debt_rows = read_sheet_data(wb, 2)

# 再从欠款数据sheet补充（包含在Sheet1中无业绩但在职的销售员，如张宸睿）
for r in debt_rows:
    dept = get_dept_norm(r)
    if dept is None:
        continue
    seller = str(r.get('销售员名称', '') or '').strip()
    status = str(r.get('销售员状态', '') or '').strip()
    if seller and status == '在职' and seller not in ('李国栋', '白雨'):
        dept_active_sellers[dept].add(seller)

dept_sales_count = {k: len(v) for k, v in dept_active_sellers.items()}
total_active_sellers = len(set().union(*dept_active_sellers.values()))

# ============================
# 欠款数据处理（修正版：按天数分类）
# ============================

dept_debt_d30 = defaultdict(float)
dept_debt_d30_90 = defaultdict(float)
dept_debt_d90_180 = defaultdict(float)
dept_debt_d180 = defaultdict(float)
dept_debt_total = defaultdict(float)
dept_debt_days_weighted = defaultdict(float)

sales_debt_d30 = defaultdict(lambda: defaultdict(float))
sales_debt_d30_90 = defaultdict(lambda: defaultdict(float))
sales_debt_d90_180 = defaultdict(lambda: defaultdict(float))
sales_debt_d180 = defaultdict(lambda: defaultdict(float))
sales_debt_total = defaultdict(lambda: defaultdict(float))

risk_customer_raw = defaultdict(lambda: {'amount': 0, 'dept': ''})

for r in debt_rows:
    dept = get_dept_norm(r)
    if dept is None:
        continue
    cycle = str(r.get('欠款周期', '') or '').strip()
    amt = safe_float(r.get('欠款金额', 0))
    days = safe_float(r.get('欠款天数', 0))
    seller = str(r.get('销售员名称', '') or '').strip()
    customer = str(r.get('客户名称', '') or '').strip()
    
    dept_debt_total[dept] += amt
    dept_debt_days_weighted[dept] += amt * days
    if seller:
        sales_debt_total[dept][seller] += amt
    
    # 按天数分类
    if cycle == '30天内':
        dept_debt_d30[dept] += amt
        if seller:
            sales_debt_d30[dept][seller] += amt
    elif cycle == '180天+':
        dept_debt_d180[dept] += amt
        if seller:
            sales_debt_d180[dept][seller] += amt
        if customer:
            risk_customer_raw[customer]['amount'] += amt
            risk_customer_raw[customer]['dept'] = dept
    elif cycle == '30天+':
        if days <= 90:
            dept_debt_d30_90[dept] += amt
            if seller:
                sales_debt_d30_90[dept][seller] += amt
        else:
            # 90-180天
            dept_debt_d90_180[dept] += amt
            if seller:
                sales_debt_d90_180[dept][seller] += amt
            if customer:
                risk_customer_raw[customer]['amount'] += amt
                risk_customer_raw[customer]['dept'] = dept

def wan(x):
    return round(x / 10000, 2)

def to_wan_dict(d):
    return {k: wan(v) for k, v in d.items()}

dept_debt_d30_w = to_wan_dict(dept_debt_d30)
dept_debt_d30_90_w = to_wan_dict(dept_debt_d30_90)
dept_debt_d90_180_w = to_wan_dict(dept_debt_d90_180)
dept_debt_d180_w = to_wan_dict(dept_debt_d180)
dept_debt_total_w = to_wan_dict(dept_debt_total)

# 高风险客户排名
risk_list = sorted([
    {'customer': k, 'dept': v['dept'], 'amount': wan(v['amount'])}
    for k, v in risk_customer_raw.items()
    if v['amount'] > 0
], key=lambda x: -x['amount'])[:15]

# ============================
# 认款数据（回款）
# ============================
_, rec_rows = read_sheet_data(wb, 3)
dept_rec_total = defaultdict(float)
dept_rec_days_weighted = defaultdict(float)
sales_rec_total = defaultdict(lambda: defaultdict(float))

for r in rec_rows:
    amt = safe_float(r.get('认款协同金额', 0))
    if amt <= 0:
        continue
    
    dept = None
    d1 = str(r.get('一级部门', '') or '').strip()
    d3 = str(r.get('三级部门', '') or '').strip()
    d2 = str(r.get('二级部门', '') or '').strip()
    if d1 == '中西部大区':
        if d3 and d3 != 'None':
            dept = normalize_dept(d3)
        elif d2 and d2 != 'None':
            dept = normalize_dept(d2)
    
    if dept is None:
        order_no = str(r.get('业绩单号', '') or '').strip()
        dept = order_dept_map.get(order_no)
    if dept is None:
        continue
    
    seller = str(r.get('销售员名称', '') or '').strip()
    if not seller:
        order_no = str(r.get('业绩单号', '') or '').strip()
        seller = order_seller_map.get(order_no, '')
    
    dept_rec_total[dept] += amt
    if seller:
        sales_rec_total[dept][seller] += amt
    
    perf_date = r.get('业绩日期')
    rec_date = r.get('认款时间') or r.get('回款日期')
    if perf_date and rec_date:
        try:
            if not hasattr(perf_date, 'date'):
                perf_date = datetime.strptime(str(perf_date)[:10], '%Y-%m-%d')
            if not hasattr(rec_date, 'date'):
                rec_date = datetime.strptime(str(rec_date)[:10], '%Y-%m-%d')
            diff = (rec_date - perf_date).days
            if diff >= 0:
                dept_rec_days_weighted[dept] += amt * diff
        except:
            pass

dept_rec_total_w = to_wan_dict(dept_rec_total)

# ============================
# 回款周期计算
# ============================
dept_cycle = {}
for dept in set(list(dept_debt_total.keys()) + list(dept_rec_total.keys())):
    dbt = dept_debt_total.get(dept, 0)
    rec = dept_rec_total.get(dept, 0)
    dbt_w = dept_debt_days_weighted.get(dept, 0)
    rec_w = dept_rec_days_weighted.get(dept, 0)
    total = dbt + rec
    if total > 0:
        dept_cycle[dept] = round((dbt_w + rec_w) / total, 1)

total_dbt_raw = sum(dept_debt_total.values())
total_rec_raw = sum(dept_rec_total.values())
total_cycle = round(
    (sum(dept_debt_days_weighted.values()) + sum(dept_rec_days_weighted.values())) /
    (total_dbt_raw + total_rec_raw), 1
) if (total_dbt_raw + total_rec_raw) > 0 else 0

# ============================
# 总计KPI
# ============================
total_perf26 = round(sum(dept_perf26.get(d, 0) for d in ALL_DEPTS), 2)
total_perf25 = round(sum(dept_perf25.get(d, 0) for d in ALL_DEPTS), 2)
total_target = dept_targets.get('中西部大区', 5586.0)
total_completion = round(total_perf26 / total_target * 100, 1) if total_target else 0
total_yoy = round((total_perf26 - total_perf25) / total_perf25 * 100, 1) if total_perf25 else 0
total_debt = round(sum(dept_debt_total_w.get(d, 0) for d in ALL_DEPTS), 2)
total_d30 = round(sum(dept_debt_d30_w.get(d, 0) for d in ALL_DEPTS), 2)
total_d30_90 = round(sum(dept_debt_d30_90_w.get(d, 0) for d in ALL_DEPTS), 2)
total_d90_180 = round(sum(dept_debt_d90_180_w.get(d, 0) for d in ALL_DEPTS), 2)
total_d180 = round(sum(dept_debt_d180_w.get(d, 0) for d in ALL_DEPTS), 2)
total_rec = round(sum(dept_rec_total_w.get(d, 0) for d in ALL_DEPTS), 2)

# 获取数据日期
max_date = None
for r in perf26_rows:
    d = r.get('业绩日期')
    if d and hasattr(d, 'date'):
        if max_date is None or d > max_date:
            max_date = d
data_date = max_date.strftime('%Y-%m-%d') if max_date else '2026-06-08'
today = datetime.now().strftime('%Y-%m-%d')

print(f"26Q1业绩: {total_perf26}万, 25Q1: {total_perf25}万")
print(f"目标: {total_target}万, 完成率: {total_completion}%")
print(f"同比: {total_yoy}%")
print(f"欠款: {total_debt}万 (30天内:{total_d30}, 30-90天:{total_d30_90}, 90-180天:{total_d90_180}, 180天+:{total_d180})")
print(f"回款: {total_rec}万")
print(f"在职销售员: {total_active_sellers}")
print(f"平均回款周期: {total_cycle}天")
print(f"各部门周期: {dict(sorted(dept_cycle.items(), key=lambda x: -x[1]))}")

# ============================
# 部门数据
# ============================
dept_data = []
for dept in ALL_DEPTS:
    v26 = dept_perf26.get(dept, 0)
    v25 = dept_perf25.get(dept, 0)
    yoy = round((v26 - v25) / v25 * 100, 1) if v25 else 0
    target = dept_targets.get(dept, 0)
    completion = round(v26 / target * 100, 1) if target else 0
    sales = dept_sales_count.get(dept, 0)
    d30 = dept_debt_d30_w.get(dept, 0)
    d30_90 = dept_debt_d30_90_w.get(dept, 0)
    d90_180 = dept_debt_d90_180_w.get(dept, 0)
    d180 = dept_debt_d180_w.get(dept, 0)
    total_d = dept_debt_total_w.get(dept, 0)
    collect = dept_rec_total_w.get(dept, 0)
    cycle = dept_cycle.get(dept, 0)
    
    if v26 == 0 and total_d == 0 and v25 == 0:
        continue
    
    dept_data.append({
        'dept': dept,
        'v26': v26, 'v25': v25, 'yoy': yoy,
        'target': round(target, 1), 'completion': completion,
        'sales': sales,
        'd30': d30, 'd30_90': d30_90, 'd90_180': d90_180, 'd180': d180,
        'total_debt': total_d, 'collect': collect, 'cycle': cycle
    })

# ============================
# 销售员业绩明细
# ============================
sales_detail_data = {}
for dept in ALL_DEPTS:
    sellers = set()
    for d in [sales_perf26, sales_collect26, sales_debt_total]:
        sellers.update(d[dept].keys())
    detail = []
    for seller in sellers:
        perf = wan(sales_perf26[dept].get(seller, 0))
        collect = wan(sales_collect26[dept].get(seller, 0))
        d30 = wan(sales_debt_d30[dept].get(seller, 0))
        d30_90 = wan(sales_debt_d30_90[dept].get(seller, 0))
        d90_180 = wan(sales_debt_d90_180[dept].get(seller, 0))
        d180 = wan(sales_debt_d180[dept].get(seller, 0))
        total_dbt = wan(sales_debt_total[dept].get(seller, 0))
        if perf != 0 or total_dbt != 0:
            detail.append({
                'name': seller, 'perf': perf, 'collect': collect,
                'debt': total_dbt, 'orders': 0,
                'd30': d30, 'd30_90': d30_90, 'd90_180': d90_180, 'd180': d180,
                'total_debt': total_dbt, 'collect_amt': collect,
            })
    if detail:
        sales_detail_data[dept] = sorted(detail, key=lambda x: -x['perf'])

# ============================
# 销售员回款周期明细（从认款数据重新计算）
# ============================
seller_debt_weighted = defaultdict(lambda: defaultdict(float))
seller_debt_amt = defaultdict(lambda: defaultdict(float))
for r in debt_rows:
    dept = get_dept_norm(r)
    if dept is None:
        continue
    seller = str(r.get('销售员名称', '') or '').strip()
    if not seller:
        continue
    amt = safe_float(r.get('欠款金额', 0))
    days = safe_float(r.get('欠款天数', 0))
    seller_debt_weighted[dept][seller] += amt * days
    seller_debt_amt[dept][seller] += amt

seller_rec_amt = defaultdict(lambda: defaultdict(float))
seller_rec_weighted = defaultdict(lambda: defaultdict(float))
for r in rec_rows:
    amt = safe_float(r.get('认款协同金额', 0))
    if amt <= 0:
        continue
    dept = None
    d1 = str(r.get('一级部门', '') or '').strip()
    d3 = str(r.get('三级部门', '') or '').strip()
    d2 = str(r.get('二级部门', '') or '').strip()
    if d1 == '中西部大区':
        if d3 and d3 != 'None':
            dept = normalize_dept(d3)
        elif d2 and d2 != 'None':
            dept = normalize_dept(d2)
    if dept is None:
        order_no = str(r.get('业绩单号', '') or '').strip()
        dept = order_dept_map.get(order_no)
    if dept is None:
        continue
    seller = str(r.get('销售员名称', '') or '').strip()
    if not seller:
        order_no = str(r.get('业绩单号', '') or '').strip()
        seller = order_seller_map.get(order_no, '')
    if not seller:
        continue
    seller_rec_amt[dept][seller] += amt
    perf_date = r.get('业绩日期')
    rec_date = r.get('认款时间') or r.get('回款日期')
    if perf_date and rec_date:
        try:
            if not hasattr(perf_date, 'date'):
                perf_date = datetime.strptime(str(perf_date)[:10], '%Y-%m-%d')
            if not hasattr(rec_date, 'date'):
                rec_date = datetime.strptime(str(rec_date)[:10], '%Y-%m-%d')
            diff = (rec_date - perf_date).days
            if diff >= 0:
                seller_rec_weighted[dept][seller] += amt * diff
        except:
            pass

sales_cycle_detail = {}
for dept in ALL_DEPTS:
    sellers = set(list(seller_debt_amt[dept].keys()) + list(seller_rec_amt[dept].keys()))
    detail = []
    for seller in sellers:
        dbt = seller_debt_amt[dept].get(seller, 0)
        rec = seller_rec_amt[dept].get(seller, 0)
        dbt_w = seller_debt_weighted[dept].get(seller, 0)
        rec_w = seller_rec_weighted[dept].get(seller, 0)
        total = dbt + rec
        if total > 0:
            cycle = round((dbt_w + rec_w) / total, 1)
        else:
            cycle = 0
        if dbt > 0 or rec > 0:
            detail.append({
                'name': seller,
                'debt_amt': wan(dbt),
                'rec_amt': wan(rec),
                'cycle': cycle,
                'debt_weighted': round(dbt_w, 0),
                'rec_weighted': round(rec_w, 0),
            })
    if detail:
        sales_cycle_detail[dept] = sorted(detail, key=lambda x: -x['cycle'])

# ============================
# 保存结果
# ============================
result = {
    'data_date': data_date,
    'today': today,
    'kpi': {
        'total_perf26': total_perf26,
        'total_target': total_target,
        'total_completion': total_completion,
        'total_yoy': total_yoy,
        'total_perf25': total_perf25,
        'total_active_sellers': total_active_sellers,
        'total_debt': total_debt,
        'debt_d90_plus': round(total_d90_180 + total_d180, 2),
    },
    'debt_kpi': {
        'total': total_debt,
        'd30': total_d30,
        'd30_ratio': round(total_d30 / total_debt * 100, 1) if total_debt else 0,
        'd30_90': total_d30_90,
        'd30_90_ratio': round(total_d30_90 / total_debt * 100, 1) if total_debt else 0,
        'd90_180': total_d90_180,
        'd180': total_d180,
    },
    'cycle_kpi': {
        'avg_cycle': total_cycle,
        'max_cycle_dept': max(dept_cycle.items(), key=lambda x: x[1])[0] if dept_cycle else '',
        'max_cycle': max(dept_cycle.values()) if dept_cycle else 0,
        'min_cycle_dept': min(dept_cycle.items(), key=lambda x: x[1])[0] if dept_cycle else '',
        'min_cycle': min(dept_cycle.values()) if dept_cycle else 0,
        'over90_depts': [k for k, v in dept_cycle.items() if v > 90],
        'over90_count': len([k for k, v in dept_cycle.items() if v > 90]),
    },
    'dept_data': dept_data,
    'sales_detail_data': sales_detail_data,
    'sales_cycle_detail': sales_cycle_detail,
    'risk_customers': risk_list,
}

OUTPUT_PATH = os.path.join(BASE_DIR, 'dashboard_data.json')
with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, indent=2, default=str)

print(f"\n✅ 数据已保存 → {OUTPUT_PATH}")
print(f"  部门数: {len(dept_data)}")
print(f"  高风险客户: {len(risk_list)}")
