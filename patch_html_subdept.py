#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""给 index.html 中三个销售员表格增加三级部门列"""
import re

with open('index.html', 'r', encoding='utf-8') as f:
    content = f.read()

# ========== 1. renderSalesPerf ==========
# 表头：在"销售员</th>" 后增加 "三级部门</th>"
old = '<th style="padding:8px 6px;color:#00d4ff;">销售员</th>\n            <th style="padding:8px 6px;color:#00d4ff;text-align:right;">26Q1业绩(万)</th>'
new = '<th style="padding:8px 6px;color:#00d4ff;">销售员</th>\n            <th style="padding:8px 6px;color:#00d4ff;">三级部门</th>\n            <th style="padding:8px 6px;color:#00d4ff;text-align:right;">26Q1业绩(万)</th>'
content = content.replace(old, new)

# 行模板：在销售员名单元格后增加三级部门单元格
old = '${s.name}${hasCust ? \' 📂\' : \'\'}</td>\n            <td style="padding:7px 5px;color:${color};text-align:right;font-weight:600;">${s.perf.toFixed(2)}</td>'
new = '${s.name}${hasCust ? \' 📂\' : \'\'}</td>\n            <td style="padding:7px 5px;color:#8892b0;">${s.sub_dept || \'其他\'}</td>\n            <td style="padding:7px 5px;color:${color};text-align:right;font-weight:600;">${s.perf.toFixed(2)}</td>'
content = content.replace(old, new)

# 合计行：在"合计"单元格后增加一个空单元格（对应三级部门列）
old = '<td style="padding:8px 6px;color:#00d4ff;">合计（${list.length}人）</td>\n            <td style="padding:8px 6px;color:#00ff88;text-align:right;">${tp}</td>'
new = '<td style="padding:8px 6px;color:#00d4ff;">合计（${list.length}人）</td>\n            <td style="padding:8px 6px;color:#8892b0;">—</td>\n            <td style="padding:8px 6px;color:#00ff88;text-align:right;">${tp}</td>'
content = content.replace(old, new)

# ========== 2. renderSalesDebt ==========
# 表头
old = '<th style="padding:8px 6px;color:#00d4ff;">销售员</th>\n            <th style="padding:8px 6px;color:#00d4ff;text-align:right;">合计欠款(万)</th>'
new = '<th style="padding:8px 6px;color:#00d4ff;">销售员</th>\n            <th style="padding:8px 6px;color:#00d4ff;">三级部门</th>\n            <th style="padding:8px 6px;color:#00d4ff;text-align:right;">合计欠款(万)</th>'
content = content.replace(old, new)

# 行模板
old = '${s.name}${hasCust ? \' 📂\' : \'\'}</td>\n            <td style="padding:7px 5px;color:${dc};text-align:right;font-weight:600;">${s.total_debt.toFixed(2)}</td>'
new = '${s.name}${hasCust ? \' 📂\' : \'\'}</td>\n            <td style="padding:7px 5px;color:#8892b0;">${s.sub_dept || \'其他\'}</td>\n            <td style="padding:7px 5px;color:${dc};text-align:right;font-weight:600;">${s.total_debt.toFixed(2)}</td>'
content = content.replace(old, new)

# 合计行
old = '<td style="padding:8px 6px;color:#00d4ff;">合计（${list.length}人）</td>\n            <td style="padding:8px 6px;color:#ff4757;text-align:right;">${td}</td>'
new = '<td style="padding:8px 6px;color:#00d4ff;">合计（${list.length}人）</td>\n            <td style="padding:8px 6px;color:#8892b0;">—</td>\n            <td style="padding:8px 6px;color:#ff4757;text-align:right;">${td}</td>'
content = content.replace(old, new)

# ========== 3. renderSalesCycle ==========
# 表头
old = '<th style="padding:8px 6px;color:#00d4ff;">销售员</th>\n            <th style="padding:8px 6px;color:#00d4ff;text-align:right;">认款金额(万)</th>'
new = '<th style="padding:8px 6px;color:#00d4ff;">销售员</th>\n            <th style="padding:8px 6px;color:#00d4ff;">三级部门</th>\n            <th style="padding:8px 6px;color:#00d4ff;text-align:right;">认款金额(万)</th>'
content = content.replace(old, new)

# 行模板
old = '${s.name}${hasCust ? \' 📂\' : \'\'}</td>\n            <td style="padding:7px 5px;color:#00ff88;text-align:right;">${(s.rec_amt/10000).toFixed(2)}</td>'
new = '${s.name}${hasCust ? \' 📂\' : \'\'}</td>\n            <td style="padding:7px 5px;color:#8892b0;">${s.sub_dept || \'其他\'}</td>\n            <td style="padding:7px 5px;color:#00ff88;text-align:right;">${(s.rec_amt/10000).toFixed(2)}</td>'
content = content.replace(old, new)

# 合计行
old = '<td style="padding:8px 6px;color:#00d4ff;">合计（${list.length}人）</td>\n            <td style="padding:8px 6px;color:#00ff88;text-align:right;">${(tc/10000).toFixed(2)}</td>'
new = '<td style="padding:8px 6px;color:#00d4ff;">合计（${list.length}人）</td>\n            <td style="padding:8px 6px;color:#8892b0;">—</td>\n            <td style="padding:8px 6px;color:#00ff88;text-align:right;">${(tc/10000).toFixed(2)}</td>'
content = content.replace(old, new)

with open('index.html', 'w', encoding='utf-8') as f:
    f.write(content)

print('Done! Modified index.html to add sub_dept column.')

# 验证
with open('index.html', 'r', encoding='utf-8') as f:
    c = f.read()

checks = [
    ('renderSalesPerf header', '<th style="padding:8px 6px;color:#00d4ff;">三级部门</th>' in c and 'renderSalesPerf' in c),
    ('renderSalesDebt header', '<th style="padding:8px 6px;color:#00d4ff;">三级部门</th>' in c and 'renderSalesDebt' in c),
    ('renderSalesCycle header', '<th style="padding:8px 6px;color:#00d4ff;">三级部门</th>' in c and 'renderSalesCycle' in c),
    ('sub_dept in Perf rows', 's.sub_dept' in c),
]

for name, ok in checks:
    print(f'  {"OK" if ok else "FAIL"}: {name}')

if all(ok for _, ok in checks):
    print('\nSUCCESS: All replacements done!')
else:
    print('\nWARNING: Some replacements may have failed.')
