import re

with open('gen_modal_dashboard.py', 'r', encoding='utf-8') as f:
    content = f.read()

# 1. renderSalesPerf 表头：在'销售员</th>' 后增加 '三级部门</th>'
old1 = '<th style="padding:8px 6px;color:#00d4ff;">销售员</th>\n            <th style="padding:8px 6px;color:#00d4ff;text-align:right;">26Q1业绩(万)</th>'
new1 = '<th style="padding:8px 6px;color:#00d4ff;">销售员</th>\n            <th style="padding:8px 6px;color:#00d4ff;">三级部门</th>\n            <th style="padding:8px 6px;color:#00d4ff;text-align:right;">26Q1业绩(万)</th>'
content = content.replace(old1, new1)

# 1b. renderSalesPerf 行模板：在销售员名单元格后增加三级部门单元格
old1b = '${s.name}${hasCust ? \' 📂\' : \'\'}</td>\n            <td style="padding:7px 5px;color:${color};text-align:right;font-weight:600;">${s.perf.toFixed(2)}</td>'
new1b = '${s.name}${hasCust ? \' 📂\' : \'\'}</td>\n            <td style="padding:7px 5px;color:#8892b0;">${s.sub_dept || \'其他\'}</td>\n            <td style="padding:7px 5px;color:${color};text-align:right;font-weight:600;">${s.perf.toFixed(2)}</td>'
content = content.replace(old1b, new1b)

# 2. renderSalesDebt 表头
old2 = '<th style="padding:8px 6px;color:#00d4ff;">销售员</th>\n            <th style="padding:8px 6px;color:#00d4ff;text-align:right;">合计欠款(万)</th>'
new2 = '<th style="padding:8px 6px;color:#00d4ff;">销售员</th>\n            <th style="padding:8px 6px;color:#00d4ff;">三级部门</th>\n            <th style="padding:8px 6px;color:#00d4ff;text-align:right;">合计欠款(万)</th>'
content = content.replace(old2, new2)

# 2b. renderSalesDebt 行模板
old2b = '${s.name}${hasCust ? \' 📂\' : \'\'}</td>\n            <td style="padding:7px 5px;color:${dc};text-align:right;font-weight:600;">${s.total_debt.toFixed(2)}</td>'
new2b = '${s.name}${hasCust ? \' 📂\' : \'\'}</td>\n            <td style="padding:7px 5px;color:#8892b0;">${s.sub_dept || \'其他\'}</td>\n            <td style="padding:7px 5px;color:${dc};text-align:right;font-weight:600;">${s.total_debt.toFixed(2)}</td>'
content = content.replace(old2b, new2b)

# 3. renderSalesCycle 表头
old3 = '<th style="padding:8px 6px;color:#00d4ff;">销售员</th>\n            <th style="padding:8px 6px;color:#00d4ff;text-align:right;">认款金额(万)</th>'
new3 = '<th style="padding:8px 6px;color:#00d4ff;">销售员</th>\n            <th style="padding:8px 6px;color:#00d4ff;">三级部门</th>\n            <th style="padding:8px 6px;color:#00d4ff;text-align:right;">认款金额(万)</th>'
content = content.replace(old3, new3)

# 3b. renderSalesCycle 行模板
old3b = '${s.name}${hasCust ? \' 📂\' : \'\'}</td>\n            <td style="padding:7px 5px;color:#00ff88;text-align:right;">${(s.rec_amt/10000).toFixed(2)}</td>'
new3b = '${s.name}${hasCust ? \' 📂\' : \'\'}</td>\n            <td style="padding:7px 5px;color:#8892b0;">${s.sub_dept || \'其他\'}</td>\n            <td style="padding:7px 5px;color:#00ff88;text-align:right;">${(s.rec_amt/10000).toFixed(2)}</td>'
content = content.replace(old3b, new3b)

with open('gen_modal_dashboard.py', 'w', encoding='utf-8') as f:
    f.write(content)

print('Done! Modified gen_modal_dashboard.py to add sub_dept column.')

# 验证替换是否成功
with open('gen_modal_dashboard.py', 'r', encoding='utf-8') as f:
    new_content = f.read()

print('Checking replacements...')
if old1 in new_content:
    print('WARNING: old1 still found (replacement may have failed)')
else:
    print('Replacement 1 (renderSalesPerf header) OK')

if old1b in new_content:
    print('WARNING: old1b still found')
else:
    print('Replacement 1b (renderSalesPerf row) OK')

if old2 in new_content:
    print('WARNING: old2 still found')
else:
    print('Replacement 2 (renderSalesDebt header) OK')

if old2b in new_content:
    print('WARNING: old2b still found')
else:
    print('Replacement 2b (renderSalesDebt row) OK')

if old3 in new_content:
    print('WARNING: old3 still found')
else:
    print('Replacement 3 (renderSalesCycle header) OK')

if old3b in new_content:
    print('WARNING: old3b still found')
else:
    print('Replacement 3b (renderSalesCycle row) OK')

# 检查是否包含新三級部门列
if '三级部门' in new_content:
    print('SUCCESS: "三级部门" column found in file!')
else:
    print('WARNING: "三级部门" column NOT found in file!')
