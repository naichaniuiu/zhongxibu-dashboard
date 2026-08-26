# 项目长期记忆 - 中西部大区销售看板 (Q2)

## 数据口径（2026-08-26 更新）
- **业绩**：只统计"是否核算B端业绩=是"；**包含离职销售员**（2026-08-26 用户确认，此前剔除）
- **欠款**：整个文件所有行直接求和；**包含离职销售员**（2026-08-26 用户确认，此前剔除）；石红银仍排除（当前文件 0 行）
- **25Q2 对比数据**：含离职（4780.23万）；表无 B 端列
- **在职销售**：用"在职销售人数"sheet（HR数据）
- **逾期** = 欠款天数 > 30天
- **回款周期** = Σ(认款金额×账龄) ÷ 回款总金额
- TODAY 基准日需每日更新为当天日期

## 项目结构
- 仓库：naichaniuiu/zhongxibu-dashboard，主分支 main，GitHub Pages: https://naichaniuiu.github.io/zhongxibu-dashboard/
- 数据源：D:/业绩 欠款看板 Q2.xlsx（6个sheet: 25Q2业绩/26Q2业绩/欠款数据/认款数据/目标拆分/在职销售人数）
- 核心脚本：gen_q2_dashboard.py（Excel → HTML），输出 中西部大区26财年Q2数据看板_弹窗下钻版.html → 覆盖 index.html
- 推送命令：CRED=$(head -1 ~/.git-credentials); GIT_TERMINAL_PROMPT=0 git -c credential.helper= push "${CRED}/naichaniuiu/zhongxibu-dashboard.git" main
  （必须加 -c credential.helper= 防止 GCM 挂起）
- Q1 看板在 C:\Users\wm881\WorkBuddy\20260513090923，主分支 master，仓库 naichaniuiu/sales-dashboard

## 用户偏好
- 看板数据需与 Excel 全量求和一致；发现差异先查根因再修复
- 偏好表格展示数据对比和分步状态汇总
