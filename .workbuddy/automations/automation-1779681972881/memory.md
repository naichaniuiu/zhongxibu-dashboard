# Automation Execution Memory

## 2026-08-25 (Tue) 09:50
- Q1 dashboard (20260513090923): Steps 1-4 all ran successfully. generate_dashboard.py extracted data, generate_html.py generated HTML, cycle drill and debt/perf drill injected. Git push to sales-dashboard repo FAILED due to GitHub Push Protection (detected PAT in push_q2_dashboard.py/push_via_api.py history). Working tree clean, no new data changes.
- Q2 dashboard (zhongxibu-dashboard): gen_q2_dashboard.py ran successfully with TODAY=2026-08-25. Remote was already at 2da3fb2 with today's data. After reset --hard to remote SHA and regenerate, working tree was clean (data identical). No new commit needed. GitHub Pages already up to date.
- KPI: 26Q2 actual=1110.09wan, 25Q2(B-end filtered)=4463.94wan, total debt=2035.20wan, overdue(>30d)=1539.03wan, active sellers=59, avg cycle=63.2d
- Note: Remote SHA changed between ls-remote (dc31734) and fetch (2da3fb2). Always use actual fetch result for reset target.
