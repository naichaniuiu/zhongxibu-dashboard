# 中西部大区数据看板 - 使用说明

## 管理者访问地址（固定链接）

**https://naichaniuiu.github.io/zhongxibu-dashboard/**

> 收藏此链接，每天打开即可看到最新数据。GitHub Pages 更新后约 1-2 分钟刷新。

---

## 数据更新方式（每日自动 + 手动兜底）

### 方式一：每日自动更新（推荐，无需操作）

WorkBuddy 通过 **Windows 计划任务** 在每天早上 7:30 自动跑完整流程：

1. 读取 `D:\业绩 欠款看板 Q2.xlsx`
2. 生成看板 JSON + HTML（含所有图表和交互）
3. 推送到 GitHub Pages

**前提条件**：
- 电脑在 7:30 时处于开机状态（错过可手动触发 `schtasks /Run /TN ZhongxibuQ2DashboardUpdate`）
- 已用当前 Windows 用户登录（SSH 密钥在 `%USERPROFILE%\.ssh\`）

**查看日志**：每次运行在 `deploy\update_log_YYYYMMDD.txt` 留一份记录。

**查看任务**：`schtasks /Query /TN ZhongxibuQ2DashboardUpdate /V`

---

### 方式二：手动立即更新（紧急时使用）

如需立刻同步最新 Excel（例如领导临时要看数据）：

1. 确保最新 Excel 文件在 `D:\业绩 欠款看板 Q2.xlsx`
2. 在项目根目录双击 `auto_update.bat`
   （路径：`C:\Users\1\WorkBuddy\2026-09-29-14-07-32\auto_update.bat`）
3. 等待几秒，浏览器打开 https://naichaniuiu.github.io/zhongxibu-dashboard/ 查看

---

## Excel 文件路径

**当前**：`D:\业绩 欠款看板 Q2.xlsx`

如需修改，请告诉我新路径，我同步更新 `auto_update.bat`。

---

## 文件说明（项目仓库结构）

| 路径 | 说明 |
|------|------|
| `deploy/index.html` | 看板页面（GitHub Pages 入口，已自包含 1MB） |
| `deploy/README.md` | 本文件 |
| `deploy/.gitignore` | 排除本机日志等 |

> **不在本仓库**（每台机器专属，不上传）：
> - `build_data.py` / `inject.py` / `dashboard_template.html` 等 Python 脚本
> - `auto_update.bat`（含本机绝对路径，Windows 计划任务入口）
> - `update_log_*.txt`（运行日志）

---

## 常见问题

**Q：管理者打不开链接？**
A：确认链接是 `https://naichaniuiu.github.io/zhongxibu-dashboard/`

**Q：自动更新没执行？**
A：
1. `schtasks /Query /TN ZhongxibuQ2DashboardUpdate` 看任务状态
2. 确认 7:30 时电脑开机且已登录
3. 看 `deploy\update_log_YYYYMMDD.txt` 最近一次日志
4. 手动触发：`schtasks /Run /TN ZhongxibuQ2DashboardUpdate`

**Q：Excel 路径能改吗？**
A：可以，但需要同步修改 `auto_update.bat` 里 `EXCEL` 那行

**Q：定时任务能改时间吗？**
A：可以，`schtasks /Change /TN ZhongxibuQ2DashboardUpdate /ST HH:MM`