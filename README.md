# 中西部大区数据看板 - 使用说明

## 管理者访问地址（固定链接）

**https://naichaniuiu.github.io/zhongxibu-dashboard/**

> 收藏此链接，每天打开即可看到最新数据。GitHub Pages 更新后约 1-2 分钟刷新。

---

## 每日更新方式

### 方式一：WorkBuddy 自动更新（推荐）

已设置每日定时自动更新，无需手动操作。WorkBuddy 会在每天指定时间自动：
1. 读取 `D:\业绩 欠款看板 Q2.xlsx`
2. 运行 `gen_q2_dashboard.py` 生成看板 HTML
3. 推送到 GitHub Pages

> 前提：电脑开机 + WorkBuddy 运行中 + Excel 文件已更新

### 方式二：手动一键更新

1. 确保最新 Excel 文件在 `D:\业绩 欠款看板 Q2.xlsx`
2. 打开此文件夹，**双击 `一键更新Q2.vbs`**
3. 等待弹窗提示 **"Update Done"** → 点击确定

---

## 文件说明

| 文件 | 说明 |
|------|------|
| `index.html` | 看板页面（GitHub Pages 入口） |
| `gen_q2_dashboard.py` | Q2 看板生成脚本（读取 Excel → 生成 HTML） |
| `一键更新Q2.vbs` | Q2 一键更新脚本（双击运行） |

---

## 常见问题

**Q：管理者无法打开链接？**
A：检查链接是否为 `https://naichaniuiu.github.io/zhongxibu-dashboard/`

**Q：自动更新没有执行？**
A：确保电脑开机、WorkBuddy 运行中。如需手动更新，双击 `一键更新Q2.vbs`

**Q：Excel 文件路径可以改吗？**
A：目前固定为 `D:\业绩 欠款看板 Q2.xlsx`，如需修改请告诉我
