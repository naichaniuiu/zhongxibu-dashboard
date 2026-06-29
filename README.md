# 中西部大区数据看板 - 使用说明

## 管理者访问地址（固定链接）

**https://naichaniuiu.github.io/zhongxibu-dashboard/**

> 收藏此链接，每天打开即可看到最新数据。GitHub Pages 更新后约 1-2 分钟刷新。

---

## 每日更新操作（你执行）

### 步骤 1：下载最新 Excel
将最新的 `业绩 欠款看板.xlsx` 下载到：
```
C:\Users\wm881\Downloads\业绩 欠款看板.xlsx
```
（覆盖旧文件即可）

### 步骤 2：双击更新脚本
打开文件夹：
```
C:\Users\wm881\WorkBuddy\2026-06-09-17-16-53\zhongxibu-dashboard\
```
**双击 `一键更新.vbs`**

等待弹窗提示 **"Update Done"** → 点击确定即可。

脚本会自动完成：
1. 读取 Excel 数据
2. 生成最新看板 HTML
3. 推送到 GitHub（管理者页面自动更新）

---

## 文件说明

| 文件 | 说明 |
|------|------|
| `index.html` | 看板页面（GitHub Pages 入口） |
| `process_data_v2.py` | Excel 数据提取脚本 |
| `extract_customers.py` | 客户维度数据提取脚本 |
| `gen_modal_dashboard.py` | 看板 HTML 生成脚本 |
| `一键更新.vbs` | 一键更新脚本（双击运行） |
| `update_log.txt` | 更新日志（出错时查看） |

---

## 常见问题

**Q：管理者无法打开链接？**
A：检查链接是否为 `https://naichaniuiu.github.io/zhongxibu-dashboard/`，GitHub Pages 首次启用可能需要 5-10 分钟生效。

**Q：双击更新脚本没反应？**
A：查看同目录下的 `update_log.txt` 文件，里面有详细错误信息。

**Q：Excel 文件路径可以改吗？**
A：目前固定为 `C:\Users\wm881\Downloads\业绩 欠款看板.xlsx`，如需修改请告诉我。

**Q：Token 过期了怎么办？**
A：重新生成 GitHub Token，然后修改 `一键更新.vbs` 中的 Token 部分，告诉我帮你改。
