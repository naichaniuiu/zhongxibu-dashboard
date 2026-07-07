#!/bin/bash
# ============================================================
#  中西部大区 Q2 看板 - 每日自动更新脚本
#  1. 运行 gen_q2_dashboard.py 生成看板 HTML
#  2. 复制 HTML 到 git 仓库
#  3. Git commit + push 到 GitHub Pages
# ============================================================

set -e

PYTHON="C:/Users/wm881/.workbuddy/binaries/python/envs/default/Scripts/python.exe"
SCRIPT="C:/Users/wm881/WorkBuddy/2026-07-07-10-45-23/gen_q2_dashboard.py"
REPO="C:/Users/wm881/WorkBuddy/zhongxibu-dashboard"
HTML_SRC="C:/Users/wm881/WorkBuddy/2026-07-07-10-45-23/中西部大区26财年Q2数据看板_弹窗下钻版.html"
HTML_DST="$REPO/index.html"
DATA_FILE="D:/业绩 欠款看板 Q2.xlsx"

echo "=== Q2 Dashboard Update: $(date) ==="

# Step 1: Check data file
if [ ! -f "$DATA_FILE" ]; then
    echo "ERROR: Data file not found: $DATA_FILE"
    exit 1
fi
echo "[1/4] Data file found: $DATA_FILE"

# Step 2: Generate dashboard HTML
echo "[2/4] Generating dashboard HTML..."
"$PYTHON" "$SCRIPT"
if [ $? -ne 0 ]; then
    echo "ERROR: Failed to generate dashboard HTML"
    exit 1
fi

# Step 3: Copy HTML to repo
echo "[3/4] Copying HTML to git repo..."
cp "$HTML_SRC" "$HTML_DST"
echo "  Copied: $HTML_SRC -> $HTML_DST"

# Step 4: Git commit and push
echo "[4/4] Pushing to GitHub..."
cd "$REPO"
git add -A

# Check if there are changes to commit
if git diff --cached --quiet; then
    echo "No changes to commit. Data is the same as previous update."
    exit 0
fi

git commit -m "Auto-update Q2 dashboard: $(date '+%Y-%m-%d %H:%M')"
git push origin main
echo "=== Update completed successfully ==="
echo "Managers can view at: https://naichaniuiu.github.io/zhongxibu-dashboard/"
