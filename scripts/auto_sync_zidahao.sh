#!/bin/bash
# 自达号看板自动同步流水线
# 由 launchd WatchPaths 触发（Excel 保存 → 自动执行）
# 步骤：① 提取数据 → ② 构建独立页面 → ③ git push

set -e
cd "/Users/xiaocao/CC/自达号业绩看板"

LOG_FILE="data/sync_zidahao.log"
EXCEL_PATH="/Users/xiaocao/Desktop/蕉下文件/业绩追击/by月业绩/6月业绩/6月业绩追击（纯直播）.xlsx"
LOCK_FILE="/tmp/zidahao_sync.lock"

# ---- 消抖：两次执行必须间隔 >= 10 秒 ----
NOW=$(date +%s)
if [ -f "$LOCK_FILE" ]; then
    LAST=$(cat "$LOCK_FILE" 2>/dev/null || echo 0)
    GAP=$((NOW - LAST))
    if [ $GAP -lt 10 ]; then
        # 太近，exit 0（不报错）
        exit 0
    fi
fi
echo "$NOW" > "$LOCK_FILE"

# ---- 双重消抖：等 5 秒后检查文件是否还在变化 ----
SIZE_BEFORE=$(stat -f%z "$EXCEL_PATH" 2>/dev/null || echo 0)
sleep 5
SIZE_AFTER=$(stat -f%z "$EXCEL_PATH" 2>/dev/null || echo 0)
if [ "$SIZE_BEFORE" != "$SIZE_AFTER" ]; then
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] Excel 仍在写入中，跳过（$SIZE_BEFORE → $SIZE_AFTER）" | tee -a "$LOG_FILE"
    exit 0
fi

echo "" | tee -a "$LOG_FILE"
echo "═══════════════════════════════════════" | tee -a "$LOG_FILE"
echo "[$(date '+%Y-%m-%d %H:%M:%S')] 自达号看板自动同步开始" | tee -a "$LOG_FILE"

# ---- ① 数据提取 ----
echo "[$(date '+%Y-%m-%d %H:%M:%S')] ① 数据提取..." | tee -a "$LOG_FILE"
python3 scripts/sync_zidahao.py 2>&1 | tee -a "$LOG_FILE"
if [ ${PIPESTATUS[0]} -ne 0 ]; then
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] ✗ 数据提取失败，停止流水线" | tee -a "$LOG_FILE"
    exit 1
fi

# ---- ② 构建独立页面 ----
echo "[$(date '+%Y-%m-%d %H:%M:%S')] ② 构建独立页面..." | tee -a "$LOG_FILE"
python3 scripts/build_zidahao_standalone.py 2>&1 | tee -a "$LOG_FILE"
if [ ${PIPESTATUS[0]} -ne 0 ]; then
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] ✗ 构建失败，停止流水线" | tee -a "$LOG_FILE"
    exit 1
fi

# ---- ③ git push ----
echo "[$(date '+%Y-%m-%d %H:%M:%S')] ③ 推送至 GitHub..." | tee -a "$LOG_FILE"
git add data/zidahao_data.json public/zidahao.html public/index.html 2>&1 | tee -a "$LOG_FILE"
if git diff --cached --quiet; then
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] ✓ 无变更，跳过推送" | tee -a "$LOG_FILE"
else
    git commit -m "📊 自达号看板数据更新 $(date '+%Y-%m-%d %H:%M')" 2>&1 | tee -a "$LOG_FILE"
    git push origin main 2>&1 | tee -a "$LOG_FILE"
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] ✓ 推送完成" | tee -a "$LOG_FILE"
fi

echo "[$(date '+%Y-%m-%d %H:%M:%S')] 自达号看板同步完成 ✓" | tee -a "$LOG_FILE"

# 清理 lock 文件
rm -f "$LOCK_FILE"
