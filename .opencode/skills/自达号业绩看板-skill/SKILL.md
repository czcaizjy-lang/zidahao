---
name: 自达号业绩看板-skill
description: 自达号专项业绩看板的搭建、数据刷新、修改、下钻交互、问题排查（基于 6月总自达号业绩.xlsx → JSON → Chart.js 看板）。按 7 个子机构（花开/集米/太古/九三/直属/半兆/久酒）展示自达号业绩。Trigger keywords 自达号看板、自达号业绩、zidahao、半兆自达号、久酒自达号、刷新自达号、自达号下钻。
---

# 自达号专项业绩看板

独立的、只聚焦自达号的业绩看板。数据源为 `6月总自达号业绩.xlsx`，展示 7 个子机构的业绩趋势、占比和达人明细。

## 项目结构

```
/Users/xiaocao/CC/自达号业绩看板/
├── 6月总自达号业绩.xlsx            # 上游数据源（两个 Sheet：6月数据 + Sheet2）
├── zidahao.html                   # 前端看板页面（本地开发用，fetch 加载 data/zidahao_data.json）
├── scripts/
│   ├── sync_zidahao.py            # 数据提取脚本（读 Excel → 生成 zidahao_data.json）
│   ├── build_zidahao_standalone.py # 构建独立页面（内嵌 JSON 数据，写入 public/zidahao.html）
│   └── auto_sync_zidahao.sh       # 自动同步脚本（launchd WatchPaths 触发，含消抖）
├── data/
│   ├── zidahao_data.json          # 生成的 JSON 数据文件
│   └── sync_zidahao.log           # auto_sync 同步日志
├── public/
│   └── zidahao.html               # 独立部署页面（内嵌数据，GH Pages 入口）
└── .opencode/skills/
    └── 自达号业绩看板-skill/
        └── SKILL.md
```

## 上游数据源

**Excel 文件**：`/Users/xiaocao/CC/自达号业绩看板/6月总自达号业绩.xlsx`

| Sheet | 内容 | 关键列 |
|-------|------|--------|
| **6月数据** | 日流水（~14K 行） | col 2 昵称、col 3 抖音号、col 4 日期、col 6 时长、col 26 GMV、col 27 支付、col 32 退款、col 34 佣金、col 44 投放绑定、col 45 投放被投 |
| **Sheet2** | 花名册（176 达人） | col 1 昵称、col 2 抖音号、col 3 **机构归属**、col 4-8 汇总数值、col 12-20 机构汇总 |

### 7 个子机构

| 子机构 | 达人数量 |
|--------|---------|
| 花开自达号 | 34 |
| 集米自达号 | 63 |
| 太古自达号 | 34 |
| 九三自达号 | 2 |
| 直属自达号 | 6 |
| 半兆自达号 | 35 |
| 久酒自达号 | 1 |

**对比原版**：新增了 **半兆自达号** 和 **久酒自达号** 两个子机构（原版 5 个 → 现版 7 个）。

## 关键口径

- **所有数值从日流水实时计算**：GMV、支付、退款、消耗、佣金、开播天数、日均时长等均从「6月数据」聚合
- **结算GMV** = 支付金额 - 退款GMV
- **每日消耗** = 优先取 col 45（投放消耗店铺被投），为 0 时回退到 col 44（投放消耗店铺绑定）
- **安全过滤**：GMV=0 的场次，消耗不计入日汇总
- **ROI** = GMV / 消耗（月度 OR 分日）
- **趋势窗口** = 以日流水最新日期为终点前推 30 天
- **分日子机构 ROI** = Σ(子机构当日所有达人 GMV) / Σ(子机构当日所有达人消耗)，加权计算
- **达人昵称优先级**：日流水 B 列 > 花名册 A 列

## 快速使用

### 更新数据（手动）

```bash
cd "/Users/xiaocao/CC/自达号业绩看板"
python3 scripts/sync_zidahao.py
```

### 本地开发预览

```bash
cd "/Users/xiaocao/CC/自达号业绩看板"
python3 -m http.server 8977
# 访问 http://localhost:8977/zidahao.html
```

### 自动同步 + 部署（launchd WatchPaths）

```
Excel 保存 → macOS 内核 WatchPaths 检测到变化
         → launchd 启动 auto_sync_zidahao.sh
              → 消抖 5~10 秒（双重防抖）
              → ① sync_zidahao.py（提取数据）
              → ② build_zidahao_standalone.py（生成独立页面）
              → ③ git push（推送到 GitHub Pages）
         → 脚本退出，恢复零占用
```

### launchd 任务管理

| 任务名 | 作用 | 触发方式 | 端口 |
|--------|------|---------|------|
| `com.dashboard.zidahao-sync` | Excel 保存 → 同步 + 构建 + 推送 | WatchPaths 文件变化 | - |
| `com.dashboard.zidahao-server` | 本地 HTTP 服务（自达号专用） | RunAtLoad 开机自启 | **8977** |

```bash
# 查看任务状态
launchctl list | grep zidahao

# 手动触发一次同步（测试用）
bash /Users/xiaocao/CC/自达号业绩看板/scripts/auto_sync_zidahao.sh

# 查看同步日志
tail -f /Users/xiaocao/CC/自达号业绩看板/data/sync_zidahao.log

# 卸载/重载
launchctl unload ~/Library/LaunchAgents/com.dashboard.zidahao-sync.plist
launchctl load ~/Library/LaunchAgents/com.dashboard.zidahao-sync.plist
```

### macOS 权限要求

launchd 启动的进程默认无权访问桌面文件。**必须将 `/bin/bash` 加入「系统设置 → 隐私与安全性 → 完全磁盘访问权限」**。

## 看板功能

KPI 卡片（GMV / 支付 / 结算 / 消耗 / ROI），ROI 卡片可点击查看整体 ROI 趋势。

### 图表

- **自达号总业绩趋势**（折线图）: 总GMV、支付金额、退款金额，近 30 天滚动窗口；点击数据点弹出自达号达人当日 vs 前日变化明细表（支持全部/仅下降/仅上涨筛选）
- **子机构趋势**（折线图）: **7 条子机构线**（含半兆、久酒），支持点击线/图例聚焦单个子机构，点击数据点弹出该子机构达人变化明细
- **子机构业绩占比**（饼图/环形图）: 7 个子机构扇区，点击弹出该子机构 Top5 达人分天趋势

### 交互

- **主题切换**: Header 右侧 🌙/☀️ 按钮，深色/浅色双主题，保存到 localStorage（key: `zdhDashboardTheme`）
- **子机构趋势图聚焦**: 点击单条线只显示该子机构，再点恢复全部
- **自达号总业绩下探**（`showZidahaoDrill()`）: 点击总趋势图数据点，弹出所有自达号达人当日 vs 前日变化明细，按掉量降序
- **自达号子机构下探**（`showZidahaoSubDrill()`）: 点击子机构趋势图数据点，弹出该子机构达人变化明细，按掉量降序
- **达人 GMV 下钻**: 点击达人列表中「直播GMV」单元格弹出该达人每日支付曲线
- **达人 ROI 下钻**: 点击达人列表中 ROI 列弹出该达人每日 ROI 趋势
- **饼图下钻**: 点击子机构扇区弹出 Top5 达人分天曲线
- **ROI KPI 下钻**: 点击 ROI 卡片弹出整体 ROI 趋势折线图
- **搜索/过滤**: 按达人名称/抖音号搜索，按子机构下拉过滤
- **排序**: 点击列表头排序

### 达人表列

`['主播昵称', '机构', '开播天数', '日均开播时长（小时）', '直播GMV', '直播支付GMV', '直播退款GMV', '直播结算GMV', '结算率', 'ROI', '佣金支出', '投放消耗金额']`

### 关键 JS 全局变量

| 变量 | 用途 |
|------|------|
| `ZDH_SUB_AGENCIES` | 7 个子机构数组 |
| `ZDH_PIE_COLORS` / `ZDH_LINE_COLORS` | 子机构配色（7 色） |
| `window.__totalChart` | 总业绩趋势图表实例 |
| `window.__agencyChart` / `window.__agencyFocusMode` | 子机构趋势图表 / 聚焦状态 |
| `window.__agencyPieChart` | 子机构饼图实例 |
| `window.__drillChart` | 达人下钻弹窗图表实例 |
| `window.__agencyDrillChart` | 机构 Top5 达人趋势弹窗图表 |
| `window.__zidahaoRoiDrillChart` | 自达号 ROI 趋势弹窗图表 |
| `showZidahaoDrill(dateIndex)` | 自达号总业绩下探函数 |
| `showZidahaoSubDrill(subName, dateIndex)` | 自达号子机构下探函数 |

## 部署上线

### 双模式架构

| 模式 | 入口 | 端口 | 数据加载 | 用途 |
|------|------|------|---------|------|
| **本地开发** | `http://localhost:8977/zidahao.html` | 8977 | fetch `data/zidahao_data.json` | 开发调试 |
| **线上看板** | GitHub Pages URL | - | 内嵌 JSON（自包含） | 生产使用 |

### 线上地址

GitHub Pages（待创建仓库后确定）：
- `https://<username>.github.io/<repo>/` → `public/index.html`
- `https://<username>.github.io/<repo>/zidahao.html` → `public/zidahao.html`

### 本地开发预览

```bash
# 服务器已由 launchd 自动启动（com.dashboard.zidahao-server）
# 手动启动（如需要）：
cd "/Users/xiaocao/CC/自达号业绩看板"
python3 -m http.server 8977

# 访问
open http://localhost:8977/zidahao.html
```

### 部署架构

```
本地 Mac                                  GitHub Pages
┌──────────────────────┐                  ┌──────────────────────────┐
│ Excel (.xlsx)        │                  │  public/                  │
│      ↓               │    git push      │  ├── index.html (入口)    │
│ sync_zidahao.py      │ ───────────────→ │  ├── zidahao.html        │
│      ↓               │                  │  └── (内嵌 JSON 数据)     │
│ build_standalone.py  │                  └──────────────────────────┘
│      ↓               │                        ↑
│ auto_sync.sh         │              GitHub Actions (.github/workflows/static.yml)
│ (launchd WatchPaths) │
│                      │
│ http.server :8977 ───┤  本地预览（fetch 模式）
└──────────────────────┘
```

### 手动部署

```bash
cd "/Users/xiaocao/CC/自达号业绩看板"

# 第 1 步：提取数据
python3 scripts/sync_zidahao.py
# （自动调用 build_zidahao_standalone.py 生成 public/index.html + public/zidahao.html）

# 第 2 步：推送到 GitHub Pages（触发 GitHub Actions 自动部署）
git add data/zidahao_data.json public/zidahao.html public/index.html
git commit -m "📊 自达号看板数据更新 $(date +'%Y-%m-%d %H:%M')"
git push origin main
```

## 常见问题

### 1. 数据加载失败

控制台报 `Failed to load zidahao data`。

**排查**:
- 确认 `data/zidahao_data.json` 文件存在
- 运行 `python3 scripts/sync_zidahao.py` 重新生成

### 2. Excel Sheet 找不到

脚本报 `KeyError: 'Worksheet xxx does not exist.'`

**解决**: 确认 Excel 文件包含 `6月数据` 和 `Sheet2` 两个 Sheet。

### 3. 子机构数据为 0

**排查**: 确认 Sheet2 col 3（机构归属）列的值正确，且对应抖音号在 6月数据 中有日流水。

### 4. 自动同步不工作

```bash
# 检查 launchd 任务状态
launchctl list | grep zidahao

# 查看同步日志
tail -20 /Users/xiaocao/CC/自达号业绩看板/data/sync_zidahao.log

# 检查 macOS 权限：系统设置 → 隐私与安全性 → 完全磁盘访问权限 → 确认 /bin/bash 已开启

# 手动测试
bash /Users/xiaocao/CC/自达号业绩看板/scripts/auto_sync_zidahao.sh

# 重载 launchd
launchctl unload ~/Library/LaunchAgents/com.dashboard.zidahao-sync.plist
launchctl load ~/Library/LaunchAgents/com.dashboard.zidahao-sync.plist
```

### 5. 跨月数据更新

上游 Excel 路径中文件名包含月份（`6月总自达号业绩.xlsx`），跨月后需更新：
- Excel 文件名
- `sync_zidahao.py` 中的 `XLSX_PATH`
- `auto_sync_zidahao.sh` 中的 `EXCEL_PATH`
- `~/Library/LaunchAgents/com.dashboard.zidahao-sync.plist` 中的 `WatchPaths`

## Git 托管

### 仓库

独立的 GitHub 仓库（与原 `czcaizjy-lang/-` 完全隔离）。初始化步骤：

```bash
# 1. 在 github.com 创建新仓库（如 zidahao-dashboard）
# 2. 添加 remote 并推送
cd "/Users/xiaocao/CC/自达号业绩看板"
git remote add origin git@github.com:<username>/<repo>.git
git push -u origin main
# 3. 在仓库 Settings → Pages → Source 选择 "GitHub Actions"
```

### GitHub Actions 自动部署

`.github/workflows/static.yml`：推送到 `main` → 自动部署 `public/` 到 GitHub Pages。

## 与原业绩看板的关系

本看板与 `CC/每日业绩自动统计/` 下的「蕉下直播业绩追击看板」是**完全独立**的项目：

| 维度 | 原版蕉下看板 | 新版自达号看板 |
|------|------------|---------------|
| 数据源 | `6月业绩追击（纯直播）.xlsx` | `6月总自达号业绩.xlsx` |
| 范围 | 全部 + 自达号切换 | **仅自达号** |
| 子机构 | 5 个 | **7 个（+半兆、久酒）** |
| 本地端口 | **8976** | **8977** |
| 本地 URL | `localhost:8976/dashboard.html` | `localhost:8977/zidahao.html` |
| Git 仓库 | `czcaizjy-lang/-` | 独立仓库 |
| launchd sync | `com.dashboard.sync` | `com.dashboard.zidahao-sync` |
| launchd server | `com.dashboard.local-server` | `com.dashboard.zidahao-server` |
| Skill | 业绩看板-skill | 自达号业绩看板-skill |
| GH Pages | `public/index.html` | `public/index.html` + `public/zidahao.html` |

数据和服务**完全物理隔离**：不同端口、不同 Git 仓库、不同 launchd 任务、不同数据源。
