# 法考打卡

2026 法考备考追踪 · 45 天精讲计划 · 每日三段式学习 · 打卡 + 进度追踪

> **TL;DR (English):** A local-first Flask web app that tracks daily study check-ins against a 45-day exam prep plan, split into morning / afternoon / evening / review slots. On first run it creates a local SQLite database and seeds the schedule automatically. All check-ins and backups stay on your machine — nothing is uploaded anywhere. Run `启动.bat` and open <http://127.0.0.1:5000>.

---

## 怎么用

### 电脑上有 Python 的话

1. 解压到任意文件夹
2. 双击 `启动.bat`，自动安装依赖并启动
3. 浏览器会自动打开 `http://127.0.0.1:5000`

### 没装 Python 的话

先去 [python.org](https://www.python.org/downloads/) 下载安装 Python 3.7 以上版本，安装时勾选"Add Python to PATH"，然后双击 `启动.bat`。

### 每天怎么用

1. 打开网站，仪表盘显示今天该学什么（上午/下午/晚上/复习四个时段）
2. 学完一个时段打个勾
3. 晚上全部学完点 Check In
4. 学快了或拖了一天，点 Prev/Next 调整计划

### 备份

点仪表盘右上角 Backup 按钮。备份文件在 `data/backups/` 里。

### 暗色模式

点右上角月亮图标切换。

---

## 学习时段

| 时段 | 时间 | 做什么 |
|---|---|---|
| 上午 | 09:00 — 12:00 | 看视频 + 记笔记（~170 分钟有效时间） |
| 下午 | 13:30 — 17:30 | 看视频 + 做真金题（~230 分钟） |
| 晚上 | 19:00 — 21:00 | 做真金题（~110 分钟） |
| 复习 | 21:00 — 22:00 | 回顾今日错题和笔记，不学新内容 |

视频建议 1.5—2 倍速观看，计划里列的是原始时长。

---

## 数据

数据本地自备，本仓库不含任何个人学习数据：

- 首次运行自动创建 `data/` 目录和 `data/fakao.db`（SQLite 文件），并按内置计划生成 45 天日程
- 打卡记录与备份（`data/backups/`）只保存在本地，不上传任何地方
- 想调整计划：修改 `scripts/gen_excel_v3.py` / `database/seed.py` 后删除 `data/fakao.db`，重启即自动重新生成
- `data/` 已加入 `.gitignore`，请勿把个人备考数据提交到公开仓库
