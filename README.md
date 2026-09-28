English · [简体中文](./README.zh-CN.md)

# Fakao Check-in (法考打卡)

A local-first Flask web app that tracks daily study check-ins against a 45-day prep plan for the 2026 National Legal Professional Examination (法考), split into morning / afternoon / evening / review slots. 45-day intensive plan · daily three-part study routine · check-ins + progress tracking. On first run it creates a local SQLite database and seeds the schedule automatically. All check-ins and backups stay on your machine — nothing is uploaded anywhere. Run `启动.bat` and open <http://127.0.0.1:5000>.

## How to use

### If you have Python on your computer

1. Extract to any folder
2. Double-click `启动.bat` — it installs the dependencies and starts automatically
3. The browser opens `http://127.0.0.1:5000` automatically

### If Python is not installed

Download and install Python 3.7 or later from [python.org](https://www.python.org/downloads/), tick "Add Python to PATH" during installation, then double-click `启动.bat`.

### Daily usage

1. Open the site; the dashboard shows what to study today (four slots: morning / afternoon / evening / review)
2. Tick a slot once it is done
3. In the evening, click Check In after finishing everything
4. Ahead of or behind schedule? Click Prev/Next to shift the plan

### Backup

Click the Backup button at the top right of the dashboard. Backup files live in `data/backups/`.

### Dark mode

Click the moon icon at the top right to toggle.

---

## Study sessions

| Slot | Time | What to do |
|---|---|---|
| Morning | 09:00 — 12:00 | Watch videos + take notes (~170 min effective time) |
| Afternoon | 13:30 — 17:30 | Watch videos + past exam questions (~230 min) |
| Evening | 19:00 — 21:00 | Past exam questions (~110 min) |
| Review | 21:00 — 22:00 | Review today's mistakes and notes; no new content |

Videos are best watched at 1.5–2x speed; the plan lists original durations.

---

## Data

Bring your own data; this repository contains no personal study data:

- The first run automatically creates the `data/` directory and `data/fakao.db` (SQLite file), and generates the 45-day schedule from the built-in plan
- Check-in records and backups (`data/backups/`) stay on this machine only — nothing is uploaded anywhere
- To adjust the plan: edit `scripts/gen_excel_v3.py` / `database/seed.py`, delete `data/fakao.db`, and restart — it regenerates automatically
- `data/` is in `.gitignore`; never commit personal exam-prep data to a public repository
