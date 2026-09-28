"""Generate Phase 1 study plan Excel from PDF data."""
import json
from datetime import date, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Load data
with open('data/phase1_daily.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

# Subject colors (light fills)
SUBJ_COLORS = {
    '刑法':   'FFCDD2',  # red light
    '民法':   'BBDEFB',  # blue light
    '民诉法': 'C8E6C9',  # green light
    '刑诉法': 'FFE0B2',  # orange light
    '行政法': 'E1BEE7',  # purple light
    '商法':   'B2DFDB',  # teal light
    '理论法': 'FFF9C4',  # yellow light
}

wb = Workbook()

# ── Sheet 1: 日程总览 ──
ws1 = wb.active
ws1.title = "日程总览"

# Styles
header_font = Font(name='微软雅黑', bold=True, size=11, color='FFFFFF')
header_fill = PatternFill(start_color='C4553A', end_color='C4553A', fill_type='solid')
body_font = Font(name='微软雅黑', size=10)
bold_font = Font(name='微软雅黑', size=10, bold=True)
thin_border = Border(
    left=Side(style='thin', color='D0D0D0'),
    right=Side(style='thin', color='D0D0D0'),
    top=Side(style='thin', color='D0D0D0'),
    bottom=Side(style='thin', color='D0D0D0'),
)
center = Alignment(horizontal='center', vertical='center', wrap_text=True)
left_wrap = Alignment(horizontal='left', vertical='center', wrap_text=True)

def style_header(ws, row, cols):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin_border

def style_row(ws, row, cols, fill_color=None):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = body_font
        cell.border = thin_border
        if fill_color:
            cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')

# Title
ws1.merge_cells('A1:G1')
ws1.cell(row=1, column=1, value='2026法考备考 · 第一轮精讲计划').font = Font(name='微软雅黑', bold=True, size=14)
ws1.cell(row=1, column=1).alignment = Alignment(horizontal='center', vertical='center')
ws1.row_dimensions[1].height = 30

ws1.merge_cells('A2:G2')
ws1.cell(row=2, column=1, value=f'日期: 2026.06.06 — 2026.07.28 | 共53天 | 每日约6-8小时视频 | 科目顺序: 刑→民→民诉→刑诉→行政→商法→理论').font = Font(name='微软雅黑', size=9, color='888888')
ws1.cell(row=2, column=1).alignment = Alignment(horizontal='center', vertical='center')

# Headers
headers = ['日期', '星期', '科目', '天数', '今日专题', '视频(分钟)', '视频(小时)']
for c, h in enumerate(headers, 1):
    ws1.cell(row=4, column=c, value=h)
style_header(ws1, 4, len(headers))

# Data
weekdays_cn = ['周一','周二','周三','周四','周五','周六','周日']

current_subj = None
subj_day = 0
for i, entry in enumerate(data):
    row = 5 + i
    d = datetime.strptime(entry['date'], '%Y-%m-%d').date()
    wd = weekdays_cn[d.weekday()]

    if entry['subject'] != current_subj:
        current_subj = entry['subject']
        subj_day = 1
    else:
        subj_day += 1

    color = SUBJ_COLORS.get(entry['subject'], None)

    ws1.cell(row=row, column=1, value=d.strftime('%Y.%m.%d'))
    ws1.cell(row=row, column=2, value=wd)
    ws1.cell(row=row, column=3, value=entry['subject'])
    ws1.cell(row=row, column=4, value=f"D{subj_day}/{entry['subject_code']}D{entry['day_num']}")
    ws1.cell(row=row, column=5, value=entry['topics'])
    ws1.cell(row=row, column=6, value=entry['video_minutes'])
    ws1.cell(row=row, column=7, value=round(entry['video_minutes'] / 60, 1))

    style_row(ws1, row, len(headers), color)
    ws1.cell(row=row, column=1).alignment = center
    ws1.cell(row=row, column=2).alignment = center
    ws1.cell(row=row, column=3).alignment = center
    ws1.cell(row=row, column=3).font = bold_font
    ws1.cell(row=row, column=4).alignment = center
    ws1.cell(row=row, column=5).alignment = left_wrap
    ws1.cell(row=row, column=6).alignment = center
    ws1.cell(row=row, column=7).alignment = center
    ws1.row_dimensions[row].height = 22

# Column widths
ws1.column_dimensions['A'].width = 13
ws1.column_dimensions['B'].width = 7
ws1.column_dimensions['C'].width = 10
ws1.column_dimensions['D'].width = 12
ws1.column_dimensions['E'].width = 72
ws1.column_dimensions['F'].width = 12
ws1.column_dimensions['G'].width = 12

# Summary at bottom
summary_row = 5 + len(data) + 2
ws1.merge_cells(f'A{summary_row}:G{summary_row}')
ws1.cell(row=summary_row, column=1, value='各科汇总').font = Font(name='微软雅黑', bold=True, size=12)

subj_summary = {}
for entry in data:
    s = entry['subject']
    if s not in subj_summary:
        subj_summary[s] = {'days': 0, 'mins': 0}
    subj_summary[s]['days'] += 1
    subj_summary[s]['mins'] += entry['video_minutes']

for j, (s, info) in enumerate(subj_summary.items()):
    r = summary_row + 1 + j
    color = SUBJ_COLORS.get(s, None)
    ws1.merge_cells(f'A{r}:B{r}')
    ws1.cell(row=r, column=1, value=f'  {s}')
    ws1.cell(row=r, column=3, value=f'{info["days"]}天')
    ws1.cell(row=r, column=4, value=f'约{info["mins"]/60:.0f}小时视频')
    ws1.cell(row=r, column=5, value=f'日均{info["mins"]/info["days"]:.0f}分钟')
    for c in range(1, 6):
        ws1.cell(row=r, column=c).font = bold_font
        ws1.cell(row=r, column=c).border = thin_border
        if color:
            ws1.cell(row=r, column=c).fill = PatternFill(start_color=color, end_color=color, fill_type='solid')
    ws1.row_dimensions[r].height = 22

total_row = summary_row + 1 + len(subj_summary)
ws1.merge_cells(f'A{total_row}:B{total_row}')
total_mins = sum(e['video_minutes'] for e in data)
ws1.cell(row=total_row, column=1, value=f'  合计: {len(data)}天, 约{total_mins/60:.0f}小时视频')
ws1.cell(row=total_row, column=1).font = Font(name='微软雅黑', bold=True, size=11)
for c in range(1, 6):
    ws1.cell(row=total_row, column=c).border = thin_border

# ── Sheet 2: 分科详情 ──
ws2 = wb.create_sheet("分科详情")

ws2.merge_cells('A1:E1')
ws2.cell(row=1, column=1, value='各科每日学习内容详表').font = Font(name='微软雅黑', bold=True, size=13)

row2 = 3
for code in ['criminal_law','civil_law','civil_proc','crim_proc','admin_law','commercial','theory']:
    subject_entries = [e for e in data if e['subject_code'] == code]
    if not subject_entries:
        continue

    sname = subject_entries[0]['subject']

    ws2.merge_cells(f'A{row2}:E{row2}')
    ws2.cell(row=row2, column=1, value=f'{sname} ({len(subject_entries)}天)').font = Font(name='微软雅黑', bold=True, size=11, color='C4553A')
    ws2.row_dimensions[row2].height = 24
    row2 += 1

    # Sub-headers
    sub_headers = ['日期', '天数', '专题内容', '视频(min)', '建议时间段']
    for c, h in enumerate(sub_headers, 1):
        ws2.cell(row=row2, column=c, value=h)
    style_header(ws2, row2, len(sub_headers))
    row2 += 1

    color = SUBJ_COLORS.get(sname)
    for entry in subject_entries:
        d = datetime.strptime(entry['date'], '%Y-%m-%d').date()
        mins = entry['video_minutes']
        # Suggest time slots based on video length
        if mins <= 180:
            slots = '上午(9:00-12:00)即可完成'
        elif mins <= 350:
            slots = '上午(9:00-12:00) + 下午(13:30-16:00)'
        else:
            slots = '上午(9:00-12:00) + 下午(13:30-17:30)'

        ws2.cell(row=row2, column=1, value=d.strftime('%m/%d'))
        ws2.cell(row=row2, column=2, value=f'第{entry["day_num"]}天')
        ws2.cell(row=row2, column=3, value=entry['topics'])
        ws2.cell(row=row2, column=4, value=mins)
        ws2.cell(row=row2, column=5, value=slots)
        style_row(ws2, row2, len(sub_headers), color)
        ws2.cell(row=row2, column=1).alignment = center
        ws2.cell(row=row2, column=2).alignment = center
        ws2.cell(row=row2, column=3).alignment = left_wrap
        ws2.cell(row=row2, column=4).alignment = center
        ws2.cell(row=row2, column=5).alignment = center
        ws2.row_dimensions[row2].height = 22
        row2 += 1

    row2 += 1  # blank row between subjects

ws2.column_dimensions['A'].width = 10
ws2.column_dimensions['B'].width = 8
ws2.column_dimensions['C'].width = 75
ws2.column_dimensions['D'].width = 12
ws2.column_dimensions['E'].width = 38

# ── Sheet 3: 空白的每日打卡表 ──
ws3 = wb.create_sheet("打卡记录")

ws3.merge_cells('A1:H1')
ws3.cell(row=1, column=1, value='每日学习打卡记录').font = Font(name='微软雅黑', bold=True, size=13)

check_headers = ['日期', '科目', '计划视频(min)', '实际完成', '上午(9-12)', '下午(1:30-5:30)', '晚上(7-10)', '备注/心情']
for c, h in enumerate(check_headers, 1):
    ws3.cell(row=3, column=c, value=h)
style_header(ws3, 3, len(check_headers))

for i, entry in enumerate(data):
    row = 4 + i
    d = datetime.strptime(entry['date'], '%Y-%m-%d').date()
    color = SUBJ_COLORS.get(entry['subject'])
    ws3.cell(row=row, column=1, value=d.strftime('%m/%d'))
    ws3.cell(row=row, column=2, value=entry['subject'])
    ws3.cell(row=row, column=3, value=entry['video_minutes'])
    ws3.cell(row=row, column=4, value='☐')  # checkbox placeholder
    ws3.cell(row=row, column=5, value='')
    ws3.cell(row=row, column=6, value='')
    ws3.cell(row=row, column=7, value='')
    ws3.cell(row=row, column=8, value='')
    style_row(ws3, row, len(check_headers), color)
    for c in [1,2,3,4]:
        ws3.cell(row=row, column=c).alignment = center
    ws3.row_dimensions[row].height = 20

ws3.column_dimensions['A'].width = 8
ws3.column_dimensions['B'].width = 8
ws3.column_dimensions['C'].width = 12
ws3.column_dimensions['D'].width = 10
ws3.column_dimensions['E'].width = 18
ws3.column_dimensions['F'].width = 18
ws3.column_dimensions['G'].width = 18
ws3.column_dimensions['H'].width = 20

# Save
out = '法考一轮复习计划.xlsx'
wb.save(out)
print(f'Saved to: {out}')
print(f'Sheets: 日程总览({len(data)}行) + 分科详情(7科) + 打卡记录({len(data)}行)')
