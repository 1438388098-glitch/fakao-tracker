"""Generate Phase 1 study plan Excel with proper 3-slot daily schedule.

Daily schedule (effective study time):
  09:00-12:00  → 170min  看视频+记笔记
  13:30-17:30  → 230min  看视频+做真题
  19:00-21:00  → 110min  做真题（视频在上午+下午已完成）
  21:00-22:00  →  60min  复习今日错题和笔记（不学新东西）
  Total: 510min study + 60min review
  Note: 视频建议1.5-2倍速，列出的分钟是视频原始时长
"""
import re
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ─── Subject data: [(day_label, total_video_min, [topic1, topic2, ...]), ...]
# Each topic string includes "(NN min)" or "(NNmin)" for its video duration.
# The generator splits topics across the 3 study slots proportionally.

CRIMINAL = [  # 刑法 8天 ~2650min
    ('刑法论 + 犯罪构成·客观要件(上)', 340, [
        '刑法概述·基本原则(90min)','刑法适用范围(40min)','行为理论(80min)',
        '因果关系·客观归责(90min)','通读精讲卷第1-2章(40min)'
    ]),
    ('客观要件(下) + 主观要件', 350, [
        '犯罪主体·特殊身份(80min)','故意理论(90min)','过失(60min)',
        '目的犯·动机(40min)','认识错误(30min)','违法阻却事由概述(50min)'
    ]),
    ('违法性阻却 + 犯罪形态', 330, [
        '正当防卫(100min)','紧急避险(70min)','其他违法阻却(30min)',
        '既遂·未遂(80min)','中止·预备(50min)'
    ]),
    ('共同犯罪', 300, [
        '共犯基础理论(90min)','间接正犯(80min)','共犯脱离(40min)',
        '共犯处罚原则(90min)'
    ]),
    ('罪数形态 + 刑罚论', 320, [
        '罪数形态(90min)','刑罚体系·量刑(80min)','累犯·自首·立功(80min)',
        '数罪并罚·缓刑·减刑·假释·追诉时效(70min)'
    ]),
    ('侵犯人身权利罪', 340, [
        '故意杀人·伤害(80min)','强奸·强制猥亵·侮辱(90min)',
        '非法拘禁·绑架(60min)','拐卖妇女儿童(40min)',
        '遗弃·侮辱诽谤(30min)','人身犯罪罪名对比(40min)'
    ]),
    ('侵犯财产罪', 350, [
        '抢劫罪·抢夺罪(100min)','盗窃罪(90min)','诈骗罪(50min)',
        '敲诈勒索(40min)','侵占·职务侵占·挪用(70min)'
    ]),
    ('公共安全+经济+社会管理+贪污渎职', 320, [
        '危害公共安全罪重点(90min)','经济秩序罪重点(80min)',
        '妨害社会管理秩序罪(60min)','贪污贿赂罪(60min)','渎职罪(30min)'
    ]),
]

CIVIL = [  # 民法 8天 ~2600min
    ('民事法律关系 + 自然人 + 法人', 240, [
        '民事法律关系(60min)','自然人·能力+监护(70min)',
        '法人概述+分类(60min)','非法人组织(30min)','民事法律行为概述(20min)'
    ]),
    ('民事法律行为(上)', 300, [
        '意思表示(80min)','成立+生效要件(70min)','附条件+附期限(30min)',
        '民事法律行为体系整理(30min)','无效民事法律行为概述(90min)'
    ]),
    ('民事法律行为效力', 350, [
        '无效民事法律行为·深度(80min)','可撤销民事法律行为(90min)',
        '效力待定(40min)','效力形态总结对比(40min)','代理概述(60min)',
        '无权代理+表见代理(40min)'
    ]),
    ('诉讼时效 + 物权总则 + 物权变动(上)', 310, [
        '诉讼时效(60min)','物权基本原则(40min)','物权变动概述(50min)',
        '登记制度(50min)','交付(30min)','善意取得(50min)','取得时效(30min)'
    ]),
    ('物权变动(下) + 所有权 + 占有', 310, [
        '所有权(60min)','业主建筑物区分所有权(40min)','相邻关系+共有(30min)',
        '占有(50min)','抵押权概述+设立(90min)','特殊抵押(40min)'
    ]),
    ('担保物权', 390, [
        '抵押权实现(50min)','质权(90min)','留置权(40min)',
        '担保物权竞合(30min)','担保制度司法解释要点(60min)',
        '非典型担保(40min)','担保体系对比总结(80min)'
    ]),
    ('债法总则 + 合同订立 + 履行', 340, [
        '债的法定分类(50min)','合同基本原则(40min)','缔约过失(40min)',
        '要约+承诺(60min)','格式条款(30min)','抗辩权·同时+不安+顺序(60min)',
        '代位权+撤销权(60min)'
    ]),
    ('合同变更转让+终止+违约+买卖合同', 360, [
        '合同变更与转让(70min)','合同解除(50min)','清偿+抵销+提存(30min)',
        '违约责任(80min)','买卖合同(60min)','供用/赠与/借款/租赁/承揽(70min)'
    ]),
]

CIVIL_PROC = [  # 民诉 5天 ~1710min
    ('诉的基本理论 + 管辖', 340, [
        '民事诉讼概论(30min)','诉的基本理论(100min)','基本原则(60min)',
        '级别管辖(40min)','地域管辖(80min)','裁定管辖+管辖权异议(30min)'
    ]),
    ('当事人 + 证据', 350, [
        '当事人(60min)','共同诉讼(40min)','第三人(50min)','诉讼代理人(20min)',
        '证据种类(60min)','证据分类(50min)','证明对象+责任+标准(70min)'
    ]),
    ('证明程序 + 保全 + 一审', 340, [
        '证明程序(50min)','保全+先予执行(40min)','调解(60min)',
        '一审起诉+受理(60min)','审理+撤诉+缺席(60min)',
        '强制措施+期间送达(50min)','特别程序概述(20min)'
    ]),
    ('简易+小额+二审+再审', 340, [
        '简易程序(60min)','小额诉讼(20min)','二审程序(100min)',
        '审判监督程序(80min)','督促+公示催告(40min)','上诉vs再审对比(40min)'
    ]),
    ('执行 + 仲裁 + 民诉总复习', 340, [
        '执行程序·开始+管辖+措施(100min)','执行救济(40min)',
        '仲裁概述+协议(60min)','仲裁程序+司法监督(60min)',
        '民诉法体系回顾(80min)'
    ]),
]

CRIM_PROC = [  # 刑诉 7天 ~2340min
    ('刑诉法概述 + 专门机关 + 管辖', 340, [
        '刑诉法概念+任务+基本原则(100min)','专门机关+诉讼参与人(90min)',
        '立案管辖(30min)','审判管辖(80min)','刑诉法体系概览(40min)'
    ]),
    ('回避 + 辩护 + 代理', 330, [
        '回避(30min)','辩护制度概述(40min)','辩护人权利(100min)',
        '法律援助(60min)','拒绝辩护(30min)','刑事代理(20min)',
        '辩护制度总结(50min)'
    ]),
    ('刑事证据', 350, [
        '证据种类(80min)','证据分类(60min)','非法证据排除规则(100min)',
        '证明对象+责任+标准(80min)','证据规则综合(30min)'
    ]),
    ('强制措施', 330, [
        '拘传(20min)','取保候审(80min)','监视居住(40min)',
        '拘留(50min)','逮捕(100min)','强制措施体系对比(40min)'
    ]),
    ('附带民诉 + 立案 + 侦查', 320, [
        '附带民事诉讼(60min)','期间+送达(20min)','立案(30min)',
        '侦查行为(80min)','侦查终结+补充侦查(70min)','侦查程序总结(60min)'
    ]),
    ('审查起诉 + 一审 + 二审', 340, [
        '审查起诉(80min)','公诉一审(100min)','自诉+简易+速裁(60min)',
        '二审程序(100min)'
    ]),
    ('死刑复核 + 再审 + 执行 + 特别程序', 330, [
        '死刑复核(60min)','审判监督(60min)','执行程序(80min)',
        '未成年特别程序(40min)','和解+缺席+没收+强制医疗(40min)',
        '刑诉法体系总复习(50min)'
    ]),
]

ADMIN = [  # 行政法 5天 ~1890min
    ('行政法概述 + 主体 + 公务员', 370, [
        '行政法概述+基本原则(90min)','行政主体·行政机关(90min)',
        '被授权组织(30min)','公务员法(120min)','行政法体系概览(40min)'
    ]),
    ('抽象+具体行政行为', 370, [
        '行政法规+规章(100min)','具体行政行为概念+分类(70min)',
        '具体行为效力(60min)','成立+生效(40min)','行政许可概述(50min)',
        '许可设定+实施(50min)'
    ]),
    ('行政许可 + 行政处罚', 380, [
        '行政许可监督检查(40min)','行政处罚种类+设定(70min)',
        '处罚实施程序(60min)','处罚执行(50min)','行政强制措施(50min)',
        '行政强制执行(80min)','许可vs处罚对比(30min)'
    ]),
    ('行政强制 + 信息公开 + 行政复议', 380, [
        '政府信息公开(60min)','行政复议范围+复议机关(80min)',
        '复议程序+决定(90min)','行政协议(30min)','其他行政行为(20min)',
        '强制措施vs执行对比(40min)','行政诉讼概述(60min)'
    ]),
    ('行政诉讼(全) + 行政复议收尾', 390, [
        '受案范围+管辖(80min)','诉讼参加人(70min)','证据(60min)',
        '起诉受理(60min)','审理+裁判(60min)','执行(40min)',
        '行政复议与诉讼衔接(20min)'
    ]),
]

COMMERCIAL = [  # 商法 6天 ~1820min
    ('公司法·基础+治理', 310, [
        '公司分类+体系(50min)','法人人格(40min)','股东权利概述(80min)',
        '公司治理结构(100min)','股东知情权(40min)'
    ]),
    ('公司法·股东权利+治理续', 310, [
        '代表诉讼(40min)','董监高义务(60min)','公司融资(40min)',
        '合并分立(50min)','解散+清算(60min)','公司决议效力(60min)'
    ]),
    ('公司法·股权转让+股份公司+国有公司', 300, [
        '有限公司股权转让(70min)','股份公司设立(60min)','上市公司(50min)',
        '国有出资公司(40min)','公司组织架构总结(40min)','公司总复习(40min)'
    ]),
    ('合伙企业法+个独+外资', 290, [
        '普通合伙·设立+财产+执行(90min)','有限合伙(60min)',
        '个独企业法(30min)','外商投资法(30min)','合伙vs公司对比(40min)',
        '商主体体系总结(40min)'
    ]),
    ('破产法', 310, [
        '破产原因+申请+受理(80min)','管理人(40min)','债权申报(60min)',
        '债务人财产(50min)','重整+和解(70min)','破产程序总结(10min)'
    ]),
    ('保险法+信托+证券+商法总复习', 300, [
        '保险合同(60min)','人身+财产保险(60min)','信托法基础(30min)',
        '证券法要点(40min)','票据法要点(30min)','商法体系总复习(80min)'
    ]),
]

THEORY = [  # 理论法 6天 ~2170min
    ('法理学·概念+价值+要素', 370, [
        '法的特征(50min)','法的本质(40min)','法的作用(50min)',
        '法的价值(50min)','法律规则(80min)','法律原则(40min)',
        '法律概念(20min)','法理学导论总结(40min)'
    ]),
    ('法理学·渊源+效力+关系+责任+实施', 350, [
        '法的渊源+分类(100min)','法的效力(50min)','法律关系(60min)',
        '法律责任(40min)','法的制定+实施(100min)'
    ]),
    ('法理学·适用+解释+宪法基础', 370, [
        '法适用一般原理(60min)','法律解释(80min)','法律漏洞填补(40min)',
        '宪法概念+原则(70min)','宪法制定+修改+解释(60min)',
        '宪法实施监督(60min)'
    ]),
    ('宪法·国家制度+选举+特区+自治', 370, [
        '国体+政体+经济制度(80min)','选举制度(90min)',
        '特别行政区制度(70min)','民族区域+基层群众自治(50min)',
        '国家制度总结(80min)'
    ]),
    ('宪法·权利+机构+法治思想', 370, [
        '公民基本权利(100min)','人大+国务院(80min)',
        '监察委+法院+检察院(60min)','法治思想核心要义(80min)',
        '宪法全体系复习(50min)'
    ]),
    ('法律史+司法制度+职业道德+总复习', 340, [
        '中国法律史·周至清(60min)','民国+新中国(40min)',
        '司法制度(50min)','法律职业道德(40min)','法律文书要点(20min)',
        '法理+宪法快速回顾(60min)','理论法全体系总结(70min)'
    ]),
]

# ─── Build schedule ───
SUBJ = [
    ('刑法', CRIMINAL), ('民法', CIVIL), ('民诉', CIVIL_PROC),
    ('刑诉', CRIM_PROC), ('行政法', ADMIN), ('商法', COMMERCIAL),
    ('理论法', THEORY),
]

SUBJ_COLORS = {
    '刑法':'FFCDD2','民法':'BBDEFB','民诉':'C8E6C9',
    '刑诉':'FFE0B2','行政法':'E1BEE7','商法':'B2DFDB','理论法':'FFF9C4',
}

start = date(2026, 6, 6)
rows = []
current = start

for sname, daily in SUBJ:
    for i, (label, total_min, topics) in enumerate(daily):
        rows.append({'date':current,'subject':sname,'day':i+1,
                     'label':label,'total_min':total_min,'topics':topics})
        current += timedelta(days=1)

end = current - timedelta(days=1)

# ─── Slot allocation: split topics across 3 study slots ───
# Morning ~35%, Afternoon ~45%, Evening ~20% of video content
# Remaining time in each slot filled with practice questions (真金题)

def split_topics(topics, total_min):
    """Split topic list into [morning], [afternoon], [evening].
    Fill morning first (~170min capacity), then afternoon (~230min), rest to evening (~110min).
    Each slot gets video content + practice questions to fill remaining time.
    """
    MORNING_CAP = 170
    AFTERNOON_CAP = 230
    EVENING_CAP = 110

    def mins(t):
        m = re.findall(r'(\d+)\s*min', t)
        return sum(int(x) for x in m) if m else 40

    morning, afternoon, evening = [], [], []
    idx = 0
    accum_m = 0

    # Fill morning: take topics up to ~170min
    while idx < len(topics) and accum_m < MORNING_CAP:
        t = topics[idx]
        m = mins(t)
        if accum_m + m <= MORNING_CAP + 10:
            morning.append(t)
            accum_m += m
            idx += 1
        else:
            break

    # Fill afternoon: take remaining topics up to ~230min
    accum_a = 0
    while idx < len(topics) and accum_a < AFTERNOON_CAP:
        t = topics[idx]
        m = mins(t)
        if accum_a + m <= AFTERNOON_CAP + 10:
            afternoon.append(t)
            accum_a += m
            idx += 1
        else:
            break

    # Evening: remaining topics (should fit within ~110min)
    accum_e = 0
    while idx < len(topics):
        t = topics[idx]
        m = mins(t)
        evening.append(t)
        accum_e += m
        idx += 1

    # Add practice questions to fill remaining time
    m_video = accum_m
    if m_video < MORNING_CAP - 20:
        morning.append(f'真金题·对应章节(~{MORNING_CAP - m_video}min)')
    if accum_a < AFTERNOON_CAP - 20:
        afternoon.append(f'真金题·对应章节(~{min(120, AFTERNOON_CAP - accum_a)}min)')
    if accum_e < EVENING_CAP - 20 and evening:
        evening.append(f'真金题·对应章节(~{EVENING_CAP - accum_e}min)')
    # If evening is empty, add practice
    if not evening:
        evening.append(f'真金题·当日剩余章节练习(~{EVENING_CAP}min)')

    return morning, afternoon, evening

# ─── Excel ───
wb = Workbook()

hfont = Font(name='微软雅黑', bold=True, size=10, color='FFFFFF')
hfill = PatternFill(start_color='5D4037', end_color='5D4037', fill_type='solid')
bfont = Font(name='微软雅黑', size=9)
bbold = Font(name='微软雅黑', size=9, bold=True)
tfont = Font(name='微软雅黑', bold=True, size=14, color='5D4037')
stfont = Font(name='微软雅黑', size=9, color='888888')
border = Border(left=Side('thin','D0D0D0'),right=Side('thin','D0D0D0'),
                top=Side('thin','D0D0D0'),bottom=Side('thin','D0D0D0'))
ca = Alignment(horizontal='center', vertical='center', wrap_text=True)
la = Alignment(horizontal='left', vertical='center', wrap_text=True)

WD = ['周一','周二','周三','周四','周五','周六','周日']

# ═══ Sheet 1: 精讲阶段日程表 ═══
ws = wb.active
ws.title = '精讲阶段日程表'
ws.merge_cells('A1:G1')
ws.cell(row=1,column=1,value='2026法考 · 第一轮精讲学习计划').font=tfont
ws.row_dimensions[1].height=30

ws.merge_cells('A2:G2')
ws.cell(row=2,column=1,value=f'每天: 9-12看视频(170min) | 1:30-5:30视频+真题(230min) | 7-9做真题(110min) | 9-10复习错题(60min) | 视频建议1.5-2倍速 | {len(rows)}天 | {start.strftime("%Y.%m.%d")}—{end.strftime("%Y.%m.%d")}').font=stfont

hdrs = ['日期','星期','科目','Day','今日学习专题','各时段具体安排','视频\n(min)']
for c,h in enumerate(hdrs,1):
    ws.cell(row=4,column=c,value=h)
for c in range(1,len(hdrs)+1):
    ws.cell(row=4,column=c).font=hfont
    ws.cell(row=4,column=c).fill=hfill
    ws.cell(row=4,column=c).alignment=ca
    ws.cell(row=4,column=c).border=border
ws.row_dimensions[4].height=30

for i,entry in enumerate(rows):
    row=5+i
    d=entry['date']
    wd=WD[d.weekday()]
    color=SUBJ_COLORS.get(entry['subject'])

    morning, afternoon, evening = split_topics(entry['topics'], entry['total_min'])

    # Build slot text
    lines = []
    lines.append('【上午 9:00-12:00 (~170min)】')
    for t in morning:
        lines.append(f'  • {t}')
    lines.append('【下午 13:30-17:30 (~230min)】')
    for t in afternoon:
        lines.append(f'  • {t}')
    lines.append('【晚上 19:00-21:00 做真题 (~110min)】')
    for t in evening:
        lines.append(f'  • {t}')
    lines.append('【晚上 21:00-22:00 复习 (~60min)】')
    lines.append('  不学新内容。回顾今日错题，快速过一遍笔记。背诵类内容（刑法罪名、民法构成要件等）此时记忆效果最好')
    slot_text = '\n'.join(lines)

    ws.cell(row=row,column=1,value=d.strftime('%Y.%m.%d'))
    ws.cell(row=row,column=2,value=wd)
    ws.cell(row=row,column=3,value=entry['subject'])
    ws.cell(row=row,column=4,value=f'第{entry["day"]}天')
    ws.cell(row=row,column=5,value=entry['label'])
    ws.cell(row=row,column=6,value=slot_text)
    ws.cell(row=row,column=7,value=entry['total_min'])

    for c in range(1,len(hdrs)+1):
        cell=ws.cell(row=row,column=c)
        cell.font=bfont
        cell.border=border
        if c<=4 or c==7:
            cell.alignment=ca
        else:
            cell.alignment=la
        if color:
            cell.fill=PatternFill(start_color=color,end_color=color,fill_type='solid')
    ws.cell(row=row,column=3).font=bbold
    ws.row_dimensions[row].height=max(20,13*(len(lines)+1))

ws.column_dimensions['A'].width=11
ws.column_dimensions['B'].width=6
ws.column_dimensions['C'].width=7
ws.column_dimensions['D'].width=6
ws.column_dimensions['E'].width=30
ws.column_dimensions['F'].width=70
ws.column_dimensions['G'].width=7

# Summary
sr=5+len(rows)+2
ws.merge_cells(f'A{sr}:G{sr}')
ws.cell(row=sr,column=1,value='各科汇总').font=Font(name='微软雅黑',bold=True,size=12)
for sname,_ in SUBJ:
    ents=[r for r in rows if r['subject']==sname]
    sr+=1
    color=SUBJ_COLORS[sname]
    ws.merge_cells(f'A{sr}:B{sr}')
    ws.cell(row=sr,column=1,value=f'  {sname}')
    ws.cell(row=sr,column=3,value=f'{len(ents)}天')
    ws.cell(row=sr,column=4,value=f'约{sum(e["total_min"] for e in ents)/60:.0f}h视频')
    ws.cell(row=sr,column=5,value=f'日均{sum(e["total_min"] for e in ents)/len(ents):.0f}min')
    for c in range(1,8):
        ws.cell(row=sr,column=c).font=bbold
        ws.cell(row=sr,column=c).border=border
        if color:
            ws.cell(row=sr,column=c).fill=PatternFill(start_color=color,end_color=color,fill_type='solid')

# ═══ Sheet 2: 每日打卡 ═══
ws2=wb.create_sheet('每日打卡(可打印)')
ws2.merge_cells('A1:J1')
ws2.cell(row=1,column=1,value='2026法考 · 每日自律打卡').font=tfont
ch2=['日期','科目','Day','计划\n(min)','☐上午\n9-12完成','☐下午\n1:30-5:30完成','☐晚上\n7-9完成','☐复习\n9-10完成','心情\n1-5','备注']
for c,h in enumerate(ch2,1):
    ws2.cell(row=3,column=c,value=h)
for c in range(1,len(ch2)+1):
    ws2.cell(row=3,column=c).font=hfont
    ws2.cell(row=3,column=c).fill=hfill
    ws2.cell(row=3,column=c).alignment=ca
    ws2.cell(row=3,column=c).border=border
ws2.row_dimensions[3].height=32

for i,entry in enumerate(rows):
    row=4+i
    color=SUBJ_COLORS.get(entry['subject'])
    ws2.cell(row=row,column=1,value=entry['date'].strftime('%m/%d'))
    ws2.cell(row=row,column=2,value=entry['subject'])
    ws2.cell(row=row,column=3,value=entry['day'])
    ws2.cell(row=row,column=4,value=entry['total_min'])
    for c in range(5,9):
        ws2.cell(row=row,column=c,value='☐')
    ws2.cell(row=row,column=9,value='')
    ws2.cell(row=row,column=10,value='')
    for c in range(1,11):
        ws2.cell(row=row,column=c).font=bfont
        ws2.cell(row=row,column=c).border=border
        ws2.cell(row=row,column=c).alignment=ca
        if color:
            ws2.cell(row=row,column=c).fill=PatternFill(start_color=color,end_color=color,fill_type='solid')
    ws2.row_dimensions[row].height=20

ws2.column_dimensions['A'].width=8
ws2.column_dimensions['B'].width=7
ws2.column_dimensions['C'].width=5
ws2.column_dimensions['D'].width=7
for c in ['E','F','G','H']:
    ws2.column_dimensions[c].width=13
ws2.column_dimensions['I'].width=6
ws2.column_dimensions['J'].width=12

# Print summary
out='法考一轮复习计划.xlsx'
wb.save(out)

print(f'Phase 1: {len(rows)} days ({start} → {end})')
for sname,_ in SUBJ:
    ents=[r for r in rows if r['subject']==sname]
    tmins=sum(e['total_min'] for e in ents)
    print(f'  {sname}: {len(ents)}天, {tmins}min={tmins/60:.0f}h视频, 日均{tmins/len(ents):.0f}min')
print(f'  总计: {sum(r["total_min"] for r in rows)/60:.0f}h')
print(f'\nSaved: {out}')
print('Sheets: 日程表(含三时段安排+复习) + 每日打卡(四大勾选项)')
