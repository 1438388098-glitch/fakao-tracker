"""Build Phase 1 daily schedule from PDF-extracted curriculum data."""
import json
import os
from datetime import date, timedelta

# ── Subject daily plans extracted from PDF 打卡表 ──
# Each: (day_label, topics_summary, video_minutes)

subjects = {}

# 民法 (孟献贵) - 14 days from PDF, user allocates 13
subjects['civil_law'] = {
    'name': '民法', 'code': 'civil_law', 'days': 13,
    'daily': [
        ('D1', '专题01:民事法律关系 + 专题02:自然人(01-02)', 123),
        ('D2', '专题02:自然人(03-06) + 法人(01-02)', 138),
        ('D3', '专题02:法人(03) + 非法人组织 + 专题03:民事法律行为(01)', 162),
        ('D4', '专题03:民事法律行为效力(01-04)', 144),
        ('D5', '专题03:民事法律行为效力(05) + 代理(01-02)', 105),
        ('D6', '专题04:代理(02) + 专题05:诉讼时效(01-02)', 152),
        ('D7', '专题06:物权基本原则 + 专题07:物权变动(01-02)', 128),
        ('D8', '专题07:物权变动(03) + 专题08:所有权 + 专题09:担保物权(01)', 167),
        ('D9', '专题09:担保物权-抵押(02-03) + 质权(04)', 184),
        ('D10', '专题09:留置权 + 专题10:占有 + 专题11:债的法定分类(01-02)', 133),
        ('D11', '专题12:合同基本原则 + 专题13:缔约过失 + 专题14:合同订立(01-02)', 162),
        ('D12', '专题15:合同履行(01-02) + 专题16:合同保全(01-02)', 188),
        ('D13', '专题17:合同变更转让(01-02) + 专题18:合同终止(01) + 专题19:违约责任(01-02)', 200),
    ]
}

# 民诉 (戴鹏) - P1-P44 segments regrouped into 6 days
subjects['civil_proc'] = {
    'name': '民诉法', 'code': 'civil_proc', 'days': 6,
    'daily': [
        ('D1', '诉的基本理论 + 基本原则 + 管辖(级别/地域/裁定/管辖权异议)', 280),
        ('D2', '当事人 + 共同诉讼 + 第三人 + 诉讼代理人', 290),
        ('D3', '证据(种类/分类) + 证明(对象/责任/标准/程序) + 保全+先予执行', 280),
        ('D4', '强制措施 + 期间送达 + 调解 + 一审普通程序(起诉/审理/撤诉/缺席)', 270),
        ('D5', '简易程序 + 小额诉讼 + 二审程序 + 审判监督程序', 290),
        ('D6', '特别程序 + 督促/公示催告 + 执行程序 + 仲裁(协议/程序/司法监督)', 300),
    ]
}

# 刑诉 (左宁) - ~48h, user allocates 7 days
subjects['crim_proc'] = {
    'name': '刑诉法', 'code': 'crim_proc', 'days': 7,
    'daily': [
        ('D1', '刑诉法概述 + 专门机关与诉讼参与人 + 管辖(立案/审判)', 340),
        ('D2', '回避 + 辩护制度(辩护人权利/法律援助/拒绝辩护) + 刑事代理', 330),
        ('D3', '刑事证据(证据种类+分类+规则+证明对象/责任/标准)', 350),
        ('D4', '强制措施(拘传/取保候审/监视居住/拘留/逮捕)', 330),
        ('D5', '附带民事诉讼 + 期间送达 + 立案 + 侦查(侦查行为/侦查终结/补充侦查)', 320),
        ('D6', '审查起诉 + 一审程序(公诉/自诉/简易/速裁) + 二审程序', 340),
        ('D7', '死刑复核 + 审判监督 + 执行 + 特别程序(未成年/和解/缺席/没收/强制医疗)', 310),
    ]
}

# 行政法 (李佳) - ~45h, user allocates 6 days
subjects['admin_law'] = {
    'name': '行政法', 'code': 'admin_law', 'days': 6,
    'daily': [
        ('D1', '行政法概述 + 行政主体(行政机关/被授权组织) + 公务员法', 330),
        ('D2', '抽象行政行为(行政法规/规章) + 具体行政行为(概念/分类/效力)', 290),
        ('D3', '行政许可(设定/实施/监督检查) + 行政处罚(种类/设定/实施/执行)', 280),
        ('D4', '行政强制(措施/执行) + 其他行政行为 + 政府信息公开', 330),
        ('D5', '行政复议(范围/复议机关/程序/决定) + 行政诉讼概述+基本原则', 280),
        ('D6', '行政诉讼(受案范围/管辖/参加人/证据/程序/裁判/执行)', 340),
    ]
}

# 商法 (郄鹏恩) - user allocates 6 days, focus on core
subjects['commercial'] = {
    'name': '商法', 'code': 'commercial', 'days': 6,
    'daily': [
        ('D1', '公司法:公司分类+法人人格+股东权利+公司治理(第1-4章)', 300),
        ('D2', '公司法:股东知情权+代表诉讼+董事监事高管+公司融资+合并分立(第5-9章)', 310),
        ('D3', '公司法:股权转让+股份公司+上市公司+国有出资公司+组织架构', 300),
        ('D4', '合伙企业法(普通合伙/有限合伙) + 个人独资企业法 + 外商投资法', 290),
        ('D5', '破产法:破产原因+申请受理+管理人+债权申报+债务人财产+重整和解', 310),
        ('D6', '保险法(合同/人身/财产) + 信托法 + 证券法要点 + 票据法要点', 280),
    ]
}

# 刑法 (柏浪涛) - reconstructed from standard 2026 curriculum, ~55h, user allocates 8 days
subjects['criminal_law'] = {
    'name': '刑法', 'code': 'criminal_law', 'days': 8,
    'daily': [
        ('D1', '刑法论(概述+基本原则+适用范围) + 犯罪构成-客观要件(行为/结果/因果关系)', 340),
        ('D2', '犯罪构成-客观要件(主体/特殊身份) + 主观要件(故意/过失/目的/动机)', 350),
        ('D3', '违法性阻却事由(正当防卫/紧急避险/其他) + 犯罪形态(既遂/未遂/中止/预备)', 330),
        ('D4', '共同犯罪(共犯分类/间接正犯/共犯脱离/共犯处罚原则)', 300),
        ('D5', '罪数形态 + 刑罚论(体系/量刑/累犯/自首/立功/数罪并罚/缓刑/减刑/假释/追诉时效)', 320),
        ('D6', '侵犯人身权利罪(杀人/伤害/强奸/非法拘禁/绑架/拐卖/遗弃/侮辱诽谤)', 340),
        ('D7', '侵犯财产罪(抢劫/抢夺/盗窃/诈骗/敲诈/侵占/职务侵占/故意毁坏财物/挪用)', 350),
        ('D8', '危害公共安全罪 + 经济秩序罪重点 + 妨害社会管理秩序 + 贪污贿赂罪 + 渎职罪', 320),
    ]
}

# 理论法 (马峰) - ~50h, ~7 days
subjects['theory'] = {
    'name': '理论法', 'code': 'theory', 'days': 7,
    'daily': [
        ('D1', '法理学:法的概念(特征/本质/作用) + 法的价值 + 法的要素(规则/原则/概念)', 330),
        ('D2', '法理学:法的渊源与分类 + 法的效力(时间/空间/对象) + 法律关系 + 法律责任', 310),
        ('D3', '法理学:法的制定与实施 + 法适用原理 + 法律解释 + 法律漏洞填补', 320),
        ('D4', '宪法学:宪法基本理论(概念/原则/制定/修改/解释/实施监督) + 国家基本制度', 340),
        ('D5', '宪法学:选举制度 + 特别行政区制度 + 民族区域自治 + 基层群众自治', 310),
        ('D6', '宪法学:公民基本权利与义务 + 国家机构(人大/国务院/监察委/法院/检察院)', 320),
        ('D7', '法治思想(核心要义+实践要求) + 中国法律史(周至清/民国) + 司法制度与法律职业道德', 300),
    ]
}

# 三国法 (杨帆) - compact, time-permitting (not in Phase 1 core)
subjects['intl_law'] = {
    'name': '三国法', 'code': 'intl_law', 'days': 4,
    'daily': [
        ('D1', '国际公法:导论+主体+空间划分+个人+外交领事+条约+争端解决+战争法', 320),
        ('D2', '国际私法:总论+冲突规范+适用法(民事主体/物权/合同/侵权/婚姻家庭)', 290),
        ('D3', '国际私法:争议解决(仲裁/诉讼)+司法协助(送达/取证/承认执行)', 260),
        ('D4', '国际经济法:货物买卖+运输保险+支付(托收/信用证)+贸易管制+WTO+投资+税收', 280),
    ]
}

# ── Build Phase 1 daily plan ──
# Order: 刑法→民法→民诉→刑诉→行政→商法→理论
# (三国/经济法/知产 left for Phase 2 or later)
# 起始日期：默认取运行当天，可用环境变量 FAKAO_START_DATE 固定

order = ['criminal_law', 'civil_law', 'civil_proc', 'crim_proc',
         'admin_law', 'commercial', 'theory']

start = date(2026, 6, 6)
daily_plans = []
current = start

for code in order:
    s = subjects[code]
    allocated = s['days']
    for i in range(allocated):
        label, topics, minutes = s['daily'][i]
        daily_plans.append({
            'date': current.isoformat(),
            'subject': s['name'],
            'subject_code': code,
            'day_num': i + 1,
            'day_label': label,
            'topics': topics,
            'video_minutes': minutes,
            'phase': '精讲阶段',
        })
        current += timedelta(days=1)

print(f"Phase 1: {len(daily_plans)} days")
print(f"Range: {daily_plans[0]['date']} to {daily_plans[-1]['date']}")
print()

# Print daily outline
for dp in daily_plans:
    print(f"{dp['date']}  {dp['subject']}  D{dp['day_num']}  {dp['video_minutes']}min")
    print(f"          {dp['topics'][:100]}")

# Save
out_path = 'data/phase1_daily.json'
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(daily_plans, f, ensure_ascii=False, indent=2)
print(f"\nSaved to {out_path}")

# Summary per subject
print("\n=== Summary ===")
for code in order:
    s = subjects[code]
    total_min = sum(d[2] for d in s['daily'][:s['days']])
    total_h = total_min / 60
    print(f"{s['name']}: {s['days']}天, 约{total_h:.0f}h视频, {total_min}min")
