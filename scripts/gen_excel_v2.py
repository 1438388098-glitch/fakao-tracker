"""Generate Phase 1 study plan Excel — with proper time slot allocation.
User schedule:
  09:00-12:00  → 170 min effective (3h - breaks)
  13:30-17:30  → 230 min effective (4h - breaks)
  19:00-21:00  → 110 min effective (2h - breaks)
  21:00-22:00  → 复习今日内容 (not new video)
Total: ~510 min new content/day
"""
import json
from datetime import date, datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re

def _parse_min(text):
    """Estimate minutes from a task description."""
    matches = re.findall(r'(\d+)min', text)
    if matches:
        return sum(int(m) for m in matches)
    topics = [t.strip() for t in text.split('+')]
    total = 0
    for t in topics:
        m = re.search(r'(\d+)', t)
        if m:
            total += int(m.group(1))
        else:
            total += 40
    return total

# ── Subject day plans (from PDFs) ──
# Each: (day_label, topics_detail, total_video_minutes, [morning_topics, afternoon_topics, evening_topics])

subjects = {}

subjects['criminal_law'] = {  # 刑法 8天 ~44h
    'name': '刑法', 'code': 'criminal_law', 'days': 8,
    'daily': [
        ('刑法论 + 犯罪构成·客观要件(上)', ['刑法概述·基本原则·适用范围(90min)', '犯罪构成·行为理论(80min)', '因果关系(90min)'], 340,
         ['刑法概述+基本原则(90min)', '适用范围(40min)', '行为理论入门(40min)'],
         ['行为理论深入(40min)', '因果关系(90min)', '真金题配套(100min)'],
         ['主观要件预习(60min)', '今日错题回顾(50min)']),
        # D2-D8 similarly...
        ('犯罪构成·客观(下) + 主观要件', ['犯罪主体·特殊身份(80min)', '故意·过失(120min)', '目的犯·动机(90min)'], 350,
         ['犯罪主体+特殊身份(80min)', '故意理论(90min)'],
         ['过失理论(30min)', '目的与动机(90min)', '真金题配套(110min)'],
         ['违法阻却预习(60min)', '今日错题回顾(50min)']),
        ('违法阻却 + 犯罪形态', ['正当防卫(100min)', '紧急避险及其他(80min)', '既遂·未遂·中止·预备(150min)'], 330,
         ['正当防卫(100min)', '紧急避险(70min)'],
         ['其他违法阻却(80min)', '犯罪形态上(150min)'],
         ['罪数形态预习(60min)', '今日错题回顾(50min)']),
        ('共同犯罪', ['共犯分类与基础理论(100min)', '间接正犯(80min)', '共犯脱离+处罚原则(120min)'], 300,
         ['共犯分类+基础理论(100min)', '间接正犯(70min)'],
         ['间接正犯续(10min)', '共犯脱离(80min)', '处罚原则(120min)'],
         ['刑罚论预习(60min)', '今日错题回顾(50min)']),
        ('罪数形态 + 刑罚论', ['罪数形态(80min)', '刑罚体系+量刑(120min)', '累犯·自首·立功+执行消灭(120min)'], 320,
         ['罪数形态(80min)', '刑罚体系(90min)'],
         ['量刑制度(30min)', '累犯自首立功(60min)', '数罪并罚+缓刑(140min)'],
         ['分则预习(60min)', '今日错题回顾(50min)']),
        ('侵犯人身权利罪', ['故意杀人·伤害(80min)', '强奸·强制猥亵(80min)', '非法拘禁·绑架·拐卖(90min)', '遗弃·侮辱诽谤(90min)'], 340,
         ['杀人+伤害(80min)', '强奸+猥亵(90min)'],
         ['非拘+绑架(80min)', '拐卖(50min)', '真金题(100min)'],
         ['遗弃+侮辱(40min)', '财产罪预习(60min)', '错题回顾(10min)']),
        ('侵犯财产罪', ['抢劫·抢夺(100min)', '盗窃·诈骗(100min)', '敲诈·侵占·其他(150min)'], 350,
         ['抢劫+抢夺(100min)', '盗窃(90min)'],
         ['诈骗(50min)', '敲诈+侵占(80min)', '故意毁坏+挪用(100min)'],
         ['经济秩序罪预习(60min)', '错题回顾(50min)']),
        ('危害公共安全 + 经济 + 社会管理 + 贪污渎职', ['危害公共安全(90min)', '经济秩序重点(100min)', '贪污贿赂(80min)', '渎职+收尾(50min)'], 320,
         ['危害公共安全(90min)', '经济秩序罪(80min)'],
         ['社会管理秩序罪(80min)', '贪污贿赂(80min)', '渎职罪(30min)'],
         ['总复习串联(60min)', '错题回顾(50min)']),
    ]
}

subjects['civil_law'] = {  # 民法 13天 ~33h
    'name': '民法', 'code': 'civil_law', 'days': 13,
    'daily': [
        ('民事法律关系 + 自然人(上)', ['民事法律关系(60min)', '自然人(70min)'], 123,
         ['民事法律关系(60min)', '自然人概念(60min)'],
         ['自然人监护(40min)', '宣告失踪+死亡(30min)', '配套真题(80min)'],
         ['预习法人(50min)', '今日错题回顾(50min)']),
        ('自然人(下) + 法人', ['自然人-监护宣告(80min)', '法人-概述(60min)'], 138,
         ['自然人监护宣告(80min)', '法人概述(30min)'],
         ['法人分类(30min)', '法人民事能力(40min)', '配套真题(60min)'],
         ['预习代理(60min)', '今日错题回顾(50min)']),
        ('法人 + 非法人组织 + 民事法律行为', ['法人变更终止(50min)', '非法人组织(30min)', '民事法律行为概述(80min)'], 162,
         ['法人变更终止(50min)', '非法人组织(30min)', '民事法律行为(70min)'],
         ['意思表示(20min)', '民事法律行为成立生效(70min)', '配套真题(70min)'],
         ['预习效力(60min)', '今日错题回顾(50min)']),
        ('民事法律行为效力', ['效力概述+有效要件(50min)', '无效民事法律行为(50min)', '可撤销(40min)', '效力待定(20min)'], 144,
         ['效力概述+有效要件(50min)', '无效民事法律行为(50min)', '可撤销(30min)'],
         ['可撤销续(20min)', '效力待定(20min)', '配套真题(90min)'],
         ['预习代理(60min)', '错题回顾(50min)']),
        ('代理 + 诉讼时效', ['代理概述(60min)', '无权代理+表见代理(50min)', '诉讼时效(60min)'], 105,
         ['代理概述(60min)', '无权代理(30min)'],
         ['表见代理(20min)', '诉讼时效(60min)', '配套真题(70min)'],
         ['预习物权(60min)', '错题回顾(50min)']),
        ('物权基本原则 + 物权变动(上)', ['物权基本原则(60min)', '物权变动-登记(70min)'], 128,
         ['物权基本原则(60min)', '物权变动概述(40min)'],
         ['登记制度(30min)', '交付(20min)', '配套真题(40min)', '额外阅读物权法(60min)'],
         ['预习担保物权(60min)', '错题回顾(50min)']),
        ('物权变动(下) + 所有权 + 担保物权(上)', ['物权变动-善意取得(40min)', '所有权(60min)', '担保物权-抵押(70min)'], 167,
         ['善意取得(40min)', '所有权(60min)', '抵押概述(30min)'],
         ['抵押权(40min)', '抵押实现(30min)', '配套真题(80min)'],
         ['预习质权(60min)', '错题回顾(50min)']),
        ('担保物权·抵押 + 质权', ['抵押权深度(90min)', '质权(90min)'], 184,
         ['抵押权深度(90min)', '质权概述(40min)'],
         ['动产质权+权利质权(50min)', '配套真题(80min)', '留置权预习(30min)'],
         ['预习占有(60min)', '错题回顾(50min)']),
        ('留置权 + 占有 + 债的法定分类', ['留置权(40min)', '占有(50min)', '债的法定分类(50min)'], 133,
         ['留置权(40min)', '占有(50min)', '债的分类(40min)'],
         ['债的分类续(10min)', '无因管理(30min)', '不当得利(30min)', '配套真题(60min)'],
         ['预习合同法(60min)', '错题回顾(50min)']),
        ('合同基本原则 + 缔约过失 + 合同订立', ['合同基本原则(60min)', '缔约过失(40min)', '合同订立-要约承诺(60min)'], 162,
         ['合同基本原则(60min)', '缔约过失(40min)', '合同订立-要约(20min)'],
         ['要约承诺(40min)', '合同成立时间地点(50min)', '格式条款(30min)', '配套真题(60min)'],
         ['预习合同履行(60min)', '错题回顾(50min)']),
        ('合同履行 + 合同保全', ['合同履行-抗辩权(60min)', '合同履行-代位权撤销权(80min)'], 188,
         ['合同履行概述(20min)', '同时履行抗辩(20min)', '不安抗辩(30min)', '顺序履行(30min)'],
         ['代位权(40min)', '撤销权(40min)', '配套真题(100min)'],
         ['预习合同变更(60min)', '错题回顾(50min)']),
        ('合同变更转让 + 合同终止', ['合同变更(30min)', '债权转让+债务承担(60min)', '清偿+抵销+提存(40min)', '解除(30min)'], 158,
         ['合同变更(30min)', '债权转让(60min)', '债务承担(30min)'],
         ['清偿抵销提存(40min)', '合同解除(30min)', '配套真题(100min)'],
         ['预习违约责任(60min)', '错题回顾(50min)']),
        ('违约责任 + 买卖合同', ['违约责任(80min)', '买卖合同(60min)', '供用/赠与/借款/租赁/承揽等(80min)'], 200,
         ['违约责任(80min)', '买卖合同(60min)'],
         ['供用/赠与/借款(40min)', '租赁/承揽(40min)', '配套真题(60min)', '民法快速回顾(50min)'],
         ['总复习(60min)', '错题回顾(50min)']),
    ]
}

# For remaining subjects, create structured entries with time-slot split
# Each entry: (label, [all topics list], total_min, [morning], [afternoon], [evening])

_civil_proc_daily = [
    ('诉的基本理论 + 基本原则 + 管辖(上)',
     ['诉的基本理论(100min)', '基本原则(100min)', '管辖-级别+地域(80min)'], 280,
     ['诉的基本理论(100min)', '基本原则入门(50min)'],
     ['基本原则续(50min)', '级别管辖(40min)', '地域管辖(40min)', '配套真题(60min)'],
     ['预习管辖下(60min)', '错题回顾(50min)']),
    ('管辖(下) + 当事人 + 共同诉讼 + 第三人',
     ['裁定管辖+管辖权异议(90min)', '当事人(100min)', '共同诉讼(40min)', '第三人(60min)'], 290,
     ['裁定管辖+异议(90min)', '当事人资格(80min)'],
     ['当事人续(20min)', '共同诉讼(40min)', '第三人(60min)', '配套真题(80min)'],
     ['预习证据(60min)', '错题回顾(50min)']),
    ('证据 + 证明 + 保全与先予执行',
     ['证据种类+分类(100min)', '证明对象/责任/标准(100min)', '保全+先予执行(80min)'], 280,
     ['证据种类(100min)', '证明对象(70min)'],
     ['证明责任(30min)', '证明标准(40min)', '保全+先予执行(80min)', '配套真题(60min)'],
     ['预习程序(60min)', '错题回顾(50min)']),
    ('强制措施 + 期间送达 + 调解 + 一审程序',
     ['强制措施(30min)', '期间送达(50min)', '调解(70min)', '一审起诉/受理/审理(120min)'], 270,
     ['强制措施+期间送达(80min)', '调解(70min)'],
     ['一审起诉+受理(60min)', '审理+撤诉+缺席(60min)', '配套真题(80min)'],
     ['预习简易程序(60min)', '错题回顾(50min)']),
    ('简易程序 + 小额 + 二审 + 审判监督',
     ['简易程序(80min)', '小额诉讼(30min)', '二审程序(100min)', '审判监督(80min)'], 290,
     ['简易程序(80min)', '小额诉讼(30min)', '二审程序(60min)'],
     ['二审续(40min)', '审判监督(80min)', '配套真题(80min)'],
     ['预习特别程序(60min)', '错题回顾(50min)']),
    ('特别程序 + 督促/公示催告 + 执行 + 仲裁',
     ['特别程序(50min)', '督促+公示催告(50min)', '执行程序(100min)', '仲裁(100min)'], 300,
     ['特别程序(50min)', '督促+公示催告(50min)', '执行程序(70min)'],
     ['执行续(30min)', '仲裁协议+程序(80min)', '司法监督仲裁(30min)', '配套真题(90min)'],
     ['民诉总复习(60min)', '错题回顾(50min)']),
]

_crim_proc_daily = [
    ('刑诉法概述 + 专门机关与参与人 + 管辖',
     ['刑诉法概述(100min)', '专门机关+参与人(130min)', '管辖-立案+审判(110min)'], 340,
     ['刑诉法概述(100min)', '专门机关+参与人(100min)'],
     ['参与人续(30min)', '管辖:立案管辖(50min)', '审判管辖(60min)', '配套真题(80min)'],
     ['预习回避+辩护(60min)', '错题回顾(50min)']),
    ('回避 + 辩护(上)',
     ['回避(40min)', '辩护制度-概述+辩护人(100min)', '辩护人权利(80min)', '法律援助+拒绝辩护(110min)'], 330,
     ['回避(40min)', '辩护制度概述(80min)', '辩护人初步(50min)'],
     ['辩护人权利(80min)', '法律援助(80min)', '拒绝辩护(30min)', '配套真题(60min)'],
     ['预习证据(60min)', '错题回顾(50min)']),
    ('辩护(下) + 刑事代理 + 证据(上)',
     ['辩护制度深度(60min)', '刑事代理(30min)', '证据种类+分类(120min)', '证据规则(140min)'], 350,
     ['辩护制度深度(60min)', '刑事代理(30min)', '证据种类(80min)'],
     ['证据分类(40min)', '证据规则(140min)', '配套真题(70min)'],
     ['预习证明(60min)', '错题回顾(50min)']),
    ('证据(下) + 强制措施(上)',
     ['证明对象/责任/标准(120min)', '强制措施-拘传/取保/监视(110min)'], 330,
     ['证明对象(40min)', '证明责任(40min)', '证明标准(40min)', '拘传+取保候审(60min)'],
     ['监视居住(50min)', '配套真题(100min)', '强制措施续(30min)'],
     ['预习拘留逮捕(60min)', '错题回顾(50min)']),
    ('强制措施(下) + 附带民事诉讼 + 立案 + 侦查',
     ['拘留+逮捕(110min)', '附带民事诉讼(60min)', '立案(40min)', '侦查-行为(60min)', '侦查终结(50min)'], 320,
     ['拘留+逮捕(110min)', '附带民事诉讼(60min)'],
     ['立案(40min)', '侦查行为(60min)', '侦查终结(50min)', '配套真题(90min)'],
     ['预习起诉(60min)', '错题回顾(50min)']),
    ('审查起诉 + 一审 + 二审',
     ['审查起诉(80min)', '公诉一审(80min)', '自诉+简易+速裁(80min)', '二审程序(100min)'], 340,
     ['审查起诉(80min)', '公诉一审(80min)', '自诉程序(40min)'],
     ['简易+速裁(40min)', '二审程序(100min)', '配套真题(80min)'],
     ['预习死刑复核(60min)', '错题回顾(50min)']),
    ('死刑复核 + 审判监督 + 执行 + 特别程序',
     ['死刑复核(60min)', '审判监督(80min)', '执行(80min)', '特别程序(90min)'], 310,
     ['死刑复核(60min)', '审判监督(80min)', '执行程序(50min)'],
     ['执行续(30min)', '特别程序(90min)', '配套真题(90min)'],
     ['刑诉总复习(60min)', '错题回顾(50min)']),
]

_admin_daily = [
    ('行政法概述 + 行政主体 + 公务员',
     ['行政法概述(90min)', '行政主体-机关+授权(120min)', '公务员法(120min)'], 330,
     ['行政法概述(90min)', '行政主体:行政机关(90min)'],
     ['授权组织(30min)', '公务员法(120min)', '配套真题(60min)'],
     ['预习抽象行为(60min)', '错题回顾(50min)']),
    ('抽象行政行为 + 具体行政行为(上)',
     ['抽象行政行为·行政法规规章(100min)', '具体行政行为概念分类效力(100min)'], 290,
     ['抽象行政行为(100min)', '具体行政行为概念(70min)'],
     ['具体行为分类(30min)', '效力(70min)', '配套真题(80min)'],
     ['预习行政许可(60min)', '错题回顾(50min)']),
    ('具体行政行为(下) + 行政许可',
     ['具体行政行为程序+效力(50min)', '行政许可设定(80min)', '许可实施+监督(150min)'], 280,
     ['具体行为程序+效力(50min)', '行政许可设定(80min)', '实施程序(40min)'],
     ['监督管理(60min)', '配套真题(100min)'],
     ['预习行政处罚(60min)', '错题回顾(50min)']),
    ('行政处罚 + 行政强制',
     ['行政处罚种类+设定(90min)', '处罚实施+执行(90min)', '行政强制措施(60min)', '行政强制执行(90min)'], 330,
     ['行政处罚种类+设定(90min)', '处罚实施(90min)'],
     ['处罚执行(30min)', '行政强制措施(60min)', '强制执行(90min)', '配套真题(60min)'],
     ['预习信息公开(60min)', '错题回顾(50min)']),
    ('行政复议 + 行政诉讼概述 + 基本原则',
     ['行政复议范围+机关(80min)', '复议程序+决定(90min)', '行政诉讼概述(40min)', '基本原则(70min)'], 280,
     ['行政复议范围(40min)', '复议机关(40min)', '复议程序(90min)'],
     ['复议决定(30min)', '行政诉讼概述(40min)', '基本原则(70min)', '配套真题(60min)'],
     ['预习诉讼程序(60min)', '错题回顾(50min)']),
    ('行政诉讼·管辖受案 + 程序证据裁判 + 执行',
     ['受案范围+管辖(100min)', '诉讼参加人+证据(100min)', '程序+裁判+执行(140min)'], 340,
     ['受案范围(60min)', '管辖(40min)', '诉讼参加人(70min)'],
     ['证据(60min)', '程序(60min)', '裁判(60min)', '执行(40min)'],
     ['行政法总复习(60min)', '错题回顾(50min)']),
]

_commercial_daily = [
    ('公司法:公司分类+治理(第1-4章)',
     ['公司分类+体系(60min)', '法人人格+股东权利(120min)', '公司治理结构(120min)'], 300,
     ['公司分类(60min)', '法人人格(60min)', '股东权利入门(50min)'],
     ['股东权利深入(70min)', '公司治理(120min)', '配套真题(40min)'],
     ['预习股东权利下(60min)', '错题回顾(50min)']),
    ('公司法:股东权利+治理(第5-9章)',
     ['知情权+代表诉讼(90min)', '董事监事高管(80min)', '融资+合并分立(100min)'], 310,
     ['知情权+代表诉讼(90min)', '董监高(80min)'],
     ['公司融资(50min)', '合并分立(50min)', '配套真题(80min)'],
     ['预习股权转让(60min)', '错题回顾(50min)']),
    ('公司法:股权转让+股份公司+国有公司',
     ['股权转让—有限(80min)', '股份公司设立+治理(90min)', '上市公司(60min)', '国有出资(70min)'], 300,
     ['股权转让(80min)', '股份公司设立(90min)'],
     ['股份公司治理(60min)', '国有出资公司(70min)', '配套真题(40min)'],
     ['预习合伙企业(60min)', '错题回顾(50min)']),
    ('合伙企业法 + 个人独资 + 外商投资',
     ['普通合伙(120min)', '有限合伙(80min)', '个独+外资(90min)'], 290,
     ['普通合伙(120min)', '有限合伙(80min)'],
     ['个独企业(40min)', '外商投资法(50min)', '配套真题(60min)'],
     ['预习破产法(60min)', '错题回顾(50min)']),
    ('破产法:破产程序全流程',
     ['破产原因+申请+受理(100min)', '管理人+债权申报(90min)', '债务人财产+重整+和解(120min)'], 310,
     ['破产原因+申请(80min)', '受理+管理人(90min)'],
     ['债权申报(90min)', '重整和解(80min)', '配套真题(50min)'],
     ['预习保险法(60min)', '错题回顾(50min)']),
    ('保险法 + 信托 + 证券要点',
     ['保险合同法(80min)', '人身保险+财产保险(80min)', '信托法基础(50min)', '证券法要点(70min)'], 280,
     ['保险合同法(80min)', '人身保险(80min)'],
     ['财产保险(30min)', '信托法(50min)', '证券法要点(70min)', '配套真题(60min)'],
     ['商法总复习(60min)', '错题回顾(50min)']),
]

_theory_daily = [
    ('法理学:法的概念 + 价值 + 要素',
     ['法的概念:特征/本质/作用(140min)', '法的价值(50min)', '法的要素:规则/原则/概念(140min)'], 330,
     ['法的概念(上)(90min)', '法的概念(下:作用)(50min)'],
     ['法的价值(50min)', '法的要素:规则(90min)', '配套真题(40min)'],
     ['预习法的渊源(60min)', '错题回顾(50min)']),
    ('法理学:渊源 + 效力 + 关系 + 责任',
     ['法的渊源与分类(100min)', '法的效力(60min)', '法律关系(80min)', '法律责任(70min)'], 310,
     ['法的渊源(100min)', '法的效力(60min)'],
     ['法律关系(80min)', '法律责任(70min)', '配套真题(70min)'],
     ['预习法的实施(60min)', '错题回顾(50min)']),
    ('法理学:制定实施 + 适用 + 解释 + 漏洞',
     ['法的制定与实施(100min)', '法适用的一般原理(80min)', '法律解释(80min)', '法律漏洞填补(60min)'], 320,
     ['法的制定(60min)', '法的实施(40min)', '适用原理(70min)'],
     ['法律解释(80min)', '法律漏洞(60min)', '配套真题(70min)'],
     ['预习宪法(60min)', '错题回顾(50min)']),
    ('宪法学:基本理论 + 国家基本制度',
     ['宪法概念/原则/制定/修改(120min)', '宪法解释+实施监督(80min)', '国家基本制度(140min)'], 340,
     ['宪法概念+原则(80min)', '制定+修改(90min)'],
     ['解释+监督(80min)', '国家基本制度(140min)'],
     ['预习选举制度(60min)', '错题回顾(50min)']),
    ('宪法学:选举 + 特别行政区 + 自治 + 基层',
     ['选举制度(120min)', '特别行政区(90min)', '民族区域+基层群众(100min)'], 310,
     ['选举制度(120min)', '特别行政区制度(90min)'],
     ['民族区域自治(50min)', '基层群众自治(50min)', '配套真题(60min)'],
     ['预习公民权利(60min)', '错题回顾(50min)']),
    ('宪法学:公民权利 + 国家机构',
     ['公民基本权利与义务(140min)', '人大/国务院(100min)', '监察委/法院/检察院(80min)'], 320,
     ['公民基本权利(140min)', '人大制度(90min)'],
     ['国务院+监察委(60min)', '法院+检察院(50min)', '配套真题(40min)'],
     ['预习法治思想(60min)', '错题回顾(50min)']),
    ('法治思想 + 中国法律史 + 司法制度与职业道德',
     ['法治思想核心要义(100min)', '中国法律史(80min)', '司法制度与职业道德(120min)'], 300,
     ['法治思想(100min)', '中国法律史(80min)'],
     ['司法制度(60min)', '法律职业道德(60min)', '配套真题(70min)'],
     ['法理+宪法总复习(60min)', '错题回顾(50min)']),
]

# Map codes to daily plans
all_daily = {
    'criminal_law': subjects['criminal_law']['daily'],
    'civil_law': subjects['civil_law']['daily'],
    'civil_proc': _civil_proc_daily,
    'crim_proc': _crim_proc_daily,
    'admin_law': _admin_daily,
    'commercial': _commercial_daily,
    'theory': _theory_daily,
}

subj_meta = {
    'criminal_law': ('刑法', 8),
    'civil_law': ('民法', 13),
    'civil_proc': ('民诉法', 6),
    'crim_proc': ('刑诉法', 7),
    'admin_law': ('行政法', 6),
    'commercial': ('商法', 6),
    'theory': ('理论法', 7),
}

order = ['criminal_law','civil_law','civil_proc','crim_proc','admin_law','commercial','theory']

SUBJ_COLORS = {
    '刑法': 'FFCDD2', '民法': 'BBDEFB', '民诉法': 'C8E6C9',
    '刑诉法': 'FFE0B2', '行政法': 'E1BEE7', '商法': 'B2DFDB', '理论法': 'FFF9C4',
}

# ── Build full schedule ──
start = date(2026, 6, 6)
rows = []
current = start

for code in order:
    name, days = subj_meta[code]
    daily_list = all_daily[code]
    for i in range(days):
        label, all_topics, total_min, morning, afternoon, evening = daily_list[i]
        # Calculate slot minutes
        morning_min = sum(_parse_min(m) for m in morning)
        afternoon_min = sum(_parse_min(m) for m in afternoon)
        evening_min = sum(_parse_min(m) for m in evening)
        rows.append({
            'date': current,
            'subject': name,
            'code': code,
            'day': i+1,
            'label': label,
            'topics': '; '.join(all_topics),
            'total_min': total_min,
            'morning': morning,
            'afternoon': afternoon,
            'evening': evening,
            'morning_min': morning_min,
            'afternoon_min': afternoon_min,
            'evening_min': evening_min,
        })
        current += timedelta(days=1)

# ── Create Excel ──
wb = Workbook()

# Styles
header_font = Font(name='微软雅黑', bold=True, size=10, color='FFFFFF')
header_fill = PatternFill(start_color='5D4037', end_color='5D4037', fill_type='solid')
sub_header_fill = PatternFill(start_color='8D6E63', end_color='8D6E63', fill_type='solid')
time_font = Font(name='微软雅黑', size=9)
body_font = Font(name='微软雅黑', size=9)
bold_font = Font(name='微软雅黑', size=9, bold=True)
title_font = Font(name='微软雅黑', bold=True, size=14, color='5D4037')
subtitle_font = Font(name='微软雅黑', size=9, color='888888')
thin_border = Border(
    left=Side(style='thin', color='D0D0D0'),
    right=Side(style='thin', color='D0D0D0'),
    top=Side(style='thin', color='D0D0D0'),
    bottom=Side(style='thin', color='D0D0D0'),
)
center = Alignment(horizontal='center', vertical='center', wrap_text=True)
left_wrap = Alignment(horizontal='left', vertical='center', wrap_text=True)

def style_header_row(ws, row, cols, fill=None):
    for c in range(1, cols+1):
        cell = ws.cell(row=row, column=c)
        cell.font = header_font
        cell.fill = fill or header_fill
        cell.alignment = center
        cell.border = thin_border

def style_row(ws, row, cols, fill_color=None):
    for c in range(1, cols+1):
        cell = ws.cell(row=row, column=c)
        cell.font = body_font
        cell.border = thin_border
        cell.alignment = left_wrap
        if fill_color:
            cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')

weekdays_cn = ['周一','周二','周三','周四','周五','周六','周日']

# ═══ Sheet 1: 精讲阶段日程表 ═══
ws = wb.active
ws.title = "精讲阶段日程表"

ws.merge_cells('A1:H1')
ws.cell(row=1, column=1, value='2026法考 · 第一轮精讲计划').font = title_font
ws.row_dimensions[1].height = 30

ws.merge_cells('A2:H2')
ws.cell(row=2, column=1, value='每天 9:00-12:00 | 13:30-17:30 | 19:00-22:00(21:00-22:00复习) | 共53天').font = subtitle_font

# Headers
headers = ['日期', '星期', '科目', '第N天', '今日学习专题', '上午 (9:00-12:00)\n~170min', '下午 (13:30-17:30)\n~230min', '晚上 (19:00-22:00)\n19-21学习 ~110min | 21-22复习']
for c, h in enumerate(headers, 1):
    ws.cell(row=4, column=c, value=h)
style_header_row(ws, 4, len(headers))
ws.row_dimensions[4].height = 36

for i, entry in enumerate(rows):
    row = 5 + i
    d = entry['date']
    wd = weekdays_cn[d.weekday()]
    color = SUBJ_COLORS.get(entry['subject'])

    ws.cell(row=row, column=1, value=d.strftime('%Y.%m.%d'))
    ws.cell(row=row, column=2, value=wd)
    ws.cell(row=row, column=3, value=entry['subject'])
    ws.cell(row=row, column=4, value=f'D{entry["day"]}')
    ws.cell(row=row, column=5, value=entry['label'])
    ws.cell(row=row, column=6, value='\n'.join(f'• {m}' for m in entry['morning']))
    ws.cell(row=row, column=7, value='\n'.join(f'• {m}' for m in entry['afternoon']))
    ws.cell(row=row, column=8, value='\n'.join(f'• {m}' for m in entry['evening']))

    style_row(ws, row, len(headers), color)
    for c in [1,2,3,4]:
        ws.cell(row=row, column=c).alignment = center
    ws.cell(row=row, column=3).font = bold_font

    # Row height based on max lines
    max_lines = max(len(entry['morning']), len(entry['afternoon']), len(entry['evening']), 2)
    ws.row_dimensions[row].height = max(22, 13 * max_lines + 8)

# Col widths
ws.column_dimensions['A'].width = 11
ws.column_dimensions['B'].width = 6
ws.column_dimensions['C'].width = 8
ws.column_dimensions['D'].width = 7
ws.column_dimensions['E'].width = 38
ws.column_dimensions['F'].width = 32
ws.column_dimensions['G'].width = 34
ws.column_dimensions['H'].width = 36

# Summary
sr = 5 + len(rows) + 2
ws.merge_cells(f'A{sr}:H{sr}')
ws.cell(row=sr, column=1, value='各科汇总').font = Font(name='微软雅黑', bold=True, size=12)
for code in order:
    name, days = subj_meta[code]
    ents = [e for e in rows if e['code'] == code]
    total_min = sum(e['total_min'] for e in ents)
    sr += 1
    color = SUBJ_COLORS[name]
    ws.merge_cells(f'A{sr}:B{sr}')
    ws.cell(row=sr, column=1, value=f'  {name}')
    ws.cell(row=sr, column=3, value=f'{days}天')
    ws.cell(row=sr, column=4, value=f'约{total_min/60:.0f}小时')
    ws.cell(row=sr, column=5, value=f'日均{total_min/days:.0f}分钟')
    for c in range(1,6):
        ws.cell(row=sr, column=c).font = bold_font
        ws.cell(row=sr, column=c).border = thin_border
        if color:
            ws.cell(row=sr, column=c).fill = PatternFill(start_color=color, end_color=color, fill_type='solid')

# ═══ Sheet 2: 打卡表 (打印版) ═══
ws2 = wb.create_sheet("每日打卡(可打印)")

ws2.merge_cells('A1:I1')
ws2.cell(row=1, column=1, value='2026法考 · 每日自律打卡').font = title_font

ch = ['日期','科目','D','计划\n(min)','上午完成\n9:00-12:00','下午完成\n13:30-17:30','晚上完成\n19:00-21:00','复习完成\n21:00-22:00','心情\n1-5']
for c, h in enumerate(ch, 1):
    ws2.cell(row=3, column=c, value=h)
style_header_row(ws2, 3, len(ch))
ws2.row_dimensions[3].height = 36

for i, entry in enumerate(rows):
    row = 4 + i
    color = SUBJ_COLORS.get(entry['subject'])
    ws2.cell(row=row, column=1, value=entry['date'].strftime('%m/%d'))
    ws2.cell(row=row, column=2, value=entry['subject'])
    ws2.cell(row=row, column=3, value=entry['day'])
    ws2.cell(row=row, column=4, value=entry['total_min'])
    for c in range(5, 9):
        ws2.cell(row=row, column=c, value='☐')
    ws2.cell(row=row, column=9, value='')
    style_row(ws2, row, len(ch), color)
    for c in range(1,10):
        ws2.cell(row=row, column=c).alignment = center
    ws2.row_dimensions[row].height = 20

ws2.column_dimensions['A'].width = 8
ws2.column_dimensions['B'].width = 7
ws2.column_dimensions['C'].width = 4
ws2.column_dimensions['D'].width = 7
for c in ['E','F','G','H']:
    ws2.column_dimensions[c].width = 13
ws2.column_dimensions['I'].width = 6

# ═══ Sheet 3: 时间分配说明 ═══
ws3 = wb.create_sheet("时间分配说明")
ws3.merge_cells('A1:D1')
ws3.cell(row=1, column=1, value='每日时间分配详解').font = title_font

info = [
    ('上午 9:00-12:00', '3小时', '约170分钟有效学习', '看精讲视频为主。每45分钟休息5分钟。建议2倍速（机构老师语速偏慢），170分钟≈340分钟视频量'),
    ('午休 12:00-13:30', '1.5小时', '—', '吃饭+休息，保证下午精力'),
    ('下午 13:30-17:30', '4小时', '约230分钟有效学习', '继续看视频+做配套真题。真金题按章节做，做完立刻对答案'),
    ('晚餐 17:30-19:00', '1.5小时', '—', '吃饭+散步/运动，大脑切换状态'),
    ('晚上 19:00-21:00', '2小时', '约110分钟有效学习', '收尾当日剩余视频内容，或整理笔记、标记重点'),
    ('晚上 21:00-22:00', '1小时', '复习时间', '不学新内容。回顾今日错题+快速过一遍笔记。记忆类科目（刑法罪名、民法构成要件）留到此时背诵'),
    ('', '', '合计 ~510分钟/天', '≈ 8.5小时有效学习时间'),
]
for i, (label, duration, effective, note) in enumerate(info):
    r = 3 + i
    ws3.cell(row=r, column=1, value=label).font = bold_font
    ws3.cell(row=r, column=2, value=duration).font = body_font
    ws3.cell(row=r, column=3, value=effective).font = Font(name='微软雅黑', size=10, bold=True, color='C4553A')
    ws3.cell(row=r, column=4, value=note).font = body_font
    for c in range(1,5):
        ws3.cell(row=r, column=c).border = thin_border
    ws3.row_dimensions[r].height = 30
    ws3.cell(row=r, column=1).alignment = center
    ws3.cell(row=r, column=2).alignment = center
    ws3.cell(row=r, column=3).alignment = center

ws3.column_dimensions['A'].width = 22
ws3.column_dimensions['B'].width = 10
ws3.column_dimensions['C'].width = 22
ws3.column_dimensions['D'].width = 70

# Save
out = r'D:\Claudeworkspace\法考一轮复习计划_2026.xlsx'
wb.save(out)
print(f'Done: {out}')
print(f'Row count: {len(rows)} days')
# Verify slot totals match daily totals
for i, entry in enumerate(rows):
    slot_total = entry['morning_min'] + entry['afternoon_min'] + entry['evening_min']
    if abs(slot_total - entry['total_min']) > 10:
        print(f'WARNING: {entry["date"]} {entry["subject"]} total={entry["total_min"]} slots={slot_total}')

