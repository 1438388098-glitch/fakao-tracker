"""Seed 8 subjects + Phase 1 (45 days, 3 daily slots) from v3 plan data."""
import json, os, sys
from datetime import date, timedelta, datetime
from .db import db
from .models import Subject, DailyPlan

SUBJECTS = [
    ("criminal_law", "刑法",   "⚖️", "#E74C3C", 1),
    ("civil_law",    "民法",   "📜", "#3498DB", 2),
    ("civil_proc",   "民诉法", "📋", "#2ECC71", 3),
    ("crim_proc",    "刑诉法", "🔍", "#E67E22", 4),
    ("admin_law",    "行政法", "🏛️", "#9B59B6", 5),
    ("commercial",   "商经知", "💼", "#1ABC9C", 6),
    ("theory",       "理论法", "📖", "#F39C12", 7),
    ("intl_law",     "三国法", "🌏", "#3498DB", 8),
]

# Lazy import from gen_excel_v3
def _get_v3_data():
    scripts_dir = os.path.join(os.path.dirname(__file__), '..', 'scripts')
    sys.path.insert(0, scripts_dir)
    from gen_excel_v3 import CRIMINAL, CIVIL, CIVIL_PROC, CRIM_PROC, ADMIN, COMMERCIAL, THEORY, SUBJ, split_topics
    return CRIMINAL, CIVIL, CIVIL_PROC, CRIM_PROC, ADMIN, COMMERCIAL, THEORY, SUBJ, split_topics


def seed_subjects():
    for code, name, emoji, color, order in SUBJECTS:
        db.session.add(Subject(code=code, name=name, emoji=emoji, color=color, sort_order=order))
    db.session.flush()


def _sub_id(code):
    s = Subject.query.filter_by(code=code).first()
    return s.id if s else None


def seed_daily_plans():
    CRIMINAL, CIVIL, CIVIL_PROC, CRIM_PROC, ADMIN, COMMERCIAL, THEORY, SUBJ, split_topics = _get_v3_data()

    # Subject name -> code mapping
    name_to_code = {
        '刑法': 'criminal_law', '民法': 'civil_law', '民诉': 'civil_proc',
        '刑诉': 'crim_proc', '行政法': 'admin_law', '商法': 'commercial',
        '理论法': 'theory',
    }

    import os
    from datetime import date as _date, timedelta as _td
    start = (_date.fromisoformat(os.environ["FAKAO_START_DATE"])
             if os.environ.get("FAKAO_START_DATE") else _date.today())
    current = start
    order_idx = 0

    for sname, daily_data in SUBJ:
        code = name_to_code[sname]
        subj_id = _sub_id(code)

        for i, (label, total_min, topics) in enumerate(daily_data):
            morning, afternoon, evening = split_topics(topics, total_min)

            # Build short task descriptions
            morning_str = '\n'.join(f'• {t}' for t in morning)
            afternoon_str = '\n'.join(f'• {t}' for t in afternoon)
            evening_str = '\n'.join(f'• {t}' for t in evening)

            # First lines for brief display
            morning_brief = morning[0] if morning else ''
            afternoon_brief = afternoon[0] if afternoon else ''
            evening_brief = evening[0] if evening else ''

            dp = DailyPlan(
                plan_date=current,
                phase_name="精讲阶段",
                subject_id=subj_id,
                task_title=f"{sname} · 第{i+1}天",
                task_type="精讲",
                planned_minutes=total_min,
                morning_task=morning_brief[:500],
                afternoon_task=afternoon_brief[:500],
                evening_task=evening_brief[:500],
                morning_items=morning_str,
                afternoon_items=afternoon_str,
                evening_items=evening_str,
                day_label=label,
                sort_order=order_idx,
            )
            db.session.add(dp)
            current += timedelta(days=1)
            order_idx += 1

    # Add remaining phases (simplified)
    _add_phase2(start + _td(days=45), order_idx)
    _add_phase3(start + _td(days=86), order_idx + 33)
    _add_exam_and_subjective(start + _td(days=98), start + _td(days=134), order_idx + 45)


def _add_phase2(start, order):
    """真题+背诵阶段 (7/21 - 8/30, ~41 days)"""
    schedule = [
        ("criminal_law","刑法",3), ("civil_law","民法",4), ("crim_proc","刑诉法",3),
        ("civil_proc","民诉法",3), ("admin_law","行政法",3), ("commercial","商经知",4),
        ("theory","理论法",4), ("intl_law","三国法",3), (None,"跨科融合",3),
    ]
    current = start
    idx = order
    for code, name, days in schedule:
        for _ in range(days):
            subj_id = _sub_id(code) if code else None
            db.session.add(DailyPlan(
                plan_date=current, phase_name="真题+背诵阶段", subject_id=subj_id,
                task_title=f"{name}真题+背诵", task_type="真题+背诵", planned_minutes=420,
                morning_task=f"上午: {name}真金题刷题", afternoon_task=f"下午: {name}背诵版精读+记忆",
                evening_task="晚上: 错题整理+薄弱点回看", sort_order=idx,
            ))
            current += timedelta(days=1)
            idx += 1


def _add_phase3(start, order):
    """冲刺阶段 (8/31 - 9/11, 12 days)"""
    days = [
        ("全真模考1-公法卷","模考:公法卷(9:00-12:00)","逐题复盘公法卷","错题整理"),
        ("全真模考1-私法卷","模考:私法卷(14:30-17:30)","逐题复盘私法卷","两卷错题汇总分析"),
        ("全真模考2-公法卷","模考:公法卷(9:00-12:00)","逐题复盘公法卷","错题整理"),
        ("全真模考2-私法卷","模考:私法卷(14:30-17:30)","逐题复盘私法卷","两卷错题汇总"),
        ("错题清零+易混对比","错题本全量重做","易混对比表集中记忆","高频错题三刷"),
        ("新增考点突击","2026大纲新增考点精读+配套题","新增考点笔记整理","理论+三国+刑诉狂背"),
        ("全真模考3-公法卷","模考:公法卷(9:00-12:00)","逐题复盘","理论法/三国法狂背"),
        ("全真模考3-私法卷","模考:私法卷(14:30-17:30)","逐题复盘","考前串讲+查漏"),
        ("考前最后冲刺1","轻量模考保持手感","重点法条快速过","早睡+心态调整"),
        ("考前最后冲刺2","考前串讲","最后查漏补缺","早睡休息"),
        ("客观题考试Day1","公法卷9:00-12:00","私法卷14:30-17:30","休息"),
        ("客观题考试Day2","备用","主观题准备","休息"),
    ]
    cur = start
    for i, (title, morn, aft, eve) in enumerate(days):
        db.session.add(DailyPlan(
            plan_date=cur, phase_name="冲刺阶段", subject_id=None,
            task_title=title, task_type="模考", planned_minutes=480,
            morning_task=morn, afternoon_task=aft, evening_task=eve, sort_order=order+i,
        ))
        cur += timedelta(days=1)


def _add_exam_and_subjective(exam_obj, exam_subj, order):
    """Exam days + 主观题阶段."""
    # Exam days
    db.session.add(DailyPlan(
        plan_date=exam_obj, phase_name="客观题考试", subject_id=None,
        task_title="客观题考试 Day1", task_type="考试", planned_minutes=480,
        morning_task="公法卷 9:00-12:00", afternoon_task="私法卷 14:30-17:30",
        evening_task="休息", sort_order=order,
    ))
    db.session.add(DailyPlan(
        plan_date=exam_obj + timedelta(days=1), phase_name="客观题考试", subject_id=None,
        task_title="客观题考试 Day2", task_type="考试", planned_minutes=480,
        morning_task="备用考试日", afternoon_task="准备主观题", evening_task="休息",
        sort_order=order+1,
    ))
    # Subjective (simplified)
    subj_start = exam_subj - timedelta(days=34)
    subjects_4 = [("criminal_law","刑法",7),("civil_law","民法+民诉+商法",7),
                  ("admin_law","行政法",7),("theory","法治思想+论述",7)]
    cur = subj_start
    idx = order + 2
    for code, name, days in subjects_4:
        for _ in range(days):
            db.session.add(DailyPlan(
                plan_date=cur, phase_name="主观题阶段", subject_id=_sub_id(code),
                task_title=f"主观题-{name}", task_type="主观题", planned_minutes=480,
                morning_task=f"{name}主观题练习(计时)", afternoon_task="对照答案修改+法条定位",
                evening_task="论述素材背诵+案例复盘", sort_order=idx,
            ))
            cur += timedelta(days=1)
            idx += 1
    # Final stretch
    for i, title in enumerate(["全真模考1","全真模考2","全真模考3","查漏补缺","考前调整","考前准备"]):
        db.session.add(DailyPlan(
            plan_date=cur, phase_name="主观题阶段", subject_id=None,
            task_title=f"主观题{title}", task_type="模考", planned_minutes=480,
            morning_task=f"模考(9:00-14:00)" if "模考" in title else "核心法条强化",
            afternoon_task="对照答案修改" if "模考" in title else "论述素材最终记忆",
            evening_task="错题复盘" if "模考" in title else "早睡准备",
            sort_order=idx+i,
        ))
        cur += timedelta(days=1)
    # Exam day
    db.session.add(DailyPlan(
        plan_date=exam_subj, phase_name="主观题考试", subject_id=None,
        task_title="主观题考试", task_type="考试", planned_minutes=300,
        morning_task="主观题考试 9:00-14:00(240分钟)", afternoon_task="考试完成!",
        evening_task="休息", sort_order=idx+6,
    ))


def seed_all():
    seed_subjects()
    seed_daily_plans()
    db.session.commit()
