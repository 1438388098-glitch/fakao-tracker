from datetime import date, timedelta
from collections import defaultdict
from flask import Blueprint, render_template
from database.models import Subject, DailyPlan, CheckIn

schedule_bp = Blueprint("schedule", __name__)

WEEKDAY_CN = ['一','二','三','四','五','六','日']


@schedule_bp.route("/schedule")
def view():
    today = date.today()

    # All plans ordered by date
    all_plans = DailyPlan.query.order_by(DailyPlan.plan_date, DailyPlan.sort_order).all()

    # Group by phase
    phase_plans = defaultdict(list)
    for p in all_plans:
        phase_plans[p.phase_name].append(p)
    phase_names = list(phase_plans.keys())

    # For each phase, group by month then by week
    phase_data = {}
    for pn in phase_names:
        plans = phase_plans[pn]
        # Group by month
        months = defaultdict(list)
        for p in plans:
            month_key = f"{p.plan_date.year}年{p.plan_date.month}月"
            months[month_key].append(p)

        month_weeks = []
        for month_label, month_plans in months.items():
            weeks = _group_into_weeks(month_plans)
            month_weeks.append((month_label, weeks))
        phase_data[pn] = month_weeks

    # Checkins for status dots
    all_dates = [p.plan_date for p in all_plans]
    checkins = CheckIn.query.filter(CheckIn.check_date.in_(all_dates)).all()
    checkin_map = {c.check_date: c for c in checkins}

    return render_template(
        "schedule.html",
        today=today,
        phase_names=phase_names,
        phase_data=phase_data,
        checkin_map=checkin_map,
    )


def _group_into_weeks(plans):
    """Group plans into weeks. Each week has 7 day slots."""
    weeks = []
    for p in plans:
        dow = p.plan_date.weekday()
        if not weeks or dow == 0:
            weeks.append({"start": p.plan_date, "days": []})
        if not weeks:
            weeks.append({"start": p.plan_date, "days": []})
        weeks[-1]["days"].append(p)

    # Fill each week to 7 days
    result = []
    for w in weeks:
        monday = w["start"] - timedelta(days=w["start"].weekday())
        full_days = []
        for i in range(7):
            d = monday + timedelta(days=i)
            plan = None
            for p in w["days"]:
                if p.plan_date == d:
                    plan = p
                    break
            full_days.append({"date": d, "plan": plan})
        result.append({"start": monday, "end": monday + timedelta(days=6), "days": full_days})
    return result
