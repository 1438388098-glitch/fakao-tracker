from datetime import date, timedelta
from flask import Blueprint, render_template, jsonify, request
from database.db import db
from database.models import Subject, DailyPlan, CheckIn

dashboard_bp = Blueprint("dashboard", __name__)

EXAM_DATE = date(2026, 9, 12)


@dashboard_bp.route("/dashboard")
def home():
    today = date.today()
    days_left = (EXAM_DATE - today).days

    # Today's plan (usually 1 per day)
    today_plan = DailyPlan.query.filter_by(plan_date=today).first()

    # Today's checkin
    checkin = CheckIn.query.filter_by(check_date=today).first()

    # Weekly calendar: Monday..Sunday of current week
    monday = today - timedelta(days=today.weekday())
    week_days = [monday + timedelta(days=i) for i in range(7)]

    week_plans = {}
    for d in week_days:
        plan = DailyPlan.query.filter_by(plan_date=d).first()
        week_plans[d] = plan

    week_checkins = {}
    checkins = CheckIn.query.filter(CheckIn.check_date.in_(week_days)).all()
    for c in checkins:
        week_checkins[c.check_date] = c

    # Streak counter
    streak = _calc_streak(today)

    # Weekly summary
    week_stats = _week_stats(week_days, week_checkins)

    # Phase progress
    phase_stats = _phase_progress(today)

    subjects = {s.id: s for s in Subject.query.all()}

    return render_template(
        "dashboard.html",
        today=today,
        days_left=days_left,
        today_plan=today_plan,
        checkin=checkin,
        week_days=week_days,
        week_plans=week_plans,
        week_checkins=week_checkins,
        streak=streak,
        week_stats=week_stats,
        phase_stats=phase_stats,
        subjects=subjects,
    )


def _phase_progress(today):
    """Calculate days completed / total for each phase."""
    from sqlalchemy import func

    phases = ["精讲阶段", "真题+背诵阶段", "冲刺阶段", "主观题阶段"]
    stats = []
    for p in phases:
        total = DailyPlan.query.filter_by(phase_name=p).count()
        done = DailyPlan.query.filter(
            DailyPlan.phase_name == p,
            DailyPlan.plan_date <= today
        ).count()
        # More accurate: count check-ins for this phase
        phase_plans = DailyPlan.query.filter_by(phase_name=p).all()
        phase_dates = [dp.plan_date for dp in phase_plans]
        checked = 0
        if phase_dates:
            checked = CheckIn.query.filter(
                CheckIn.check_date.in_(phase_dates),
                CheckIn.completed == True  # noqa
            ).count()
        stats.append({
            "name": p,
            "total": total,
            "done": done,
            "checked": checked,
            "pct": round(checked / total * 100) if total > 0 else 0,
        })
    return stats


def _calc_streak(today):
    """Count consecutive check-in days ending at today or yesterday."""
    checkins = CheckIn.query.filter(
        CheckIn.completed == True  # noqa
    ).order_by(CheckIn.check_date.desc()).all()

    if not checkins:
        return 0

    # Start counting from today (if checked in) or yesterday
    expected = today
    if checkins[0].check_date < today:
        expected = today - timedelta(days=1)

    count = 0
    for c in checkins:
        if c.check_date == expected:
            count += 1
            expected -= timedelta(days=1)
        elif c.check_date < expected:
            break
    return count


def _week_stats(week_days, week_checkins):
    """Calculate this week's stats."""
    checked = sum(1 for d in week_days if week_checkins.get(d) and week_checkins[d].completed)
    total = len(week_days)
    minutes = sum(
        (week_checkins[d].actual_minutes or 0) for d in week_days
        if week_checkins.get(d) and week_checkins[d].completed
    )
    rest = sum(1 for d in week_days if d > date.today())
    return {"checked": checked, "total": total, "minutes": minutes, "remaining": rest}


# ── Reschedule API ──

@dashboard_bp.route("/api/reschedule", methods=["POST"])
def reschedule():
    """Shift + compact: move all uncompleted plans from today onward, then fill gaps."""
    data = request.get_json() or {}
    days = data.get("days", 0)
    if days == 0:
        return jsonify({"ok": False, "error": "days is required"})

    today = date.today()

    # Collect completed dates (locked in place)
    completed_dates = set(
        c.check_date for c in CheckIn.query.filter(
            CheckIn.check_date >= today,
            CheckIn.completed == True  # noqa
        ).all()
    )

    # Find all plans from today onward, sorted by date
    all_future = DailyPlan.query.filter(
        DailyPlan.plan_date >= today
    ).order_by(DailyPlan.plan_date).all()

    # Separate locked (completed or today) vs movable
    locked = {}   # date -> plan
    movable = []  # plans that can be rescheduled
    for p in all_future:
        if p.plan_date == today or p.plan_date in completed_dates:
            locked[p.plan_date] = p
        else:
            movable.append(p)

    # Apply shift to movable plans
    for p in movable:
        p.plan_date = p.plan_date + timedelta(days=days)

    # Compact: reassign dates sequentially to fill gaps
    # Re-fetch sorted movable plans after shift
    movable.sort(key=lambda p: p.plan_date)

    # Build occupied dates (locked + today)
    occupied = set(locked.keys())
    # Also mark today as occupied if a plan is locked there
    occupied.add(today)

    # Reassign movable plans to earliest available slots
    cursor = today + timedelta(days=1)
    shifted = 0
    for p in movable:
        # Find next available date (skip occupied + already-assigned)
        while cursor in occupied:
            cursor += timedelta(days=1)
        if p.plan_date != cursor:
            p.plan_date = cursor
            shifted += 1
        occupied.add(cursor)
        cursor += timedelta(days=1)

    db.session.commit()
    return jsonify({"ok": True, "shifted": shifted, "days": days})


@dashboard_bp.route("/api/status")
def status():
    """Check if user is ahead/on-track/behind."""
    today = date.today()
    # Count days checked in vs days planned (up to today)
    planned = DailyPlan.query.filter(DailyPlan.plan_date <= today).count()
    checked = CheckIn.query.filter(
        CheckIn.check_date <= today,
        CheckIn.completed == True  # noqa
    ).count()
    # Count how many future plans are past their original subject
    today_plan = DailyPlan.query.filter_by(plan_date=today).first()
    today_checkin = CheckIn.query.filter_by(check_date=today).first()

    return jsonify({
        "ok": True,
        "planned": planned,
        "checked": checked,
        "ahead_by": checked - planned if checked >= planned else 0,
        "today_has_plan": today_plan is not None,
        "today_checked": today_checkin is not None and today_checkin.completed,
    })


# ── Timer API (simple) ──
active_timers = {}  # In production, use DB StudySession


@dashboard_bp.route("/api/timer/start", methods=["POST"])
def timer_start():
    import json
    from datetime import datetime
    from flask import request
    data = request.get_json() or {}
    subject_id = data.get("subject_id")
    timer_id = str(len(active_timers) + 1)
    active_timers[timer_id] = {
        "start": datetime.utcnow().isoformat(),
        "subject_id": subject_id,
    }
    return jsonify({"ok": True, "timer_id": timer_id})


@dashboard_bp.route("/api/timer/stop", methods=["POST"])
def timer_stop():
    import json
    from datetime import datetime
    from flask import request
    data = request.get_json() or {}
    timer_id = data.get("timer_id")
    if timer_id and timer_id in active_timers:
        t = active_timers.pop(timer_id)
        elapsed = (datetime.utcnow() - datetime.fromisoformat(t["start"])).total_seconds()
        return jsonify({"ok": True, "elapsed_seconds": int(elapsed)})
    return jsonify({"ok": False, "error": "timer not found"}), 404
