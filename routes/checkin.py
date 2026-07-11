import os, shutil, csv, io
from datetime import date, datetime
from flask import Blueprint, render_template, request, jsonify, send_file
from database.db import db
from database.models import CheckIn, DailyPlan

checkin_bp = Blueprint("checkin", __name__)


@checkin_bp.route("/api/checkin", methods=["POST"])
def do_checkin():
    data = request.get_json() or {}
    today = date.today()
    mood = data.get("mood", 0)
    notes = data.get("notes", "")

    existing = CheckIn.query.filter_by(check_date=today).first()
    repeated = False
    if existing:
        existing.completed = True
        if mood: existing.mood = mood
        if notes: existing.notes = notes
        repeated = True
    else:
        ci = CheckIn(
            check_date=today,
            completed=True,
            actual_minutes=data.get("actual_minutes", 0),
            notes=notes,
            mood=mood,
        )
        db.session.add(ci)

    db.session.commit()
    return jsonify({"ok": True, "repeated": repeated})


@checkin_bp.route("/api/checkin/undo", methods=["POST"])
def undo_checkin():
    today = date.today()
    ci = CheckIn.query.filter_by(check_date=today).first()
    if ci:
        ci.completed = False
        db.session.commit()
    return jsonify({"ok": True})


@checkin_bp.route("/history")
def history():
    # All checkins
    all_checkins = (
        CheckIn.query
        .order_by(CheckIn.check_date.desc())
        .limit(120)
        .all()
    )

    # Stats
    today = date.today()
    total_days = CheckIn.query.filter_by(completed=True).count()
    total_minutes = db.session.query(
        db.func.sum(CheckIn.actual_minutes)
    ).filter_by(completed=True).scalar() or 0

    # On-track: days with both plan AND checkin, out of total planned days up to today
    planned_dates = db.session.query(DailyPlan.plan_date).filter(
        DailyPlan.plan_date <= today,
        DailyPlan.task_type != "休息"
    ).all()
    planned_days = len(planned_dates)
    planned_date_set = {p[0] for p in planned_dates}
    completed_on_plan = CheckIn.query.filter(
        CheckIn.completed == True,  # noqa
        CheckIn.check_date.in_(planned_date_set)
    ).count() if planned_date_set else 0
    on_track_pct = round(completed_on_plan / planned_days * 100, 1) if planned_days > 0 else 0

    checkin_by_date = {c.check_date: c for c in all_checkins}

    return render_template(
        "history.html",
        checkins=all_checkins,
        checkin_by_date=checkin_by_date,
        total_days=total_days,
        total_minutes=total_minutes,
        planned_days=planned_days,
        on_track_pct=on_track_pct,
        today=today,
    )


@checkin_bp.route("/api/backup")
def backup_data():
    """Export all checkins as CSV + backup database file."""
    import zipfile

    # Create backup directory
    backup_dir = os.path.join(os.path.dirname(__file__), '..', 'data', 'backups')
    os.makedirs(backup_dir, exist_ok=True)
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')

    # CSV export
    csv_path = os.path.join(backup_dir, f'checkins_{ts}.csv')
    checkins = CheckIn.query.order_by(CheckIn.check_date).all()
    with open(csv_path, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(['date', 'completed', 'minutes', 'mood', 'notes'])
        for c in checkins:
            w.writerow([c.check_date.isoformat(), int(c.completed), c.actual_minutes, c.mood, c.notes])

    # DB copy
    db_path = os.path.join(os.path.dirname(__file__), '..', 'data', 'fakao.db')
    db_copy = os.path.join(backup_dir, f'fakao_{ts}.db')
    shutil.copy2(db_path, db_copy)

    # Clean old backups (keep last 10)
    files = sorted(os.listdir(backup_dir))
    for old in files[:-20]:  # keep 10 csv + 10 db
        os.remove(os.path.join(backup_dir, old))

    return jsonify({
        "ok": True,
        "csv": csv_path,
        "db": db_copy,
        "checkins": len(checkins),
        "timestamp": ts,
    })
