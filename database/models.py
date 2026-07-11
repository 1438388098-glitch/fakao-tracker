from datetime import date, datetime
from .db import db


class Subject(db.Model):
    __tablename__ = "subjects"

    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(20), unique=True, nullable=False)
    name = db.Column(db.String(50), nullable=False)
    emoji = db.Column(db.String(10), default="")
    color = db.Column(db.String(7), default="#C4553A")  # hex color for calendar badge
    sort_order = db.Column(db.Integer, default=0)

    chapters = db.relationship("Chapter", backref="subject", lazy="select",
                               order_by="Chapter.sort_order")
    daily_plans = db.relationship("DailyPlan", backref="subject", lazy="select")


class Chapter(db.Model):
    __tablename__ = "chapters"

    id = db.Column(db.Integer, primary_key=True)
    subject_id = db.Column(db.Integer, db.ForeignKey("subjects.id"), nullable=False)
    title = db.Column(db.String(200), nullable=False)
    sort_order = db.Column(db.Integer, default=0)


class DailyPlan(db.Model):
    __tablename__ = "daily_plans"

    id = db.Column(db.Integer, primary_key=True)
    plan_date = db.Column(db.Date, nullable=False, index=True)
    phase_name = db.Column(db.String(50), nullable=False)  # 精讲/真题+背诵/冲刺/主观题
    subject_id = db.Column(db.Integer, db.ForeignKey("subjects.id"), nullable=True)
    task_title = db.Column(db.String(200), nullable=False)
    task_type = db.Column(db.String(20), nullable=False)  # 精讲/真题/背诵/模考/休息
    planned_minutes = db.Column(db.Integer, default=0)
    morning_task = db.Column(db.String(500), default="")
    afternoon_task = db.Column(db.String(500), default="")
    evening_task = db.Column(db.String(500), default="")
    morning_items = db.Column(db.Text, default="")  # JSON list of morning slot items
    afternoon_items = db.Column(db.Text, default="")  # JSON list of afternoon slot items
    evening_items = db.Column(db.Text, default="")  # JSON list of evening slot items
    day_label = db.Column(db.String(200), default="")  # Short label like "刑法论 + 犯罪构成"
    sort_order = db.Column(db.Integer, default=0)


class CheckIn(db.Model):
    __tablename__ = "checkins"

    id = db.Column(db.Integer, primary_key=True)
    check_date = db.Column(db.Date, unique=True, nullable=False, index=True)
    completed = db.Column(db.Boolean, default=False)
    actual_minutes = db.Column(db.Integer, default=0)
    notes = db.Column(db.Text, default="")
    mood = db.Column(db.Integer, default=0)  # 1-5, 0=未设置
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    sessions = db.relationship("StudySession", backref="checkin", lazy="select")


class StudySession(db.Model):
    __tablename__ = "study_sessions"

    id = db.Column(db.Integer, primary_key=True)
    checkin_id = db.Column(db.Integer, db.ForeignKey("checkins.id"), nullable=False)
    subject_id = db.Column(db.Integer, db.ForeignKey("subjects.id"), nullable=True)
    start_time = db.Column(db.DateTime, nullable=False)
    end_time = db.Column(db.DateTime, nullable=True)
    notes = db.Column(db.Text, default="")

    subject = db.relationship("Subject", lazy="select")
