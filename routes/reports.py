from datetime import datetime, timezone

from flask import Blueprint, render_template, request
from flask_login import login_required
from sqlalchemy import func, extract

from extensions import db
from models import PaymentRequest, PaymentRequestItem, Branch, Category
from utils import format_naira

reports_bp = Blueprint("reports", __name__)


@reports_bp.route("/reports")
@login_required
def reports():
    now = datetime.now(timezone.utc)
    year_filter = int(request.args.get("year", now.year))

    monthly_data = (
        db.session.query(
            extract("month", PaymentRequest.created_at).label("month"),
            Branch.name.label("branch"),
            func.sum(PaymentRequest.requested_amount).label("requested"),
            func.sum(PaymentRequest.approved_amount).label("approved"),
            func.count(PaymentRequest.id).label("count"),
        )
        .join(Branch, PaymentRequest.branch_id == Branch.id)
        .filter(extract("year", PaymentRequest.created_at) == year_filter)
        .group_by("month", Branch.name)
        .order_by("month")
        .all()
    )

    cat_data = (
        db.session.query(
            Category.name.label("category"),
            Category.cost_type.label("cost_type"),
            func.sum(PaymentRequestItem.amount).label("total"),
            func.count(PaymentRequestItem.id).label("count"),
        )
        .join(PaymentRequestItem, PaymentRequestItem.category_id == Category.id)
        .join(PaymentRequest, PaymentRequest.id == PaymentRequestItem.request_id)
        .filter(
            PaymentRequest.status == "approved",
            extract("year", PaymentRequest.created_at) == year_filter,
        )
        .group_by(Category.name, Category.cost_type)
        .order_by(func.sum(PaymentRequestItem.amount).desc())
        .all()
    )

    status_counts = (
        db.session.query(PaymentRequest.status, func.count(PaymentRequest.id))
        .filter(extract("year", PaymentRequest.created_at) == year_filter)
        .group_by(PaymentRequest.status)
        .all()
    )
    status_map = {s: c for s, c in status_counts}

    variances = (
        PaymentRequest.query
        .filter(
            PaymentRequest.status == "approved",
            PaymentRequest.approved_amount != PaymentRequest.requested_amount,
            extract("year", PaymentRequest.created_at) == year_filter,
        )
        .order_by(PaymentRequest.reviewed_at.desc())
        .limit(20)
        .all()
    )

    total_requested = db.session.query(func.sum(PaymentRequest.requested_amount))\
        .filter(extract("year", PaymentRequest.created_at) == year_filter).scalar() or 0
    total_approved = db.session.query(func.sum(PaymentRequest.approved_amount))\
        .filter(
            PaymentRequest.status == "approved",
            extract("year", PaymentRequest.created_at) == year_filter,
        ).scalar() or 0

    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

    return render_template(
        "reports.html",
        monthly_data=monthly_data,
        cat_data=cat_data,
        status_map=status_map,
        variances=variances,
        format_naira=format_naira,
        year_filter=year_filter,
        months=months,
        now=now,
        total_requested=total_requested,
        total_approved=total_approved,
    )
