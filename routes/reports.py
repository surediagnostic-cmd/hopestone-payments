from datetime import datetime, timezone

from flask import Blueprint, render_template, request
from flask_login import login_required
from sqlalchemy import func, extract

from extensions import db
from models import PaymentRequest, PaymentRequestItem, Branch, Category, Budget
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

    # Leaf-items-only subquery (exclude parent roll-up rows to avoid double-counting)
    _parent_ids_sq = db.session.query(PaymentRequestItem.parent_id).filter(
        PaymentRequestItem.parent_id.isnot(None)
    ).subquery()

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
            ~PaymentRequestItem.id.in_(_parent_ids_sq),
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

    # ── Budget vs actual by category ───────────────────────────────────────
    budget_rows = (
        db.session.query(
            Category.id.label("cat_id"),
            Category.name.label("cat_name"),
            Category.cost_type.label("cost_type"),
            func.sum(Budget.monthly_amount).label("budget_total"),
        )
        .join(Budget, Budget.category_id == Category.id)
        .filter(Budget.year == year_filter)
        .group_by(Category.id, Category.name, Category.cost_type)
        .order_by(func.sum(Budget.monthly_amount).desc())
        .all()
    )

    # actual per category for the year
    actual_by_cat = {
        row.category: float(row.total or 0) for row in cat_data
    }

    budget_vs_actual = []
    for br in budget_rows:
        budget_amt = float(br.budget_total or 0)
        actual_amt = actual_by_cat.get(br.cat_name, 0)
        variance   = actual_amt - budget_amt
        pct        = round(actual_amt / budget_amt * 100, 1) if budget_amt > 0 else None
        budget_vs_actual.append({
            "name":      br.cat_name,
            "cost_type": br.cost_type,
            "budget":    budget_amt,
            "actual":    actual_amt,
            "variance":  variance,
            "pct":       pct,
        })

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
        budget_vs_actual=budget_vs_actual,
    )
