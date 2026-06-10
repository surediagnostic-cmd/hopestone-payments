from datetime import datetime, timezone
from decimal import Decimal
from functools import wraps

from flask import Blueprint, render_template, request, jsonify, flash, redirect, url_for
from flask_login import login_required, current_user
from sqlalchemy import func, extract

from extensions import db
from models import Budget, Branch, Category, PaymentRequest, PaymentRequestItem
from utils import format_naira

budget_bp = Blueprint("budget", __name__)

WEEKS_PER_MONTH = 4.33


def _mds_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not current_user.is_authenticated or not current_user.is_mds:
            flash("Access restricted to MD.", "danger")
            return redirect(url_for("requests.dashboard"))
        return f(*args, **kwargs)
    return decorated


# ── Budget page ─────────────────────────────────────────────────────────────

@budget_bp.route("/budget")
@login_required
@_mds_required
def budget():
    now = datetime.now(timezone.utc)
    year = int(request.args.get("year", now.year))
    branch_id = int(request.args.get("branch_id", 0))

    all_branches = Branch.query.filter_by(is_active=True).order_by(Branch.name).all()
    categories   = Category.query.filter_by(is_active=True).order_by(Category.cost_type, Category.name).all()

    # Auto-select the only branch if none chosen
    if branch_id == 0 and len(all_branches) == 1:
        branch_id = all_branches[0].id

    # Fetch saved budgets → map {(category_id, month): float}
    bq = Budget.query.filter_by(year=year)
    if branch_id:
        bq = bq.filter_by(branch_id=branch_id)
    budget_map = {(b.category_id, b.month): float(b.monthly_amount) for b in bq.all()}

    # Fetch actual approved spend → map {(category_id, month): float}
    # Leaf items only — exclude parent roll-up rows to avoid double-counting
    _parent_ids_sq = db.session.query(PaymentRequestItem.parent_id).filter(
        PaymentRequestItem.parent_id.isnot(None)
    ).subquery()

    aq = (
        db.session.query(
            PaymentRequestItem.category_id,
            extract("month", PaymentRequest.created_at).label("m"),
            func.sum(PaymentRequestItem.amount).label("total"),
        )
        .join(PaymentRequest, PaymentRequest.id == PaymentRequestItem.request_id)
        .filter(
            PaymentRequest.status == "approved",
            extract("year", PaymentRequest.created_at) == year,
            ~PaymentRequestItem.id.in_(_parent_ids_sq),
        )
    )
    if branch_id:
        aq = aq.filter(PaymentRequest.branch_id == branch_id)
    actuals_map = {
        (row.category_id, int(row.m)): float(row.total or 0)
        for row in aq.group_by(PaymentRequestItem.category_id, "m").all()
    }

    months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

    return render_template(
        "budget.html",
        year=year,
        branch_id=branch_id,
        all_branches=all_branches,
        categories=categories,
        budget_map=budget_map,
        actuals_map=actuals_map,
        months=months,
        now=now,
        format_naira=format_naira,
        weeks_per_month=WEEKS_PER_MONTH,
    )


# ── AJAX: save one cell ──────────────────────────────────────────────────────

@budget_bp.route("/budget/set", methods=["POST"])
@login_required
@_mds_required
def set_budget():
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"error": "no data"}), 400

    try:
        branch_id   = int(data["branch_id"])
        category_id = int(data["category_id"])
        year        = int(data["year"])
        month       = int(data["month"])
        amount      = max(0.0, float(data.get("amount") or 0))
    except (KeyError, ValueError) as e:
        return jsonify({"error": str(e)}), 400

    row = Budget.query.filter_by(
        branch_id=branch_id, category_id=category_id, year=year, month=month
    ).first()

    if row:
        row.monthly_amount = Decimal(str(amount))
    else:
        db.session.add(Budget(
            branch_id=branch_id, category_id=category_id,
            year=year, month=month, monthly_amount=Decimal(str(amount)),
        ))
    db.session.commit()

    # Recalculate yearly total for this category+branch
    year_total = float(
        db.session.query(func.sum(Budget.monthly_amount))
        .filter_by(branch_id=branch_id, category_id=category_id, year=year)
        .scalar() or 0
    )

    return jsonify({
        "ok": True,
        "amount": amount,
        "weekly": round(amount / WEEKS_PER_MONTH, 2),
        "year_total": year_total,
    })
