from datetime import datetime, timezone
from decimal import Decimal
from functools import wraps

from flask import Blueprint, render_template, redirect, url_for, request, flash
from flask_login import login_required, current_user
from sqlalchemy.orm import subqueryload, joinedload

from extensions import db
from models import PaymentRequest, PaymentRequestItem
from utils import format_naira

approvals_bp = Blueprint("approvals", __name__)


def _mds_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not current_user.is_authenticated or not current_user.is_mds:
            flash("Access restricted to MD.", "danger")
            return redirect(url_for("requests.dashboard"))
        return f(*args, **kwargs)
    return decorated


@approvals_bp.route("/requests/<int:req_id>/review", methods=["GET", "POST"])
@login_required
@_mds_required
def review(req_id):
    pr = (
        PaymentRequest.query
        .options(
            subqueryload(PaymentRequest.items).joinedload(PaymentRequestItem.category),
            joinedload(PaymentRequest.branch),
            joinedload(PaymentRequest.submitter),
        )
        .filter_by(id=req_id)
        .first_or_404()
    )

    if request.method == "POST":
        try:
            action = request.form.get("action")
            comment = request.form.get("comment", "").strip()
            approved_amount_str = request.form.get("approved_amount", "").strip()

            if action not in ("approve", "reject"):
                flash("Invalid action.", "danger")
                return redirect(url_for("approvals.review", req_id=req_id))

            pr.mds_comment = comment or None
            pr.reviewed_at = datetime.now(timezone.utc)

            if action == "approve":
                try:
                    approved_amount = Decimal(approved_amount_str) if approved_amount_str else pr.requested_amount
                except Exception:
                    approved_amount = pr.requested_amount
                pr.approved_amount = approved_amount
                pr.status = "approved"
                flash(f"{pr.reference} approved for ₦{float(approved_amount):,.2f}.", "success")
            else:
                pr.status = "rejected"
                flash(f"{pr.reference} has been rejected.", "warning")

            db.session.commit()
            return redirect(url_for("requests.dashboard"))

        except Exception as e:
            try:
                db.session.rollback()
            except Exception:
                pass
            flash(f"An error occurred: {str(e)}", "danger")
            return redirect(url_for("approvals.review", req_id=req_id))

    try:
        return render_template("review_request.html", pr=pr, format_naira=format_naira)
    except Exception as tpl_err:
        flash(f"Could not load review page: {tpl_err}", "danger")
        return redirect(url_for("requests.dashboard"))
