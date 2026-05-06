import bcrypt
from functools import wraps

from flask import Blueprint, render_template, redirect, url_for, request, flash
from flask_login import login_required, current_user

from extensions import db
from models import Branch, Category, User

admin_bp = Blueprint("admin", __name__)


def _mds_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not current_user.is_authenticated or not current_user.is_mds:
            flash("Access restricted to MD.", "danger")
            return redirect(url_for("requests.dashboard"))
        return f(*args, **kwargs)
    return decorated


@admin_bp.route("/admin")
@login_required
@_mds_required
def admin():
    branches = Branch.query.order_by(Branch.name).all()
    categories = Category.query.order_by(Category.cost_type, Category.name).all()
    users = User.query.order_by(User.name).all()
    all_branches = branches
    return render_template("admin.html", branches=branches, categories=categories, users=users, all_branches=all_branches)


# ── Branches ────────────────────────────────────────────────────────────────

@admin_bp.route("/admin/branches/add", methods=["POST"])
@login_required
@_mds_required
def add_branch():
    name = request.form.get("name", "").strip()
    source_account = request.form.get("source_account", "").strip()
    if not name or not source_account:
        flash("Branch name and source account are required.", "danger")
        return redirect(url_for("admin.admin"))
    if Branch.query.filter_by(name=name).first():
        flash(f'Branch "{name}" already exists.', "warning")
        return redirect(url_for("admin.admin"))
    db.session.add(Branch(name=name, source_account=source_account, is_active=True))
    db.session.commit()
    flash(f'Branch "{name}" added successfully.', "success")
    return redirect(url_for("admin.admin"))


@admin_bp.route("/admin/branches/<int:branch_id>/toggle", methods=["POST"])
@login_required
@_mds_required
def toggle_branch(branch_id):
    branch = Branch.query.get_or_404(branch_id)
    branch.is_active = not branch.is_active
    db.session.commit()
    status = "activated" if branch.is_active else "deactivated"
    flash(f'Branch "{branch.name}" {status}.', "success")
    return redirect(url_for("admin.admin"))


# ── Categories ───────────────────────────────────────────────────────────────

@admin_bp.route("/admin/categories/add", methods=["POST"])
@login_required
@_mds_required
def add_category():
    name = request.form.get("name", "").strip()
    cost_type = request.form.get("cost_type", "overhead")
    if not name:
        flash("Category name is required.", "danger")
        return redirect(url_for("admin.admin"))
    if cost_type not in ("direct_cost", "overhead"):
        cost_type = "overhead"
    if Category.query.filter_by(name=name).first():
        flash(f'Category "{name}" already exists.', "warning")
        return redirect(url_for("admin.admin"))
    db.session.add(Category(name=name, cost_type=cost_type, is_active=True))
    db.session.commit()
    flash(f'Category "{name}" added successfully.', "success")
    return redirect(url_for("admin.admin"))


@admin_bp.route("/admin/categories/<int:cat_id>/edit", methods=["POST"])
@login_required
@_mds_required
def edit_category(cat_id):
    cat = Category.query.get_or_404(cat_id)
    name = request.form.get("name", "").strip()
    cost_type = request.form.get("cost_type", "overhead")
    if not name:
        flash("Category name is required.", "danger")
        return redirect(url_for("admin.admin"))
    existing = Category.query.filter_by(name=name).first()
    if existing and existing.id != cat_id:
        flash(f'Category "{name}" already exists.', "warning")
        return redirect(url_for("admin.admin"))
    cat.name = name
    cat.cost_type = cost_type
    db.session.commit()
    flash("Category updated.", "success")
    return redirect(url_for("admin.admin"))


@admin_bp.route("/admin/categories/<int:cat_id>/toggle", methods=["POST"])
@login_required
@_mds_required
def toggle_category(cat_id):
    cat = Category.query.get_or_404(cat_id)
    cat.is_active = not cat.is_active
    db.session.commit()
    status = "activated" if cat.is_active else "deactivated"
    flash(f'Category "{cat.name}" {status}.', "success")
    return redirect(url_for("admin.admin"))


# ── Users ────────────────────────────────────────────────────────────────────

@admin_bp.route("/admin/users/add", methods=["POST"])
@login_required
@_mds_required
def add_user():
    name = request.form.get("name", "").strip()
    email = request.form.get("email", "").strip().lower()
    password = request.form.get("password", "").strip()
    role = request.form.get("role", "accountant")
    branch_ids = request.form.getlist("branch_ids")

    if not name or not email or not password:
        flash("Name, email and password are required.", "danger")
        return redirect(url_for("admin.admin"))
    if len(password) < 6:
        flash("Password must be at least 6 characters.", "danger")
        return redirect(url_for("admin.admin"))
    if User.query.filter_by(email=email).first():
        flash(f'A user with email "{email}" already exists.', "warning")
        return redirect(url_for("admin.admin"))
    if role not in ("accountant", "mds"):
        role = "accountant"

    pw_hash = bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()
    user = User(name=name, email=email, password_hash=pw_hash, role=role, is_active=True)

    for bid in branch_ids:
        branch = Branch.query.get(int(bid))
        if branch:
            user.branches.append(branch)

    db.session.add(user)
    db.session.commit()
    flash(f'User "{name}" created successfully.', "success")
    return redirect(url_for("admin.admin"))


@admin_bp.route("/admin/users/<int:user_id>/toggle", methods=["POST"])
@login_required
@_mds_required
def toggle_user(user_id):
    user = User.query.get_or_404(user_id)
    if user.id == current_user.id:
        flash("You cannot deactivate your own account.", "warning")
        return redirect(url_for("admin.admin"))
    user.is_active = not user.is_active
    db.session.commit()
    status = "activated" if user.is_active else "deactivated"
    flash(f'User "{user.name}" {status}.', "success")
    return redirect(url_for("admin.admin"))


@admin_bp.route("/admin/users/<int:user_id>/reset-password", methods=["POST"])
@login_required
@_mds_required
def reset_password(user_id):
    user = User.query.get_or_404(user_id)
    new_password = request.form.get("new_password", "").strip()
    if not new_password or len(new_password) < 6:
        flash("Password must be at least 6 characters.", "danger")
        return redirect(url_for("admin.admin"))
    user.password_hash = bcrypt.hashpw(new_password.encode(), bcrypt.gensalt()).decode()
    db.session.commit()
    flash(f'Password for "{user.name}" has been updated.', "success")
    return redirect(url_for("admin.admin"))
