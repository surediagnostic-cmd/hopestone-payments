from datetime import datetime, timezone
from flask_login import UserMixin
from extensions import db

user_branches = db.Table(
    "hop_user_branches",
    db.Column("user_id", db.Integer, db.ForeignKey("hop_users.id", ondelete="CASCADE"), primary_key=True),
    db.Column("branch_id", db.Integer, db.ForeignKey("hop_branches.id", ondelete="CASCADE"), primary_key=True),
)


class User(UserMixin, db.Model):
    __tablename__ = "hop_users"

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), nullable=False, default="accountant")  # accountant | mds
    is_active = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    branches = db.relationship("Branch", secondary=user_branches, backref="users", lazy="subquery")
    requests = db.relationship("PaymentRequest", backref="submitter", lazy=True)

    @property
    def is_mds(self):
        return self.role == "mds"

    @property
    def branch(self):
        return self.branches[0] if self.branches else None


class Branch(db.Model):
    __tablename__ = "hop_branches"

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), unique=True, nullable=False)
    source_account = db.Column(db.String(120), nullable=False)
    is_active = db.Column(db.Boolean, default=True)

    requests = db.relationship("PaymentRequest", backref="branch", lazy=True)


class Category(db.Model):
    __tablename__ = "hop_categories"

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), unique=True, nullable=False)
    cost_type = db.Column(db.String(20), nullable=False, default="overhead")  # direct_cost | overhead
    is_active = db.Column(db.Boolean, default=True)


class PaymentRequest(db.Model):
    __tablename__ = "hop_payment_requests"

    id = db.Column(db.Integer, primary_key=True)
    reference = db.Column(db.String(30), unique=True, nullable=False)
    date = db.Column(db.Date, nullable=False)
    branch_id = db.Column(db.Integer, db.ForeignKey("hop_branches.id"), nullable=False)
    beneficiary_name = db.Column(db.String(200), nullable=False)
    beneficiary_account = db.Column(db.String(30), nullable=False)
    beneficiary_bank = db.Column(db.String(100), nullable=False)
    bank_code = db.Column(db.String(10), nullable=True)
    requested_amount = db.Column(db.Numeric(14, 2), nullable=False)
    approved_amount = db.Column(db.Numeric(14, 2), nullable=True)
    mds_comment = db.Column(db.Text, nullable=True)
    status = db.Column(db.String(20), nullable=False, default="pending")  # pending | approved | rejected
    upload_status = db.Column(db.String(20), nullable=False, default="not_uploaded")  # not_uploaded | uploaded
    receipt_filename = db.Column(db.String(255), nullable=True)
    submitted_by_id = db.Column(db.Integer, db.ForeignKey("hop_users.id"), nullable=False)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    reviewed_at = db.Column(db.DateTime, nullable=True)

    items = db.relationship(
        "PaymentRequestItem",
        foreign_keys="[PaymentRequestItem.request_id]",
        backref="request",
        lazy="subquery",
        cascade="all, delete-orphan",
    )

    @staticmethod
    def generate_reference():
        now = datetime.now(timezone.utc)
        prefix = f"PAY-{now.strftime('%Y%m')}"
        existing = db.session.query(PaymentRequest.reference)\
            .filter(PaymentRequest.reference.like(f"{prefix}-%")).all()
        max_num = 0
        for (ref,) in existing:
            try:
                max_num = max(max_num, int(ref.rsplit("-", 1)[-1]))
            except (ValueError, IndexError):
                pass
        return f"{prefix}-{str(max_num + 1).zfill(3)}"

    @property
    def status_color(self):
        return {"pending": "warning", "approved": "success", "rejected": "danger"}.get(self.status, "secondary")

    @property
    def status_label(self):
        return {"pending": "Pending", "approved": "Approved", "rejected": "Rejected"}.get(self.status, self.status.title())


class Budget(db.Model):
    __tablename__ = "hop_budgets"

    id          = db.Column(db.Integer, primary_key=True)
    branch_id   = db.Column(db.Integer, db.ForeignKey("hop_branches.id"),   nullable=False)
    category_id = db.Column(db.Integer, db.ForeignKey("hop_categories.id"), nullable=False)
    year        = db.Column(db.Integer, nullable=False)
    month       = db.Column(db.Integer, nullable=False)  # 1-12
    monthly_amount = db.Column(db.Numeric(14, 2), nullable=False, default=0)

    __table_args__ = (
        db.UniqueConstraint("branch_id", "category_id", "year", "month", name="uq_hop_budget"),
    )

    branch   = db.relationship("Branch",   lazy="joined")
    category = db.relationship("Category", lazy="joined")

    @property
    def weekly_amount(self):
        return round(float(self.monthly_amount) / 4.33, 2)


class PaymentRequestItem(db.Model):
    __tablename__ = "hop_payment_request_items"

    id = db.Column(db.Integer, primary_key=True)
    request_id = db.Column(
        db.Integer,
        db.ForeignKey("hop_payment_requests.id", ondelete="CASCADE"),
        nullable=False,
    )
    parent_id = db.Column(
        db.Integer,
        db.ForeignKey("hop_payment_request_items.id"),
        nullable=True,
    )
    description = db.Column(db.String(500), nullable=False)
    note        = db.Column(db.String(1000), nullable=True)
    category_id = db.Column(db.Integer, db.ForeignKey("hop_categories.id"), nullable=False)
    quantity = db.Column(db.Integer, nullable=False, default=1)
    rate = db.Column(db.Numeric(14, 2), nullable=False)
    amount = db.Column(db.Numeric(14, 2), nullable=False)

    category = db.relationship("Category", lazy="joined")
    # Self-referential: one parent item → many sub-items
    children = db.relationship(
        "PaymentRequestItem",
        foreign_keys="[PaymentRequestItem.parent_id]",
        backref=db.backref("parent_item", remote_side="[PaymentRequestItem.id]"),
        lazy="select",
        cascade="all, delete-orphan",
    )
