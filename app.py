import os
from flask import Flask, render_template
from sqlalchemy.pool import NullPool
from extensions import db, login_manager


def create_app():
    app = Flask(__name__)

    app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "dev-secret-hopestone-change-in-prod")
    app.config["SQLALCHEMY_DATABASE_URI"] = os.environ.get("DATABASE_URL", "sqlite:///hopestone.db")
    app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

    if "postgresql" in app.config["SQLALCHEMY_DATABASE_URI"]:
        app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {
            "poolclass": NullPool,
            "connect_args": {
                "sslmode": "require",
                "options": "-c statement_timeout=30000",
            },
            "execution_options": {"prepared_statement_cache_size": 0},
        }

    db.init_app(app)
    login_manager.init_app(app)
    login_manager.login_view = "auth.login"
    login_manager.login_message = "Please log in to access this page."
    login_manager.login_message_category = "info"

    from models import User

    @login_manager.user_loader
    def load_user(user_id):
        return User.query.get(int(user_id))

    from routes.auth import auth_bp
    from routes.requests import requests_bp
    from routes.approvals import approvals_bp
    from routes.reports import reports_bp
    from routes.admin import admin_bp
    from routes.budget import budget_bp

    app.register_blueprint(auth_bp)
    app.register_blueprint(requests_bp)
    app.register_blueprint(approvals_bp)
    app.register_blueprint(reports_bp)
    app.register_blueprint(admin_bp)
    app.register_blueprint(budget_bp)

    @app.errorhandler(500)
    def internal_error(error):
        import traceback
        tb = traceback.format_exc()
        print(f"[500 ERROR]\n{tb}", flush=True)
        try:
            db.session.rollback()
        except Exception:
            pass
        try:
            return render_template("500.html", error=str(error), traceback=tb), 500
        except Exception:
            return (
                f"<pre style='padding:20px;background:#1a3a5c;color:#3a9fd8;"
                f"font-family:monospace;font-size:13px;'>"
                f"500 Internal Server Error\n\n{tb}</pre>"
            ), 500

    with app.app_context():
        try:
            _run_migrations()
        except Exception as e:
            print(f"[startup] migration warning: {e}", flush=True)
        try:
            db.create_all()
        except Exception as e:
            print(f"[startup] db.create_all warning: {e}", flush=True)
        try:
            _seed_defaults()
        except Exception as e:
            print(f"[startup] seed warning: {e}", flush=True)

    return app


def _run_migrations():
    from sqlalchemy import text
    db_url = os.environ.get("DATABASE_URL", "")
    if "postgresql" not in db_url:
        return

    with db.engine.connect() as conn:
        result = conn.execute(text(
            "SELECT column_name FROM information_schema.columns "
            "WHERE table_name='hop_categories' AND column_name='cost_type'"
        ))
        if not result.fetchone():
            conn.execute(text(
                "ALTER TABLE hop_categories ADD COLUMN cost_type VARCHAR(50) DEFAULT 'overhead'"
            ))
            conn.commit()
            print("[migration] Added cost_type to hop_categories", flush=True)

        result = conn.execute(text(
            "SELECT column_name FROM information_schema.columns "
            "WHERE table_name='hop_payment_requests' AND column_name='receipt_filename'"
        ))
        if not result.fetchone():
            try:
                conn.execute(text(
                    "ALTER TABLE hop_payment_requests ADD COLUMN receipt_filename VARCHAR(255)"
                ))
                conn.commit()
                print("[migration] Added receipt_filename to hop_payment_requests", flush=True)
            except Exception:
                pass

        result = conn.execute(text(
            "SELECT column_name FROM information_schema.columns "
            "WHERE table_name='hop_payment_request_items' AND column_name='parent_id'"
        ))
        if not result.fetchone():
            try:
                conn.execute(text(
                    "ALTER TABLE hop_payment_request_items ADD COLUMN parent_id INTEGER "
                    "REFERENCES hop_payment_request_items(id)"
                ))
                conn.commit()
                print("[migration] Added parent_id to hop_payment_request_items", flush=True)
            except Exception:
                pass


def _seed_defaults():
    from models import Branch, Category, User
    import bcrypt

    if Branch.query.count() == 0:
        db.session.add(Branch(
            name="Hopestone Hospital (Ikeja)",
            source_account="Sterling Bank",
            is_active=True,
        ))
        db.session.commit()
        print("[seed] Branch seeded", flush=True)

    if Category.query.count() == 0:
        categories = [
            Category(name="Drugs", cost_type="direct_cost", is_active=True),
            Category(name="Hospital Consumables", cost_type="direct_cost", is_active=True),
            Category(name="Specialist Payments", cost_type="direct_cost", is_active=True),
            Category(name="Locum", cost_type="direct_cost", is_active=True),
            Category(name="Outsourced Tests", cost_type="direct_cost", is_active=True),
            Category(name="Staff Salaries", cost_type="overhead", is_active=True),
            Category(name="Repairs & Maintenance", cost_type="overhead", is_active=True),
            Category(name="Utilities", cost_type="overhead", is_active=True),
            Category(name="Administrative", cost_type="overhead", is_active=True),
            Category(name="Others", cost_type="overhead", is_active=True),
        ]
        db.session.add_all(categories)
        db.session.commit()
        print("[seed] Categories seeded", flush=True)

    if User.query.count() == 0:
        pw_hash = bcrypt.hashpw("Admin@1234".encode(), bcrypt.gensalt()).decode()
        admin = User(
            name="MD Admin",
            email="admin@hopestone.ng",
            password_hash=pw_hash,
            role="mds",
            is_active=True,
        )
        db.session.add(admin)
        db.session.flush()
        branch = Branch.query.first()
        if branch:
            admin.branches.append(branch)
        db.session.commit()
        print("[seed] Admin user seeded: admin@hopestone.ng / Admin@1234", flush=True)


app = create_app()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5001))
    app.run(debug=True, port=port)
