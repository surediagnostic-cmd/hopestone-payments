# Hopestone Payment Tracker

A web app for managing and approving bank payment requests at Hopestone Hospital. Accountants submit requests; the MD reviews, approves or rejects them; approved payments are tracked through bank upload.

---

## Tech Stack

| Layer | Choice |
|-------|--------|
| Backend | Python 3.11 + Flask 3.0 |
| Database | Supabase (PostgreSQL) via SQLAlchemy |
| Auth | Flask-Login + bcrypt |
| Export/Import | openpyxl |
| Charts | Chart.js (CDN) |
| Hosting | Railway |

---

## Default Login (first deploy)

| Field | Value |
|-------|-------|
| Email | `admin@hopestone.ng` |
| Password | `Admin@1234` |

**Change this password immediately after first login** via Settings → Users → Reset Password.

---

## Deployment — Railway + Supabase

### Step 1 — Push to GitHub

```bash
git init
git add .
git commit -m "Initial commit — Hopestone Payment Tracker"
# Create a new repo on GitHub, then:
git remote add origin https://github.com/YOUR_ORG/hopestone-payment-tracker.git
git push -u origin main
```

### Step 2 — Create Railway project

1. Go to [railway.app](https://railway.app) → **New Project** → **Deploy from GitHub repo**
2. Select the new repo
3. Railway will detect the `Procfile` and deploy automatically

### Step 3 — Set environment variables in Railway

Go to your Railway project → **Variables** tab → add these:

| Variable | Value |
|----------|-------|
| `DATABASE_URL` | Session Pooler URL from Supabase (**see below**) |
| `SECRET_KEY` | Any random 32+ character string |

> **Email notifications are disabled** in this version. No mail variables needed.

### Step 4 — Get the correct Supabase DATABASE_URL

⚠️ Railway is **IPv4-only**. Supabase's direct connection URL uses IPv6 and will fail with "Network is unreachable". You **must** use the **Session Pooler** URL.

**How to get it:**
1. Supabase Dashboard → **Settings** → **Database** → **Connect**
2. Select **Session pooler** + **URI** format
3. Copy the URL — it looks like:
   ```
   postgresql://postgres.PROJECTREF:PASSWORD@REGION.pooler.supabase.com:5432/postgres
   ```
4. Paste it as `DATABASE_URL` in Railway

> **Note:** This app uses table names prefixed with `hop_` (e.g. `hop_users`, `hop_payment_requests`) so it is safe to share the same Supabase project as Sure Diagnostics — there are no table name conflicts.

### Step 5 — Redeploy

After setting variables, Railway auto-redeploys. Check the **Logs** tab — you should see:

```
[seed] Branch seeded
[seed] Categories seeded
[seed] Admin user seeded: admin@hopestone.ng / Admin@1234
```

---

## First-Time Setup Checklist

- [ ] Log in as `admin@hopestone.ng` / `Admin@1234`
- [ ] Go to **Settings** → reset the admin password
- [ ] Create accountant users and assign them to branches
- [ ] Verify branches and categories are correct (Settings page)
- [ ] Have an accountant submit a test request and approve it as MD
- [ ] Test the Excel export downloads correctly
- [ ] Redeploy with a trivial commit — confirm data persists (Supabase working)

---

## Database Tables (all prefixed `hop_`)

| Table | Purpose |
|-------|---------|
| `hop_users` | Staff accounts (accountants + MD) |
| `hop_branches` | Hospital branches |
| `hop_categories` | Payment categories (direct cost / overhead) |
| `hop_user_branches` | Many-to-many: users ↔ branches |
| `hop_payment_requests` | Payment requests |
| `hop_payment_request_items` | Line items per request |

---

## Local Development

```bash
python -m venv venv
source venv/bin/activate      # Windows: venv\Scripts\activate
pip install -r requirements.txt
python app.py
# Visit http://localhost:5000
```

SQLite is used automatically for local dev (no DATABASE_URL needed).

---

## Roles

| Role | Can Do |
|------|--------|
| **Accountant** | Submit requests, bulk upload, mark approved requests as uploaded, view own requests |
| **MD / Admin** | All accountant actions + review/approve/reject all requests, manage users/branches/categories, view reports |

---

## Adding Email Notifications Later

To add email notifications, install Flask-Mail and add to `requirements.txt`:
```
Flask-Mail==0.10.0
```

Then add to Railway variables:
```
MAIL_SERVER=smtp.gmail.com
MAIL_PORT=587
MAIL_USERNAME=your@gmail.com
MAIL_PASSWORD=your-app-password
MAIL_DEFAULT_SENDER=Hopestone <your@gmail.com>
MDS_EMAIL=md@hopestone.ng
```

Email sending must always run in a **background thread** — never block the HTTP response (gunicorn timeout is 30s).
