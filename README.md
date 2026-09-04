# Training Management System (TMS)

A production-ready Django system for managing Sectors, Trades, Levels,
Trainers, Modules, Learning Outcomes, Indicative Contents, Lesson Plans and
Logos — opening with a public landing dashboard that presents **Login**,
**Scheme of Work** and **Lesson Plans** as the three main entry points.

## Windows Setup

```bat
:: 1. Create and activate a virtual environment
python -m venv venv
venv\Scripts\activate

:: 2. Install dependencies
pip install -r requirements.txt

:: 3. Configure environment
copy .env.example .env
:: (edit .env if you want PostgreSQL instead of the default SQLite)

:: 4. Run migrations
python manage.py migrate

:: 5. Create an admin account
python manage.py createsuperuser

:: 6. Run the dev server
python manage.py runserver
```

Then open http://127.0.0.1:8000/ — this is the new public landing page with
three tiles (Login, Scheme of Work, Lesson Plans). Signing in from there
takes you into the full dashboard.

## Switching to PostgreSQL

In `.env`, set:

```
DB_ENGINE=postgres
DB_NAME=training_management
DB_USER=postgres
DB_PASSWORD=your_password
DB_HOST=localhost
DB_PORT=5432
```

Make sure the database exists first (`createdb training_management` or via pgAdmin),
then run `python manage.py migrate` again.

## Deploying to Vercel

This project uses Vercel's zero-configuration Django support — Vercel detects
`manage.py`, reads `WSGI_APPLICATION` from settings, installs
`requirements.txt`, and runs `collectstatic` automatically (static files are
served from `STATIC_ROOT` via the Vercel CDN).

1. Push this repo to GitHub (see `.gitignore` — `.env`, `db.sqlite3`, and
   `venv/` are never committed).
2. In the Vercel dashboard, **Add New → Project** and import the GitHub repo.
3. Add these Environment Variables in the Vercel project settings
   (Settings → Environment Variables), for both Production and Preview:
   - `SECRET_KEY` — a long random string
   - `DEBUG` — `False`
   - `DATABASE_URL` — a Postgres connection string (e.g. from Neon, Supabase,
     or the Vercel Postgres integration). Leave unset to fall back to SQLite,
     which is fine for a quick demo but **not** for real data on Vercel,
     since the filesystem is ephemeral between deployments.
   - `GROQ_API_KEY` — your Groq API key
   - `GROQ_MODEL` — optional, defaults to `openai/gpt-oss-120b`
4. Deploy. After the first deploy, run migrations against the same database
   from your local machine (with `DATABASE_URL` pointed at the same Postgres
   instance):
   ```bash
   python manage.py migrate
   python manage.py createsuperuser
   ```

`ALLOWED_HOSTS` and `CSRF_TRUSTED_ORIGINS` already default to allow any
`*.vercel.app` domain (including preview deployments), so no extra
configuration is needed there unless you attach a custom domain — in that
case add it via the `ALLOWED_HOSTS` and `CSRF_TRUSTED_ORIGINS` environment
variables (comma-separated).

## Project structure

```
tms/
├── core/                  # models, views, forms, urls, admin
│   └── migrations/
├── training_management/   # project settings, root urls
├── templates/
│   ├── base.html           # sidebar shell shared by all authenticated pages
│   ├── registration/login.html
│   └── core/
│       ├── landing.html        # public landing page (Login / Scheme of Work / Lesson Plans)
│       ├── dashboard.html
│       ├── crud_list.html      # generic list (used by every model)
│       ├── crud_form.html      # generic create/edit form
│       └── crud_delete.html    # generic delete confirmation
├── static/
│   ├── css/style.css       # dark navy/gold/teal design system
│   └── js/main.js          # sidebar drawer, alerts, search, validation, etc.
├── requirements.txt
├── .env.example
└── manage.py
```

## Landing page & routing

- `/` — public landing page (`core.LandingView`), the first screen every
  visitor sees. It shows three large tiles: **Login**, **Scheme of Work**
  and **Lesson Plans**. No authentication is required to view it.
- `/login/` — the existing styled login form, reached from the Login tile
  (or automatically when an unauthenticated visitor clicks Scheme of Work
  or Lesson Plans — they're sent to login with `?next=` set, then bounced
  straight to the module they wanted).
- `/dashboard/` — the authenticated Scheme of Work dashboard (sectors,
  trades, levels, modules, trainers, learning outcomes, indicative contents).
- `/lesson-plans/` — the new Lesson Plans module (list/create/edit/delete),
  login-required, following the same generic CRUD pattern as every other
  model.

## Generic CRUD pattern

Every model (Sector, Trade, Level, TradeLevel, Trainer, Module,
LearningOutcome, IndicativeContent, LessonPlan, Logo) is wired through four shared base
views in `core/views.py` (`BaseListView`, `BaseCreateView`, `BaseUpdateView`,
`BaseDeleteView`) and three shared templates
(`crud_list.html`, `crud_form.html`, `crud_delete.html`).
To add a new field to a model, just add it to the model, the corresponding
`ModelForm` in `core/forms.py`, and the `headers`/`columns` list on that
model's List view — no new templates needed.

## Front-end interactivity (static/js/main.js)

- **Mobile sidebar drawer** — hamburger button + overlay, slides the sidebar
  in/out on screens under 900px, auto-closes when a nav link is tapped.
- **Auto-dismissing alerts** — Django messages and form banners fade out
  after 5 seconds (skip this by adding the `alert-persistent` class).
- **Password visibility toggle** — reusable on any `.password-field` wrapper.
- **Quick search** — an `#tableSearch` input filters any CRUD list table
  client-side, with an empty-state message when nothing matches.
- **Loading button state** — forms with `.js-loading-submit` disable their
  submit button and show a spinner to prevent double submits.
- **Client-side validation** — forms with `.needs-validation` get Bootstrap's
  validity-check-and-highlight behavior before submitting.
- **`TMS.toast(message, variant)`** — a small helper for one-off JS-triggered
  notifications, exposed on `window.TMS`.

## Django admin

The built-in Django admin is also available at `/django-admin/` for quick
data inspection, registered in `core/admin.py`.

## Authentication

Phase 1 uses Django's built-in `User`/session auth for administrators only
(`django.contrib.auth`). The `Trainer` model's `password_hash` field is
already wired up (via `TrainerForm`, using `django.contrib.auth.hashers`) so
trainer self-service login can be added later without changing the schema.
# TRAINING-MANAGEMENT-SYSTEM
