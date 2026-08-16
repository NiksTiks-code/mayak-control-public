<p align="center">
  <img src="static/images/mayak-control-logo.png" alt="MAYAK logo" width="220">
</p>

<h1 align="center">MAYAK — Construction Progress Control</h1>

<p align="center">
  A field-ready web application for tracking engineering work across construction objects, systems, rooms and contractors.
</p>

<p align="center">
  <strong>English</strong> · <a href="README_RU.md">Русский</a>
</p>

> This repository is a sanitized portfolio snapshot. It contains synthetic demo data and newly captured demo screenshots only. Production source history, configuration and customer data remain private.

## What MAYAK does

MAYAK gives customers and contractors one operational view of construction readiness. Work is tracked by contractor, object, engineering system and room, with role-aware actions and live completion summaries.

- **Construction Progress Matrix** — an interactive room-by-room chessboard with completed, review, denied and pending states.
- **Universal Excel Importer** — preserves worksheet geometry, merged cells, entrances, floors, apartments and non-residential rooms, with a preview-before-commit flow.
- **Role-based access** — separate Super Admin, customer/admin and contractor permissions.
- **Photo evidence** — UUID-based uploads linked to room and common-work records, with rollback if JSON persistence fails.
- **Object readiness** — live totals, completion percentage, status distribution and engineering-system context.
- **Rozliv tracking** — progress for single, upper and lower distribution layouts.
- **PDF reporting** — printable object/system status summaries.

## Screenshots

All screenshots below were captured from a local sanitized environment. Names, addresses, rooms and progress values are synthetic.

### Dashboard

![Sanitized MAYAK dashboard](docs/screenshots/dashboard.jpg)

### Construction Progress Matrix

![Sanitized construction progress matrix](docs/screenshots/construction-progress-matrix.jpg)

### Universal Excel Import

![Sanitized Excel import page](docs/screenshots/excel-import.jpg)

### Room Details

![Sanitized room details](docs/screenshots/room-details.jpg)

### Object Readiness

![Sanitized object readiness summary](docs/screenshots/object-readiness.jpg)

### Mobile Dashboard

<p align="center">
  <img src="docs/screenshots/mobile-dashboard.jpg" alt="Sanitized MAYAK mobile dashboard" width="390">
</p>

## Production Stability / Load Testing

During production load testing, a critical JSON read-modify-write race condition was identified across three Gunicorn workers. It was fixed with interprocess `fcntl.flock`, rereading the latest JSON after acquiring the lock, and an atomic temp-file write using `flush`, `fsync` and `os.replace()`.

Post-fix validation:

| Scenario | Result |
| --- | --- |
| 20 concurrent writes × 5 series | **100/100 persisted**, **0 lost updates** |
| 50 concurrent writes × 5 series | **250/250 persisted**, **0 lost updates** |
| 5 / 10 / 20 / 50 simultaneous photo uploads | All files and JSON references persisted; **0 orphan**, **0 corrupted** |
| 50 concurrent mixed operations | **50/50 successful**, **0 lost updates** |
| Read-only HTTPS, 50 concurrent users | **0 errors**; p95 **411 ms**, p99 **570 ms**, max **724 ms** |
| Writes, 50 concurrent operations | p95 **120.9 ms**, p99 **128.3 ms**, max **129.9 ms** |

**PASS — concurrent JSON writes fixed**

**Under the tested workload, MAYAK is safe for 50 concurrent users.** This does not claim unlimited scalability or validate workloads beyond the scenarios above.

Writes are serialized by a filesystem lock, and simultaneous updates to the same room use last-write-wins semantics. JSON remains the current scaling boundary. If data volume or write frequency grows substantially, the next architectural step should be a transactional database such as PostgreSQL.

## Security architecture

- Flask session secrets are supplied through the runtime environment; startup fails closed when the required secret is absent.
- Password authentication is hash-only through Werkzeug; there is no plaintext fallback or built-in credential.
- User mutations and business-data writes use interprocess locks, reread-after-lock and atomic replacement.
- Runtime JSON, uploads, environment files, reports, caches and archives are excluded from Git.
- Demo account passwords are entered locally through `getpass` and are never included in this repository.

## Tech stack

- Python, Flask, Werkzeug and Gunicorn
- Jinja2 templates, HTML, CSS and vanilla JavaScript
- OpenPyXL for structure-preserving Excel imports
- ReportLab for PDF reports
- JSON persistence with `fcntl`-based interprocess locking and atomic file replacement

## Local demo

The demo targets Linux or macOS because its current locking layer uses `fcntl`.

```bash
git clone https://github.com/NiksTiks-code/mayak-control-public.git
cd mayak-control-public

python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

python scripts/init_demo.py
export MAYAK_SECRET_KEY="$(python -c 'import secrets; print(secrets.token_urlsafe(48))')"
python app.py
```

Open `http://127.0.0.1:5001`. The initializer requests new passwords for `demo_admin` and `demo_contractor` using hidden interactive input. No default password is provided.

The repository also includes [`demo/mayak_demo_import.xlsx`](demo/mayak_demo_import.xlsx), a synthetic workbook demonstrating entrances, floors, apartments, non-residential rooms, merged cells and preserved geometry.

## Repository contents

```text
app.py                         Flask application and safe JSON write paths
excel_importer.py              Structure-preserving Excel importer
templates/                     Jinja2 user interface
static/                        CSS, JavaScript and MAYAK branding
demo/                          Synthetic JSON examples and demo workbook
docs/screenshots/              Sanitized screenshots from the local demo
scripts/init_demo.py           Interactive local demo initializer
.env.example                   Environment variable names only
```
