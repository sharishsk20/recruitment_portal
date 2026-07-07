# TÜV SÜD — Recruitment Portal

A Python (Flask) + HTML web portal to enter candidates, filter them, and export to Excel.
Styled in TÜV SÜD's colours (deep blue + grey + white).

## What it does
- Add candidates through a web form (name, role, experience, certifications, ISO status, source, stage, status, notes)
- Live filters: search by name/certification, filter by role, ISO certified, minimum experience, source, and status
- Data stored in a real database (SQLite — `candidates.db`, created automatically)
- **Export to Excel**: downloads whatever is currently filtered as a formatted `.xlsx`

## How to run

1. Make sure you have Python 3.8+ installed.
2. In this folder, install dependencies:
   ```
   pip install -r requirements.txt
   ```
3. Start the server:
   ```
   python app.py
   ```
4. Open your browser to:
   ```
   http://127.0.0.1:5000
   ```

## How to use
- Click **Load Sample Data** to see it working with 4 example rows.
- Use the filters, then click **⬇ Export to Excel** — only the filtered rows are exported.
- Click **Clear All** to empty the database before entering real candidates.

## Project structure
```
recruitment_portal/
├── app.py               # Flask backend + SQLite + Excel export
├── requirements.txt     # Python dependencies
├── templates/
│   └── index.html       # Frontend (TÜV SÜD styling)
└── candidates.db        # Auto-created on first run
```

## Notes
- This runs locally on your machine. For multiple recruiters accessing one shared
  database over a network, it would need to be deployed to a server (e.g. Render,
  Railway, or an internal host) — the code is structured to support that.
- Sample rows are labelled `[Sample]` on purpose. Clear them before using real data.
- TÜV SÜD's exact brand Pantone is not published publicly; the deep-blue + grey +
  white palette here matches their public brand identity (blue octagon, grey corporate area).
