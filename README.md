# mcss-25-26-ELA_Results

Group totals for McIntosh County Schools' 2025-26 Georgia Milestones ELA results, with 2024-25 comparisons.

**No student-level data belongs here.** Names, student IDs, and birth dates stay out. The `raw/` folder is ignored by git for that reason.

- `workbook/` - one Excel workbook of totals for every school and grade: achievement levels, year over year, matched students, domains, writing, and MCA Lexile.
- `dashboard/` - the source for the "MCSS ELA Milestones 2025-26" Claude dashboard: the page (`index.html`), its script (`app.js`), its colors and type (`look.json`), and the data files it reads (`data/`). It runs inside Claude's dashboard, not as a standalone website.
- `analysis/` - the scripts that calculate every number. With the roster and longitudinal files placed in `raw/`, `analysis/run_all.sh` rebuilds the dashboard data and the workbook.

2025-26 was the first year of the new ELA test. Achievement levels and scale scores are not compared across years; Lexile and growth percentile are.
