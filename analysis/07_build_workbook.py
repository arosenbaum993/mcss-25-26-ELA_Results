"""Step 7. Builds workbook/MCSS_ELA_Milestones_2025-26_Totals.xlsx.

Run from the repository root after steps 1-6. The workbook holds group totals only:
no student names, IDs, or birth dates. Percentages and changes are Excel formulas,
so they recalculate if a count is corrected.
"""
import json
import statistics as st
import warnings

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

warnings.simplefilter("ignore")
OUT = "workbook/MCSS_ELA_Milestones_2025-26_Totals.xlsx"
LINE = {3: 520, 4: 740, 5: 830, 6: 925, 7: 970, 8: 1010}
A = {int(k): v for k, v in json.load(open("analysis/out/agg.json")).items()}
rows_of = lambda name: json.load(open(f"dashboard/data/{name}.json"))["rows"]

FONT = Font(name="Arial", size=10)
BOLD = Font(name="Arial", size=10, bold=True)
TITLE = Font(name="Arial", size=13, bold=True)
HEAD = PatternFill("solid", start_color="F1E7C8", end_color="F1E7C8")
wb = openpyxl.Workbook()
wb.remove(wb.active)


def sheet(name, title, header, widths):
    ws = wb.create_sheet(name)
    ws["A1"] = title
    ws["A1"].font = TITLE
    for j, h in enumerate(header, 1):
        c = ws.cell(3, j, h)
        c.font, c.fill = BOLD, HEAD
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[3].height = 32
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "A4"
    return ws


def put(ws, r, values, fmts=None, bold=False):
    for j, v in enumerate(values, 1):
        c = ws.cell(r, j, v)
        c.font = BOLD if bold else FONT
        if fmts and fmts.get(j):
            c.number_format = fmts[j]


def note(ws, r, text):
    ws.cell(r, 1, text).font = Font(name="Arial", size=9, italic=True)


# ---------------- Achievement levels ----------------
ws = sheet("Achievement Levels", "ELA achievement levels, 2025-26",
           ["School", "Group", "Tested", "Beginning", "Developing", "Proficient", "Distinguished",
            "% Beginning", "% Developing", "% Proficient", "% Distinguished", "Proficient or above", "% Proficient or above"],
           [10, 22, 9, 11, 11, 11, 13, 12, 12, 12, 13, 13, 14])
PCT = {8: "0.0%", 9: "0.0%", 10: "0.0%", 11: "0.0%", 13: "0.0%"}
lv = {r["id"]: r for r in rows_of("levels")}
r = 4
first = {}
for school, ids in (("TGES", ["g3", "g4", "g5"]), ("MCMS", ["g6", "g7", "g8"]), ("MCA", ["lc2"])):
    first[school] = r
    for i in ids:
        x = lv[i]
        put(ws, r, [school, x["name"], f"=SUM(D{r}:G{r})", x["beginning"], x["developing"], x["proficient"], x["distinguished"],
                    f"=D{r}/$C{r}", f"=E{r}/$C{r}", f"=F{r}/$C{r}", f"=G{r}/$C{r}", f"=F{r}+G{r}", f"=L{r}/C{r}"], PCT)
        r += 1
    if len(ids) > 1:
        a, b = first[school], r - 1
        put(ws, r, [school, f"{school} total"] + [f"=SUM({c}{a}:{c}{b})" for c in "CDEFG"]
            + [f"=D{r}/$C{r}", f"=E{r}/$C{r}", f"=F{r}/$C{r}", f"=G{r}/$C{r}", f"=F{r}+G{r}", f"=L{r}/C{r}"], PCT, bold=True)
        r += 1
note(ws, r + 1, "Grades 3-8: Spring 2026 EOG rosters. Lit & Comp II: Winter 2025 and Spring 2026 EOC combined (89 students).")
note(ws, r + 2, "2025-26 was the first year of the new ELA test. Do not compare these levels with earlier years.")

# ---------------- Year over year ----------------
ws = sheet("Year over Year", "Grades 3-8: 2024-25 next to 2025-26",
           ["Grade", "Students 2024-25", "Students 2025-26", "Median Lexile 2024-25", "Median Lexile 2025-26", "Lexile change",
            "Grade-level Lexile line", "At or above line 2024-25", "At or above line 2025-26", "% at line 2024-25", "% at line 2025-26",
            "Change (points)", "Growth percentile 2024-25", "Growth percentile 2025-26", "Growth change"],
           [9, 11, 11, 13, 13, 10, 12, 13, 13, 11, 11, 10, 13, 13, 10])
for k, g in enumerate(range(3, 9)):
    r = 4 + k
    d, p = A[g], A[g]["prior"]
    sg = d["sgp"] is not None
    put(ws, r, [f"Grade {g}", p["n"], d["n"], p["med_lex"], d["med_lex"], f"=E{r}-D{r}", LINE[g], p["at_line"], d["at_line"],
                f"=H{r}/B{r}", f"=I{r}/C{r}", f"=(K{r}-J{r})*100", p["sgp"] if sg else "none", d["sgp"] if sg else "none",
                f"=N{r}-M{r}" if sg else "none"],
        {6: "+0;-0;0", 10: "0.0%", 11: "0.0%", 12: "+0.0;-0.0;0.0", 15: "+0.0;-0.0;0.0"})
note(ws, 11, "Each row compares a grade with the students who were in that grade the year before.")
note(ws, 12, "Lexile and growth percentile carry across the change to the new test. Achievement levels and scale scores do not.")
note(ws, 13, "2024-25 figures: district longitudinal file. 2025-26 figures: spring rosters. Grade 3 has no growth percentile.")

# ---------------- Matched students ----------------
ws = sheet("Matched Students", "The same students, one year later (matched by student ID)",
           ["Group", "Students in 2025-26 grade", "Students matched", "Median Lexile 2024-25", "Median Lexile 2025-26",
            "Median gain", "Students who grew", "% who grew"], [22, 13, 11, 13, 13, 10, 11, 10])
for k, g in enumerate(range(4, 9)):
    r = 4 + k
    c = A[g]["cohort"]
    put(ws, r, [f"Grade {g - 1} to grade {g}", A[g]["n"], c["matched"], c["before"], c["now"], c["gain"], c["grew"], f"=G{r}/C{r}"],
        {8: "0.0%"})
note(ws, 10, "Median gain is the median of each student's own change, so it can differ from the gap between the two medians.")

# ---------------- Domains ----------------
ws = sheet("Domains", "Domain results, 2025-26: students at Met Target and Below Target",
           ["Group", "Domain", "Students", "Met Target", "Below Target", "% Met Target", "% Below Target"], [16, 52, 10, 11, 12, 12, 13])
DOM = [("it_css", "Interpreting Texts: Context / Structure & Style"), ("it_tech", "Interpreting Texts: Techniques"),
       ("it_ra", "Interpreting Texts: Research & Analysis"), ("it_gv", "Interpreting Texts: Grammar Conventions / Vocabulary"),
       ("ct_css", "Constructing Texts: Context / Structure & Style"), ("ct_tra", "Constructing Texts: Techniques / Research & Analysis"),
       ("ct_gv", "Constructing Texts: Grammar Conventions / Vocabulary")]
r = 4
for g in range(3, 9):
    for key, label in DOM:
        x = A[g]["dom"][key]
        put(ws, r, [f"Grade {g}", label, A[g]["n"], x["met"], x["below"], f"=D{r}/C{r}", f"=E{r}/C{r}"], {6: "0.0%", 7: "0.0%"})
        r += 1
mca = openpyxl.load_workbook("raw/mca/LitComp2_2025-26_MCA_Analysis.xlsx", data_only=True)
dom = [x for x in mca["Domains"].iter_rows(values_only=True)]
head, body = dom[0], [x for x in dom[1:] if x[0]]
for col in (6, 7, 8, 9, 11, 12, 13):
    vals = [x[col] for x in body]
    put(ws, r, ["Lit & Comp II", head[col], len(body), vals.count("Met Target"), vals.count("Below Target"), f"=D{r}/C{r}", f"=E{r}/C{r}"],
        {6: "0.0%", 7: "0.0%"})
    r += 1
note(ws, r + 1, "Met Target means a 75% or better chance of scoring Proficient on that domain. Approaching Target is the remainder.")

# ---------------- Writing ----------------
ws = sheet("Writing", "Extended writing task: students at each total score (0 to 8), 2025-26",
           ["Group"] + [f"Score {s}" for s in range(9)] + ["Papers scored"], [16] + [9] * 9 + [13])
W = rows_of("writing")
for k, (i, name) in enumerate((("g3", "Grade 3"), ("g5", "Grade 5"), ("g8", "Grade 8"), ("lc2", "Lit & Comp II"))):
    r = 4 + k
    counts = {x["score"]: x["students"] for x in W if x["id"] == i}
    put(ws, r, [name] + [counts.get(s, 0) for s in range(9)] + [f"=SUM(B{r}:J{r})"])
note(ws, 9, "Papers that could not be scored are left out: 17 in grade 3, 7 in grade 5, and 1 in grade 8.")
note(ws, 10, "Lit & Comp II counts two papers reported as Blank or Copied as a score of 0. Grades 4, 6, and 7 have no extended writing score.")

# ---------------- MCA Lexile ----------------
ws = sheet("MCA Lexile", "MCA Lexile: 2024-25 American Lit next to 2025-26 Lit & Comp II",
           ["Measure", "2024-25", "2025-26", "Change", "% 2024-25", "% 2025-26"], [30, 11, 11, 10, 11, 11])
prior = sorted(x[6] for x in mca["2024-25 Am Lit"].iter_rows(min_row=2, values_only=True) if x[0])
now = sorted(x[3] for x in mca["Students"].iter_rows(min_row=2, values_only=True) if x[0])
q = lambda v, i: st.quantiles(v, n=4)[i]
put(ws, 4, ["Students", len(prior), len(now), "=C4-B4"])
for k, (label, f) in enumerate((("Median Lexile", st.median), ("Average Lexile", lambda v: round(st.mean(v))),
                                ("Lowest quarter begins at", lambda v: q(v, 0)), ("Top quarter begins at", lambda v: q(v, 2)),
                                ("Lowest Lexile", min), ("Highest Lexile", max))):
    r = 5 + k
    put(ws, r, [label, f(prior), f(now), f"=C{r}-B{r}"], {4: "+0;-0;0"})
r = 12
for label, test in (("Below 900L", lambda v: v < 900), ("900L to 1079L", lambda v: 900 <= v < 1080),
                    ("1080L to 1184L", lambda v: 1080 <= v < 1185), ("1185L to 1299L", lambda v: 1185 <= v < 1300),
                    ("1300L to 1399L", lambda v: 1300 <= v < 1400), ("1400L and above", lambda v: v >= 1400),
                    ("At or above 1080L", lambda v: v >= 1080), ("At or above 1185L", lambda v: v >= 1185)):
    put(ws, r, [label, sum(map(test, prior)), sum(map(test, now)), f"=C{r}-B{r}", f"=B{r}/B$4", f"=C{r}/C$4"],
        {4: "+0;-0;0", 5: "0.0%", 6: "0.0%"})
    r += 1
note(ws, r + 1, "2024-25 counts 93 American Lit test records, mostly 11th graders. 2025-26 counts 89 Lit & Comp II students.")
note(ws, r + 2, "The two groups took different tests a grade apart, so compare them on fixed Lexile lines, not the state reading status label.")

# ---------------- Notes ----------------
ws = sheet("Notes", "About this workbook", ["Topic", "Detail"], [26, 110])
for k, (a, b) in enumerate([
    ("What is here", "Group totals for 2025-26 Georgia Milestones ELA at Todd Grant (grades 3-5), MCMS (grades 6-8), and MCA (Lit & Comp II), with 2024-25 comparisons."),
    ("What is not here", "No student names, student IDs, or birth dates. Student-level files stay out of this repository."),
    ("New test", "2025-26 was the first year of the new ELA test. Achievement levels and scale scores are not compared across years."),
    ("Fair comparisons", "Lexile and growth percentile are compared across years. Grade-level Lexile lines: 520L, 740L, 830L, 925L, 970L, 1010L for grades 3-8."),
    ("Grade 7 line", "The 970L line for grade 7 sits between the highest Below Grade Level score (950L) and the lowest Grade Level or Above score (985L) on the roster."),
    ("MCA 2024-25", "Uses the 93-record American Lit file. Two students in it tested in both December and May."),
    ("Rebuilding", "Run analysis/run_all.sh from the repository root with the student-level files in raw/."),
]):
    put(ws, 4 + k, [a, b])
    ws.cell(4 + k, 2).alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(4 + k, 1).font = BOLD

wb.save(OUT)
print("wrote", OUT)
