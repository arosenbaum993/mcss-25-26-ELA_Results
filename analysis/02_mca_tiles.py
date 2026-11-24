"""Step 2. The four MCA Lexile gain tiles and the first bright-spot tiles.
The Lexile figures were calculated with 05_mca_lexile_check.py and entered here.

Run from the repository root. Reads student-level files from raw/ (never committed)
and writes group totals only.
"""
import json
SRC="Lexile column for every student in LitComp 2 2025-26 MCA Analysis.xlsx (89 students, Winter + Spring) and 2024-25 American Lit Milestones.xlsx (93 test records). Percentages are rounded to whole numbers; the lowest-quarter figure is the 25th percentile."
rows=[
 dict(key="at1080",order=1,label="read at or above 1080L",now="82%",before="78%",change="Up 4 points"),
 dict(key="at1185",order=2,label="read at or above 1185L",now="65%",before="59%",change="Up 6 points"),
 dict(key="q1",order=3,label="where the lowest quarter of readers begins",now="1138L",before="1080L",change="58L higher"),
 dict(key="low",order=4,label="lowest Lexile in the group",now="685L",before="405L",change="280L higher")]
json.dump({"rows":rows,"meta":[{"label":"Source","value":SRC}]},open("dashboard/data/lexile.json","w"),indent=1)
B=lambda k,o,big,label,note: dict(key=k,order=o,big=big,label=label,note=note)
rows=[B("growth",1,"53, 56, 57","median growth percentiles in grades 6, 7, and 8","Growth was typical to strong in every grade."),
 B("g8growth",2,"41","Grade 8 students showed High Growth","Median growth percentile of 57."),
 B("g7read",3,"85%","of Grade 7 students read on grade level","Median Lexile 1115L."),
 B("g8lang",4,"75%","of Grade 8 students earned full credit on Language","Clean, correct sentences in the argument essay."),
 B("lc2goal",5,"52%","of Lit & Comp II students scored Proficient or above","That clears the MCA improvement plan goal of 50%."),
 B("lc2beg",6,"8%","of Lit & Comp II students scored Beginning","The improvement plan goal was 20% or fewer.")]
json.dump({"rows":rows,"meta":[]},open("dashboard/data/bright.json","w"),indent=1)
