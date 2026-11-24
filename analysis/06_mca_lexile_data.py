"""Step 6. Writes the Lexile deep-dive files for MCA. The figures were calculated with
05_mca_lexile_check.py and entered here.

Run from the repository root. Reads student-level files from raw/ (never committed)
and writes group totals only.
"""
import json
SRC="Lexile column for every student in LitComp 2 2025-26 MCA Analysis.xlsx (89 students, Winter + Spring 2025-26) and 2024-25 American Lit Milestones.xlsx (93 test records, mostly 11th graders). Percentages are rounded to whole numbers."
def dump(name,rows): json.dump({"rows":rows,"meta":[{"label":"Source","value":SRC}]},open(f"dashboard/data/{name}.json","w"),indent=1)
dump("lex_bands",[dict(key=k,order=i+1,band=b,n_before=nb,pct_before=pb,n_now=nn,pct_now=pn) for i,(k,b,nb,pb,nn,pn) in enumerate([
 ("b900","Below 900L",8,9,2,2),("b1079","900L to 1079L",12,13,14,16),("b1184","1080L to 1184L",18,19,15,17),("b1299","1185L to 1299L",10,11,18,20),("b1399","1300L to 1399L",24,26,15,17),("b1400","1400L and above",21,23,25,28)])])
dump("lex_stats",[dict(key=k,order=i+1,measure=m,before=b,now=n,change=c) for i,(k,m,b,n,c) in enumerate([
 ("n","Students tested","93","89",""),("median","Median Lexile","1285L","1260L","25L lower"),("mean","Average Lexile","1242L","1283L","41L higher"),
 ("q1","Lowest quarter begins at","1080L","1138L","58L higher"),("q3","Top quarter begins at","1380L","1430L","50L higher"),
 ("low","Lowest Lexile","405L","685L","280L higher"),("high","Highest Lexile","1800L","1795L","About the same"),
 ("b900","Below 900L","8 (9%)","2 (2%)","6 fewer students"),("a1080","At or above 1080L","73 (78%)","73 (82%)","Up 4 points"),
 ("a1185","At or above 1185L","55 (59%)","58 (65%)","Up 6 points"),("a1400","At or above 1400L","21 (23%)","25 (28%)","Up 5 points")])])
dump("lex_terms",[dict(key=k,order=i+1,year=y,term=t,students=n,median=m,at1080=a) for i,(k,y,t,n,m,a) in enumerate([
 ("w2425","2024-25","Winter",46,"1333L","87%"),("s2425","2024-25","Spring",47,"1160L","70%"),("w2526","2025-26","Winter",26,"1148L","62%"),("s2526","2025-26","Spring",63,"1330L","90%")])])
dump("lex_levels",[dict(key=k,order=i+1,level=l,students=n,median=m,at1080=a) for i,(k,l,n,m,a) in enumerate([
 ("beg","Beginning",7,"925L","0 of 7"),("dev","Developing",36,"1160L","27 of 36"),("pro","Proficient",34,"1368L","34 of 34"),("dis","Distinguished",12,"1690L","12 of 12")])])
dump("lex_notes",[dict(key=f"n{i+1}",order=i+1,kind=k,change=c,detail=d) for i,(k,c,d) in enumerate([
 ("Gain","Far fewer very low readers","Students reading below 900L fell from 8 to 2."),
 ("Gain","The lowest quarter moved up","It now begins at 1138L, up from 1080L, and the lowest Lexile in the group rose from 405L to 685L."),
 ("Gain","More readers at the top","Students at 1400L or above grew from 21 (23%) to 25 (28%), and the top quarter now begins at 1430L, up from 1380L."),
 ("Gain","More readers over both grade-level lines","82% read at or above 1080L (was 78%), and 65% read at or above 1185L (was 59%)."),
 ("Gain","Average Lexile is up","The average rose from 1242L to 1283L."),
 ("Steady","The middle shifted slightly","The median is 1260L, 25 points under the prior year's 1285L. This group tested a grade younger. Fewer students landed in the 1300s (15, down from 24) and more landed between 1185L and 1299L (18, up from 10)."),
 ("Watch","The stronger term flipped","In 2024-25 the winter group read higher than spring (median 1333L against 1160L). In 2025-26 spring read higher than winter (1330L against 1148L). The gap follows which students test each term, so check each term's readers early."),
 ("Next step","Most Developing students can already read the text","27 of the 36 Developing students read at or above 1080L. For them the next step is writing and analysis, not reading level."),
 ("Next step","Beginning students need reading support first","All 7 Beginning students read below 1080L, from 685L to 1065L.")])])
