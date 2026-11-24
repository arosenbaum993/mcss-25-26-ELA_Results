"""Step 5 (check). Prints the MCA Lexile figures for 2024-25 American Lit and 2025-26 Lit & Comp II:
bands, quartiles, winter and spring, and Lexile by achievement level. Prints totals only.

Run from the repository root. Reads student-level files from raw/ (never committed)
and writes group totals only.
"""
import openpyxl, warnings, statistics as st, json
from collections import Counter
warnings.simplefilter("ignore")
wb=openpyxl.load_workbook("raw/mca/LitComp2_2025-26_MCA_Analysis.xlsx",data_only=True)
p=[r for r in wb["2024-25 Am Lit"].iter_rows(min_row=2,values_only=True) if r[0]]
s=[r for r in wb["Students"].iter_rows(min_row=2,values_only=True) if r[0]]
P=[dict(lex=r[6],level=r[5],adm=r[2],grade=r[3]) for r in p]
S=[dict(lex=r[3],level=r[5],adm="Winter" if r[0]=="Winter" else "Spring",cls=r[0]) for r in s]
bands=[("Below 900L",0,899),("900L to 1079L",900,1079),("1080L to 1184L",1080,1184),("1185L to 1299L",1185,1299),("1300L to 1399L",1300,1399),("1400L and above",1400,9999)]
def stats(X):
    l=sorted(x["lex"] for x in X); n=len(l); q=st.quantiles(l,n=4)
    return dict(n=n,median=st.median(l),mean=round(st.mean(l)),q1=q[0],q3=q[2],lo=l[0],hi=l[-1],a1080=sum(v>=1080 for v in l),a1185=sum(v>=1185 for v in l),a1300=sum(v>=1300 for v in l),a1400=sum(v>=1400 for v in l),b900=sum(v<900 for v in l))
for name,X in (("2024-25",P),("2025-26",S)):
    print(name,stats(X)); print("  bands",[(b[0],sum(b[1]<=x["lex"]<=b[2] for x in X),round(100*sum(b[1]<=x["lex"]<=b[2] for x in X)/len(X))) for b in bands])
    for a in ("Winter","Spring"): print("  ",a,stats([x for x in X if x["adm"]==a]))
    for lv in ("Beginning","Developing","Proficient","Distinguished"):
        L=[x for x in X if x["level"]==lv]; 
        if L: print("   level",lv,len(L),"median",st.median(x["lex"] for x in L),"min",min(x["lex"] for x in L),"max",max(x["lex"] for x in L),">=1080:",sum(x["lex"]>=1080 for x in L),">=1185:",sum(x["lex"]>=1185 for x in L))
print("2024-25 grades:",Counter(x["grade"] for x in P)); print("2025-26 classes:",Counter(x["cls"] for x in S))
print("2024-25 grade11 only:",stats([x for x in P if x["grade"]==11]))
