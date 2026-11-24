"""Step 4. Adds Todd Grant grades 3-5 to the dashboard files and builds the year-over-year
tiles, the grade-by-grade comparison table, and the matched-student table.

Run from the repository root. Reads student-level files from raw/ (never committed)
and writes group totals only.
"""
import json
A=json.load(open("analysis/out/agg.json")); A={int(k):v for k,v in A.items()}
import pandas as pd, glob, warnings; warnings.simplefilter("ignore")
rnd=lambda x:int(x+0.5)
near={}
for p in glob.glob("raw/tges/*.xlsx"):
    r=pd.read_excel(p,header=6); r=r[r["Student Name"].notna()]; s=pd.to_numeric(r["ELA Scale Score¹"]); near[int(r["Grade"].iloc[0])]=int(((s>=660)&(s<=674)).sum())
SRC25="2025-26 student rosters: Georgia Milestones Spring 2026 ELA EOG extracts for Todd Grant grades 3-5 and MCMS grades 6-8 (Nov 24, 2026)"
SRCL="Longitudinal_Data.xlsx, ELA rows for 2024-25"
def dump(name,rows,src): json.dump({"rows":rows,"meta":[{"label":"Source","value":src}]},open(f"dashboard/data/{name}.json","w"),indent=1)
old=lambda n: json.load(open(f"dashboard/data/{n}.json"))["rows"]
# levels
def lv(id,name,school,now,order,b,d,p,x,median):
    n=b+d+p+x; pc=lambda v: rnd(100*v/n)
    return dict(id=id,name=name,school=school,now=now,order=order,tested=n,beginning=b,developing=d,proficient=p,distinguished=x,prof_plus=p+x,pct_beginning=pc(b),pct_developing=pc(d),pct_proficient=pc(p),pct_distinguished=pc(x),pct_prof_plus=pc(p+x),median_scale=median,proficient_cut=675)
L=old("levels"); newl=[]; tot=[0,0,0,0]
for g,now in ((3,"Now in 4th grade"),(4,"Now in 5th grade"),(5,"Now in 6th grade at MCMS")):
    l=A[g]["levels"]; v=[l["Beginning"],l["Developing"],l["Proficient"],l["Distinguished"]]; tot=[a+b for a,b in zip(tot,v)]
    newl.append(lv(f"g{g}",f"Grade {g}","TGES",now,g,*v,A[g]["median_ss"]))
newl.append(lv("tges","Grades 3-5","TGES","All three grades",0,*tot,None))
dump("levels",newl+L,SRC25+"; Lit & Comp II row from LitComp 2 2025-26 MCA Analysis.xlsx")
# facts
F=lambda id,o,display,label,note: dict(key=f"{id}-{o}",id=id,order=o,display=display,label=label,note=note)
pct=lambda a,b: f"{rnd(100*a/b)}%"
nf=[F("g3",1,"646","median scale score","29 points under the Proficient cut of 675."),F("g3",2,pct(A[3]["on_level"],125),"read on grade level",f"{A[3]['on_level']} of 125 students, based on Lexile."),
 F("g3",3,str(near[3]),"students just under Proficient","Scored 660-674."),F("g3",4,"47","students scored Proficient or Distinguished","In their first year taking the Milestones."),
 F("g4",1,"673","median scale score","Two points under the Proficient cut of 675."),F("g4",2,pct(A[4]["on_level"],99),"read on grade level",f"{A[4]['on_level']} of 99 students, based on Lexile."),
 F("g4",3,str(near[4]),"students just under Proficient","Scored 660-674."),F("g4",4,"27","students showed High Growth","Median growth percentile of 42.5."),
 F("g5",1,"661","median scale score","14 points under the Proficient cut of 675."),F("g5",2,pct(A[5]["on_level"],108),"read on grade level",f"{A[5]['on_level']} of 108 students, based on Lexile."),
 F("g5",3,str(near[5]),"students just under Proficient","Scored 660-674."),F("g5",4,"24","students showed High Growth","Median growth percentile of 39.")]
dump("facts",nf+old("facts"),SRC25+"; grades 6-8 and Lit & Comp II figures as before")
# gaps
NAMES={"it_css":("Structure and context in reading","Interpreting Texts: Context / Structure & Style"),"it_tech":("Author's craft in reading","Interpreting Texts: Techniques"),"it_ra":("Reading across sources","Interpreting Texts: Research & Analysis"),"it_gv":("Grammar and vocabulary in reading","Interpreting Texts: Grammar Conventions / Vocabulary"),"ct_css":("Organizing their writing","Constructing Texts: Context / Structure & Style"),"ct_tra":("Elaboration in writing","Constructing Texts: Techniques / Research & Analysis"),"ct_gv":("Grammar in their own writing","Constructing Texts: Grammar Conventions / Vocabulary")}
ng=[]
for g in (3,4,5):
    d=A[g]["dom"]; n=A[g]["n"]; best=max(d,key=lambda k:d[k]["met"]); worst=sorted([k for k in d if k!=best],key=lambda k:-d[k]["below"])[:3]
    ng.append(dict(key=f"g{g}-1",id=f"g{g}",order=1,kind="Strength",area=NAMES[best][0],domain=NAMES[best][1],result=f"{rnd(100*d[best]['met']/n)}% Met Target",meaning=f"The strongest of the seven domains ({d[best]['met']} students)."))
    for i,k in enumerate(worst): ng.append(dict(key=f"g{g}-{i+2}",id=f"g{g}",order=i+2,kind="Gap",area=NAMES[k][0],domain=NAMES[k][1],result=f"{rnd(100*d[k]['below']/n)}% Below Target",meaning=f"{d[k]['below']} of {n} students Below Target."))
dump("gaps",ng+old("gaps"),SRC25+"; grades 6-8 and Lit & Comp II rows as before")
# skills
S=old("skills"); ns=[]
for sk,dk,suf in (("Elaboration in writing","ct_tra","elab"),("Reading across sources","it_ra","sources")):
    dom=[r for r in S if r["skill"]==sk][0]["domain"]
    for g in (3,4,5): ns.append(dict(key=f"g{g}-{suf}",id=f"g{g}",group=f"Grade {g}",skill=sk,domain=dom,pct_met=rnd(100*A[g]["dom"][dk]["met"]/A[g]["n"])))
    ns+= [r for r in S if r["skill"]==sk]
dump("skills",ns,SRC25+"; Lit & Comp II from LitComp 2 2025-26 MCA Analysis.xlsx")
# writing
W=old("writing"); nw=[]
for g in (3,5):
    for s in range(9): nw.append(dict(key=f"g{g}-{s}",id=f"g{g}",score=s,students=A[g]["writing"].get(str(s),A[g]["writing"].get(s,0))))
dump("writing",nw+W,SRC25+"; Lit & Comp II essay totals from LitComp 2 2025-26 MCA Analysis.xlsx")
# yoy
Y=[]; T=[]; C=[]
def y(id,o,now,label,change,before,tone): Y.append(dict(key=f"{id}-{o}",id=id,order=o,now=now,label=label,change=change,before=before,tone=tone))
def delta(a,b,unit,flat=2):
    d=a-b
    if abs(d)<flat: return "About the same","steady"
    if unit=="pts": return (f"Up {d} points","gain") if d>0 else (f"Down {-d} points","lower")
    if unit=="L": return (f"{d}L higher","gain") if d>0 else (f"{-d}L lower","lower")
    return (f"Up {d:g}","gain") if d>0 else (f"Down {-d:g}","lower")
for g in range(3,9):
    d=A[g]; p=d["prior"]; id=f"g{g}"; o=1
    if "cohort" in d:
        c=d["cohort"]; y(id,o,f"{rnd(c['gain'])}L",f"median Lexile gain since grade {g-1}, same students",f"{rnd(100*c['grew']/c['matched'])}% of {c['matched']} matched students grew","","gain"); o+=1
        C.append(dict(key=f"c{g}",order=g,cohort=f"Grade {g-1} to grade {g}",matched=f"{c['matched']} of {d['n']}",before=f"{rnd(c['before'])}L",now=f"{rnd(c['now'])}L",gain=f"{rnd(c['gain'])}L",grew=f"{rnd(100*c['grew']/c['matched'])}%"))
    a,b=rnd(100*d["at_line"]/d["n"]),rnd(100*p["at_line"]/p["n"]); ch,t=delta(rnd(abs(100*d["at_line"]/d["n"]-100*p["at_line"]/p["n"]))*(1 if d["at_line"]/d["n"]>p["at_line"]/p["n"] else -1),0,"pts"); y(id,o,f"{a}%","read at or above the grade-level Lexile line",ch,f"{b}%",t); o+=1
    la,lb=rnd(d["med_lex"]),rnd(p["med_lex"]); ch2,t2=delta(la,lb,"L",10); y(id,o,f"{la}L","median Lexile",ch2,f"{lb}L",t2); o+=1
    row=dict(key=f"t{g}",order=g,grade=f"Grade {g}",lex_before=f"{lb}L",lex_now=f"{la}L",line_before=f"{b}%",line_now=f"{a}%",sgp_before="No growth score",sgp_now="No growth score")
    if d["sgp"] is not None:
        ch3,t3=delta(d["sgp"],p["sgp"],"n",1); y(id,o,f"{d['sgp']:g}","median growth percentile",ch3,f"{p['sgp']:g}",t3); row.update(sgp_before=f"{p['sgp']:g}",sgp_now=f"{d['sgp']:g}")
    T.append(row)
y("all",1,"5 of 5","grades with growth scores raised their median growth percentile","Grades 4 through 8","","gain")
y("all",2,"175L","median Lexile gain for last year's 6th graders, now in 7th","93% of 90 matched students grew","","gain")
y("all",3,"160L","median Lexile gain for both elementary groups we could follow","Grade 3 to 4 and grade 4 to 5","","gain")
y("all",4,"66%","of 6th graders read at the grade-level Lexile line","Up 16 points","50%","gain")
dump("yoy",Y,SRC25+" and "+SRCL+". Grade-level Lexile lines: 520L, 740L, 830L, 925L, 970L, 1010L for grades 3-8. Matched students share a student ID in both years."); dump("yoy_table",T,SRC25+" and "+SRCL); dump("cohorts",C,SRC25+" and "+SRCL+", matched by student ID")
# bright
B=old("bright"); B=[b for b in B if b["key"]!="growth"]
for i,b in enumerate(B): b["order"]=i+4
nb=[dict(key="growthup",order=1,big="5 of 5",label="grades with growth scores raised their median growth percentile",note="Grades 4 through 8, compared with 2024-25."),
    dict(key="cohort67",order=2,big="175L",label="median Lexile gain for last year's 6th graders",note="93% of the students we could match grew."),
    dict(key="cohortel",order=3,big="160L",label="median Lexile gain for both elementary groups",note="Students followed from grade 3 to 4 and from grade 4 to 5.")]
dump("bright",nb+B,"Year-over-year figures from the 2025-26 rosters and Longitudinal_Data.xlsx; other figures from the grades 6-8 deck and the Lit & Comp II workbook; improvement plan targets from the MCA SIP 2025-26 goals supplied by A. Rosenbaum.")
print('updated dashboard/data for all grades')
