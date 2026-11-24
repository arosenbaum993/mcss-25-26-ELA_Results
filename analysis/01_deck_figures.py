"""Step 1. Figures entered by hand from the Nov 18, 2026 decks and grade plans (grades 6-8)
and from the Lit & Comp II workbook: achievement levels, key numbers, strengths and gaps,
two shared skills, and essay score counts. Later steps add to these files.

Run from the repository root. Reads student-level files from raw/ (never committed)
and writes group totals only.
"""
import json, os
os.makedirs('dashboard/data',exist_ok=True)
SRC6="MCMS ELA EOG rosters, Spring 2026, as summarized in the Grades 6-8 District Data Review deck, grade plans and PLC decks (Nov 18, 2026)"
SRCL="LitComp 2 2025-26 MCA Analysis.xlsx (Winter Dec 2025 + Spring 2026 EOC rosters)"
def lv(id,name,school,now,order,b,d,p,x,median):
    n=b+d+p+x; pc=lambda v: round(100*v/n)
    return dict(id=id,name=name,school=school,now=now,order=order,tested=n,beginning=b,developing=d,proficient=p,distinguished=x,prof_plus=p+x,
                pct_beginning=pc(b),pct_developing=pc(d),pct_proficient=pc(p),pct_distinguished=pc(x),pct_prof_plus=pc(p+x),median_scale=median,proficient_cut=675)
levels=[lv("g6","Grade 6","MCMS","Now in 7th grade",1,8,38,43,7,678), lv("g7","Grade 7","MCMS","Now in 8th grade",2,5,50,34,4,669),
        lv("g8","Grade 8","MCMS","Now in 9th grade at MCA",3,8,46,39,8,673), lv("lc2","Lit & Comp II","MCA","EOC course, mostly 10th grade",4,7,36,34,12,676),
        lv("mcms","Grades 6-8","MCMS","All three grades",5,21,134,116,19,None)]
groups=[lv("winter","Winter 2025-26","MCA","",1,5,14,4,3,None), lv("spring","Spring 2025-26","MCA","",2,2,22,30,9,None),
        lv("combined","2025-26 combined","MCA","",3,7,36,34,12,None), lv("prior","2024-25 American Lit","MCA","",4,29,33,28,3,None)]
for g in groups:
    for k in ("school","now","median_scale","proficient_cut"): g.pop(k)
F=lambda id,o,display,label,note: dict(key=f"{id}-{o}",id=id,order=o,display=display,label=label,note=note)
facts=[F("g6",1,"678","median scale score","Three points over the Proficient cut of 675."),
 F("g6",2,"33","students reading below grade level","One in three, based on Lexile."),
 F("g6",3,"11","students just under Proficient","Scored 660-674, inside the test's margin of error."),
 F("g6",4,"41%","Below Target on author's craft","The most of any grade."),
 F("g7",1,"669","median scale score","Six points under the Proficient cut of 675."),
 F("g7",2,"85%","read on grade level","Median Lexile 1115L. Reading level is not the barrier."),
 F("g7",3,"21","students just under Proficient","Scored 660-674, inside the test's margin of error."),
 F("g7",4,"30 of 41","on-level Developing readers Below Target in writing","They can read it. They can't yet write about it."),
 F("g8",1,"673","median scale score","Two points under the Proficient cut of 675."),
 F("g8",2,"45","students scored exactly 2-2-2 on the argument","Including 26 of the 39 Proficient students."),
 F("g8",3,"75%","earned full credit on Language","Conventions are a strength."),
 F("g8",4,"22","students just under Proficient","Scored 660-674, inside the test's margin of error."),
 F("lc2",1,"676","median scale score","One point over the Proficient cut of 675."),
 F("lc2",2,"82%","read at grade level or above","73 of 89 students at 1080L or higher."),
 F("lc2",3,"16","students just under Proficient","Scored 660-674."),
 F("lc2",4,"30","students scored exactly 2-2-2 on the essay","21 students earned the full 8.")]
G=lambda id,o,kind,area,domain,result,meaning: dict(key=f"{id}-{o}",id=id,order=o,kind=kind,area=area,domain=domain,result=result,meaning=meaning)
gaps=[G("g6",1,"Strength","Grammar and vocabulary in reading","Interpreting Texts: Grammar Conventions / Vocabulary","19% Below Target","Students know the rules."),
 G("g6",2,"Gap","Author's craft in reading","Interpreting Texts: Techniques / Periods & Movements","41% Below Target","Main gap for Developing students (29 of 38)."),
 G("g6",3,"Gap","Elaboration in writing","Constructing Texts: Techniques / Research & Analysis","27% Met Target","Even Proficient students fall Below."),
 G("g6",4,"Gap","Grammar in their own writing","Constructing Texts: Grammar Conventions / Vocabulary","38% Below Target","The rules aren't showing up in writing."),
 G("g7",1,"Strength","Reading on grade level","Reading Status / Lexile","85% at grade level","Students can access the text."),
 G("g7",2,"Gap","Elaboration in writing","Constructing Texts: Techniques / Research & Analysis","44% Below Target","Only 24% Met Target."),
 G("g7",3,"Gap","Reading across sources","Interpreting Texts: Research & Analysis","37% Below Target","Includes 10 of 34 Proficient students."),
 G("g7",4,"Gap","Grammar and vocabulary","Grammar Conventions / Vocabulary (both domains)","33-38% Below Target","Weak in reading and in writing."),
 G("g8",1,"Strength","Language and conventions","Argument trait: Language Usage & Conventions (0-2)","75% full credit","Students write clean sentences."),
 G("g8",2,"Gap","Explaining evidence","Argument trait: Evidence & Elaboration (0-3)","33% scored 0-1","Only 18% earned a 3."),
 G("g8",3,"Gap","Organizing the argument","Argument trait: Purpose & Organization (0-3)","20% earned a 3","Most stop at a 2."),
 G("g8",4,"Gap","Reading across sources","Interpreting Texts: Research & Analysis","25% Met Target","The argument is built on this skill."),
 G("lc2",1,"Strength","Reading on grade level","Reading Status / Lexile (1080L+)","82% at grade level or above","Most students can access the text."),
 G("lc2",2,"Gap","Grammar and vocabulary in reading","Interpreting Texts: Grammar / Vocabulary","25% Met Target","The lowest of the seven domains."),
 G("lc2",3,"Gap","Author's craft in reading","Interpreting Texts: Techniques / Periods & Movements","33% Met Target","Carries the most test weight: 14 of 60 points."),
 G("lc2",4,"Gap","The winter section","All domains","27% Proficient or above, against 62% in spring","Winter also averaged 4.0 of 8 on the essay, against 5.4 in spring.")]
skills=[]
for sk,dom,vals in (("Elaboration in writing","Constructing Texts: Techniques / Research & Analysis",(27,24,36,42)),("Reading across sources","Interpreting Texts: Research & Analysis",(33,29,25,33))):
    for (id,name),v in zip((("g6","Grade 6"),("g7","Grade 7"),("g8","Grade 8"),("lc2","Lit & Comp II")),vals):
        skills.append(dict(key=f"{id}-{'elab' if sk.startswith('Elab') else 'sources'}",id=id,group=name,skill=sk,domain=dom,pct_met=v))
writing=[]
for id,d in (("g8",{0:7,1:0,2:2,3:11,4:11,5:4,6:45,7:2,8:18}),("lc2",{0:8,1:3,2:1,3:17,4:3,5:6,6:30,7:0,8:21})):
    for s,n in d.items(): writing.append(dict(key=f"{id}-{s}",id=id,score=s,students=n))
assert sum(r["students"] for r in writing if r["id"]=="lc2")==89 and sum(r["students"] for r in writing if r["id"]=="g8")==100
out={"levels":(levels,SRC6+"; Lit & Comp II row from "+SRCL+". Level counts for grades 6-8 are worked out from the totals the decks report."),
     "lc2_groups":(groups,SRCL+"; 2024-25 row from 2024-25 American Lit Milestones.xlsx"),
     "facts":(facts,SRC6+"; Lit & Comp II from "+SRCL),"gaps":(gaps,SRC6+"; Lit & Comp II from "+SRCL),
     "skills":(skills,SRC6+"; Lit & Comp II from "+SRCL),"writing":(writing,"Grade 8 argument totals from the Grades 6-8 District Data Review deck; Lit & Comp II essay totals from "+SRCL)}
for k,(rows,src) in out.items():
    json.dump({"rows":rows,"meta":[{"label":"Source","value":src}]},open(f"dashboard/data/{k}.json","w"),indent=1)
print({k:len(v[0]) for k,v in out.items()}); print([(l["id"],l["pct_beginning"],l["pct_developing"],l["pct_proficient"],l["pct_distinguished"],l["pct_prof_plus"]) for l in levels+groups])
