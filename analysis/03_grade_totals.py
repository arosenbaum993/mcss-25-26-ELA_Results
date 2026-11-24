"""Step 3. Reads the 2025-26 spring rosters for grades 3-8 and the 2024-25 ELA rows of the
longitudinal file, and writes one block of totals per grade: achievement levels, Lexile,
growth, domain results, writing scores, and gains for students matched across both years.

Run from the repository root. Reads student-level files from raw/ (never committed)
and writes group totals only.
"""
import glob, warnings, json, pandas as pd, statistics as st
warnings.simplefilter("ignore")
LINE={3:520,4:740,5:830,6:925,7:970,8:1010}
DOM={"Interpreting Texts: Context / Structure & Style Domain Achievement":"it_css","Interpreting Texts: Techniques Domain Achievement":"it_tech","Interpreting Texts: Techniques / Periods & Movements Domain Achievement":"it_tech","Interpreting Texts: Research & Analysis Domain Achievement":"it_ra","Interpreting Texts: Grammar Conventions / Vocabulary Domain Achievement":"it_gv","Constructing Texts: Context / Structure & Style Domain Achievement":"ct_css","Constructing Texts: Techniques / Research & Analysis Domain Achievement":"ct_tra","Constructing Texts: Grammar Conventions / Vocabulary Domain Achievement":"ct_gv"}
R={}
for p in glob.glob("raw/tges/*.xlsx")+glob.glob("raw/mcms/*.xlsx"):
    r=pd.read_excel(p,header=6); r=r[r["Student Name"].notna()].copy(); g=int(r["Grade"].iloc[0]); R[g]=r
    r["lex"]=r["Lexile Score"].astype(str).str.rstrip("L").astype(float); r["ss"]=pd.to_numeric(r["ELA Scale Score¹"]); r["lvl"]=r["ELA Achievement Level Description"].str.split().str[0]; r["sgp"]=pd.to_numeric(r["ELA Growth Percentile"],errors="coerce")
L=pd.read_excel("raw/longitudinal/Longitudinal_Data.xlsx",sheet_name="Data"); e=L[(L.subject=="ELA")&(L.spring_year==2025)].copy(); e["grade"]=e.grade.astype(int)
out={}
for g in sorted(R):
    r=R[g]; n=len(r); c=r.lvl.value_counts(); a=e[e.grade==g]
    d=dict(n=n,levels={k:int(c.get(k,0)) for k in ["Beginning","Developing","Proficient","Distinguished"]},median_ss=float(r.ss.median()),cut=float(r[r.lvl=="Proficient"].ss.min()),max_dev=float(r[r.lvl=="Developing"].ss.max()),
       near=int(((r.ss>=r[r.lvl=="Proficient"].ss.min()-15)&(r.lvl=="Developing")).sum()),on_level=int((r["Reading Status"]=="Grade Level or Above").sum()),med_lex=float(r.lex.median()),at_line=int((r.lex>=LINE[g]).sum()),
       sgp=None if r.sgp.isna().all() else float(r.sgp.median()),growth=r["ELA Growth Level"].str.strip().value_counts().to_dict(),
       dom={v:dict(met=int((r[k]=="Met Target").sum()),below=int((r[k]=="Below Target").sum())) for k,v in DOM.items() if k in r.columns},
       prior=dict(n=len(a),med_lex=float(a.lexile.median()),at_line=int((a.lexile>=LINE[g]).sum()),sgp=None if a.sgp.isna().all() else float(a.sgp.median())))
    if "Extended Writing Task - Purpose" in r.columns:
        w=r[["Extended Writing Task - Purpose","Extended Writing Task - Evidence","Extended Writing Task - Language"]]
        num=w.apply(lambda s:pd.to_numeric(s.astype(str).str.split().str[0],errors="coerce")); tot=num.sum(axis=1,min_count=3)
        d["writing"]={int(k):int(v) for k,v in tot.dropna().astype(int).value_counts().sort_index().items()}; d["writing_unscored"]=int(tot.isna().sum()); d["codes"]=w.iloc[:,0][num.iloc[:,0].isna()].value_counts().to_dict()
        d["trait_avg"]=[round(float(x),2) for x in num.mean()]; d["lang_full"]=int((num.iloc[:,2]==2).sum()); d["scored"]=int(tot.notna().sum())
    if g>3:
        b=e[e.grade==g-1][["gtid","lexile"]]; m=b.merge(pd.DataFrame(dict(gtid=r["GTID"].astype("int64"),lex=r.lex)),on="gtid"); gain=m.lex-m.lexile
        d["cohort"]=dict(matched=len(m),gain=float(gain.median()),grew=int((gain>0).sum()),before=float(m.lexile.median()),now=float(m.lex.median()))
    out[g]=d
json.dump(out,open("analysis/out/agg.json","w"),indent=1)
print('wrote analysis/out/agg.json for grades',sorted(out))
