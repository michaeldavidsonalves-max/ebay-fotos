import re,os,csv,glob,json,sys
from python_calamine import CalamineWorkbook
repo=sys.argv[1]; arch=f"{repo}/listings-archive"
readme=open(f"{arch}/README.md").read()
lots=[re.search(r"\[`([^`]+)`\]",l).group(1) for l in readme.splitlines() if l.startswith("| [`")]
dates={re.search(r"\[`([^`]+)`\]",l).group(1): re.findall(r"\d{4}-\d{2}-\d{2}",l.split(" | ")[-1]+l)[0] for l in readme.splitlines() if l.startswith("| [`")}
extra=[f for f in os.listdir(arch) if f.endswith(".xlsx") and f not in lots]
# ItemIDs from upload reports
ids={}
for f in glob.glob(os.path.expanduser("~/Downloads/*-20[0-9][0-9]-*.csv")):
    if not re.search(r"-(Sep|Oct|Nov|Dec)-2026-",f): continue
    for r in csv.DictReader(open(f,encoding="utf-8-sig")):
        if r.get("Status") in ("Success","Warning") and r.get("ItemID","").strip() and r.get("CustomLabel","").strip():
            ids[r["CustomLabel"].strip()]=(r["ItemID"].strip(), os.path.basename(f))
LIVE_FILES={"ja-ao-vivo-consolidado-com-precos-2026-10-02.xlsx":"Ao vivo (consolidado confirmado 02/10)",
 "novos-revisados-consolidado-pid-2026-10-02.xlsx":"Ao vivo (confirmado 03/10)",
 "novos-revisados-reenvio-12itens-2026-10-03.xlsx":"Ao vivo (reenvio confirmado 03/10)",
 "onthefigure-batch2-dc-marvel-plus-doctordoom-2026-11-25.xlsx":"Ao vivo (confirmado pelo dono 25/11)"}
items={}
def key(f):
    m=re.search(r"\d{4}-\d{2}-\d{2}",f); d=m.group(0) if m else "0000"
    return (d, 0 if f in extra else 1, lots.index(f) if f in lots else 0)
for lot in sorted(lots+extra,key=key):
    try: rows=CalamineWorkbook.from_path(f"{arch}/{lot}").get_sheet_by_name('Listings').to_python()
    except Exception as e: print("skip",lot,e); continue
    h=[str(x) for x in rows[3]]
    def col(name):
        for i,x in enumerate(h):
            if x.split("(")[0].strip().lower().startswith(name.lower()): return i
    iS,iT,iP,iC,iU,iQ,iM,iMPN=col("Custom label"),col("Title"),col("Start price"),col("Condition ID"),col("P:UPC"),col("Quantity"),col("Item photo URL"),col("C:MPN")
    for r in rows[4:]:
        if str(r[0]).strip()!="Add": continue
        sku=str(r[iS]).strip()
        if not sku: continue
        prev=items.get(sku,{})
        price=r[iP] if r[iP] not in ("",None) else prev.get("price","")
        photos=[u for u in str(r[iM]).split("|") if u.strip()]
        items[sku]=dict(sku=sku,title=r[iT],price=price,cond=("New" if str(r[iC]).startswith("1000") else "Used" if str(r[iC]).startswith("3000") else str(r[iC])),
            upc=str(r[iU]),qty=r[iQ],cover=photos[0] if photos else "",nphotos=len(photos),
            lot=lot,date=dates.get(lot,"(fora do índice)"),lots=prev.get("lots",[])+[lot],
            live=LIVE_FILES.get(lot) or prev.get("live",""))
out=[]
for s,d in items.items():
    if s in ids: st,iid,src="Ao vivo (ItemID no relatório do eBay)",ids[s][0],ids[s][1]
    elif d["live"]: st,iid,src=d["live"],"",""
    else: st,iid,src="Não confirmado",""," "
    desc=f"descriptions/{s.lower()}.html"
    d.update(status=st,itemid=iid,report=src,desc=desc if os.path.exists(f"{repo}/{desc}") else "(dentro da planilha do lote)")
    out.append(d)
json.dump(out,open(sys.argv[2],"w"),default=str)
from collections import Counter
print(len(out),"SKUs únicos"); print(Counter(d["status"] for d in out))
print("ItemIDs from reports:",len(ids))
