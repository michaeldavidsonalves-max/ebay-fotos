# Gera MASTER-anuncios-ebay.xlsx/.csv a partir do JSON de gerar_master_dados.py
# Uso: python3 gerar_master_planilha.py master.json <pasta listings-archive>
import json,sys,csv,collections
from openpyxl import Workbook
from openpyxl.styles import Font,PatternFill
from openpyxl.worksheet.table import Table,TableStyleInfo
d=json.load(open(sys.argv[1])); outdir=sys.argv[2]
# Duplicados: SKU que sobra -> SKU que fica (produto já ao vivo)
DUP={"DCCDCG43521-10":"IS-AQUAMAN-COMICS-DELUXE","BATM6631620-10":"IS-BATMAN-66-DELUXE",
"090026":"IS-BATMAN-DARK-KNIGHT-DELUXE","DCCDCG15419-10":"IS-BATMAN-EDDY-BARROWS",
"DCCBAT39120-10":"IS-CATWOMAN-BATMAN-RETURNS","DCCW8427420-10":"IS-WONDERWOMAN-DELUXE-GOLD-WW84",
"160006":"IS-MCOMICS-HULK","BATM6631720-10":"IS-JOKER-DELUXE-BATMAN66","DCCJLE57421-10":"IS-MERA-ZSJL-2026",
"DCCJLE41321-10":"IS-SUPERMAN-BLACKSUIT-JL","DCCJLE56521-10":"IS-WONDERWOMAN-ZSJL",
"DCCDCG29820-10":"IS-BATMAN-ROBIN-DELUXE","DCCW8433120-10":"IS-WONDERWOMAN-YOUNGDIANA-WW84"}
OBS={"IS-PENIPARKER-SPDR-SVM":"⚠ Possível anúncio repetido AO VIVO: mesmo UPC de IS-PENIPARKER-SPDR-110 – conferir no eBay e encerrar um",
"IS-PENIPARKER-SPDR-110":"⚠ Possível anúncio repetido AO VIVO: mesmo UPC de IS-PENIPARKER-SPDR-SVM – conferir no eBay e encerrar um",
"IS-DEADPOOL-1006":"Produto diferente de IS-DEADPOOL-XFORCE-CCXP2017, mas com título idêntico",
"IS-DEADPOOL-XFORCE-CCXP2017":"Produto diferente de IS-DEADPOOL-1006, mas com título idêntico",
"IS-MK-LOT4-GORO-RAIDEN-SUBZERO-SCORPION":"Kit (Goro+Raiden+Sub-Zero+Scorpion) – os 4 também estão anunciados avulsos"}
skus={x["sku"] for x in d}
assert all(k in skus and v in skus for k,v in DUP.items()), [k for k,v in DUP.items() if k not in skus or v not in skus]
keep=[];excl=[]
for x in d:
    x["obs"]=OBS.get(x["sku"],"")
    if x["sku"] in DUP: x["excl"]=f"Repetido – mesmo produto de {DUP[x['sku']]} (que já está ao vivo)"; excl.append(x)
    elif x["lot"]=="star-wars-cyclops-2026-10-09.xlsx" and x["status"]=="Não confirmado": x["excl"]="Removido no v2 do lote Star Wars"; excl.append(x)
    else: keep.append(x)
rank=lambda s:0 if s.startswith("Ao vivo") else 1
keep.sort(key=lambda x:(rank(x["status"]),x["date"],x["title"]))
GH="https://github.com/michaeldavidsonalves-max/ebay-fotos/blob/main/"
cols=[("SKU","sku"),("Título","title"),("Preço (US$)","price"),("Condição","cond"),("Qtd","qty"),("UPC","upc"),("Status no eBay","status"),("eBay ItemID","itemid"),("Lote (planilha de origem)","lot"),("Data do lote","date"),("Nº fotos","nphotos"),("Foto de capa","cover"),("Descrição HTML","desc"),("Observação","obs")]
H=Font(bold=True,color="FFFFFF"); HF=PatternFill("solid",fgColor="0E2F4F")
def sheet(ws,data,extra=None):
    cc=cols+([("Motivo da exclusão","excl")] if extra else [])
    ws.append([c for c,_ in cc])
    for x in data:
        row=[]
        for _,k in cc:
            v=x.get(k,"")
            if k=="price" and v not in ("",None): v=float(v)
            if k=="qty" and v not in ("",None):
                try: v=int(float(v))
                except: pass
            row.append(v)
        ws.append(row); r=ws.max_row
        ws.cell(r,9).hyperlink=GH+"listings-archive/"+x["lot"]
        if str(x["desc"]).startswith("descriptions/"): ws.cell(r,13).hyperlink=GH+x["desc"]
        if x["cover"]: ws.cell(r,12).hyperlink=x["cover"]
        if x["itemid"]: ws.cell(r,8).hyperlink="https://www.ebay.com/itm/"+x["itemid"]
        ws.cell(r,7).fill=PatternFill("solid",fgColor="D9EAD3" if x["status"].startswith("Ao vivo") else "FFF2CC")
        if x["obs"].startswith("⚠"): ws.cell(r,14).fill=PatternFill("solid",fgColor="F4CCCC")
        ws.cell(r,3).number_format='#,##0.00'
    for c in ws[1]: c.font=H; c.fill=HF
    for i,w in enumerate([34,70,11,9,6,16,34,15,48,12,8,40,40,60,55],1): ws.column_dimensions[chr(64+i)].width=w
    ws.freeze_panes="C2"
    return len(cc)
wb=Workbook(); m=wb.active; m.title="Master"
n=sheet(m,keep); t=Table(displayName="Master",ref=f"A1:{chr(64+n)}{m.max_row}"); t.tableStyleInfo=TableStyleInfo(name="TableStyleLight9",showRowStripes=True); m.add_table(t)
e=wb.create_sheet("Excluídos (repetidos)"); sheet(e,excl,True)
s=wb.create_sheet("Resumo",0)
s["A1"]="ONTHEFIGURE – Master de anúncios eBay"; s["A1"].font=Font(bold=True,size=14,color="0E2F4F")
s["A2"]="1 linha por PRODUTO (sem repetidos). Selecione na aba Master o que subir; a planilha do lote (coluna I) tem a linha pronta para o eBay."
s["A2"].font=Font(italic=True,color="666666")
s.append([]); s.append(["Status","Produtos","Valor anunciado (US$)"]); [setattr(c,"font",Font(bold=True)) for c in s[4]]
for st,n in sorted(collections.Counter(x["status"] for x in keep).items(),key=lambda z:-z[1]):
    s.append([st,n,f'=SUMIFS(Master!C:C,Master!G:G,A{s.max_row+1})'])
s.append(["Total produtos únicos",f"=SUM(B5:B{s.max_row})",f"=SUM(C5:C{s.max_row})"]); [setattr(c,"font",Font(bold=True)) for c in s[s.max_row]]
for r in range(5,s.max_row+1): s.cell(r,3).number_format='#,##0'
s.append([]); s.append(["Excluídos por repetição / removidos (aba própria)",len(excl)])
s.append(["Alertas para conferir no eBay",sum(1 for x in keep if x["obs"].startswith("⚠"))])
s.append([]); s.append(["Por lote","Produtos","Ao vivo","Não confirmado"]); [setattr(c,"font",Font(bold=True)) for c in s[s.max_row]]
bl=collections.defaultdict(lambda:[0,0,0])
for x in keep:
    b=bl[(x["date"],x["lot"])]; b[0]+=1; b[1]+=x["status"].startswith("Ao vivo"); b[2]+=x["status"].startswith("Não")
for (dt,l),v in sorted(bl.items()): s.append([f"{dt}  {l}"]+v)
s.column_dimensions["A"].width=75; s.column_dimensions["B"].width=12; s.column_dimensions["C"].width=22; s.column_dimensions["D"].width=16
wb.save(f"{outdir}/MASTER-anuncios-ebay.xlsx")
with open(f"{outdir}/MASTER-anuncios-ebay.csv","w",newline="") as f:
    w=csv.writer(f); w.writerow([c for c,_ in cols]); [w.writerow([x.get(k,"") for _,k in cols]) for x in keep]
# checagem final de unicidade
import re
assert len({x["sku"] for x in keep})==len(keep)
ups=collections.Counter(x["upc"].lstrip("0") for x in keep if re.fullmatch(r"\d{8,14}",x["upc"]))
print("produtos:",len(keep),"| excluídos:",len(excl),"| status:",dict(collections.Counter(x["status"][:7] for x in keep)))
print("UPCs ainda repetidos (produtos diferentes/kit/UPC emprestado):",sum(1 for v in ups.values() if v>1))
