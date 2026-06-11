#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Genera Reporte Labos 2026 manteniendo el formato del original (5 hojas)
   en .xlsx y una versión HTML con pestañas. Agrupado por cliente."""

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import html as _html

MESES = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO","JULIO",
         "AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]
ACTIVOS = {"ENERO","FEBRERO","MARZO","ABRIL"}

# Valor unitario Avanter por mes (laboratorios viejos)
AVU = {"ENERO":74.0,"FEBRERO":74.0,"MARZO":None,"ABRIL":93.0}

def av(cod, prog, lab, fc, status, e, f, m, a, precio="Facturado vía Avanter - valor unitario por TRX"):
    """Fila Avanter: TRX ene/feb/mar/abr -> monto = trx*VU."""
    md = {}
    for mes, trx in (("ENERO",e),("FEBRERO",f),("MARZO",m),("ABRIL",a)):
        if trx is None or AVU[mes] is None:
            md[mes] = None if trx is None else (trx, None)
        else:
            md[mes] = (trx, round(trx*AVU[mes],2))
    return dict(cod=cod, prog=prog, lab=lab, propio="NO", avp="NO", fc=fc,
                precio=precio, act="Trimestral por IPC Indec", status=status, m=md)

def dr(cod, prog, lab, propio, avp, precio, act, status, m):
    return dict(cod=cod, prog=prog, lab=lab, propio=propio, avp=avp, fc="",
                precio=precio, act=act, status=status, m=m)

# ---------------- DATOS POR CLIENTE ----------------
GRUPOS = [
 ("ASTRAZENECA", [
   dr("5005","Programa Elegir Salud (AZ)","ASTRAZENECA","SI","SI",
      "Uso plataforma MAT + TRX x valor unitario","Mensual por IPC Caba","PRODUCTIVO",
      {"ENERO":(15730,3164020.00),"FEBRERO":(13198,2976652.00),
       "MARZO":(14775,3447758.00),"ABRIL":(14849,3898957.00)}),
 ]),
 ("BAYER", [
   dr("5015","Programa Más Simple (Bayer)","BAYER","SI","NO",
      "Automático por 0,50% PVP. Con mínimo","Mensual","PRODUCTIVO",
      {"ENERO":(1154,972221.92),"FEBRERO":(1090,941267.72),
       "MARZO":(1181,1044493.17),"ABRIL":(1178,1075303.74)}),
 ]),
 ("SIDUS", [
   dr("5040","Sidus Dermo (Bonos + Plataforma digital)","SIDUS DERMO","SI","SI",
      "Valor fijo por operación + plataforma + WhatsApp","Mensual por IPC Indec. Valor por TRX","PRODUCTIVO",
      {"ENERO":(2522,1282897.00),"FEBRERO":(2281,1255491.00),
       "MARZO":(2333,1330813.00),"ABRIL":(4879,2298686.15)}),
   dr("5041","Programa Convida (Sidus Farma)","SIDUS FARMA","SI","NO",
      "Automático por 1,00% PVP. Con mínimo $766.721,67","Mensual","PRODUCTIVO",
      {"ENERO":(688,766721.67),"FEBRERO":None,
       "MARZO":(615,766721.67),"ABRIL":(894,766721.67)}),
 ]),
 ("PANALAB", [
   dr("5080","Panalab Digital (Bonos + Plataforma)","PANALAB","SI","SI",
      "On line 0,6% PVP + servicio + WhatsApp","Mensual","PRODUCTIVO",
      {"ENERO":(4196,7099101.00),"FEBRERO":(3565,6836393.00),
       "MARZO":(4586,8378215.00),"ABRIL":(4586,9065584.00)}),
   dr("5081","Panalab Papel Manual","PANALAB","SI","NO",
      "Registros procesados + cajas","Mensual por IPC Caba","PRODUCTIVO",
      {"ENERO":(63271,6443666.51),"FEBRERO":(82475,8320434.77),
       "MARZO":(90524,9289060.32),"ABRIL":(103401,10793227.50)}),
 ]),
 ("COLGATE", [
   dr("5028","Bonos de Descuento Colgate","COLGATE","SI","NO",
      "Automático por 1,00% PVP","Mensual","PRODUCTIVO",
      {"ENERO":(18127,1730226.41),"FEBRERO":None,
       "MARZO":(23875,2338884.94),"ABRIL":(20351,1876522.43)}),
 ]),
 ("MAX VISION", [
   dr("5025","Programa Max Compromiso (Max Vision)","MAX VISION","SI","NO",
      "Automático por 1,00% PVP","Mensual","PRODUCTIVO",
      {"ENERO":(6309,1313290.74),"FEBRERO":None,
       "MARZO":(6019,1368036.90),"ABRIL":(6040,1378201.65)}),
 ]),
 ("BIU", [
   dr("5070","Bonos BIU Cosmeceuticals","BIU","SI","SI",
      "Valor fijo por operación","Trimestral por IPC CABA","PRODUCTIVO",
      {"ENERO":(22,1628.00),"FEBRERO":(20,1480.00),
       "MARZO":(21,1554.00),"ABRIL":(9,837.00)}),
 ]),
 ("CEODERMA", [
   dr("5120","Programa Eximia Cepage (Ceoderma)","CEODERMA","SI","NO",
      "Automático por 0,60% PVP","Mensual","PRODUCTIVO",
      {"ENERO":(769,159805.20),"FEBRERO":None,"MARZO":None,"ABRIL":None}),
 ]),
 ("HALEON / HEALTH CARE", [
   dr("5026","Bonos Descuento Haleon + Luar (Health Care)","HALEON","SI","NO",
      "Valor fijo por operación","Fijo. Se ve con Health Care Group","PRODUCTIVO",
      {"ENERO":None,"FEBRERO":(None,871875.50),"MARZO":None,"ABRIL":(None,952203.02)}),
 ]),
 ("LAZAR (nuevo 2026)", [
   dr("5300","Lazar (Fase 1 Startup)","LAZAR","SI","NO",
      "Abono startup; luego % PVP con mínimo","Mensual","STARTUP",
      {"ENERO":(None,500000.00),"FEBRERO":None,"MARZO":None,"ABRIL":(252,344079.51)}),
 ]),
 ("PROVINCIA ART (nuevo 2026)", [
   dr("5310","Provincia ART","PROVINCIA ART","SI","NO",
      "0,35% sobre monto vendido","Mensual","PRODUCTIVO",
      {"ENERO":(8166,1073509.30),"FEBRERO":None,
       "MARZO":(8692,1293685.98),"ABRIL":(8849,1471007.21)}),
 ]),
 ("SERDATA (nuevo 2026)", [
   dr("5320","Serdata (Bonos)","SERDATA","SI","NO",
      "Valor unitario por bono procesado","Mensual","PRODUCTIVO",
      {"ENERO":None,"FEBRERO":(76676,4544586.52),"MARZO":None,"ABRIL":None}),
 ]),
 ("AVANTER", [
   av("5004","Programa TEVAcuidar (TEVA)","TEVA","SI - ETICOS","PRODUCTIVO",147,135,None,150),
   av("5010","Bonos Loreal","LOREAL","SI - DINAMICAS - BONOS","PRODUCTIVO",91364,74721,None,67206),
   av("5011","Programa Beneficiarte (Ferring)","FERRING","SI - ETICOS","PRODUCTIVO",453,394,None,484),
   av("5014","Programa Rossmore","ROSSMORE","SI - ETICOS","PRODUCTIVO",1549,1437,None,1683),
   av("5020","Bonos Eucerin","BEIERSDORF","SI - DINAMICAS - BONOS","PRODUCTIVO",42157,30505,None,29074),
   av("5022","Ofertas Casasco","CASASCO","SI - DINAMICAS - BONOS","PRODUCTIVO",34587,28811,None,25273),
   av("5023","Programa Farmaspen (Aspen)","ASPEN","SI - ETICOS","PRODUCTIVO",1991,1778,None,2285),
   av("5024","Nutricia Siempre Juntos (Bagó)","BAGO","SI - ETICOS","PRODUCTIVO",387,329,None,421),
   av("5027","Programa Vari Te Acerca (Varifarma)","VARIFARMA","SI - ETICOS","PRODUCTIVO",7,6,None,8),
   av("5050","Bonos Andrómaco","ANDROMACO","SI - DINAMICAS - BONOS","PRODUCTIVO",554837,336212,None,260641),
   av("5060","Bonos Isdin","ISDIN","SI - DINAMICAS - BONOS","PRODUCTIVO",40175,25065,None,18937),
   av("5090","Bonos Galderma","GALDERMA","SI - DINAMICAS - BONOS","PRODUCTIVO",10966,7941,None,9934),
   av("5100","Bonos Bernabo","BERNABO","SI - DINAMICAS - BONOS","PRODUCTIVO",4458,4207,None,4464),
   av("5110","Bonos Eurolab","EUROLAB","SI - DINAMICAS - BONOS","PRODUCTIVO",26534,24819,None,31434),
   av("5130","Bonos Cassara","CASSARA","SI - DINAMICAS - BONOS","PRODUCTIVO",8603,7894,None,9823),
   av("5140","Bonos Bagó","BAGO","SI - DINAMICAS - BONOS","PRODUCTIVO",54581,35663,None,15759),
   av("5150","Bonos Caviahue","CAVIAHUE","SI - DINAMICAS - BONOS","PRODUCTIVO",1002,982,None,2593),
   av("5160","Bonos By Derm","BYDERM","SI - DINAMICAS - BONOS","PRODUCTIVO",5461,5116,None,6100),
   av("5170","Bonos Bioderma","BIODERMA","SI - DINAMICAS - BONOS","PRODUCTIVO",1129,1053,None,1462),
   av("5180","Bonos Biferdil","BIFERDIL","SI - DINAMICAS - BONOS","STARTUP",10,12,None,22),
   av("5190","Bonos Fresenius","FRESENIUS","SI - DINAMICAS - BONOS","STARTUP",47,38,None,107),
   av("5200","Bonos SC Johnson","SC JHONSON","SI - DINAMICAS - BONOS","STARTUP",16,20,None,6),
   av("5210","Bonos Pharmatrix","PHARMATRIX","SI - DINAMICAS - BONOS","STARTUP",85,80,None,320),
   av("5220","Bonos Megalabs","MEGALABS","SI - DINAMICAS - BONOS","STARTUP",0,0,None,None),
   av("5230","Bonos Siegfried","SIEGFRIED","SI - DINAMICAS - BONOS","STARTUP",0,None,None,None),
   av("5240","Bonos Raisse","RAISSE","SI - DINAMICAS - BONOS","STARTUP",167,356,None,None),
   av("5250","Bonos Siscom (Rayito de Sol)","SISCOM","SI - DINAMICAS - BONOS","STARTUP",0,0,None,1),
   dr("5260","Bonos ENA (alta 22/04/2026)","ENA","NO","NO",
      "Alta laboratorio nuevo - Avanter","Mensual","STARTUP",
      {"ENERO":None,"FEBRERO":None,"MARZO":None,"ABRIL":(None,547420.80)}),
   dr("5270","Bonos Convatec (alta 24/04/2026)","CONVATEC","NO","NO",
      "Alta laboratorio nuevo - Avanter","Mensual","STARTUP",
      {"ENERO":None,"FEBRERO":None,"MARZO":None,"ABRIL":(None,547420.80)}),
 ]),
]

# -------- Totales por mes (para % TRX MES) --------
tot_trx = {mes:0 for mes in MESES}
for _, labs in GRUPOS:
    for L in labs:
        for mes in MESES:
            v = L["m"].get(mes)
            if v and v[0]:
                tot_trx[mes] += v[0]
tot_anual_trx = sum(tot_trx[m] for m in ACTIVOS)

def lab_anual(L):
    trx=sum((L["m"].get(m) or (0,0))[0] or 0 for m in ACTIVOS)
    monto=sum((L["m"].get(m) or (0,0))[1] or 0 for m in ACTIVOS)
    return trx, monto

# =================== XLSX ===================
wb = openpyxl.Workbook()
ws = wb.active; ws.title="Tabla"

C_MES   = PatternFill("solid", fgColor="4D94D8")
C_HEAD  = PatternFill("solid", fgColor="A6C9EB")
C_GRUPO = PatternFill("solid", fgColor="1F4E79")
C_TOT   = PatternFill("solid", fgColor="FFE699")
C_INACT = PatternFill("solid", fgColor="F2F2F2")
WHITE = Font(bold=True, color="FFFFFF")
BOLD  = Font(bold=True)
CEN   = Alignment(horizontal="center", vertical="center", wrap_text=True)
thin  = Side(style="thin", color="D0D0D0")
BORDER= Border(left=thin,right=thin,top=thin,bottom=thin)

FMT_MONTO = '"$"\\ #,##0.00;[Red]\\-"$"\\ #,##0.00'
FMT_TRX = '#,##0'
FMT_PCT = '0.0000%'

DESC = ["Código","Programa","Laboratorio","Labortorio Propio","Avanter Proveedor",
        "FC POR AVANTER","Precio","Actualización","Status"]
NDESC = len(DESC)  # 9

# Fila 1: meses (merge de 4) + ANUAL
for i,mes in enumerate(MESES):
    c0 = NDESC + 1 + i*4
    ws.merge_cells(start_row=1,start_column=c0,end_row=1,end_column=c0+3)
    cell=ws.cell(1,c0,mes); cell.fill=C_MES; cell.font=WHITE; cell.alignment=CEN
c_anual = NDESC+1+12*4
ws.merge_cells(start_row=1,start_column=c_anual,end_row=1,end_column=c_anual+3)
cell=ws.cell(1,c_anual,"ANUAL (ene-abr)"); cell.fill=C_MES; cell.font=WHITE; cell.alignment=CEN

# Fila 2: encabezados
for j,h in enumerate(DESC):
    cell=ws.cell(2,j+1,h); cell.fill=C_HEAD; cell.font=BOLD; cell.alignment=CEN
sub=["TRX","% TRX MES","MONTO FC","MONTO X TRX"]
for blk in range(13):
    for k,h in enumerate(sub):
        cell=ws.cell(2,NDESC+1+blk*4+k,h); cell.fill=C_HEAD; cell.font=BOLD; cell.alignment=CEN

r=3
TOTAL_W = NDESC+13*4
def write_lab(L):
    global r
    vals=[L["cod"],L["prog"],L["lab"],L["propio"],L["avp"],L["fc"],L["precio"],L["act"],L["status"]]
    for j,v in enumerate(vals):
        c=ws.cell(r,j+1,v)
        if j in (0,3,4,8): c.alignment=Alignment(horizontal="center")
    # meses (solo escribimos celdas con valor)
    for i,mes in enumerate(MESES):
        c0=NDESC+1+i*4
        v=L["m"].get(mes)
        trx=monto=None
        if v: trx,monto=v
        if trx is not None:
            cT=ws.cell(r,c0,trx); cT.number_format=FMT_TRX
            if tot_trx[mes]:
                cP=ws.cell(r,c0+1,trx/tot_trx[mes]); cP.number_format=FMT_PCT
        if monto is not None:
            cM=ws.cell(r,c0+2,monto); cM.number_format=FMT_MONTO
            if trx:
                cX=ws.cell(r,c0+3,monto/trx); cX.number_format='"$"\\ #,##0.00'
    # anual
    atrx,amonto=lab_anual(L)
    c0=c_anual
    if atrx:
        cT=ws.cell(r,c0,atrx); cT.number_format=FMT_TRX; cT.font=BOLD
        if tot_anual_trx:
            cP=ws.cell(r,c0+1,atrx/tot_anual_trx); cP.number_format=FMT_PCT
    if amonto:
        cM=ws.cell(r,c0+2,round(amonto,2)); cM.number_format=FMT_MONTO; cM.font=BOLD
        if atrx:
            cX=ws.cell(r,c0+3,amonto/atrx); cX.number_format='"$"\\ #,##0.00'
    r+=1

for cliente,labs in GRUPOS:
    ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=TOTAL_W)
    c=ws.cell(r,1,f"CLIENTE: {cliente}"); c.fill=C_GRUPO; c.font=WHITE
    c.alignment=Alignment(horizontal="left",vertical="center"); r+=1
    for L in labs: write_lab(L)

# Fila TOTAL general
ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=NDESC)
c=ws.cell(r,1,"TOTAL GENERAL"); c.fill=C_TOT; c.font=BOLD
for i,mes in enumerate(MESES):
    c0=NDESC+1+i*4
    cT=ws.cell(r,c0); cT.fill=C_TOT; cT.font=BOLD; cT.number_format=FMT_TRX
    if mes in ACTIVOS and tot_trx[mes]: cT.value=tot_trx[mes]
    mtot=sum((L["m"].get(mes) or (0,0))[1] or 0 for _,labs in GRUPOS for L in labs)
    cM=ws.cell(r,c0+2); cM.fill=C_TOT; cM.font=BOLD; cM.number_format=FMT_MONTO
    if mes in ACTIVOS and mtot: cM.value=mtot
    ws.cell(r,c0+1).fill=C_TOT; ws.cell(r,c0+3).fill=C_TOT
c0=c_anual
ws.cell(r,c0,tot_anual_trx).number_format=FMT_TRX; ws.cell(r,c0).fill=C_TOT; ws.cell(r,c0).font=BOLD
amt=sum(lab_anual(L)[1] for _,labs in GRUPOS for L in labs)
ws.cell(r,c0+2,round(amt,2)).number_format=FMT_MONTO; ws.cell(r,c0+2).fill=C_TOT; ws.cell(r,c0+2).font=BOLD
ws.cell(r,c0+1).fill=C_TOT; ws.cell(r,c0+3).fill=C_TOT

# anchos y freeze
widths={1:10,2:34,3:16,4:14,5:14,6:22,7:34,8:26,9:13}
for col,w in widths.items(): ws.column_dimensions[get_column_letter(col)].width=w
for col in range(NDESC+1,TOTAL_W+1): ws.column_dimensions[get_column_letter(col)].width=13
ws.freeze_panes="D3"
ws.row_dimensions[2].height=30

# ---- ANALISI ----
wa=wb.create_sheet("ANALISI")
wa["A1"]="ANÁLISIS 2026 (ene-abr)"; wa["A1"].font=Font(bold=True,size=13)
hdr=["Métrica"]+[m.capitalize() for m in ["ENERO","FEBRERO","MARZO","ABRIL"]]+["Prom/Total"]
for j,h in enumerate(hdr):
    c=wa.cell(3,j+1,h); c.fill=C_HEAD; c.font=BOLD; c.alignment=CEN
def prom_dir():
    rows=[]
    # promedio $/TRX directos y avanter por mes
    for nombre,filt in [("Valor promedio $/TRX - Clientes directos",lambda cl: cl!="AVANTER"),
                        ("Valor promedio $/TRX - Avanter",lambda cl: cl=="AVANTER")]:
        fila=[nombre]; tot_t=tot_m=0
        for mes in ["ENERO","FEBRERO","MARZO","ABRIL"]:
            t=m=0
            for cl,labs in GRUPOS:
                if not filt(cl): continue
                for L in labs:
                    v=L["m"].get(mes)
                    if v and v[0] and v[1]: t+=v[0]; m+=v[1]
            fila.append(round(m/t,2) if t else None); tot_t+=t; tot_m+=m
        fila.append(round(tot_m/tot_t,2) if tot_t else None)
        rows.append(fila)
    return rows
r2=4
for fila in prom_dir():
    for j,v in enumerate(fila):
        c=wa.cell(r2,j+1,v)
        if j>=1 and isinstance(v,(int,float)): c.number_format='"$"\\ #,##0.00'
    r2+=1
# montos totales por cliente
wa.cell(r2+1,1,"FACTURACIÓN TOTAL POR CLIENTE (ene-abr)").font=BOLD; r2+=2
wa.cell(r2,1,"Cliente").fill=C_HEAD; wa.cell(r2,1).font=BOLD
wa.cell(r2,2,"TRX acum").fill=C_HEAD; wa.cell(r2,2).font=BOLD
wa.cell(r2,3,"Monto FC acum").fill=C_HEAD; wa.cell(r2,3).font=BOLD; r2+=1
for cl,labs in GRUPOS:
    t=sum(lab_anual(L)[0] for L in labs); m=sum(lab_anual(L)[1] for L in labs)
    wa.cell(r2,1,cl); wa.cell(r2,2,t).number_format=FMT_TRX
    wa.cell(r2,3,round(m,2)).number_format=FMT_MONTO; r2+=1
for col,w in {1:42,2:14,3:20,4:16,5:16,6:16}.items(): wa.column_dimensions[get_column_letter(col)].width=w

# ---- Hoja 2 (parámetros) ----
h2=wb.create_sheet("Hoja 2")
h2["A1"]="Parámetros (heredado 2025 - ajustar si cambió)"; h2["A1"].font=BOLD
params=[("Descuento",60,5),("Set up",1000000,""),("Bonos",150,""),("Dinámica",135,"")]
for i,(k,v1,v2) in enumerate(params):
    h2.cell(3+i,1,k); h2.cell(3+i,2,v1); h2.cell(3+i,3,v2)
h2.column_dimensions["A"].width=18

# ---- Hoja 1 (promedios anuales) ----
h1=wb.create_sheet("Hoja 1")
prom=prom_dir()
labels=["Valor promedio tx directa","Valor promedio tx avanter"]
h1["A1"]="Valores promedio 2026 (ene-abr)"; h1["A1"].font=BOLD
for i,fila in enumerate(prom):
    h1.cell(3+i,1,fila[0]); h1.cell(3+i,2,fila[-1]).number_format='"$"\\ #,##0.00'
h1.column_dimensions["A"].width=42

# ---- Hoja2 (pivot cli_id -> TRX acum) ----
hp=wb.create_sheet("Hoja2")
hp.cell(1,1,"cli_id").font=BOLD; hp.cell(1,2,"(TRX acum ene-abr)").font=BOLD
piv=[]
for cl,labs in GRUPOS:
    for L in labs:
        t=lab_anual(L)[0]
        piv.append((L["cod"],t))
piv.sort(key=lambda x:-x[1])
for i,(cod,t) in enumerate(piv):
    hp.cell(2+i,1,cod); hp.cell(2+i,2,t).number_format=FMT_TRX
hp.column_dimensions["A"].width=12; hp.column_dimensions["B"].width=20

wb.save("Reporte_Labos_2026.xlsx")
print("XLSX OK -> Reporte_Labos_2026.xlsx  | filas Tabla:",r)
print("Totales TRX por mes:",{m:tot_trx[m] for m in ACTIVOS})

# =================== HTML ===================
def faro(x, dec=2):
    if x is None: return ""
    s=f"{x:,.{dec}f}"
    return s.replace(",","§").replace(".",",").replace("§",".")
def money(x):
    if x is None: return ""
    return "$ "+faro(x,2)
def pct(x):
    if x is None: return ""
    return faro(x*100,4)+"%"

month_th="".join(f'<th colspan="4" class="mes {"act" if m in ACTIVOS else "inact"}">{m}</th>' for m in MESES)
month_th+='<th colspan="4" class="mes anual">ANUAL (ene-abr)</th>'
sub_th=""
for m in MESES+["ANUAL"]:
    cls="act" if (m in ACTIVOS or m=="ANUAL") else "inact"
    for s in ["TRX","% TRX","MONTO FC","$/TRX"]:
        sub_th+=f'<th class="sub {cls}">{s}</th>'

rows_html=""
for cliente,labs in GRUPOS:
    rows_html+=f'<tr class="grp"><td colspan="{9+13*4}">CLIENTE: {cliente}</td></tr>'
    for L in labs:
        tds=f'<td class="cod">{L["cod"]}</td><td class="l">{_html.escape(L["prog"])}</td><td class="l">{_html.escape(L["lab"])}</td>'
        tds+=f'<td class="c">{L["propio"]}</td><td class="c">{L["avp"]}</td><td class="l sm">{_html.escape(L["fc"])}</td>'
        tds+=f'<td class="l sm">{_html.escape(L["precio"])}</td><td class="l sm">{_html.escape(L["act"])}</td><td class="c sm">{L["status"]}</td>'
        for m in MESES:
            v=L["m"].get(m); cls="" if m in ACTIVOS else "inact"
            trx=monto=None
            if v: trx,monto=v
            p = (trx/tot_trx[m]) if (trx and tot_trx[m]) else None
            x = (monto/trx) if (trx and monto) else None
            tds+=f'<td class="n {cls}">{faro(trx,0) if trx is not None else ""}</td>'
            tds+=f'<td class="n {cls}">{pct(p)}</td>'
            tds+=f'<td class="n {cls}">{money(monto)}</td>'
            tds+=f'<td class="n {cls}">{money(x)}</td>'
        atrx,amonto=lab_anual(L)
        ap=atrx/tot_anual_trx if (atrx and tot_anual_trx) else None
        ax=amonto/atrx if (atrx and amonto) else None
        tds+=f'<td class="n b">{faro(atrx,0) if atrx else ""}</td><td class="n b">{pct(ap)}</td><td class="n b">{money(amonto)}</td><td class="n b">{money(ax)}</td>'
        rows_html+=f"<tr>{tds}</tr>"
# total general
gt=f'<tr class="tot"><td colspan="9">TOTAL GENERAL</td>'
for m in MESES:
    mt=sum((L["m"].get(m) or (0,0))[1] or 0 for _,labs in GRUPOS for L in labs) if m in ACTIVOS else None
    tt=tot_trx[m] if m in ACTIVOS else None
    gt+=f'<td class="n">{faro(tt,0) if tt else ""}</td><td></td><td class="n">{money(mt) if mt else ""}</td><td></td>'
amt=sum(lab_anual(L)[1] for _,labs in GRUPOS for L in labs)
gt+=f'<td class="n">{faro(tot_anual_trx,0)}</td><td></td><td class="n">{money(amt)}</td><td></td></tr>'
rows_html+=gt

# ANALISI html
an_rows=""
for fila in prom_dir():
    an_rows+="<tr><td class='l'>"+fila[0]+"</td>"+"".join(f"<td class='n'>{money(v)}</td>" for v in fila[1:])+"</tr>"
cli_rows=""
for cl,labs in GRUPOS:
    t=sum(lab_anual(L)[0] for L in labs); m=sum(lab_anual(L)[1] for L in labs)
    cli_rows+=f"<tr><td class='l'>{cl}</td><td class='n'>{faro(t,0)}</td><td class='n'>{money(m)}</td></tr>"
piv_rows="".join(f"<tr><td class='c'>{cod}</td><td class='n'>{faro(t,0)}</td></tr>" for cod,t in piv)

HTML=f"""<!DOCTYPE html><html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Reporte Labos 2026 - por Cliente</title>
<style>
*{{box-sizing:border-box}}
body{{font-family:'Segoe UI',Roboto,Arial,sans-serif;margin:0;background:#eef2f7;color:#1c2b3a}}
header{{background:#1F4E79;color:#fff;padding:14px 22px}}
header h1{{margin:0;font-size:19px}}
header p{{margin:4px 0 0;font-size:12.5px;opacity:.85}}
.tabs{{display:flex;gap:4px;background:#15395c;padding:0 14px}}
.tabs button{{background:none;border:0;color:#cfe0f0;padding:11px 18px;cursor:pointer;font-size:13.5px;border-bottom:3px solid transparent}}
.tabs button.on{{color:#fff;border-bottom-color:#4D94D8;font-weight:600}}
.pane{{display:none;padding:14px}}
.pane.on{{display:block}}
.scroll{{overflow:auto;max-height:80vh;border:1px solid #cdd8e3;border-radius:6px;background:#fff}}
table{{border-collapse:collapse;font-size:11.5px;white-space:nowrap}}
th,td{{border:1px solid #dce4ec;padding:4px 7px}}
thead th{{position:sticky;top:0;background:#A6C9EB;color:#16314a;font-weight:700;text-align:center;z-index:3}}
thead tr:first-child th{{top:0}}
thead tr:nth-child(2) th{{top:27px}}
th.mes{{background:#4D94D8;color:#fff}}
th.mes.anual{{background:#1F4E79}}
.inact{{background:#f2f4f7 !important;color:#aeb8c2 !important}}
td.l{{text-align:left;white-space:normal;max-width:230px}}
td.sm{{font-size:10px;color:#566}}
td.c{{text-align:center}}
td.cod{{text-align:center;font-weight:600}}
td.n{{text-align:right;font-variant-numeric:tabular-nums}}
td.b{{font-weight:700;background:#fbfdff}}
tr.grp td{{background:#1F4E79;color:#fff;font-weight:700;text-align:left;position:sticky;left:0}}
tr.tot td{{background:#FFE699;font-weight:700}}
tbody tr:hover td:not(.inact){{background:#eaf3fc}}
.note{{font-size:12px;color:#5a6b7b;margin:6px 2px 12px;line-height:1.5}}
.card{{background:#fff;border:1px solid #cdd8e3;border-radius:6px;padding:16px;max-width:760px;margin-bottom:14px}}
.card h3{{margin:0 0 10px;color:#1F4E79}}
.card table{{white-space:normal}}
</style></head><body>
<header><h1>Reporte Labos 2026 — Seguimiento por Cliente</h1>
<p>Datos ene–abr 2026 · Montos en ARS sin IVA · TRX = transacciones/envases aprobados del mes · Fuente: carpeta PROFORMAS (Drive)</p></header>
<div class="tabs">
<button class="on" onclick="sel(0)">Tabla</button>
<button onclick="sel(1)">ANALISI</button>
<button onclick="sel(2)">Hoja 2</button>
<button onclick="sel(3)">Hoja 1</button>
<button onclick="sel(4)">Hoja2</button>
</div>

<div class="pane on" id="p0">
<div class="note">⚠️ <b>Avanter marzo</b> no tiene desglose por laboratorio en los archivos compartidos (s/d). <b>Febrero</b> de Colgate, MaxVision, Sidus Farma y Ceoderma sin archivo. Mayo–Diciembre quedan vacíos para completar a lo largo del año.</div>
<div class="scroll"><table>
<thead><tr><th rowspan="2">Código</th><th rowspan="2">Programa</th><th rowspan="2">Laboratorio</th><th rowspan="2">Lab.<br>Propio</th><th rowspan="2">Avanter<br>Proveedor</th><th rowspan="2">FC POR<br>AVANTER</th><th rowspan="2">Precio</th><th rowspan="2">Actualización</th><th rowspan="2">Status</th>{month_th}</tr>
<tr>{sub_th}</tr></thead>
<tbody>{rows_html}</tbody>
</table></div></div>

<div class="pane" id="p1">
<div class="card"><h3>Valores promedio $/TRX</h3><table>
<thead><tr><th>Métrica</th><th>Enero</th><th>Febrero</th><th>Marzo</th><th>Abril</th><th>Prom</th></tr></thead>
<tbody>{an_rows}</tbody></table></div>
<div class="card"><h3>Facturación total por cliente (ene–abr)</h3><table>
<thead><tr><th>Cliente</th><th>TRX acum</th><th>Monto FC acum</th></tr></thead>
<tbody>{cli_rows}</tbody></table></div>
</div>

<div class="pane" id="p2"><div class="card"><h3>Hoja 2 — Parámetros (heredado 2025)</h3><table>
<tbody><tr><td class='l'>Descuento</td><td class='n'>60</td><td class='n'>5</td></tr>
<tr><td class='l'>Set up</td><td class='n'>{faro(1000000,0)}</td><td></td></tr>
<tr><td class='l'>Bonos</td><td class='n'>150</td><td></td></tr>
<tr><td class='l'>Dinámica</td><td class='n'>135</td><td></td></tr></tbody></table>
<p class="note">Parámetros heredados del archivo 2025; ajustar si cambiaron en 2026.</p></div></div>

<div class="pane" id="p3"><div class="card"><h3>Hoja 1 — Valores promedio anuales 2026</h3><table>
<tbody>{''.join(f"<tr><td class='l'>{f[0]}</td><td class='n'>{money(f[-1])}</td></tr>" for f in prom_dir())}</tbody></table></div></div>

<div class="pane" id="p4"><div class="card"><h3>Hoja2 — TRX acumuladas por cli_id (ene–abr)</h3><table>
<thead><tr><th>cli_id</th><th>TRX acum</th></tr></thead><tbody>{piv_rows}</tbody></table></div></div>

<script>
function sel(i){{
 document.querySelectorAll('.tabs button').forEach((b,k)=>b.classList.toggle('on',k===i));
 document.querySelectorAll('.pane').forEach((p,k)=>p.classList.toggle('on',k===i));
}}
</script>
</body></html>"""
open("Reporte_Labos_2026.html","w",encoding="utf-8").write(HTML)
print("HTML OK -> Reporte_Labos_2026.html")
