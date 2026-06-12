#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Reporte Labos 2026 (ene-abr): genera XLSX (5 hojas), HTML por secciones
   y CSV para Google Sheet nativo. Agrupado por cliente.
   Fuente: soportes de facturación de la carpeta PROFORMAS (Drive)."""

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import html as _html

MESES = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO","JULIO",
         "AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]
ACTIVOS = {"ENERO","FEBRERO","MARZO","ABRIL"}
M4 = ["ENERO","FEBRERO","MARZO","ABRIL"]

# Valor unitario Avanter por mes (laboratorios viejos)
AVU = {"ENERO":74.0,"FEBRERO":74.0,"MARZO":74.0,"ABRIL":93.0}

# Totales consolidados Avanter según archivos oficiales "Facturación Avanter <mes>"
# ABRIL incluye altas ENA + Convatec ($547.420,80 c/u).
AV_CONSOL = {"ENERO":(881068,65199032.00),"FEBRERO":(587574,43480476.00),
             "MARZO":(532345,39393530.00),"ABRIL":(488187,46496232.59)}

def av(cod, prog, lab, fc, status, e, f, m, a, precio="Facturado vía Avanter - valor unitario por TRX"):
    """Fila Avanter: TRX ene/feb/mar/abr -> monto = trx*VU del mes."""
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
   dr("5005","Programa Elegir Salud (AZ)","ASTRAZENECA","SI","NO",
      "Presupuesto Elegir Salud: operación + call center + cápitas + auditorías + supervisor CS","Mensual por IPC CABA","PRODUCTIVO",
      {"ENERO":(15730,20929743.55),"FEBRERO":(13198,19689476.25),
       "MARZO":(14775,21612629.37),"ABRIL":(14849,21786516.13)}),
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
 # NOTA: Provincia ART (convenio, no laboratorio) y Serdata (proveedor, no cliente)
 # quedan EXCLUIDOS de este reporte de laboratorios por definición del negocio.
 ("AVANTER", [
   av("5004","Programa TEVAcuidar (TEVA)","TEVA","SI - ETICOS","PRODUCTIVO",147,135,172,150),
   av("5010","Bonos Loreal","LOREAL","SI - DINAMICAS - BONOS","PRODUCTIVO",91364,74721,75325,67206),
   av("5011","Programa Beneficiarte (Ferring)","FERRING","SI - ETICOS","PRODUCTIVO",453,394,518,484),
   av("5014","Programa Rossmore","ROSSMORE","SI - ETICOS","PRODUCTIVO",1549,1437,1708,1683),
   av("5020","Bonos Eucerin","BEIERSDORF","SI - DINAMICAS - BONOS","PRODUCTIVO",42157,30505,34604,29074),
   av("5022","Ofertas Casasco","CASASCO","SI - DINAMICAS - BONOS","PRODUCTIVO",34587,28811,28256,25273),
   av("5023","Programa Farmaspen (Aspen)","ASPEN","SI - ETICOS","PRODUCTIVO",1991,1778,2097,2285),
   av("5024","Nutricia Siempre Juntos (Bagó)","BAGO","SI - ETICOS","PRODUCTIVO",387,329,383,421),
   av("5027","Programa Vari Te Acerca (Varifarma)","VARIFARMA","SI - ETICOS","PRODUCTIVO",7,6,8,8),
   av("5050","Bonos Andrómaco","ANDROMACO","SI - DINAMICAS - BONOS","PRODUCTIVO",554837,336212,274190,260641),
   av("5060","Bonos Isdin","ISDIN","SI - DINAMICAS - BONOS","PRODUCTIVO",40175,25065,21680,18937),
   av("5090","Bonos Galderma","GALDERMA","SI - DINAMICAS - BONOS","PRODUCTIVO",10966,7941,10088,9934),
   av("5100","Bonos Bernabo","BERNABO","SI - DINAMICAS - BONOS","PRODUCTIVO",4458,4207,4609,4464),
   av("5110","Bonos Eurolab","EUROLAB","SI - DINAMICAS - BONOS","PRODUCTIVO",26534,24819,30701,31434),
   # Ceoderma operó vía Avanter en enero (355 TRX en el archivo oficial de Avanter)
   av("5120","Eximia Cepage (Ceoderma) - vía Avanter","CEODERMA","SI - ETICOS","PRODUCTIVO",355,None,None,None),
   av("5130","Bonos Cassara","CASSARA","SI - DINAMICAS - BONOS","PRODUCTIVO",8603,7894,9727,9823),
   av("5140","Bonos Bagó","BAGO","SI - DINAMICAS - BONOS","PRODUCTIVO",54581,35663,29614,15759),
   av("5150","Bonos Caviahue","CAVIAHUE","SI - DINAMICAS - BONOS","PRODUCTIVO",1002,982,804,2593),
   av("5160","Bonos By Derm","BYDERM","SI - DINAMICAS - BONOS","PRODUCTIVO",5461,5116,6240,6100),
   av("5170","Bonos Bioderma","BIODERMA","SI - DINAMICAS - BONOS","PRODUCTIVO",1129,1053,1289,1462),
   av("5180","Bonos Biferdil","BIFERDIL","SI - DINAMICAS - BONOS","STARTUP",10,12,14,22),
   av("5190","Bonos Fresenius","FRESENIUS","SI - DINAMICAS - BONOS","STARTUP",47,38,46,107),
   av("5200","Bonos SC Johnson","SC JHONSON","SI - DINAMICAS - BONOS","STARTUP",16,20,24,6),
   av("5210","Bonos Pharmatrix","PHARMATRIX","SI - DINAMICAS - BONOS","STARTUP",85,80,248,320),
   av("5220","Bonos Megalabs","MEGALABS","SI - DINAMICAS - BONOS","STARTUP",0,0,0,None),
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
ws = wb.active; ws.title="Tabla por Cliente"

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
NDESC = len(DESC)

for i,mes in enumerate(MESES):
    c0 = NDESC + 1 + i*4
    ws.merge_cells(start_row=1,start_column=c0,end_row=1,end_column=c0+3)
    cell=ws.cell(1,c0,mes); cell.fill=C_MES; cell.font=WHITE; cell.alignment=CEN
c_anual = NDESC+1+12*4
ws.merge_cells(start_row=1,start_column=c_anual,end_row=1,end_column=c_anual+3)
cell=ws.cell(1,c_anual,"ANUAL (ene-abr)"); cell.fill=C_MES; cell.font=WHITE; cell.alignment=CEN

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

widths={1:10,2:34,3:16,4:14,5:14,6:22,7:34,8:26,9:13}
for col,w in widths.items(): ws.column_dimensions[get_column_letter(col)].width=w
for col in range(NDESC+1,TOTAL_W+1): ws.column_dimensions[get_column_letter(col)].width=13
ws.freeze_panes="D3"
ws.row_dimensions[2].height=30

# ---- Análisis ----
wa=wb.create_sheet("Análisis")
wa["A1"]="ANÁLISIS 2026 (ene-abr)"; wa["A1"].font=Font(bold=True,size=13)
hdr=["Métrica"]+[m.capitalize() for m in M4]+["Prom/Total"]
for j,h in enumerate(hdr):
    c=wa.cell(3,j+1,h); c.fill=C_HEAD; c.font=BOLD; c.alignment=CEN
def prom_dir():
    rows=[]
    for nombre,filt in [("Valor promedio $/TRX - Clientes directos",lambda cl: cl!="AVANTER"),
                        ("Valor promedio $/TRX - Avanter",lambda cl: cl=="AVANTER")]:
        fila=[nombre]; tot_t=tot_m=0
        for mes in M4:
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
wa.cell(r2+1,1,"FACTURACIÓN TOTAL POR CLIENTE (ene-abr)").font=BOLD; r2+=2
wa.cell(r2,1,"Cliente").fill=C_HEAD; wa.cell(r2,1).font=BOLD
wa.cell(r2,2,"TRX acum").fill=C_HEAD; wa.cell(r2,2).font=BOLD
wa.cell(r2,3,"Monto FC acum").fill=C_HEAD; wa.cell(r2,3).font=BOLD; r2+=1
for cl,labs in GRUPOS:
    t=sum(lab_anual(L)[0] for L in labs); m=sum(lab_anual(L)[1] for L in labs)
    wa.cell(r2,1,cl); wa.cell(r2,2,t).number_format=FMT_TRX
    wa.cell(r2,3,round(m,2)).number_format=FMT_MONTO; r2+=1
for col,w in {1:42,2:14,3:20,4:16,5:16,6:16}.items(): wa.column_dimensions[get_column_letter(col)].width=w

# ---- Parámetros ----
h2=wb.create_sheet("Parámetros")
h2["A1"]="Parámetros de facturación 2026"; h2["A1"].font=BOLD
params=[("VU Avanter Enero/Febrero",74,""),("VU Avanter Abril",93,""),
        ("Alta lab nuevo Avanter (base dic-25)",500000,""),
        ("Alta lab nuevo Avanter (marzo, IPC acum)",547420.80,""),
        ("Mínimo Sidus Farma",766721.67,""),("Mínimo Bayer (mar-abr)",844447.62,"")]
for i,(k,v1,v2) in enumerate(params):
    h2.cell(3+i,1,k); h2.cell(3+i,2,v1); h2.cell(3+i,3,v2)
h2.column_dimensions["A"].width=40

# ---- Promedios ----
h1=wb.create_sheet("Promedios")
prom=prom_dir()
h1["A1"]="Valores promedio 2026 (ene-abr)"; h1["A1"].font=BOLD
for i,fila in enumerate(prom):
    h1.cell(3+i,1,fila[0]); h1.cell(3+i,2,fila[-1]).number_format='"$"\\ #,##0.00'
h1.column_dimensions["A"].width=42

# ---- TRX por Código ----
hp=wb.create_sheet("TRX por Código")
hp.cell(1,1,"cli_id").font=BOLD; hp.cell(1,2,"TRX acum ene-abr").font=BOLD
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
print("XLSX OK -> Reporte_Labos_2026.xlsx | filas Tabla:",r)

# =================== DATOS DERIVADOS PARA HTML ===================
DIRECT_GROUPS=[(c,l) for c,l in GRUPOS if c!="AVANTER"]
AV_LABS=dict(GRUPOS)["AVANTER"]

def disp(c):
    if " (nuevo 2026)" in c: return c.replace(" (nuevo 2026)",""), True
    return c, False

def dir_tot(mes):
    t=m=0
    for _,labs in DIRECT_GROUPS:
        for L in labs:
            v=L["m"].get(mes)
            if v:
                if v[0]: t+=v[0]
                if v[1]: m+=v[1]
    return t,m

def grupo_mes(labs,mes):
    t=m=0; has=False
    for L in labs:
        v=L["m"].get(mes)
        if v:
            has=True
            if v[0]: t+=v[0]
            if v[1]: m+=v[1]
    return (t,m) if has else None

month_rows=[]
for mes in M4:
    dt,dm=dir_tot(mes)
    avc=AV_CONSOL[mes]
    at,am=(avc if avc else (None,None))
    month_rows.append(dict(mes=mes,dt=dt,dm=dm,at=at,am=am,
                           tt=dt+(at or 0), tm=dm+(am or 0)))
GT_TRX=sum(rr["tt"] for rr in month_rows)
GT_FC=sum(rr["tm"] for rr in month_rows)
DIR_FC=sum(rr["dm"] for rr in month_rows)
DIR_TRX=sum(rr["dt"] for rr in month_rows)
AV_T=sum(v[0] for v in AV_CONSOL.values() if v)
AV_M=sum(v[1] for v in AV_CONSOL.values() if v)

CLITOT=[]
for c,labs in DIRECT_GROUPS:
    t=sum(lab_anual(L)[0] for L in labs); m=sum(lab_anual(L)[1] for L in labs)
    CLITOT.append((c,t,m,labs))
CLITOT.append(("AVANTER",AV_T,AV_M,None))
CLITOT.sort(key=lambda x:-x[2])

# ---- helpers de formato (AR) ----
def faro(x, dec=2):
    if x is None: return ""
    s=f"{x:,.{dec}f}"
    return s.replace(",","§").replace(".",",").replace("§",".")
def m0(x): return ("$ "+faro(x,0)) if x is not None else ""
def m2(x): return ("$ "+faro(x,2)) if x is not None else ""
def n0(x): return faro(x,0) if x is not None else ""
def MM(x): return "$ "+faro(x/1e6,1)+" M"
def pctf(x,dec=1): return (faro(x*100,dec)+"%") if x is not None else ""
SD='<span class="sd">s/d</span>'

# =================== HTML ===================
MESN={"ENERO":"Enero","FEBRERO":"Febrero","MARZO":"Marzo","ABRIL":"Abril"}

# ---- Resumen: KPIs ----
n_dir=len(DIRECT_GROUPS)
av_activos=sum(1 for L in AV_LABS if L["cod"] not in ("5260","5270") and lab_anual(L)[0]>0)
mejor=max(month_rows,key=lambda rr:rr["tm"])
kpis=f"""
<div class="kpis">
 <div class="kpi"><div class="kpi-t">Facturación ene–abr</div><div class="kpi-v">{MM(GT_FC)}</div>
   <div class="kpi-s">{m2(GT_FC)} · 4 meses completos</div></div>
 <div class="kpi"><div class="kpi-t">Transacciones informadas</div><div class="kpi-v">{n0(GT_TRX)}</div>
   <div class="kpi-s">TRX aprobadas ene–abr · directos {n0(DIR_TRX)} + Avanter {n0(AV_T)}</div></div>
 <div class="kpi"><div class="kpi-t">Mejor mes</div><div class="kpi-v">{MESN[mejor["mes"]]}</div>
   <div class="kpi-s">{m0(mejor["tm"])} facturados · {n0(mejor["tt"])} TRX</div></div>
 <div class="kpi"><div class="kpi-t">Cartera</div><div class="kpi-v">{n_dir} + Avanter</div>
   <div class="kpi-s">{n_dir} clientes directos · {av_activos} laboratorios activos vía Avanter · Nuevos 2026: Lazar + altas ENA y Convatec (Avanter)</div></div>
</div>"""

# ---- Resumen: reporte mensual ----
mt_rows=""
for rr in month_rows:
    am_c = m0(rr["am"]) if rr["am"] is not None else SD
    at_c = n0(rr["at"]) if rr["at"] is not None else SD
    xtrx = (rr["tm"]/rr["tt"]) if rr["tt"] else None
    nota = ' <span class="flag">incompleto</span>' if rr["am"] is None else ""
    mt_rows+=f"""<tr><td class="l"><b>{MESN[rr["mes"]]}</b>{nota}</td>
      <td class="n">{n0(rr["dt"])}</td><td class="n">{m0(rr["dm"])}</td>
      <td class="n">{at_c}</td><td class="n">{am_c}</td>
      <td class="n b">{n0(rr["tt"])}</td><td class="n b">{m0(rr["tm"])}</td>
      <td class="n">{m2(xtrx)}</td></tr>"""
mt_rows+=f"""<tr class="tot"><td class="l">TOTAL ene–abr</td>
  <td class="n">{n0(DIR_TRX)}</td><td class="n">{m0(DIR_FC)}</td>
  <td class="n">{n0(AV_T)}</td><td class="n">{m0(AV_M)}</td>
  <td class="n">{n0(GT_TRX)}</td><td class="n">{m0(GT_FC)}</td>
  <td class="n">{m2(GT_FC/GT_TRX)}</td></tr>"""

# ---- Resumen: gráfico mensual ----
maxtm=max(rr["tm"] for rr in month_rows)
chart=""
for rr in month_rows:
    wd=rr["dm"]/maxtm*100
    if rr["am"] is not None:
        wa_=rr["am"]/maxtm*100
        segs=f'<i class="seg d" style="width:{wd:.1f}%"></i><i class="seg a" style="width:{wa_:.1f}%"></i>'
        nota=""
    else:
        segs=f'<i class="seg d" style="width:{wd:.1f}%"></i><i class="seg x" style="width:22%"></i>'
        nota=' <small class="warn">Avanter s/d</small>'
    chart+=f'<div class="brow"><span class="blab">{MESN[rr["mes"]]}</span><div class="btrack">{segs}</div><span class="bval">{MM(rr["tm"])}{nota}</span></div>'
chart+="""<div class="legend"><span><i class="dot d"></i>Clientes directos</span>
<span><i class="dot a"></i>Avanter</span></div>"""

# ---- Resumen: top clientes ----
maxm=CLITOT[0][2]
top=""
for i,(c,t,m,_l) in enumerate(CLITOT,1):
    name,nuevo=disp(c)
    badge=' <span class="chip chip-new">NUEVO</span>' if nuevo else ""
    top+=f"""<tr><td class="c">{i}</td><td class="l"><b>{name}</b>{badge}</td>
     <td class="n">{m0(m)}</td><td class="n">{pctf(m/GT_FC)}</td>
     <td class="w"><div class="minibar"><i style="width:{m/maxm*100:.1f}%"></i></div></td>
     <td class="n">{n0(t) if t else "–"}</td></tr>"""

alertas="""
<div class="alert">
 <b>⚠️ Datos pendientes que afectan los totales</b>
 <ul>
  <li><b>Febrero:</b> sin soporte de Colgate, Max Vision, Sidus Farma y Ceoderma.</li>
  <li><b>Ceoderma:</b> solo enero · <b>Health Care:</b> montos sin TRX y enero pendiente · <b>Lazar:</b> sin feb–mar.</li>
 </ul>
 <small>✅ Avanter marzo ya incorporado (532.345 TRX × $74 = $ 39.393.530). El detalle completo está en «Fuentes y Datos».</small>
</div>"""

# ---- Clientes Directos: tabla detalle ----
NCOLS=4+4*3+4
dir_rows=""
for cliente,labs in DIRECT_GROUPS:
    name,nuevo=disp(cliente)
    badge=' <span class="chip chip-new">NUEVO 2026</span>' if nuevo else ""
    dir_rows+=f'<tr class="grp"><td colspan="{NCOLS}">{name}{badge}</td></tr>'
    for L in labs:
        st=L["status"]
        chip=f'<span class="chip {"chip-prod" if st=="PRODUCTIVO" else "chip-start"}">{st}</span>'
        row=f'<td class="cod">{L["cod"]}</td><td class="l"><b>{_html.escape(L["prog"])}</b><br><small>{_html.escape(L["lab"])}</small></td>'
        row+=f'<td class="l sm">{_html.escape(L["precio"])}<br><small>{_html.escape(L["act"])}</small></td><td class="c">{chip}</td>'
        for mes in M4:
            v=L["m"].get(mes)
            if v is None:
                row+=f'<td class="n">{SD}</td><td class="n">{SD}</td><td class="n"></td>'
            else:
                trx,monto=v
                x=(monto/trx) if (trx and monto) else None
                row+=f'<td class="n">{n0(trx) if trx is not None else "–"}</td><td class="n">{m0(monto)}</td><td class="n xs">{m2(x)}</td>'
        at,am=lab_anual(L)
        ax=am/at if (at and am) else None
        row+=f'<td class="n b">{n0(at) if at else "–"}</td><td class="n b">{m0(am)}</td><td class="n xs">{m2(ax)}</td><td class="n">{pctf(am/GT_FC,2) if am else ""}</td>'
        dir_rows+=f"<tr>{row}</tr>"
    if len(labs)>1:
        row=f'<td></td><td class="l b">Total {disp(cliente)[0]}</td><td></td><td></td>'
        for mes in M4:
            g=grupo_mes(labs,mes)
            if g is None: row+=f'<td class="n">{SD}</td><td class="n">{SD}</td><td></td>'
            else:
                t,m=g; x=m/t if t else None
                row+=f'<td class="n b">{n0(t)}</td><td class="n b">{m0(m)}</td><td class="n xs">{m2(x)}</td>'
        t=sum(lab_anual(L)[0] for L in labs); m=sum(lab_anual(L)[1] for L in labs)
        row+=f'<td class="n b">{n0(t)}</td><td class="n b">{m0(m)}</td><td class="n xs">{m2(m/t) if t else ""}</td><td class="n">{pctf(m/GT_FC,2)}</td>'
        dir_rows+=f'<tr class="subtot">{row}</tr>'
# filas de cierre
row=f'<td></td><td class="l b">TOTAL CLIENTES DIRECTOS</td><td></td><td></td>'
for rr in month_rows:
    row+=f'<td class="n b">{n0(rr["dt"])}</td><td class="n b">{m0(rr["dm"])}</td><td class="n xs">{m2(rr["dm"]/rr["dt"]) if rr["dt"] else ""}</td>'
row+=f'<td class="n b">{n0(DIR_TRX)}</td><td class="n b">{m0(DIR_FC)}</td><td class="n xs">{m2(DIR_FC/DIR_TRX)}</td><td class="n b">{pctf(DIR_FC/GT_FC)}</td>'
dir_rows+=f'<tr class="subtot">{row}</tr>'
row=f'<td></td><td class="l b">AVANTER (consolidado → ver sección Avanter)</td><td></td><td></td>'
for rr in month_rows:
    if rr["am"] is None: row+=f'<td class="n">{SD}</td><td class="n">{SD}</td><td></td>'
    else: row+=f'<td class="n b">{n0(rr["at"])}</td><td class="n b">{m0(rr["am"])}</td><td class="n xs">{m2(rr["am"]/rr["at"])}</td>'
row+=f'<td class="n b">{n0(AV_T)}</td><td class="n b">{m0(AV_M)}</td><td class="n xs">{m2(AV_M/AV_T)}</td><td class="n b">{pctf(AV_M/GT_FC)}</td>'
dir_rows+=f'<tr class="avrow">{row}</tr>'
row=f'<td></td><td class="l b">TOTAL GENERAL</td><td></td><td></td>'
for rr in month_rows:
    row+=f'<td class="n b">{n0(rr["tt"])}</td><td class="n b">{m0(rr["tm"])}</td><td class="n xs">{m2(rr["tm"]/rr["tt"])}</td>'
row+=f'<td class="n b">{n0(GT_TRX)}</td><td class="n b">{m0(GT_FC)}</td><td class="n xs">{m2(GT_FC/GT_TRX)}</td><td class="n b">100%</td>'
dir_rows+=f'<tr class="tot">{row}</tr>'

dir_head_months="".join(f'<th colspan="3" class="mes">{MESN[m]}</th>' for m in M4)
dir_sub="".join('<th class="sub">TRX</th><th class="sub">Monto FC</th><th class="sub">$/TRX</th>' for _ in M4)

# ---- Avanter: tabla por laboratorio ----
av_data=[]
for L in AV_LABS:
    if L["cod"] in ("5260","5270"): continue
    at,am=lab_anual(L)
    av_data.append((L,at,am))
av_data.sort(key=lambda x:-x[1])
av_rows=""
for L,at,am in av_data:
    f=L.get("fc") or ""
    tipo="Éticos" if "ETICOS" in f else ("Dinámicas / Bonos" if "DINAMICAS" in f else "—")
    cells=f'<td class="cod">{L["cod"]}</td><td class="l"><b>{_html.escape(L["prog"])}</b><br><small>{_html.escape(L["lab"])}</small></td><td class="c sm">{tipo}</td>'
    for mes in M4:
        v=L["m"].get(mes)
        if v is None or v[0] is None: cells+=f'<td class="n">–</td>'
        else: cells+=f'<td class="n">{n0(v[0])}</td>'
    share=at/AV_T if AV_T else 0
    cells+=f'<td class="n b">{n0(at)}</td>'
    cells+=f'<td class="w"><div class="minibar"><i style="width:{min(share*100,100):.2f}%"></i></div></td>'
    cells+=f'<td class="n">{pctf(share,2)}</td><td class="n">{m0(am) if am else "–"}</td>'
    av_rows+=f"<tr>{cells}</tr>"
av_rows+=f"""<tr class="tot"><td></td><td class="l">TOTAL AVANTER (facturado oficial)</td><td></td>
 <td class="n">{n0(881068)}</td><td class="n">{n0(587574)}</td><td class="n">{n0(532345)}</td><td class="n">{n0(488187)}</td>
 <td class="n">{n0(AV_T)}</td><td></td><td class="n">100%</td><td class="n">{m0(AV_M)}</td></tr>"""

altas_rows=""
for L in AV_LABS:
    if L["cod"] not in ("5260","5270"): continue
    am=lab_anual(L)[1]
    altas_rows+=f'<tr><td class="cod">{L["cod"]}</td><td class="l">{_html.escape(L["prog"])}</td><td class="c">Abril 2026</td><td class="n">{m0(am)}</td></tr>'

# ---- Análisis: matriz cliente × mes ----
mat_rows=""
for c,t,m,labs in CLITOT:
    name,nuevo=disp(c)
    badge=' <span class="chip chip-new">NUEVO</span>' if nuevo else ""
    vals=[]
    for mes in M4:
        if labs is None:
            avc=AV_CONSOL[mes]; vals.append(avc[1] if avc else None)
        else:
            g=grupo_mes(labs,mes); vals.append(g[1] if g else None)
    d = (vals[3]/vals[0]-1) if (vals[0] and vals[3]) else None
    dcls = "up" if (d is not None and d>=0) else "down"
    dtxt = ("▲ " if (d is not None and d>=0) else "▼ ")+pctf(abs(d)) if d is not None else "–"
    cells="".join(f'<td class="n">{m0(v) if v is not None else SD}</td>' for v in vals)
    mat_rows+=f'<tr><td class="l"><b>{name}</b>{badge}</td>{cells}<td class="n b">{m0(m)}</td><td class="n">{pctf(m/GT_FC)}</td><td class="n {dcls}">{dtxt}</td></tr>'
mat_rows+=f"""<tr class="tot"><td class="l">TOTAL</td>
 {"".join(f'<td class="n">{m0(rr["tm"])}</td>' for rr in month_rows)}
 <td class="n">{m0(GT_FC)}</td><td class="n">100%</td>
 <td class="n">▼ {pctf(abs(month_rows[3]["tm"]/month_rows[0]["tm"]-1))}</td></tr>"""

prom_rows=""
for rr in month_rows:
    xd=rr["dm"]/rr["dt"] if rr["dt"] else None
    vu=AVU[rr["mes"]]
    xt=rr["tm"]/rr["tt"] if rr["tt"] else None
    prom_rows+=f'<tr><td class="l">{MESN[rr["mes"]]}</td><td class="n">{m2(xd)}</td><td class="n">{m2(vu) if vu else SD}</td><td class="n b">{m2(xt)}</td></tr>'
prom_rows+=f'<tr class="tot"><td class="l">Promedio ene–abr</td><td class="n">{m2(DIR_FC/DIR_TRX)}</td><td class="n">{m2(AV_M/AV_T)}*</td><td class="n">{m2(GT_FC/GT_TRX)}</td></tr>'

FECHA="12/06/2026"
HTML=f"""<!DOCTYPE html><html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Reporte Laboratorios 2026 — World Salud</title>
<style>
:root{{--azul1:#1F4E79;--azul2:#4D94D8;--azul3:#A6C9EB;--azul4:#EAF3FC;--amar:#FFE699;--gris:#5a6b7b;--borde:#cdd8e3}}
*{{box-sizing:border-box}}
body{{font-family:'Segoe UI',Roboto,Arial,sans-serif;margin:0;background:#eef2f7;color:#1c2b3a}}
header{{background:var(--azul1);color:#fff;padding:18px 26px}}
header h1{{margin:0;font-size:21px;letter-spacing:.2px}}
header p{{margin:6px 0 0;font-size:13px;opacity:.85}}
.tabs{{display:flex;gap:2px;background:#15395c;padding:0 16px;overflow-x:auto}}
.tabs button{{background:none;border:0;color:#cfe0f0;padding:12px 18px;cursor:pointer;font-size:14px;border-bottom:3px solid transparent;white-space:nowrap}}
.tabs button.on{{color:#fff;border-bottom-color:var(--azul2);font-weight:600}}
.pane{{display:none;padding:18px;max-width:1500px;margin:0 auto}}
.pane.on{{display:block}}
h2{{color:var(--azul1);font-size:17px;margin:22px 4px 10px}}
h2:first-child{{margin-top:4px}}
.note{{font-size:12.5px;color:var(--gris);margin:6px 4px 12px;line-height:1.55}}
/* KPIs */
.kpis{{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:12px;margin-bottom:6px}}
.kpi{{background:#fff;border:1px solid var(--borde);border-top:4px solid var(--azul2);border-radius:8px;padding:14px 16px}}
.kpi-t{{font-size:12px;text-transform:uppercase;letter-spacing:.5px;color:var(--gris)}}
.kpi-v{{font-size:26px;font-weight:700;color:var(--azul1);margin:4px 0}}
.kpi-s{{font-size:11.5px;color:var(--gris);line-height:1.45}}
/* tablas */
.scroll{{overflow:auto;max-height:78vh;border:1px solid var(--borde);border-radius:8px;background:#fff}}
table{{border-collapse:collapse;font-size:12px;white-space:nowrap;width:100%}}
th,td{{border:1px solid #dce4ec;padding:5px 8px}}
thead th{{position:sticky;top:0;background:var(--azul3);color:#16314a;font-weight:700;text-align:center;z-index:3}}
thead tr:nth-child(2) th{{top:29px}}
th.mes{{background:var(--azul2);color:#fff}}
td.l{{text-align:left;white-space:normal;min-width:170px}}
td.l small{{color:var(--gris)}}
td.sm{{font-size:11px;color:#566;max-width:240px}}
td.c{{text-align:center}}
td.cod{{text-align:center;font-weight:600}}
td.n{{text-align:right;font-variant-numeric:tabular-nums}}
td.b,.b{{font-weight:700}}
td.xs{{font-size:11px;color:var(--gris)}}
td.w{{min-width:110px}}
td.inact{{background:#f4f6f9}}
.sd{{color:#b06a00;font-size:11px;font-style:italic}}
tr.grp td{{background:var(--azul1);color:#fff;font-weight:700;text-align:left;font-size:12.5px}}
tr.subtot td{{background:var(--azul4)}}
tr.avrow td{{background:#d9e8f7}}
tr.tot td{{background:var(--amar);font-weight:700}}
tbody tr:hover td:not(.inact){{background:#f2f8fe}}
tr.grp:hover td{{background:var(--azul1)}}
tr.tot:hover td{{background:var(--amar)}}
/* chips */
.chip{{display:inline-block;font-size:10px;font-weight:700;padding:2px 8px;border-radius:10px;vertical-align:middle}}
.chip-prod{{background:var(--azul4);color:var(--azul1);border:1px solid var(--azul3)}}
.chip-start{{background:var(--amar);color:#7a5b00}}
.chip-new{{background:var(--amar);color:#7a5b00;margin-left:6px}}
.flag{{background:var(--amar);color:#7a5b00;font-size:10px;font-weight:700;padding:1px 7px;border-radius:9px}}
.warn{{color:#b06a00}}
.up{{color:#1d7a3a;font-weight:600}} .down{{color:#b3261e;font-weight:600}}
/* gráfico de barras */
.chartbox{{background:#fff;border:1px solid var(--borde);border-radius:8px;padding:16px 18px}}
.brow{{display:flex;align-items:center;gap:10px;margin:9px 0}}
.blab{{width:70px;font-size:13px;font-weight:600;color:var(--azul1)}}
.btrack{{flex:1;height:22px;background:#f1f4f8;border-radius:5px;overflow:hidden;display:flex}}
.seg{{height:100%}}
.seg.d{{background:var(--azul1)}} .seg.a{{background:var(--azul2)}}
.seg.x{{background:repeating-linear-gradient(45deg,#e3e9f0,#e3e9f0 6px,#f6f8fb 6px,#f6f8fb 12px)}}
.bval{{width:200px;font-size:13px;font-weight:700;color:var(--azul1)}}
.legend{{display:flex;gap:18px;font-size:12px;color:var(--gris);margin-top:10px}}
.dot{{display:inline-block;width:11px;height:11px;border-radius:3px;margin-right:5px;vertical-align:-1px}}
.dot.d{{background:var(--azul1)}} .dot.a{{background:var(--azul2)}}
.dot.x{{background:repeating-linear-gradient(45deg,#dde,#dde 3px,#f6f8fb 3px,#f6f8fb 6px)}}
.minibar{{height:12px;background:#f1f4f8;border-radius:4px;overflow:hidden}}
.minibar i{{display:block;height:100%;background:linear-gradient(90deg,var(--azul2),var(--azul1))}}
/* alert / cards */
.alert{{background:#fffaf0;border:1px solid #f0d490;border-left:5px solid #e3a008;border-radius:8px;padding:12px 16px;font-size:13px;margin-top:14px}}
.alert ul{{margin:8px 0 4px;padding-left:20px;line-height:1.6}}
.card{{background:#fff;border:1px solid var(--borde);border-radius:8px;padding:16px 18px;margin-bottom:14px}}
.cols2{{display:grid;grid-template-columns:1fr 1fr;gap:14px}}
@media(max-width:980px){{.cols2{{grid-template-columns:1fr}}.bval{{width:auto}}}}
footer{{text-align:center;font-size:11.5px;color:var(--gris);padding:16px}}
ul.clean{{line-height:1.7;font-size:13px;padding-left:20px;margin:6px 0}}
</style></head><body>
<header><h1>📊 Reporte Laboratorios 2026 — Seguimiento por Cliente</h1>
<p>Enero – Abril 2026 · Montos en ARS netos (sin IVA) · TRX = transacciones / envases aprobados · Actualizado al {FECHA}</p></header>
<div class="tabs">
<button class="on" onclick="sel(0)">Resumen</button>
<button onclick="sel(1)">Clientes Directos</button>
<button onclick="sel(2)">Avanter</button>
<button onclick="sel(3)">Análisis</button>
<button onclick="sel(4)">Fuentes y Datos</button>
</div>

<!-- ============ RESUMEN ============ -->
<div class="pane on" id="p0">
{kpis}
<div class="alert" style="background:#eef6ff;border-color:#9cc4ec;border-left-color:var(--azul2)"><b>💡 Importante — valores netos sin IVA.</b> Todos los montos de este reporte son <b>precios netos (sin IVA)</b>, tal como figuran en los soportes de facturación. Para emitir factura se agrega <b>IVA 21%</b> y, según el cliente/jurisdicción, <b>percepciones de IIBB</b>. Ejemplo: AstraZeneca abril $ 21.786.516,13 neto → $ 26.361.684,52 con IVA.</div>
<h2>Reporte mensual</h2>
<div class="scroll" style="max-height:none"><table>
<thead><tr><th rowspan="2">Mes</th><th colspan="2" class="mes">Clientes directos</th><th colspan="2" class="mes">Avanter</th><th colspan="2" class="mes">Total</th><th rowspan="2">$ prom / TRX</th></tr>
<tr><th class="sub">TRX</th><th class="sub">Facturación</th><th class="sub">TRX</th><th class="sub">Facturación</th><th class="sub">TRX</th><th class="sub">Facturación</th></tr></thead>
<tbody>{mt_rows}</tbody></table></div>
<p class="note">El $ promedio/TRX se calcula sobre TRX informadas; incluye conceptos sin TRX (abonos, plataformas, altas).</p>

<h2>Facturación por mes</h2>
<div class="chartbox">{chart}</div>

<h2>Ranking de clientes (facturación acumulada ene–abr)</h2>
<div class="scroll" style="max-height:none"><table>
<thead><tr><th>#</th><th>Cliente</th><th>Facturación acum.</th><th>% del total</th><th>Participación</th><th>TRX acum.</th></tr></thead>
<tbody>{top}</tbody></table></div>
{alertas}
</div>

<!-- ============ CLIENTES DIRECTOS ============ -->
<div class="pane" id="p1">
<h2>Detalle por cliente y programa</h2>
<p class="note">Cada programa con su modelo de cobro y la facturación mensual. «s/d» = sin archivo/soporte para ese mes. Las filas celestes son subtotales por cliente; la amarilla, el total general. Provincia ART (convenio de servicios) y Serdata (proveedor) quedan excluidos de este reporte de laboratorios.</p>
<div class="scroll"><table>
<thead><tr><th rowspan="2">Código</th><th rowspan="2">Programa / Laboratorio</th><th rowspan="2">Modelo de cobro</th><th rowspan="2">Status</th>{dir_head_months}<th colspan="4" class="mes" style="background:var(--azul1)">ACUMULADO ene–abr</th></tr>
<tr>{dir_sub}<th class="sub">TRX</th><th class="sub">Monto FC</th><th class="sub">$/TRX</th><th class="sub">% FC total</th></tr></thead>
<tbody>{dir_rows}</tbody></table></div>
</div>

<!-- ============ AVANTER ============ -->
<div class="pane" id="p2">
<div class="kpis">
 <div class="kpi"><div class="kpi-t">Facturado Avanter ene–abr</div><div class="kpi-v">{MM(AV_M)}</div><div class="kpi-s">{m2(AV_M)} · {pctf(AV_M/GT_FC)} del total general</div></div>
 <div class="kpi"><div class="kpi-t">TRX Avanter</div><div class="kpi-v">{n0(AV_T)}</div><div class="kpi-s">Ene 881.068 · Feb 587.574 · Mar 532.345 · Abr 488.187</div></div>
 <div class="kpi"><div class="kpi-t">Valor unitario por TRX</div><div class="kpi-v">$74 → $93</div><div class="kpi-s">Ene–Mar $74,00 · Abr $93,00 (ajuste trimestral IPC INDEC)</div></div>
 <div class="kpi"><div class="kpi-t">Altas abril</div><div class="kpi-v">ENA + Convatec</div><div class="kpi-s">$ 547.420,80 c/u (alta laboratorio nuevo) = $ 1.094.841,59</div></div>
</div>
<h2>TRX por laboratorio</h2>
<p class="note">Ordenado por TRX acumuladas. Monto FC = TRX × valor unitario del mes. El % es sobre el total de TRX Avanter informadas. «–» = el laboratorio no figura en el archivo de ese mes.</p>
<div class="scroll"><table>
<thead><tr><th>Código</th><th>Laboratorio / Programa</th><th>Tipo</th><th>TRX Ene</th><th>TRX Feb</th><th>TRX Mar</th><th>TRX Abr</th><th>TRX acum.</th><th>Participación</th><th>%</th><th>Monto FC acum.</th></tr></thead>
<tbody>{av_rows}</tbody></table></div>
<h2>Altas de laboratorios nuevos</h2>
<div class="scroll" style="max-height:none"><table>
<thead><tr><th>Código</th><th>Laboratorio</th><th>Alta</th><th>Monto facturado</th></tr></thead>
<tbody>{altas_rows}</tbody></table></div>
</div>

<!-- ============ ANÁLISIS ============ -->
<div class="pane" id="p3">
<h2>Facturación mensual por cliente</h2>
<p class="note">Matriz cliente × mes con la variación de abril contra enero (sobre clientes con ambos meses informados).</p>
<div class="scroll"><table>
<thead><tr><th>Cliente</th><th>Enero</th><th>Febrero</th><th>Marzo</th><th>Abril</th><th>Acumulado</th><th>% total</th><th>Δ Abr vs Ene</th></tr></thead>
<tbody>{mat_rows}</tbody></table></div>
<h2>Valor promedio por transacción</h2>
<div class="cols2">
<div class="card"><table>
<thead><tr><th>Mes</th><th>$/TRX Directos</th><th>VU Avanter</th><th>$/TRX Global</th></tr></thead>
<tbody>{prom_rows}</tbody></table>
<p class="note">* El promedio Avanter ene–abr ({m2(AV_M/AV_T)}) supera el VU porque abril incluye las altas de ENA y Convatec. El $/TRX de los directos incluye abonos y plataformas (AZ, Sidus, Panalab digital), por eso es más alto que un valor por bono.</p></div>
<div class="card"><b style="color:var(--azul1)">Lecturas rápidas</b>
<ul class="clean">
<li><b>Avanter</b> explica {pctf(AV_M/GT_FC)} de la facturación; Andrómaco concentra más de la mitad de sus TRX.</li>
<li><b>Panalab</b> es el cliente directo más grande ({m0(66225682.10)}): el papel manual crece todos los meses (+67% TRX abr vs ene) y la plataforma digital aporta ~$6,6–8,6 M/mes por WhatsApp + servicio.</li>
<li><b>Sidus Dermo</b> casi duplicó TRX en abril (2.333 → 4.879) e incorporó armado de dinámicas.</li>
<li><b>AZ (Elegir Salud)</b> es el mayor cliente directo: ~$ 20–22 M/mes por presupuesto integral (operación + call center + cápitas + auditorías + supervisor CS). Febrero bajó a $ 19,7 M por menos cápitas activas adicionales; +4% abr vs ene por ajuste IPC CABA.</li>
<li>La caída de abril vs enero en el total (−9%) se explica por menos TRX Avanter (881 mil → 588 mil → 532 mil → 488 mil), parcialmente compensada por el VU $74 → $93; AZ se mantiene estable (~$ 21 M/mes).</li>
</ul></div>
</div>
</div>

<!-- ============ FUENTES Y DATOS ============ -->
<div class="pane" id="p4">
<h2>Datos pendientes (impactan los totales)</h2>
<div class="card"><ul class="clean">
<li>✅ <b>Avanter:</b> 4 meses completos (marzo incorporado: 532.345 TRX × $74 = $ 39.393.530).</li>
<li>🟠 <b>Febrero:</b> sin soporte de <b>Colgate</b>, <b>Max Vision</b>, <b>Sidus Farma</b> (aplicaría su mínimo de $ 766.721,67) y <b>Ceoderma</b>.</li>
<li>🟠 <b>Ceoderma:</b> solo enero disponible (feb–abr s/d). En enero además registró 355 TRX vía Avanter, ya incluidas en el consolidado Avanter.</li>
<li>🟡 <b>Health Care (Haleon + Luar):</b> montos de feb y abr sin cantidad de TRX; marzo s/d. Enero pendiente de confirmar: el archivo «Facturación Health Care» muestra componentes ene-26 por $ 449.969,36 y $ 371.875,50 que no se sumaron por ambigüedad.</li>
<li>🟡 <b>Lazar:</b> enero $ 500.000 (startup fase 1) y abril $ 344.079,51 (mínimo facturable); feb–mar s/d.</li>
<li>ℹ️ <b>Excluidos por definición:</b> Provincia ART (convenio de servicios, no laboratorio) y Serdata (proveedor de procesamiento, no cliente).</li>
</ul></div>

<h2>Parámetros de facturación vigentes</h2>
<div class="cols2">
<div class="card"><b style="color:var(--azul1)">Valores y mínimos</b>
<ul class="clean">
<li><b>VU Avanter:</b> Ene–Mar $74,00 · Abr $93,00 (ajuste trimestral por IPC INDEC).</li>
<li><b>Alta laboratorio nuevo (Avanter):</b> base dic-25 $ 500.000 → IPC ene +2,90% ($ 514.500) → feb +2,90% ($ 529.420,50) → mar +3,40% (<b>$ 547.420,80</b>, aplicado a ENA y Convatec).</li>
<li><b>Mínimo Sidus Farma:</b> $ 766.721,67 (aplicó en ene/mar/abr: el 1% PVP quedó debajo).</li>
<li><b>Mínimo Bayer:</b> $ 801.411,81 (ene–feb) / $ 844.447,62 (mar–abr). No aplicó: facturó por 0,50% PVP.</li>
<li><b>Dólar BNA (WhatsApp USD 0,10):</b> ene $1.390 · feb $1.420 · mar $1.380 · abr $1.410.</li>
<li><b>Operación AZ (Elegir Salud):</b> $ 7.961.858 (ene) → $ 8.208.676 (feb) → $ 8.422.101 (mar) → $ 8.674.764 (abr), ajuste mensual IPC CABA. Total presupuesto: $ 20,93 / 19,69 / 21,61 / 21,79 M.</li>
</ul></div>
<div class="card"><b style="color:var(--azul1)">% sobre PVP por cliente</b>
<ul class="clean">
<li>Bayer 0,50% (Firialta / Xarelto / Verquvo)</li>
<li>Panalab digital 0,60% + servicio · Ceoderma 0,60%</li>
<li>Colgate 1,00% · Max Vision 1,00% · Sidus Farma 1,00%</li>
<li>Provincia ART 0,35% sobre monto vendido</li>
<li>Panalab papel: por registro procesado + cajas (IPC CABA): registro $86,34 → $93,53; caja $ 4.063 → $ 4.401</li>
</ul></div>
</div>

<h2>Fuentes (carpeta PROFORMAS – Drive)</h2>
<div class="card"><ul class="clean">
<li><b>Avanter:</b> «Facturación Avanter Enero/Febrero/Marzo/Abril 2026.xlsx» (TRX por laboratorio + VU + altas).</li>
<li><b>Bayer:</b> «Soporte Facturación Enero/Febrero/Marzo/Abril 2026.xlsx» (resumen 0,5% PVP con mínimo).</li>
<li><b>AZ (Elegir Salud):</b> «PRESUPUESTO FACTURACION &lt;MES&gt; 2026 - PROGRAMA ELEGIR SALUD» / Proforma editable con ajuste IPC (operación + call center + cápitas + auditorías + supervisor CS).</li>
<li><b>Sidus Dermo, BIU, Panalab digital:</b> «Facturación Sidus DC, Biu, Panalab, AZ &lt;mes&gt; 2026.xlsx».</li>
<li><b>Panalab papel:</b> «Bonos Papel Panalab Manual &lt;mes&gt; 2026.xlsx».</li>
<li><b>Colgate / Max Vision / Sidus Farma / Ceoderma / Lazar:</b> soportes «Soporte Facturación &lt;MES&gt; 2026_&lt;cliente&gt;.xlsx» (carátula).</li>
<li><b>Health Care:</b> «Facturación Health Care.xlsx» + «Proformas - Resumen Valores sin adjunto».</li>
</ul>
<p class="note">Metodología: montos <b>netos sin IVA</b> tomados de las carátulas/resúmenes de cada soporte (coinciden con el «Importe sin impuestos» de cada factura). Para facturar se agrega IVA 21% + percepciones IIBB según corresponda. Avanter: monto = TRX × VU del mes (coincide con el total facturado oficial). Sidus Dermo y Panalab digital suman bonos + uso de plataforma + WhatsApp.</p></div>
</div>

<footer>Reporte Laboratorios 2026 · World Salud · Generado automáticamente desde los soportes de facturación (carpeta PROFORMAS) · Actualizado al {FECHA}</footer>
<script>
function sel(i){{
 document.querySelectorAll('.tabs button').forEach((b,k)=>b.classList.toggle('on',k===i));
 document.querySelectorAll('.pane').forEach((p,k)=>p.classList.toggle('on',k===i));
 window.scrollTo(0,0);
}}
</script>
</body></html>"""
open("Reporte_Labos_2026.html","w",encoding="utf-8").write(HTML)
print("HTML OK -> Reporte_Labos_2026.html")
print("Totales:", {rr["mes"]: (rr["tt"], round(rr["tm"],2)) for rr in month_rows})
print("GT_TRX:",GT_TRX," GT_FC:",round(GT_FC,2)," AV:",AV_T,round(AV_M,2)," DIR:",DIR_TRX,round(DIR_FC,2))

# =================== CSV (layout original) -> Google Sheet nativo ===================
import csv as _csv
def _n(x,d=2):
    if x is None: return ""
    if isinstance(x,int): return str(x)
    return f"{x:.{d}f}"
hdr=["Código","Programa","Laboratorio","Lab. Propio","Avanter Proveedor",
     "FC POR AVANTER","Precio","Actualización","Status"]
for m in MESES+["ANUAL (ene-abr)"]:
    hdr+=[f"{m} TRX",f"{m} % TRX MES",f"{m} MONTO FC",f"{m} MONTO X TRX"]
rows=[["REPORTE LABOS 2026 - Seguimiento por Cliente (ene-abr) | Montos ARS sin IVA | s/d: feb de Colgate/MaxVision/SidusFarma/Ceoderma; Ceoderma feb-abr | Excluidos: Provincia ART (convenio) y Serdata (proveedor)"],[],hdr]
for cliente,labs in GRUPOS:
    rows.append([f"CLIENTE: {cliente}"])
    for L in labs:
        row=[L["cod"],L["prog"],L["lab"],L["propio"],L["avp"],L["fc"],L["precio"],L["act"],L["status"]]
        for m in MESES:
            v=L["m"].get(m); trx=monto=None
            if v: trx,monto=v
            p=(trx/tot_trx[m]) if (trx and tot_trx[m]) else None
            x=(monto/trx) if (trx and monto) else None
            row+=[_n(trx,0),(_n(p*100,4)+"%" if p is not None else ""),_n(monto),_n(x)]
        at,am=lab_anual(L)
        ap=at/tot_anual_trx if (at and tot_anual_trx) else None
        ax=am/at if (at and am) else None
        row+=[_n(at,0),(_n(ap*100,4)+"%" if ap is not None else ""),_n(am),_n(ax)]
        rows.append(row)
with open("Reporte_Labos_2026_GoogleSheet.csv","w",newline="",encoding="utf-8") as f:
    w=_csv.writer(f)
    for r3 in rows: w.writerow(r3)
print("CSV OK -> Reporte_Labos_2026_GoogleSheet.csv")
