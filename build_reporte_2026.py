#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Genera el "Reporte Labos 2026" replicando la estructura del reporte 2025,
pero AGRUPADO POR CLIENTE (clientes directos + Avanter).

Fuente de datos: carpeta Drive "PROFORMAS" (facturación ene-abr 2026).
Meses disponibles: Enero, Febrero, Marzo, Abril 2026.

Columnas por mes (igual que 2025): TRX, Monto FC, Monto x TRX.
Montos en pesos (ARS), sin IVA. TRX = transacciones / envases aprobados del mes.
"""

import csv

MESES = ["ENE", "FEB", "MAR", "ABR"]

# Valor unitario por TRX de Avanter ("laboratorios viejos") por mes.
AVANTER_VU = {"ENE": 74.0, "FEB": 74.0, "MAR": None, "ABR": 93.0}

# ---------------------------------------------------------------------------
# Cada fila: (codigo, laboratorio, modelo, datos)
# datos = dict mes -> (trx, monto)   None = sin dato
# Para Avanter, monto se calcula = trx * VU del mes.
# ---------------------------------------------------------------------------

def avanter_row(cod, lab, trx_ene, trx_feb, trx_abr):
    d = {}
    for mes, trx in (("ENE", trx_ene), ("FEB", trx_feb), ("ABR", trx_abr)):
        if trx is None:
            d[mes] = None
        else:
            vu = AVANTER_VU[mes]
            d[mes] = (trx, round(trx * vu, 2))
    d["MAR"] = None  # marzo Avanter: no hay desglose por laboratorio disponible
    return (cod, lab, "Avanter - bono/transacción x VU", d)

# --- CLIENTES DIRECTOS ------------------------------------------------------
directos = [
    ("ASTRAZENECA", [
        ("5005", "Programa Elegir Salud (AZ)", "Uso plataforma MAT + TRX x VU", {
            "ENE": (15730, 3164020.00), "FEB": (13198, 2976652.00),
            "MAR": (14775, 3447758.00), "ABR": (14849, 3898957.00)}),
    ]),
    ("BAYER", [
        ("5015", "Programa Más Simple (Bayer)", "0,50% PVP con mínimo", {
            "ENE": (1154, 972221.92), "FEB": (1090, 941267.72),
            "MAR": (1181, 1044493.17), "ABR": (1178, 1075303.74)}),
    ]),
    ("SIDUS", [
        ("5040", "Sidus Dermo (Bonos + Plataforma digital)", "TRX x VU + uso plataforma + WhatsApp", {
            "ENE": (2522, 1282897.00), "FEB": (2281, 1255491.00),
            "MAR": (2333, 1330813.00), "ABR": (4879, 2298686.15)}),
        ("5041", "Programa Convida (Sidus Farma)", "1,00% PVP con mínimo $766.721,67", {
            "ENE": (688, 766721.67), "FEB": None,
            "MAR": (615, 766721.67), "ABR": (894, 766721.67)}),
    ]),
    ("PANALAB", [
        ("5080", "Panalab Digital (Bonos + Plataforma)", "TRX x VU + uso plataforma + WhatsApp", {
            "ENE": (4196, 7099101.00), "FEB": (3565, 6836393.00),
            "MAR": (4586, 8378215.00), "ABR": (4586, 9065584.00)}),
        ("5081", "Panalab Papel Manual", "Registros procesados + cajas (IPC CABA)", {
            "ENE": (63271, 6443666.51), "FEB": (82475, 8320434.77),
            "MAR": (90524, 9289060.32), "ABR": (103401, 10793227.50)}),
    ]),
    ("COLGATE", [
        ("5028", "Bonos de Descuento Colgate", "1,00% PVP", {
            "ENE": (18127, 1730226.41), "FEB": None,
            "MAR": (23875, 2338884.94), "ABR": (20351, 1876522.43)}),
    ]),
    ("MAX VISION", [
        ("5025", "Programa Max Compromiso (Max Vision)", "1,00% PVP", {
            "ENE": (6309, 1313290.74), "FEB": None,
            "MAR": (6019, 1368036.90), "ABR": (6040, 1378201.65)}),
    ]),
    ("BIU", [
        ("5070", "Bonos BIU Cosmeceuticals", "TRX x VU", {
            "ENE": (22, 1628.00), "FEB": (20, 1480.00),
            "MAR": (21, 1554.00), "ABR": (9, 837.00)}),
    ]),
    ("CEODERMA", [
        ("5120", "Programa Eximia Cepage (Ceoderma)", "0,60% PVP", {
            "ENE": (769, 159805.20), "FEB": None, "MAR": None, "ABR": None}),
    ]),
    ("HALEON / HEALTH CARE", [
        ("5026", "Bonos Descuento Haleon + Luar (Health Care)", "Valor fijo por operación", {
            "ENE": None, "FEB": (None, 871875.50),
            "MAR": None, "ABR": (None, 952203.02)}),
    ]),
    ("LAZAR (nuevo 2026)", [
        ("—", "Lazar (Fase 1 Startup)", "Abono/startup; luego % PVP con mínimo", {
            "ENE": (None, 500000.00), "FEB": None,
            "MAR": None, "ABR": (252, 344079.51)}),
    ]),
    ("PROVINCIA ART (nuevo 2026)", [
        ("—", "Provincia ART", "0,35% sobre monto vendido", {
            "ENE": (8166, 1073509.30), "FEB": None,
            "MAR": (8692, 1293685.98), "ABR": (8849, 1471007.21)}),
    ]),
    ("SERDATA (nuevo 2026)", [
        ("—", "Serdata (Bonos)", "Valor unitario por bono procesado", {
            "ENE": None, "FEB": (76676, 4544586.52),
            "MAR": None, "ABR": None}),
    ]),
]

# --- AVANTER (un cliente, muchos laboratorios) ------------------------------
avanter_labs = [
    avanter_row("5004", "Programa TEVAcuidar (TEVA)", 147, 135, 150),
    avanter_row("5010", "Bonos Loreal", 91364, 74721, 67206),
    avanter_row("5011", "Programa Beneficiarte (Ferring)", 453, 394, 484),
    avanter_row("5014", "Programa Rossmore", 1549, 1437, 1683),
    avanter_row("5020", "Bonos Eucerin (Beiersdorf)", 42157, 30505, 29074),
    avanter_row("5022", "Ofertas Casasco", 34587, 28811, 25273),
    avanter_row("5023", "Programa Farmaspen (Aspen)", 1991, 1778, 2285),
    avanter_row("5024", "Nutricia Siempre Juntos (Bagó)", 387, 329, 421),
    avanter_row("5027", "Programa Vari Te Acerca (Varifarma)", 7, 6, 8),
    avanter_row("5050", "Bonos Andrómaco", 554837, 336212, 260641),
    avanter_row("5060", "Bonos Isdin", 40175, 25065, 18937),
    avanter_row("5090", "Bonos Galderma", 10966, 7941, 9934),
    avanter_row("5100", "Bonos Bernabo", 4458, 4207, 4464),
    avanter_row("5110", "Bonos Eurolab", 26534, 24819, 31434),
    avanter_row("5130", "Bonos Cassara", 8603, 7894, 9823),
    avanter_row("5140", "Bonos Bagó", 54581, 35663, 15759),
    avanter_row("5150", "Bonos Caviahue", 1002, 982, 2593),
    avanter_row("5160", "Bonos By Derm", 5461, 5116, 6100),
    avanter_row("5170", "Bonos Bioderma", 1129, 1053, 1462),
    avanter_row("5180", "Bonos Biferdil", 10, 12, 22),
    avanter_row("5190", "Bonos Fresenius", 47, 38, 107),
    avanter_row("5200", "Bonos SC Johnson", 16, 20, 6),
    avanter_row("5210", "Bonos Pharmatrix", 85, 80, 320),
    avanter_row("5220", "Bonos Megalabs", 0, 0, None),
    avanter_row("5230", "Bonos Siegfried", 0, None, None),
    avanter_row("5240", "Bonos Raisse", 167, 356, None),
    avanter_row("5250", "Bonos Siscom (Rayito de Sol)", 0, 0, 1),
    # Altas abril 2026 (facturación de alta, sin TRX transaccional aún)
    ("5260", "Bonos ENA (alta 22/04/2026)", "Alta laboratorio nuevo", {
        "ENE": None, "FEB": None, "MAR": None, "ABR": (None, 547420.80)}),
    ("5270", "Bonos Convatec (alta 24/04/2026)", "Alta laboratorio nuevo", {
        "ENE": None, "FEB": None, "MAR": None, "ABR": (None, 547420.80)}),
]

# ---------------------------------------------------------------------------
def fmt_num(x):
    if x is None:
        return ""
    if isinstance(x, int):
        return str(x)
    return f"{x:.2f}"

def acumular(datos):
    trx_tot, monto_tot = 0, 0.0
    hay_trx, hay_monto = False, False
    for mes in MESES:
        v = datos.get(mes)
        if v is None:
            continue
        trx, monto = v
        if trx is not None:
            trx_tot += trx; hay_trx = True
        if monto is not None:
            monto_tot += monto; hay_monto = True
    return (trx_tot if hay_trx else None, monto_tot if hay_monto else None)

def fila(cliente, cod, lab, modelo, datos):
    row = [cliente, cod, lab, modelo]
    for mes in MESES:
        v = datos.get(mes)
        if v is None:
            row += ["", "", ""]
        else:
            trx, monto = v
            mxt = ""
            if trx and monto is not None and trx != 0:
                mxt = f"{monto/trx:.2f}"
            row += [fmt_num(trx), fmt_num(monto), mxt]
    acum_trx, acum_monto = acumular(datos)
    row += [fmt_num(acum_trx), fmt_num(acum_monto)]
    return row

# ---------------------------------------------------------------------------
header = ["Cliente", "Código", "Laboratorio / Programa", "Modelo de cobro"]
for m in MESES:
    header += [f"{m} TRX", f"{m} Monto FC", f"{m} Monto x TRX"]
header += ["ACUM TRX (ene-abr)", "ACUM Monto FC (ene-abr)"]

rows = []
# Notas / encabezado
rows.append(["REPORTE LABOS 2026 - Seguimiento por Cliente (ene-abr 2026)"])
rows.append(["Montos en ARS sin IVA. TRX = transacciones / envases aprobados del mes. Fuente: carpeta PROFORMAS Drive."])
rows.append(["NOTA: Avanter marzo no tiene desglose por laboratorio en los archivos compartidos (s/d). Feb de Colgate/MaxVision/Sidus Farma/Ceoderma sin archivo (s/d)."])
rows.append([])
rows.append(header)

# Clientes directos
for cliente, labs in directos:
    for cod, lab, modelo, datos in labs:
        rows.append(fila(cliente, cod, lab, modelo, datos))

# Avanter
for cod, lab, modelo, datos in avanter_labs:
    rows.append(fila("AVANTER", cod, lab, modelo, datos))

# Total consolidado Avanter (de los archivos de facturación)
rows.append(fila("AVANTER", "TOTAL", "TOTAL CONSOLIDADO AVANTER (facturado)",
                 "Total TRX x Valor Unitario + altas",
                 {"ENE": (881068, 65199032.00), "FEB": (587574, 43480476.00),
                  "MAR": None, "ABR": (488187, 46496232.59)}))

with open("Reporte_Labos_2026.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    for r in rows:
        w.writerow(r)

print("OK -> Reporte_Labos_2026.csv")
print(f"Filas de datos: {len(rows)-5}")
