# Reporte Labos 2026 — Instrucciones del proyecto

## Entregable (REGLA FIJA pedida por Lucas)
- El resultado a entregar es **únicamente el HTML** (no enviar xlsx/csv salvo pedido explícito).
- Subirlo a la carpeta de Drive `PROFORMAS`: **folder id `1Ecg_Ds3DyD9UpO7xcryXek234z2scOtY`**.
- Nombre del archivo: **`DSH - Reporte Labos - 2026 - ENE-<ÚLTIMO MES CARGADO>`**
  (ej.: con datos hasta abril → `DSH - Reporte Labos - 2026 - ENE-ABR`; hasta mayo → `... - ENE-MAY`).
- Subida: `mcp__Google_Drive__create_file` con `textContent` (el HTML es UTF-8),
  `contentMimeType: text/html` y `disableConversionToGoogleType: true`.

## Cómo se genera
- Script único: `build_reporte_2026.py` (genera HTML + xlsx + csv; solo se entrega el HTML).
- Para sumar un mes: cargar los datos del mes en `GRUPOS`, `AVU` y `AV_CONSOL`,
  agregar el mes a `ACTIVOS`/`M4` y correr `python3 build_reporte_2026.py`.
- Los datos salen de las carátulas/resúmenes de los soportes de facturación de la
  carpeta PROFORMAS (subcarpeta "PROFORMAS - pasar a carpeta Claude", id `1iClJUGdszxJTVv16FwoJTuIunmMfv5-F`).

## Definiciones de negocio (NO cambiar sin confirmar con Lucas)
- **Provincia ART** (convenio de servicios) y **Serdata** (proveedor de procesamiento):
  **EXCLUIDOS** del reporte de laboratorios.
- **Avanter** es un cliente que agrupa ~28 laboratorios; factura TRX × valor unitario del mes
  (ene–mar 2026: $74 · abr: $93, ajuste trimestral IPC INDEC) + altas de laboratorios nuevos
  ($547.420,80 en abril: ENA y Convatec).
- El archivo genérico «Soporte Facturación <Mes> 2026.xlsx» es **BAYER** (0,5% PVP con mínimo).
- Ceoderma operó también vía Avanter en enero (355 TRX, ya en el consolidado Avanter).
- Sidus Dermo y Panalab digital suman bonos + uso de plataforma + WhatsApp (USD 0,10 × dólar BNA).
- Montos en ARS sin IVA. TRX = transacciones/envases aprobados del mes.

## Estado de datos (al 12/06/2026, ene–abr)
- Total: $377.426.332,29 · 3.004.525 TRX. Avanter completo los 4 meses.
- La planilla «Seguimiento Facturación» (Sheet id `17RdCgz50wF5pSqS0GcAKRv4wblUrIBetP-DyF0tEako`)
  lista facturas por MES DE EMISIÓN; el servicio se factura al mes siguiente (correr 1 columna).
  De ahí salen: Colgate feb $1.740.460 · Max Vision feb $1.188.815 · Health Care ene $971.876 y
  mar $904.782 · Lazar feb–mar $500.000 c/u · Sidus Farma feb $131.941 · Ceoderma feb $0.
- Sidus Dermo abril: $3.633.862,44 / 4.445 TRX (proforma + FC 0002-00000370).
- Panalab: el reporte muestra SOLO fee/servicio; las FC grandes (~$28-31M) con reintegro de bonos
  quedan fuera (confirmado por Lucas 12/06/2026).
- AstraZeneca (Programa Elegir Salud): se carga el presupuesto integral mensual (operación +
  call center + cápitas + auditorías + supervisor CS), NO la línea de Avanter. Totales:
  ene $20.929.743,55 · feb $19.689.476,25 · mar $21.612.629,37 · abr $21.786.516,13.
  Fuente: «PRESUPUESTO FACTURACION ‹MES› 2026 - PROGRAMA ELEGIR SALUD» / Proforma editable IPC CABA.
- Faltantes restantes: Ceoderma mar–abr (feb = $0 facturado);
  TRX de Colgate feb, Max Vision feb, Sidus Farma feb, Health Care (todos) y Lazar feb–mar.
