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
- Total: $370.153.282,00 · 3.004.959 TRX. Avanter completo los 4 meses.
- AstraZeneca (Programa Elegir Salud): se carga el presupuesto integral mensual (operación +
  call center + cápitas + auditorías + supervisor CS), NO la línea de Avanter. Totales:
  ene $20.929.743,55 · feb $19.689.476,25 · mar $21.612.629,37 · abr $21.786.516,13.
  Fuente: «PRESUPUESTO FACTURACION ‹MES› 2026 - PROGRAMA ELEGIR SALUD» / Proforma editable IPC CABA.
- Faltantes: feb de Colgate/Max Vision/Sidus Farma/Ceoderma; Ceoderma mar–abr;
  Health Care ene (confirmar $449.969,36 + $371.875,50) y mar + TRX; Lazar feb–mar.
