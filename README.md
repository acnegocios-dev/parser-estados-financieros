# Prototipo local - Estados Financieros

Prototipo local para leer una balanza de Auditalo, calcular un estado de resultados
preliminar y generar evidencia XLSX antes de llevar el modulo al VPS.

Rama de preparacion actual:

`feature/estados-financieros-parser-auditalo`

Objetivo de la rama: documentar y validar el parser Auditalo antes de integrar
el modulo al Portal CI.

## Ejecutar

```bash
py -m src.prototype
```

Entrada usada por defecto:

`sample-inputs/balanza_SME170717GA0_2026_07.xls`

Salidas generadas:

- `sample-outputs/estado_resultados_al_servicios_multiples_empresariales_sa_de_cv_2026_07.xlsx`
- `sample-outputs/validation_report_al_servicios_multiples_empresariales_sa_de_cv_2026_07.json`

## API local minima

```bash
python3 -m src.api --host 127.0.0.1 --port 8080
```

Endpoints:

- `GET /health`
- `POST /process`

El `POST /process` acepta `multipart/form-data` con el campo `file` o un cuerpo
JSON con `input_path` para pruebas locales.

La respuesta y el reporte separan la inspeccion estatica de la evaluacion real:

- `formula_static_validation`: valida texto de formulas y referencias externas.
- `formula_recalculation_performed`: indica si un motor recalculo el workbook.
- `formula_recalculation_engine`: motor usado o `none`.
- `formula_evaluated_error_count`: cantidad evaluada, o `null` sin recalculo.
- `formula_cached_values_available`: indica si existen valores cacheados evaluados.

## Pruebas

```bash
python3 -m unittest discover -s tests
```

## Estado validado

- Periodo detectado: `2026-07`
- Filas normalizadas: `157`
- Cuentas hoja usadas para calculo: `126`
- Diferencia de cuadre: `-0.059663` (`abs(diferencia) < 1`)
- Validacion estatica de formulas: `true`
- Recalculo de formulas ejecutado: `false` (LibreOffice/soffice no disponible)
- Errores evaluados: `null` mientras no exista un motor de recalculo
- Valores cacheados de formulas: `false`

## Carpetas

- `docs`: analisis, contratos y evidencia local.
- `src`: parser, motor contable, generador XLSX, validacion y orquestador.
- `sample-inputs`: archivos de ejemplo de entrada.
- `sample-outputs`: archivos generados para comparacion.
