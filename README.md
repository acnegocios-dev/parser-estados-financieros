# Prototipo local - Estados Financieros

Prototipo local para leer una balanza de Auditalo, calcular un estado de resultados
preliminar y generar evidencia XLSX antes de llevar el modulo al VPS.

## Ejecutar

```bash
py -m src.prototype
```

Entrada usada por defecto:

`sample-inputs/balanza_SME170717GA0_2026_07.xls`

Salidas generadas:

- `sample-outputs/estado_resultados_al_servicios_multiples_empresariales_sa_de_cv_2026_07.xlsx`
- `sample-outputs/validation_report_al_servicios_multiples_empresariales_sa_de_cv_2026_07.json`

## Estado validado

- Periodo detectado: `2026-07`
- Filas normalizadas: `157`
- Cuentas hoja usadas para calculo: `126`
- Diferencia de cuadre: `0.0`
- Errores de formula detectados: `0`

## Carpetas

- `docs`: analisis, contratos y evidencia local.
- `src`: parser, motor contable, generador XLSX, validacion y orquestador.
- `sample-inputs`: archivos de ejemplo de entrada.
- `sample-outputs`: archivos generados para comparacion.
