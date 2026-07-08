# Evidencia local - ER XLSX

## Archivos

- Origen: `sample-inputs/balanza_SME170717GA0_2026_07.xls`
- Generado: `sample-outputs/estado_resultados_al_servicios_multiples_empresariales_sa_de_cv_2026_07.xlsx`
- Reporte: `sample-outputs/validation_report_al_servicios_multiples_empresariales_sa_de_cv_2026_07.json`
- Hoja visible generada: `ER`

## Deteccion

- Empresa detectada: `AL SERVICIOS MULTIPLES EMPRESARIALES SA DE CV`
- Periodo detectado: `2026-07`
- Encabezado ER generado: `Del 1ro de Enero al 31 de Julio 2026`

## Validacion local

- Formula cells generadas: `92`
- Formula errors encontrados por construccion/inspeccion: `0`
- Validacion de cuadre local: `abs(0.0) < 1`
- Filas normalizadas: `157`
- Cuentas hoja usadas para calculo: `126`
- Cuentas no mapeadas por el motor ER provisional: `68`
- Referencia conceptual de cuadre: `BG!L47` del manual, sin copiar el workbook manual.

La validacion no evalua formulas de Excel; revisa que las formulas generadas no
contengan tokens de error como `#REF!`, `#DIV/0!`, `#VALUE!`, ni referencias a
workbooks externos. La diferencia de cuadre local se calculo con cuentas hoja como
`sum(debe) - sum(haber)`, sin sumar filas padre duplicadas.

## Celdas con formulas

Subtotales ER:

- `ER!D20`, `ER!H20`: ingresos netos.
- `ER!D25`, `ER!H25`: utilidad bruta.
- `ER!D51`, `ER!H51`: gastos de operacion.
- `ER!D53`, `ER!H53`: utilidad o perdida de operacion.
- `ER!D58`, `ER!H58`: total otros ingresos.
- `ER!D63`, `ER!H63`: total R.I.F.
- `ER!D65`, `ER!H65`: resultado antes de impuestos.
- `ER!D70`, `ER!H70`: resultado del ejercicio.

Porcentajes ER:

- Columnas `ER!F` y `ER!J` en filas `18`, `19`, `20`, `23`, `25`, `28` a `47`,
  `50`, `51`, `53`, `56`, `57`, `58`, `61`, `62`, `63`, `65`, `67`, `68`, `70`.

## API publica usada

- `src.workbook.detect_source_metadata(path)`
- `src.workbook.build_er_workbook(dataset, metadata=None, source_path=None)`
- `src.workbook.save_er_workbook(dataset, output_path=None, metadata=None, source_path=None)`
- `src.validation.validate_generated_workbook(workbook_or_path, balance_difference=None, tolerance=1.0)`
- `src.validation.validate_balance_difference(difference, tolerance=1.0)`
- `src.validation.assert_generated_workbook_valid(workbook_or_path, balance_difference=None, tolerance=1.0)`
- `src.prototype.run_prototype(input_path)`
