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

- Formula cells generadas: `46`
- Validacion estatica de formulas: `true`
- Recalculo de formulas ejecutado: `false`
- Motor de recalculo: `none`
- Errores evaluados: `null` (no hubo motor de recalculo)
- Valores cacheados de formulas disponibles: `false`
- Validacion de cuadre local: se expone en el reporte como `difference_cuadre`, `cuadra` y `balanza_no_cuadra`.
- Filas normalizadas: `157`
- Cuentas hoja usadas para calculo: `126`
- Cuentas no mapeadas por el motor ER provisional: `68`
- Referencia conceptual de cuadre: `BG!L47` del manual, sin copiar el workbook manual.

## Reejecucion completa y comparacion contra el manual

La ejecucion completa posterior al ajuste de validacion se realizo con:

```text
python3 -m src.prototype
```

Resultado obtenido:

- `period_ym=2026-07`.
- `normalized_rows=157`; `leaf_rows_used_for_calculation=126`.
- `ER!H46=39,614.91`, igual al manual para `Varios` (`6148 + 6176 + 6195`).
- `formula_static_validation=true`; no se detectaron `#REF!`, `#DIV/0!`,
  `#VALUE!`, `#N/A` ni `#NAME?`.
- `difference_cuadre=-0.059663`; tolerancia `1.0`; `cuadra=true` y
  `validation_ok=true`.
- `balanza_no_cuadra=false`.

La salida y el reporte actualizados son los archivos indicados en la seccion
`Archivos`. El reporte conserva la advertencia de que no se ejecuto un motor de
recalculo porque LibreOffice/`soffice` no esta instalado; por tanto, la
validacion de formulas es estatica y el archivo solicita recalculo al abrirse.

La comparacion con `sample-inputs/EEFF_202602_AL_Serv_Prueba.xlsx`, hoja `ER`,
se hizo con `openpyxl` en ambos archivos. No se observaron diferencias en:

- `B9:J12`, los textos de titulo y el encabezado de columnas en fila 15;
- alturas de filas 1 a 20;
- anchos de columnas A:J;
- `defaultRowHeight`, `baseColWidth`, margenes, orientacion vertical,
  combinaciones y `freeze_panes`.

La diferencia numerica documentada frente al valor cacheado `BG!L47` del
manual es `0.030337` en valor absoluto (`-0.059663` programatico frente a
`-0.029663` cacheado); permanece dentro de la tolerancia contractual menor a
`1` y no cambia el resultado de cuadre.

La suite automatizada final ejecuto `36` pruebas y termino en `OK`. La matriz
de escenarios y sus advertencias/errores esperados esta en
[`docs/scenario_matrix.md`](scenario_matrix.md).

LibreOffice no esta disponible en PATH en esta VPS (`libreoffice`/`soffice` no
existen). Por eso la validacion estatica inspecciona que las formulas generadas
no contengan tokens de error como `#REF!`, `#DIV/0!`, `#VALUE!`, ni referencias
a workbooks externos, pero no afirma que las formulas hayan sido evaluadas.
El workbook se configura con `calcMode=auto`, `fullCalcOnLoad=true` y
`forceFullCalc=true` para recalculo al abrirlo en un motor compatible. La
validacion de cuadre se calcula programaticamente sobre la balanza normalizada y
se reporta aparte del XLSX generado.

## Celdas con formulas

Subtotales ER:

- `ER!H20`: ingresos netos, `=SUM(H18:H19)`.
- `ER!H25`: utilidad bruta, `=H20-H23`.
- `ER!H51`: gastos de operacion, `=SUM(H28:H50)`.
- `ER!H53`: utilidad o perdida de operacion, `=H25-H51`.
- `ER!H58`: total otros ingresos, `=SUM(H56:H57)`.
- `ER!H63`: total R.I.F., `=SUM(H61:H62)`.
- `ER!H65`: resultado antes de impuestos, `=H53+H58+H63`.
- `ER!H70`: resultado del ejercicio, `=H65-H67-H68`.

Porcentajes ER:

- Columna `ER!J` en filas `18`, `19`, `20`, `23`, `25`, `28` a `47`,
  `50`, `51`, `53`, `56`, `57`, `58`, `61`, `62`, `63`, `65`, `67`, `68`, `70`.
- Las columnas `ER!D` y `ER!F` se dejan sin formulas para no copiar los errores
  `#REF!` observados en el manual.

## API publica usada

- `src.workbook.detect_source_metadata(path)`
- `src.workbook.build_er_workbook(dataset, metadata=None, source_path=None)`
- `src.workbook.save_er_workbook(dataset, output_path=None, metadata=None, source_path=None)`
- `src.validation.validate_generated_workbook(workbook_or_path, balance_difference=None, tolerance=1.0)`
- `src.validation.recalculate_workbook(workbook_path, balance_difference=None, tolerance=1.0)`
- `src.validation.validate_balance_difference(difference, tolerance=1.0)`
- `src.validation.assert_generated_workbook_valid(workbook_or_path, balance_difference=None, tolerance=1.0)`
- `src.prototype.run_prototype(input_path)`
