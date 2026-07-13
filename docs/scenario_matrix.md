# Matriz de escenarios contables y de entrada

## Alcance

Esta matriz cubre el parser de balanzas Auditalo, el motor de mapeo de `ER` y la
validacion programatica de cuadre. La fuente funcional es el contexto Obsidian
`Portal_CI_Estados_Financieros`, especialmente:

- `Analisis - Ejemplos XLSX Auditalo y formato ER.md`;
- `Mapeo - Balanza Auditalo a ER.md`;
- `Contrato - Balanza XLSX de entrada.md`.

El archivo manual `EEFF_202602_AL_Serv_Prueba.xlsx` se usa como referencia de
formato y de importes esperados, no como fuente del periodo de ejecucion.

## Matriz

| ID | Escenario | Entrada | Resultado esperado | Advertencia / error | Cobertura |
| --- | --- | --- | --- | --- | --- |
| P01 | `.xls` con contenido OOXML | Extension `.xls`, firma ZIP `PK`, hoja `Balanza` | Se detecta `ooxml_zip` y se procesa con `openpyxl` | Sin advertencia | `test_detects_ooxml_content_even_with_xls_extension` |
| P02 | Periodo coincidente | Nombre `..._2026_07.xls`, contenido `Periodo: 2026-07` | Procesamiento normal; `period_ym=2026-07` | Sin error | `test_normalizes_auditalo_balanza_rows` |
| P03 | Periodo inconsistente | Nombre `..._2026_06.xls`, contenido `Periodo: 2026-07` | No se genera ER | `ValueError: Filename period ... differs` | `test_rejects_filename_period_that_differs_from_sheet` |
| P04 | RFC faltante en contenido | Nombre con RFC valido, empresa sin RFC en la hoja | Se usa RFC del nombre y se conserva la razon social | `rfc_no_encontrado_en_contenido` | `test_warns_when_content_rfc_is_missing_and_keeps_filename_rfc` |
| P05 | RFC inconsistente | RFC del nombre distinto al RFC de la hoja | No se genera ER | `ValueError: Filename RFC ... differs` | `test_rejects_filename_rfc_that_differs_from_sheet` |
| P06 | Acumuladora y detalle | `6110` junto con `6110-0001`, `6110-0002` | Solo se agregan los detalles; no se duplica la acumuladora | Sin error | `test_leaf_accounts_prevent_accumulator_double_counting` |
| P07 | Cuenta esperada ausente | Falta una cuenta mapeada, por ejemplo `6147` | La linea ER queda en `0` y el proceso continua | `cuenta_no_encontrada` | `test_missing_expected_accounts_return_zero_and_warning` |
| P08 | Saldo negativo | Cuenta de ingreso con saldo negativo o saldo financiero con regla de signo | Se conserva el signo definido por la politica del rubro | Sin error; puede producir resultado negativo | `test_negative_account_and_financial_sign_policy_are_preserved` |
| P09 | Ingresos en cero | No existe cuenta `4110` | Porcentajes de ER quedan en `0`; no hay division entre cero | `cuenta_no_encontrada` para cuentas faltantes | `test_zero_income_produces_zero_percentages_without_division_error` |
| P10 | Costo o gastos en cero | Solo hay ingresos, sin cuentas `5`/`6` requeridas | Costo y gastos quedan en `0`; utilidad bruta y resultado se calculan | Advertencias de cuentas faltantes | Cubierto por P07, P09 y flujo completo |
| P11 | Resultado negativo | Gastos/costos mayores que ingresos, como el ejemplo manual | `resultado_ejercicio < 0`; se conserva en `H70` | Sin bloqueo por signo negativo | `test_builds_er_dataset_from_sample_by_account_code` |
| P12 | Encabezado repetido | Fila `Cuenta / Saldo Inicial / Debe / Haber / SaldoFinal` dentro de datos | Se ignora la fila y no se convierte en cuenta | `encabezado_repetido_ignorado` | `test_ignores_repeated_headers_records_blank_rows_and_warns_on_duplicate_accounts` |
| P13 | Cuenta repetida | Mismo codigo de detalle en mas de un renglon | Se conservan y agregan todos los renglones | `cuenta_repetida_agregada` | `test_ignores_repeated_headers_records_blank_rows_and_warns_on_duplicate_accounts` |
| P14 | Filas vacias intermedias | Renglones sin valores entre cuentas | Se omiten del calculo y se registran en `empty_rows` | Sin error | `test_ignores_repeated_headers_records_blank_rows_and_warns_on_duplicate_accounts` |
| P15 | `Varios` compuesto | `6148`, `6176`, `6195` | `H46 = 1,173.30 + 38,441.61 + 0.00 = 39,614.91` | Si falta alguna, `cuenta_no_encontrada` y se suma lo disponible | `test_varios_includes_all_manual_composite_accounts` |
| P16 | Cuadre dentro de tolerancia | `abs(diferencia) < 1`, por ejemplo `0.999999` | `cuadra=True` | Sin error | `test_balance_tolerance_is_strictly_less_than_one` |
| P17 | Cuadre fuera de tolerancia | `abs(diferencia) >= 1`, por ejemplo `1.0` | `cuadra=False`, `balanza_no_cuadra=True` | Bloqueo de validacion | `test_balance_tolerance_is_strictly_less_than_one` |

## Advertencias y errores esperados

### Advertencias no bloqueantes

- `rfc_no_encontrado_en_contenido`: el RFC del archivo no aparece dentro de la
  hoja; se conserva el RFC del nombre para la corrida.
- `periodo_no_encontrado_en_contenido`: el periodo interno no aparece; se
  conserva el periodo del nombre. Si aparece y no coincide, el caso es error,
  no advertencia.
- `encabezado_repetido_ignorado`: se encontro una fila de encabezados despues
  del inicio de datos y se descarto.
- `cuenta_repetida_agregada`: un codigo aparece en varios renglones y se
  agregaron sus saldos, sin deduplicar datos contables.
- `cuenta_no_encontrada`: una cuenta requerida por el mapeo no existe; la linea
  se resuelve a cero y el proceso puede continuar con advertencias.

### Errores bloqueantes

- `ValueError` por periodo de nombre distinto al periodo interno.
- `ValueError` por RFC de nombre distinto al RFC interno.
- `ValueError` si no existe la hoja `Balanza`.
- `ValueError` si no se encuentran las columnas requeridas.
- `ValueError` si el contenido no es un paquete OOXML ZIP soportado.
- `balanza_no_cuadra` cuando `abs(diferencia_cuadre) >= 1`.

## Ajustes aplicados

1. Se agrego `6148` al mapeo de `Varios`; el ejemplo manual ahora coincide con
   `H46=39,614.91`.
2. El parser ahora devuelve advertencias estructuradas para RFC/periodo faltante,
   encabezados repetidos y cuentas repetidas.
3. La razon social se conserva aunque el contenido no incluya RFC.
4. Las filas vacias siguen omitiendose del calculo, pero quedan registradas para
   auditoria.
5. Se mantuvo la seleccion de cuentas hoja para evitar doble conteo de cuentas
   acumuladoras.

## Criterio de estabilidad

El modulo no debe considerarse estable hasta que:

- todos los casos P01-P17 permanezcan cubiertos;
- el ejemplo completo conserve `period_ym=2026-07`, `157` filas normalizadas y
  `126` cuentas hoja usadas;
- `Varios` entregue `39,614.91`;
- la diferencia de cuadre satisfaga estrictamente `abs(diferencia) < 1`;
- el workbook generado no contenga formulas con `#REF!`, `#DIV/0!`, `#VALUE!`,
  `#N/A` o `#NAME?`.
