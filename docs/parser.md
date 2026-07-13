# Parser de Balanza

## API publica

- `src.period.extract_period_variables(path)`: extrae RFC y periodo desde el nombre del archivo.
- `src.period.period_variables_dict(path)`: devuelve esas variables como `dict`.
- `src.parser.detect_mime_kind(path)`: detecta el contenedor real por contenido.
- `src.parser.load_ooxml_workbook(path)`: carga un paquete OOXML con `openpyxl` desde file-like, incluso si la extension es `.xls`.
- `src.parser.parse_balanza(path, sheet_name="Balanza")`: devuelve un `ParsedBalanza`.
- `src.parser.parse_balanza_dict(path, sheet_name="Balanza")`: version serializable como `dict`.

## Contrato de periodo

El nombre debe contener el patron `RFC_YYYY_MM`, por ejemplo:

`balanza_SME170717GA0_2026_07.xls`

Variables generadas:

- `source_filename`
- `rfc`
- `period_year`
- `period_month`
- `period_ym`
- `period_compact`
- `period_last_day`
- `period_label_bg`
- `period_label_er`
- `period_label_bal`

## Contrato de filas

La hoja requerida es `Balanza` y debe contener estas columnas:

- `Cuenta`
- `Saldo Inicial`
- `Debe`
- `Haber`
- `SaldoFinal`

Cada fila contable se normaliza a:

- `source_row`
- `account_raw`
- `account_code`
- `account_name`
- `top_account`
- `saldo_inicial`
- `debe`
- `haber`
- `saldo_final`

Los importes vacios se normalizan como `Decimal("0")`; en `to_dict()` salen como texto para evitar perdida de precision.

## Metadatos de hoja

`ParsedBalanza` tambien expone:

- `detected_mime_kind`
- `company_name`
- `content_rfc`
- `content_period_ym`

El periodo y RFC del contenido se comparan contra el nombre de archivo. Si no
coinciden, el parser rechaza la corrida con `ValueError`.

## Tipo real de archivo

El parser no decide por extension. Primero clasifica el contenido como:

- `ooxml_zip`
- `ole_xls`
- `html_excel`
- `unknown`

La version inicial acepta `ooxml_zip`, incluso cuando la extension visible sea
`.xls`, como ocurre con el ejemplo de Auditalo. Los otros tipos quedan
detectados pero rechazados hasta implementar lectores especificos.

## Reportes

`ParsedBalanza.empty_rows` contiene los numeros de filas vacias posteriores al encabezado.

`ParsedBalanza.structure_issues` contiene filas no vacias con codigo de cuenta invalido o importes no numericos. Las filas descriptivas antes de la primera cuenta valida se ignoran como preambulo.

No se usa la suma global de todas las filas como cuadre final porque la balanza
incluye cuentas acumuladoras y cuentas detalle.
