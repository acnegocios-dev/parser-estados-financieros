# Parser de Balanza

## API publica

- `src.period.extract_period_variables(path)`: extrae RFC y periodo desde el nombre del archivo.
- `src.period.period_variables_dict(path)`: devuelve esas variables como `dict`.
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

- `company_name`
- `content_period_ym`

El periodo del contenido se compara contra el periodo del nombre de archivo. Si
no coincide, se agrega una entrada en `structure_issues`.

## Reportes

`ParsedBalanza.empty_rows` contiene los numeros de filas vacias posteriores al encabezado.

`ParsedBalanza.structure_issues` contiene filas no vacias con codigo de cuenta invalido o importes no numericos. Las filas descriptivas antes de la primera cuenta valida se ignoran como preambulo.
