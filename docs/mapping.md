# Mapeo provisional del motor de estado de resultados

Este prototipo clasifica cuentas por `account_code`, no por posicion en Excel.
La entrada esperada es una lista de filas normalizadas con estos campos:

- `source_row`
- `account_raw`
- `account_code`
- `account_name`
- `top_account`
- `saldo_inicial`
- `debe`
- `haber`
- `saldo_final`

## Reglas editables

Las reglas viven en `src/engine.py` como `MAP_RULES`. Se evaluan en orden, de
mas especifica a mas general.

| Prefijo | Rubro ER | Naturaleza | Criterio |
| --- | --- | --- | --- |
| `4` | Ventas / ingresos netos | Acreedora | `haber - debe` |
| `5` | Costo de ventas | Deudora | `debe - haber` |
| `6` | Gastos de operacion | Deudora | `debe - haber` |
| `700`, `701`, `702`, `71` | Otros ingresos | Acreedora | `haber - debe` |
| `703`, `704`, `705`, `72` | Otros gastos | Deudora | `debe - haber` |
| `730`, `731`, `732`, `74` | Productos financieros | Acreedora | `haber - debe` |
| `733`, `734`, `735`, `75`, `76` | Gastos financieros | Deudora | `debe - haber` |
| `77`, `78`, `79` | Impuestos a la utilidad | Deudora | `debe - haber` |
| `7` | Otros gastos | Deudora | Regla conservadora de cierre |

## Signos

El estado de resultados se presenta con importes positivos para ingresos,
costos y gastos segun su naturaleza contable. Las utilidades se calculan por
resta:

- Ventas / ingresos netos: positivo si `haber > debe`.
- Costos, gastos e impuestos: positivo si `debe > haber`.
- Si una fila no trae movimientos (`debe` y `haber` en cero o vacios), el motor
  usa `saldo_final` como valor de presentacion observado en Auditalo.

## Calculos equivalentes

El motor entrega importes como equivalente de columna `H`. Los porcentajes son
equivalentes de columna `J` y siempre se calculan con base en `H18`, que en este
prototipo corresponde a `ventas_ingresos_netos`.

- `utilidad_bruta = ventas_ingresos_netos - costo_ventas`
- `resultado_operacion = utilidad_bruta - gastos_operacion`
- `resultado_antes_impuestos = resultado_operacion + otros_ingresos - otros_gastos + productos_financieros - gastos_financieros`
- `resultado_ejercicio = resultado_antes_impuestos - impuestos`

## Limitaciones actuales

El catalogo definitivo todavia no esta disponible. Por eso:

- La clase `7` se separa con prefijos comunes, pero puede requerir ajuste fino.
- Devoluciones, descuentos, bonificaciones u otras contra-cuentas dentro de
  ingresos se acumulan segun su movimiento neto, no por reglas especificas de
  contra-ingreso.
- Las cuentas sin prefijo reconocido se devuelven en `unmatched_accounts` para
  revision contable.
