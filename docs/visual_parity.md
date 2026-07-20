# Paridad visual y de impresión: `BG`, `ER` y `BAL`

## Referencias y perfiles versionados

Las especificaciones se extraen offline de
`sample-inputs/reference-2026-07-20/`; el runtime sólo carga JSON y nunca abre
un manual XLSX.

| Hoja | Spec | Versión | Rango efectivo | SHA-256 visual |
| --- | --- | --- | --- | --- |
| BG | `src/bg_style_spec.json` | `2026-07-20.bg-v2` | `A1:L47` | `991daeaa5b9f957e490e231164825640127cb850f01db865c04cfbb25e72b12c` |
| ER | `src/er_style_spec.json` | `2026-07-20.er-v2` | `A1:J70` | `991daeaa5b9f957e490e231164825640127cb850f01db865c04cfbb25e72b12c` |
| BAL | `src/bal_style_spec.json` | `2026-07-20.bal-v2` | máscara limpia `C:G` | `991daeaa5b9f957e490e231164825640127cb850f01db865c04cfbb25e72b12c` |

La máscara coloreada de BAL tiene SHA-256
`c27d3b4f40737e00a01dc83a2bc8745f6d62fd4a17cd8d78600a5ec09764dda2`.
El amarillo (`FFFFFF00`) documenta `C1:C4` y `C5:G185`; el verde
(`FF00A933`) documenta `H7:M183`. Ninguno se emite en la salida.

## Diferencias aplicadas

| Área | Resultado versionado |
| --- | --- |
| Libro | Sólo tres hojas visibles y ordenadas: `BG`, `ER`, `BAL`. |
| BG/ER | Fondo blanco sólido y cuadrícula oculta en las áreas efectivas; fuentes, bordes, formatos y geometría por celda se preservan desde el spec. |
| BG | Estructura contable por código, sin referencias físicas a BAL, `/4` ni ajuste `-0.03`; `L47=F45-L45`. |
| ER | Referencia visual vigente separada de la referencia histórica de valores cacheados; porcentajes con `IF`. |
| BAL | Reconstrucción dinámica sólo en `C:G`, bandas limpias, bordes `hair`, Arial 7 y `SUMAS IGUALES` dinámico; sin datos en `A:B` ni `H:M`. |

## Contrato de impresión

Las validaciones se ejecutan después de guardar y reabrir el archivo:

| Hoja | Área | Filas repetidas |
| --- | --- | --- |
| BG | `B7:L47` | `$7:$10` |
| ER | `B9:J70` | `$9:$15` |
| BAL | `C1:G{sum_row}`; ejemplo `C1:G165` | `$1:$6` |

Las tres requieren `fitToPage=true`, `fitToWidth=1`, `fitToHeight=0`, A4
vertical y márgenes versionados (`0.75` izquierdo/derecho/superior/inferior;
`0.25` encabezado/pie). Cuando LibreOffice y Poppler están disponibles, la
validación crea un PDF temporal, renderiza todas sus páginas y revisa títulos
repetidos, orientación, tamaño y geometría de texto. No se retiene el PDF.

## Evidencia reproducible

```bash
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest discover -s tests -v
```

Resultado más reciente: 69 pruebas aprobadas. La suite cubre hashes de
referencia, autosuficiencia de los specs, fórmulas, enlaces externos, nombres
heredados, fondos, máscara BAL, impresión serializada y regresiones contables.

El candidato local generado en `/tmp` tuvo SHA-256
`2b24e06afc849fc8e2808e01244fecd55b2cbf48d04334470556a64412645c5b`:

- 157 filas normalizadas;
- `ER!H46=39614.91`;
- diferencia programática de BG `-0.059663`;
- cero tokens de error, enlaces externos, nombres definidos y colores de
  máscara;
- `formula_recalculation_performed=false`, porque esta VPS no dispone de
  `libreoffice` ni `soffice`.

El smoke test del API candidato produjo un XLSX ZIP válido con SHA-256
`ad0a47e12d503d836d3f13be9dac372cedc5d94835cb8a732456548b8fff4934`,
coincidente con el valor reportado por la API, y perfil
`2026-07-20.exact-style-print-v2`.
