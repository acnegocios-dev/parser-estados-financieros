# Paridad visual de `ER`

Especificación versionada: `src/er_style_spec.json`

- Versión: `2026-07-13.manual-er-v1`
- Fuente: `sample-inputs/EEFF_202602_AL_Serv_Prueba.xlsx`
- SHA-256 de la fuente: `d530b541450f6fa9b1bfbece5c8cf4d811b97096884587edfde92156ef4ce0cc`
- El generador carga únicamente el JSON versionado; no abre el manual en runtime.

## Diferencias antes/después

| Área | Antes | Después |
| --- | --- | --- |
| Fuente tipográfica | Calibri | Arial, extraída del manual |
| Títulos `B9:B12` | pesos y tamaños propios del generador | estilos por celda del manual |
| Rellenos | azul, gris y verde decorativos | relleno equivalente al manual, sin decoración inventada |
| Bordes | bordes inferiores añadidos por el generador | bordes equivalentes a la celda manual |
| Formatos numéricos | formato MX uniforme del generador | formato exacto por celda manual |
| Geometría | filas finales con alturas distintas y ocultación incompleta | alturas, filas ocultas, merges, márgenes, orientación y DPI del manual |
| Estilos `B9:J70` | no equivalentes | 0 diferencias semánticas contra el manual |
| Importes de detalle | ya iguales dentro de `$0.01` | se mantienen iguales dentro de `$0.01` |

## Evidencia reproducible

```bash
python3 docs/extract_er_style_spec.py
python3 -m unittest discover -s tests -v
python3 -m src.prototype
```

La suite cubre valores de detalle contra los valores cacheados del manual,
`H46`, fórmulas/subtotales, geometría, estilos parametrizados, referencias
externas y tokens de error.

## Render

No se generó PDF/imagen porque `libreoffice` y `soffice` no están disponibles
en esta VPS. La validación disponible compara directamente propiedades de
OpenXML y estilos de `openpyxl`.

## Salida validada

- Archivo: `sample-outputs/estado_resultados_al_servicios_multiples_empresariales_sa_de_cv_2026_07.xlsx`
- SHA-256: `8b633e1e0ff5ac085b10052e21ad2c6d08d01fc01e39261023f8aa8c63b37643`
- Filas normalizadas: `157`
- `H46`: `39614.91`
- Diferencia de cuadre: `-0.059663`
- Errores de fórmula: `0`
- Referencias externas: `0`
