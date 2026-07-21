# Gobierno de perfiles contables y pruebas manuales de desarrollo

Este documento describe controles y evidencia, no contiene CSV, balanzas, XLSX,
rutas internas de carga ni credenciales. Un catálogo sin balanza representativa
no tiene validación financiera real.

## Ciclo de vida

| Estado | Puede hacer | No puede hacer |
| --- | --- | --- |
| `draft` | Concentrar hallazgos, sugerencias y evidencia de catálogo. | Seleccionarse en runtime ni autorizar un XLSX. |
| `approved` | Seleccionarse por RFC y vigencia si hashes y controles coinciden. | Sustituir controles de cobertura, materialidad y cuadre. |
| `retired` | Conservar trazabilidad histórica y referencia `supersedes`. | Seleccionarse para una nueva ejecución. |

El alta parte del catálogo autorizado y de la taxonomía offline. Un motor puede
sugerir `line_key` y nivel de confianza usando SAT, jerarquía y evidencia, pero
la sugerencia siempre queda en `draft`: **no resuelve, aprueba ni sustituye la
revisión contable**.

Para aprobar se requiere RFC, vigencia, identidad del catálogo (`source_sha256`
y `semantic_sha256`), versión de taxonomía, responsable/referencia de aprobación
y reglas sin ambigüedad ni doble asignación. Un cambio de catálogo, vigencia,
taxonomía, overlay o reglas crea una nueva versión; no se edita retroactivamente
la aprobada. El retiro marca la versión no seleccionable y conserva su sucesora.

## Matriz mínima de perfiles

Los hashes se muestran truncados de forma intencional. El RFC se enmascara fuera
del runtime y de los reportes de operación.

| Empresa autorizada | RFC | Perfil / versión | Vigencia | Source / semantic hash | Taxonomía / overlay | Estado |
| --- | --- | --- | --- | --- | --- | --- |
| AL Servicios (caso autorizado) | `SME•••••GA0` | `sme170717ga0-al-servicios` / `2026-07-v1` | desde `2026-07-01` | `bc6ea8f39768…` / `f20f736a8a06…` | `sat-rmf-2026-v1` / reglas por código y subárbol del catálogo autorizado | approved |

La taxonomía SAT/RMF es una base offline y versionada: el runtime no descarga
Internet. El catálogo es un overlay de empresa; `parent_code`, naturaleza y
guardas de sección dan contexto, mientras que el nombre sólo aporta sugerencias.
La precedencia es: override aprobado por código/subárbol, SAT exacto/familia,
jerarquía+naturaleza+guarda de sección, nombre como sugerencia y, finalmente,
`pending_classification`. El primer dígito nunca hace una asignación final.

## Controles antes de generar

1. Selección por RFC y vigencia; nunca por nombre de archivo.
2. Identidad exacta del catálogo por ambos hashes; una discrepancia bloquea.
3. Cada cuenta material con saldo distinto de cero debe asignarse exactamente una
   vez. Cuentas no asignadas materiales, ambigüedad, duplicidad o diferencia de
   sección de al menos $1 bloquean la salida.
4. Una cuenta faltante con saldo cero es advertencia, no corrección automática.
5. Cobertura reporta asignadas, no asignadas, ambiguas, duplicadas y diferencias
   por sección. La interfaz sólo permite descargar con perfil `approved`, 100%
   de cobertura y ceros en los tres últimos conteos.
6. Se preservan `presentation_sign`, contra-activos, resultado, acumuladores,
   fórmulas, impresión y estilos del perfil de generador correspondiente.

`generator_profile` y `accounting_profile` son identidades distintas. Para el
caso autorizado, el generador es `manual-eeff-three-sheet` versión
`2026-07-20.exact-style-print-v2`; el perfil contable es el de la matriz. Una
actualización visual no aprueba reglas contables, y una versión contable no
cambia por sí misma el layout del libro.

## Evidencia por empresa y cambios

Cada expediente de empresa debe conservar, fuera de Git: autorización de
catálogo, hashes, balanza representativa, periodo, revisión contable, resultado
de cobertura, controles de sección, aprobación y referencia de rollback. La
evidencia no se incluye en respuestas HTTP ni en esta documentación.

Para un paralelo controlado se exigen al menos tres empresas independientes con
RFC, catálogo `approved`, balanza representativa y revisión contable. Fixtures
sintéticos validan parser, jerarquía, selección y bloqueos; no validan cifras
reales.

Rollback de perfil: retirar la versión nueva, volver a seleccionar sólo la
versión previamente `approved` que coincida en RFC, vigencia y hashes, y dejar
la decisión/referencia en la bitácora. No se usa `reset`, `clean` ni corrección
automática de catálogo.

## Registro de desarrollo observado

Captura de solo lectura: `2026-07-21 UTC`.

| Componente | Activo observado | Salud / evidencia | Rollback |
| --- | --- | --- | --- |
| Backend de desarrollo | `feature/eeff-bg-er-bal` en `5bd5d4c` | unidad activa, PID `2941550`, health directo OK, CPU 50% y 512 MiB | sin cambio aplicado; conservar este SHA y reiniciar sólo la unidad si un corte autorizado falla |
| Backend candidato | `feature/accounting-profiles-v1` en `b4647ad` | 111 pruebas OK; publicado, no desplegado | no requiere rollback mientras no se integre |
| Frontend candidato | `development` en `84ad7a1` | 70 pruebas unitarias y build OK; publicado, no desplegado | no requiere rollback mientras no se publique |
| Producción web | `ci-frontend` | hash de `build/index.html`: `526b21ab…cc4858d`; sin ruta financiera | no fue reiniciado ni escrito |

El proxy financiero respondió `200` desde `https://dev.ci.acnegocios.mx/`
mediante `/estados-financieros/health`; la misma ruta en producción respondió
`404`. No obstante, el mount actual de `ci-dev` comparte la raíz
`/opt/ci-acnegocios.mx/frontend` con el build servido por producción. Por ello
los candidatos anteriores **no están desplegados** y no se ejecutó un smoke de
generación/descarga contra ellos.

## Guía breve de prueba manual en desarrollo

Precondición: ejecutar sólo después de aislar `ci-dev` en un worktree limpio de
`development` y desplegar el backend candidato según
[Despliegue de desarrollo](despliegue-desarrollo-prompt-11.md). Nunca usar el
build ni el contenedor de producción.

1. Abrir `https://dev.ci.acnegocios.mx/` y la sección **Estados financieros**.
   Confirmar que no se exponen rutas, CSV crudos ni RFC completo.
2. Caso exitoso autorizado: seleccionar la balanza representativa y el catálogo
   autorizado de AL Servicios desde almacenamiento privado. Pulsar **Procesar**.
   Verificar perfil contable enmascarado, versión, vigencia, taxonomía, hash
   semántico abreviado, generador, cobertura 100% y cero pendientes,
   ambigüedades y duplicados.
3. Descargar sólo cuando el estado sea `Aprobado` y `Descargable`. Abrir el
   XLSX en LibreOffice/Excel y revisar hojas `BG`, `ER`, `BAL`, fórmulas sin
   errores, área de impresión, estilos y los controles de cuadre acordados.
4. Casos bloqueados esperados: perfil inexistente/no aprobado, hash distinto,
   vigencia fuera de rango, cuenta material sin mapear, doble asignación,
   ambigüedad y diferencia de control. Cada uno debe devolver 422 accionable,
   no generar descarga ni exponer datos crudos.
5. En DevTools comprobar que la petición es multipart con `file` y `catalog`,
   que se dirige al proxy de desarrollo y que no hay errores de consola/red.

El resultado de estos pasos debe agregarse al expediente de la empresa. Si falla
un control, conservar la evidencia y usar el rollback de perfil o de despliegue;
no modificar el catálogo para forzar el resultado.

## Trazabilidad

- Decisiones y reglas: análisis de perfiles por empresa, mapeo de ER,
  especificación de salida BG/ER/BAL y contrato de balanza del vault operativo.
- Prompts de referencia: perfiles versionados por empresa, paridad de estilos e
  impresión y cierre de activación BG/ER/BAL.
- Código y pruebas: [accounting_profiles.py](../src/accounting_profiles.py),
  [test_accounting_profiles.py](../tests/test_accounting_profiles.py),
  [test_runtime_profile_api.py](../tests/test_runtime_profile_api.py) y
  [test_synthetic_catalog_guardrails.py](../tests/test_synthetic_catalog_guardrails.py).
- Operación y reversión: [despliegue-desarrollo-prompt-11.md](despliegue-desarrollo-prompt-11.md).
