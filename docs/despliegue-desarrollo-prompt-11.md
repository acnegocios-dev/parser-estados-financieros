# Despliegue de desarrollo: BG, ER y BAL

Estado observado el 2026-07-21 por consultas de solo lectura. Este documento
no autoriza cambios de infraestructura; describe el corte requerido para que
el desarrollo sea consistente y reversible.

## Mecanismo real observado

- Backend: unidad `estados-financieros-api.service`, habilitada y activa.
  - `WorkingDirectory=/opt/n8n/apps/estados-financieros`
  - `ExecStart=/usr/bin/python3 -m src.api --host 172.18.0.1 --port 8080`
  - Límites ya declarados: `CPUQuota=50%`, `MemoryMax=512M`.
- Desarrollo web: contenedor `ci-dev`, imagen `node:20-alpine`, comando
  `sh -c "npm install && npm start"`, expuesto por Traefik en
  `dev.ci.acnegocios.mx` hacia el puerto interno 3000.
  - Su bind mount activo es `/opt/ci-acnegocios.mx/frontend:/app`.
  - Ese checkout está en `main` y tiene cambios locales; no es una copia de
    `origin/development`.
  - La definición actual no declara límites de CPU/RAM ni rotación de logs.
- Producción web: `ci-frontend` sirve, en modo lectura, el host
  `/opt/ci-acnegocios.mx/frontend/build` en Nginx. No interviene en este corte.
- Proxy de desarrollo: el `setupProxy.js` presente en el checkout montado
  reescribe `/estados-financieros` y usa la variable
  `ESTADOS_FINANCIEROS_API_URL`, con fallback a `http://172.18.0.1:8080`.

## Compatibilidad que debe conservarse

El backend de perfiles acepta sólo multipart `file` + `catalog`. La UI limpia
publicada en `development` envía ambos campos. La UI actualmente montada por
`ci-dev` sólo envía `file`; por eso no se debe reiniciar ni actualizar sólo la
unidad backend antes de cambiar el origen del bind mount a una copia limpia de
`development`.

## Secuencia de integración propuesta

Ejecutar sólo con autorización de operación porque modifica el mount de
`ci-dev`, recrea ese contenedor y reinicia la unidad backend.

```bash
# 1. Preparar un worktree limpio y verificar la rama publicada.
git -C /opt/ci-acnegocios.mx/frontend fetch origin development
git -C /opt/ci-acnegocios.mx/frontend worktree add --detach /opt/ci-acnegocios.mx/frontend-development origin/development
git -C /opt/ci-acnegocios.mx/frontend-development status --short --branch

# 2. Integrar backend sin reset/clean/stash; conserva los dos outputs no rastreados.
git -C /opt/n8n/apps/estados-financieros fetch origin feature/accounting-profiles-v1
git -C /opt/n8n/apps/estados-financieros merge --ff-only origin/feature/accounting-profiles-v1

# 3. Aplicar exactamente este hunk a /opt/n8n/docker-compose.yml antes de
#    recrear ci-dev (no modifica ci-frontend):
#
#   ci-dev:
#     image: node:20-alpine
#     container_name: ci-dev
#     cpus: 0.50
#     mem_limit: 512m
#     logging:
#       driver: json-file
#       options:
#         max-size: "10m"
#         max-file: "3"
#     working_dir: /app
#     command: sh -c "npm install && npm start"
#     volumes:
#       - /opt/ci-acnegocios.mx/frontend-development:/app
#
# 4. Validar la configuración ya editada y recrear exclusivamente ci-dev.
docker compose -f /opt/n8n/docker-compose.yml config
docker compose -f /opt/n8n/docker-compose.yml up -d --no-deps ci-dev

# 5. Recargar sólo la unidad backend ya limitada.
systemctl restart estados-financieros-api.service
```

## Validación posterior

```bash
systemctl is-active estados-financieros-api.service
curl --fail --silent --show-error http://172.18.0.1:8080/health
docker compose -f /opt/n8n/docker-compose.yml ps ci-dev
curl --fail --silent --show-error https://dev.ci.acnegocios.mx/
```

La prueba funcional debe cargar balanza y catálogo autorizados por multipart.
Un catálogo inválido debe responder 422 sin rutas ni datos crudos; no constituye
una validación financiera real para catálogos sin balanza representativa.

## Rollback

1. Restaurar en Compose el bind mount previo de `ci-dev` y recrear sólo
   `ci-dev` con los mismos límites y logging aprobados.
2. Revertir los commits backend mediante commits inversos, no `reset --hard`,
   y reiniciar exclusivamente `estados-financieros-api.service`.
3. Confirmar `/health`, el estado de `ci-dev` y que `ci-frontend` no fue
   recreado ni publicado.
