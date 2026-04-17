# Roster RT-UT (Web Interna)

Sistema de roster para coordinadores, con:

- Importacion de Excel (`.xlsx/.xls`).
- Grilla editable por persona y dia.
- Alertas de conflictos (cobertura y reglas basicas).
- Alertas listas para enviar por WhatsApp/Email (sitio + rango + grupos coincidentes).
- Filtros por busqueda, categoria, residencia y grupo.
- Exportacion a Excel (.xlsx) y CSV.
- Modo consulta movil (PWA instalable con cache basico).

## Requisitos

- Node.js 20+ (probado con Node 24).

## Instalacion y uso

```bash
npm install
npm start
```

Abrir en navegador:

- `http://localhost:3000`

## Deploy gratis (Link fijo) - Netlify + Supabase

Para tener un **link fijo** y que cualquiera pueda entrar desde cualquier lugar (obra, casa, datos moviles):

- Frontend: Netlify (gratis)
- API: Netlify Functions (gratis)
- Persistencia: Supabase (gratis)

### 1) Tabla en Supabase

Crear tabla `roster_store` con columnas:

- `id` (bigint) primary key
- `data` (jsonb)
- `updated_at` (timestamptz)

Para simplificar permisos: desactiva RLS para esa tabla, o usa `SUPABASE_SERVICE_ROLE_KEY` en Netlify (recomendado).

### 2) Seed del store actual

En tu PC, define variables y ejecuta:

```powershell
$env:SUPABASE_URL="https://xxxx.supabase.co"
$env:SUPABASE_SERVICE_ROLE_KEY="xxxxx"
npm run seed:supabase
```

### 3) Netlify

El repo ya trae `netlify.toml` (publica `public/` y expone la API en Functions).

Configurar variables en Netlify:

- `STORE_BACKEND=supabase`
- `SUPABASE_URL=...`
- `SUPABASE_SERVICE_ROLE_KEY=...` (o `SUPABASE_ANON_KEY` si dejaste RLS abierto)
- `SUPABASE_TABLE=roster_store` (opcional)

Link final:

- `https://tu-sitio.netlify.app`

## Flujo recomendado

1. Cargar el Excel actual desde el panel **Ingreso de Datos**.
2. Filtrar por fechas/categoria/residencia.
3. Editar celdas con click (rota codigos), `Shift+click` para limpiar.
4. Revisar alertas en **Alertas**.
5. Exportar con **Exportar CSV**.

## API principal

- `POST /api/import-excel`
- `GET /api/summary`
- `GET /api/roster`
- `POST /api/shifts`
- `GET /api/conflicts`
- `GET /api/export/roster.csv`

## Persistencia

En modo local (archivo), la informacion queda en:

- `data/store.json`

En modo deploy (Supabase), el store se guarda en la tabla `roster_store` (fila `id=1`).

## Nota

En el primer arranque, la app intenta auto-importar el primer `.xlsx` que encuentre en la carpeta del proyecto.
