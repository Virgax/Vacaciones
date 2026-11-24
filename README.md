# Consulta de deuda de vacaciones

App web (Node 20 + Express) para que cada empleado, escaneando su código de carnet, vea cuánto
debía y cuánto ha pagado de sus días de vacaciones compensados en dinero. La base de datos es el
Excel **Control_Deuda_Vacaciones.xlsb** que RRHH mantiene en SharePoint.

```
 Lector de códigos ──► navegador (kiosco) ──► Railway: esta app ──► SharePoint (Microsoft Graph)
                                                     │
                                                     └─► copia local del Excel (failsafe sin internet)
```

Lo que ve el empleado, sacado de la hoja **Control** del libro:

| Bloque | Dato | Columna del Excel |
|---|---|---|
| Deuda inicial | Días adeudados · Monto | Días adeudados iniciales · Deuda inicial (RD$) |
| Pendiente | Días pendientes · Deuda restante | Días pendientes · Deuda restante (RD$) |
| Has pagado | Días pagados = adeudados − pendientes · Deuda pagada = inicial − restante | calculado |
| Historial | horas extras aplicadas a la deuda | hoja **Movimientos** |

No se muestra el salario. Si en el Excel una celda calculada viene vacía, la app la recalcula con las
mismas fórmulas del libro (`|días| × salario diario`, `MAX(deuda − aplicado, 0)`…).

## Cómo lee el Excel (failsafe)

1. **Al arrancar** carga primero la copia local (si existe) y luego intenta SharePoint.
2. **En cada consulta** pregunta a SharePoint (una llamada ligera de metadatos) si el archivo cambió.
   Si cambió, lo descarga, lo vuelve a leer y **guarda una copia** en `DATA_DIR`.
   Para no saturar, no vuelve a preguntar antes de `MIN_CHECK_SECONDS` (15 s por defecto).
3. **Sin internet / SharePoint caído** responde con la copia guardada. La pantalla lo indica
   ("Sin conexión · copia guardada") y el pie muestra la fecha de esa copia.
4. **Cuando vuelve la conexión**, la siguiente consulta vuelve a leer en vivo sola.
5. Opcional: `REFRESH_MINUTES` revisa SharePoint en segundo plano aunque nadie consulte.

## 1. Desplegar en Railway

1. Sube esta carpeta a un repositorio propio en GitHub y en Railway: *New Project → Deploy from GitHub repo*.
   `railway.json` y `nixpacks.toml` ya definen build, arranque y healthcheck (`/health`).
2. Añade un **volumen** montado en `/data` (Settings → Volumes) para que la copia del Excel sobreviva a los redeploys.
3. Variables (Settings → Variables):

   | Variable | Para qué |
   |---|---|
   | `ADMIN_KEY` | clave para "Recargar" y "Subir Excel" en `/admin.html` (cabecera `x-admin-key`) |
   | `DATA_DIR` | `/data` (la carpeta del volumen) |
   | `SP_FILE_URL` | el enlace del Excel en SharePoint, tal cual lo abre el navegador |
   | `GRAPH_TENANT_ID`, `GRAPH_CLIENT_ID`, `GRAPH_CLIENT_SECRET` | credenciales de la app registrada en Entra ID (paso 2) |
   | `SP_FILE_PATH` | (opcional) ruta del archivo dentro de la biblioteca, p. ej. `Documentos compartidos/RRHH/Control_Deuda_Vacaciones.xlsb`; evita buscarlo por nombre |
   | `REFRESH_MINUTES` | (opcional) refresco en segundo plano; `0` = solo al consultar |
   | `MIN_CHECK_SECONDS` | (opcional) tiempo mínimo entre comprobaciones en vivo (15) |
   | `DATA_FILE` | (opcional) usar un archivo fijo en lugar de SharePoint |

4. Genera un dominio (Settings → Networking → Generate Domain) y abre la URL en el PC del kiosco.

Si todavía no tienes las credenciales de Graph, la app funciona igual subiendo el Excel desde
`https://TU-APP.up.railway.app/admin.html` (con la `ADMIN_KEY`). En cuanto configures Graph, pasa a leer en vivo.

## 2. Dar acceso a SharePoint (Microsoft Graph)

Lo hace un administrador de Microsoft 365 una sola vez, en <https://entra.microsoft.com>:

1. **Identity → Applications → App registrations → New registration**. Nombre: `Consulta Deuda Vacaciones`.
   Tipo de cuenta: solo este directorio. Sin URI de redirección.
2. En la app: **Overview** → copia *Application (client) ID* → `GRAPH_CLIENT_ID` y *Directory (tenant) ID* → `GRAPH_TENANT_ID`.
3. **Certificates & secrets → New client secret** (caduca en 24 meses máximo; apunta la fecha). El *Value* → `GRAPH_CLIENT_SECRET`.
4. **API permissions → Add a permission → Microsoft Graph → Application permissions**:
   - lo más sencillo: `Sites.Read.All` (lee cualquier sitio; solo lectura);
   - más restringido: `Sites.Selected` y luego conceder a la app acceso de lectura solo al sitio `AirlinkDR`
     (se hace con PowerShell `Grant-PnPAzureADAppSitePermission` o con Graph `POST /sites/{id}/permissions`).
   Pulsa **Grant admin consent**.
5. En Railway rellena las variables y pulsa **Recargar desde SharePoint ahora** en `/admin.html`. El estado debe decir `online: true`
   y `via: sharepoint (ruta|enlace|busqueda)`.

Cómo encuentra el archivo (en este orden): por `SP_FILE_PATH` si está; por el enlace `SP_FILE_URL`
(`/shares/{enlace}`); y si no, buscando `Control_Deuda_Vacaciones.xlsb` en las bibliotecas del sitio
`/sites/AirlinkDR`. Si RRHH mueve o renombra el archivo, se vuelve a buscar en la siguiente consulta.

## 3. El kiosco

- Abre la URL en pantalla completa (F11) en un PC o tablet con el lector de códigos conectado (los lectores
  USB funcionan como teclado y mandan Enter al final: no hay que configurar nada).
- La caja de texto siempre tiene el foco; si se está mostrando un resultado y alguien escanea, empieza una consulta nueva.
- Acepta `DR0333`, `dr333`, `333` o `0333` (los códigos numéricos de la nómina se convierten a `DR0333`).
- El resultado se borra solo a los 45 s para que no quede a la vista de la siguiente persona.
- Hay un límite de 40 consultas por minuto por dirección IP.

## API

| Método y ruta | Descripción |
|---|---|
| `GET /api/empleado/:codigo` | datos del empleado + `movimientos[]` + `datos{via, fileModified, online}`; 404 si no está, 400 si el código no es válido |
| `GET /api/status` | origen, fecha del Excel, online/offline, número de empleados |
| `POST /api/admin/refresh` | fuerza la comprobación en SharePoint (`x-admin-key`) |
| `POST /api/admin/upload` | cuerpo = el archivo .xlsb/.xlsx (`x-admin-key`, opcional `x-file-name`) |
| `GET /health` | healthcheck de Railway |

## Desarrollo local

```bash
npm install
cp .env.example .env         # rellena ADMIN_KEY y, si las tienes, las variables de Graph
DATA_FILE=./data/Control_Deuda_Vacaciones.xlsb npm start   # o con las variables de Graph
npm test
```

Node carga `.env` con `node --env-file=.env server/index.js` si prefieres no exportar variables.

> La carpeta `data/` está en `.gitignore`: contiene nombres y deudas de empleados, no la subas al repositorio.
