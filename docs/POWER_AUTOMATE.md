# Sincronizar el Excel con Power Automate

Con Power Automate no hace falta registrar ninguna aplicación en Entra ID. Cada vez que RRHH guarda el Excel en
SharePoint, un flujo manda los datos a la app, que los valida, los guarda en su volumen y los usa desde ese momento.
La pantalla muestra "Sincronizado desde SharePoint" y la fecha de modificación del archivo.

Hay tres maneras de montar el flujo. La **variante 3** es la que está configurada; las otras dos acaban en una acción **HTTP** hacia la app:

| | Variante 1: enviar el archivo | Variante 2: enviar las filas (*Obtener filas de una tabla*) |
|---|---|---|
| Formato del Excel | `.xlsb` o `.xlsx`, tal como está | **solo `.xlsx`**, con dos **tablas** de Excel (`TablaControl`, `TablaMovimientos`) |
| Acciones del flujo | Obtener contenido del archivo + HTTP | Obtener filas ×2 + HTTP |
| Endpoint de la app | `POST /api/admin/upload` | `POST /api/admin/rows` |
| Cambios en el Excel | ninguno | convertir a `.xlsx` y crear las tablas una vez |

```
 RRHH guarda el Excel ──► SharePoint ──► Power Automate (cuando se modifica un archivo)
                                               │  variante 1: Obtener contenido del archivo
                                               │  variante 2: Obtener filas de TablaControl y TablaMovimientos
                                               ▼
                                   HTTP POST a la app en Railway ──► se guarda en el volumen /data
```

> **Licencia**: la acción **HTTP** de Power Automate es un conector *premium* en las dos variantes. Hace falta que la
> cuenta que crea el flujo tenga Power Automate Premium (o un plan por flujo). Si no la tienes, las alternativas son
> la lectura directa con Microsoft Graph (README, sección 2) o subir el Excel a mano desde `/admin.html`.

## Variante 3 (la que está en marcha): la app llama al flujo

Es la más sencilla y la que queda configurada en Railway. El flujo no envía nada: **responde** con las filas de la tabla
cuando la app lo llama. La app lo llama cuando alguien consulta (como mucho una vez por minuto), guarda la última respuesta
en su volumen y, si Power Automate o Internet fallan, sigue respondiendo con esa copia.

```
 Empleado escanea ──► app en Railway ──► POST a la URL del flujo ──► Obtener filas (TablaControl) ──► Respuesta JSON
                              │
                              └─► copia en /data (failsafe)
```

El flujo tiene tres pasos:

1. Desencadenador **Cuando se recibe una solicitud HTTP** (método POST; no hace falta esquema del cuerpo).
2. **Obtener filas de una tabla** (Excel Online (Business)): el `.xlsx` de SharePoint y la tabla `TablaControl`.
   En ⋯ → *Configuración* activa **Paginación** con umbral `5000`.
3. **Respuesta**: código `200`, cabecera `Content-Type: application/json`, cuerpo = contenido dinámico **value** de
   *Obtener filas* (la app también acepta la salida completa de la acción, con `value` dentro).

Al guardar, el desencadenador muestra la **URL de POST**. Esa URL lleva una firma (`sig=`) que da acceso al flujo:
se guarda solo en Railway, en la variable `FLOW_URL`, y no se comparte. Variables relacionadas:

| Variable | Valor |
|---|---|
| `FLOW_URL` | la URL del desencadenador HTTP del flujo |
| `FLOW_METHOD` | `POST` (por defecto) o `GET` |
| `MIN_CHECK_SECONDS` | `60`: no se vuelve a llamar al flujo antes de este tiempo, aunque haya muchas consultas |
| `REFRESH_MINUTES` | `10`: refresco en segundo plano aunque nadie consulte (0 para desactivarlo) |

Cada llamada es una ejecución del flujo (cuenta para los límites diarios de Power Automate). Con el límite de una por
minuto y el refresco cada 10 minutos, el máximo teórico son unas 1 500 ejecuciones al día; en la práctica, muchas menos.

En pantalla, la pastilla dice **SharePoint en vivo** cuando el flujo responde y **Sin conexión · copia guardada** si falla.
El pie muestra "Último cambio del Excel detectado": la hora en la que la app vio por primera vez la versión actual de la tabla.

---

Las variantes 1 y 2 de abajo son la alternativa "por empuje": el flujo se dispara al modificar el archivo y manda los datos
a la app. Sirven si en algún momento prefieres no exponer una URL de flujo.

## Datos que necesitas

| Dato | Dónde está |
|---|---|
| URL de la app | `https://deuda-vacaciones-production.up.railway.app` |
| `ADMIN_KEY` | Railway → proyecto *AIZ Deuda Vacaciones* → servicio `deuda-vacaciones` → Variables |
| Sitio de SharePoint | `https://airlinkdistribution.sharepoint.com/sites/AirlinkDR` |
| Biblioteca y carpeta del Excel | donde RRHH guarda `Control_Deuda_Vacaciones.xlsb` |

## Variante 1: enviar el archivo (10 minutos)

1. Entra en <https://make.powerautomate.com> con la cuenta de la empresa → **Crear** → **Flujo de nube automatizado**.
   Nombre: `Sincronizar Control Deuda Vacaciones`. Desencadenador: busca **SharePoint** y elige
   **Cuando se modifica un archivo (solo propiedades)**. *Crear*.

2. **Desencadenador**
   - *Dirección del sitio*: `https://airlinkdistribution.sharepoint.com/sites/AirlinkDR`
   - *Nombre de la biblioteca*: la biblioteca donde está el Excel (p. ej. *Documentos*).
   - *Carpeta* (opcional): la carpeta del Excel. Así el flujo no se dispara por otros archivos.
   - En los *Ajustes* del desencadenador (⋯ → Configuración) puedes bajar el intervalo de sondeo a 1 minuto.

3. **Condición** (recomendado si la carpeta tiene más archivos)
   - *Nuevo paso* → **Condición**.
   - Izquierda: contenido dinámico **Nombre de archivo con extensión** (del desencadenador).
   - Operador: *es igual a*. Derecha: `Control_Deuda_Vacaciones.xlsb`.
   - Los pasos 4 y 5 van dentro de la rama **Si es verdadero**.

4. **Obtener contenido del archivo** (SharePoint)
   - *Dirección del sitio*: la misma.
   - *Identificador de archivo*: contenido dinámico **Identificador** (del desencadenador).

5. **HTTP**
   - *Método*: `POST`
   - *URI*: `https://deuda-vacaciones-production.up.railway.app/api/admin/upload`
   - *Encabezados* (una fila por cada uno):

     | Clave | Valor |
     |---|---|
     | `x-admin-key` | la `ADMIN_KEY` de Railway |
     | `x-file-name` | contenido dinámico **Nombre de archivo con extensión** |
     | `x-file-modified` | contenido dinámico **Modificado** (del desencadenador) |
     | `x-source` | `power automate` |
     | `Content-Type` | `application/octet-stream` |

   - *Cuerpo*: contenido dinámico **Contenido del archivo** (de *Obtener contenido del archivo*).
   - En *Configuración* de la acción HTTP activa **Entradas seguras** y **Salidas seguras** para que la clave no
     quede visible en el historial de ejecuciones.

6. **Guardar** y **Probar** → *Manualmente*. Abre el Excel en SharePoint, cambia cualquier celda y guarda.
   En 1 o 2 minutos el flujo debe ejecutarse en verde y la acción HTTP devolver `200` con un cuerpo como:

   ```json
   { "ok": true, "info": { "empleados": 241, "conDeuda": 115, "movimientos": 41 }, "via": "power automate", "fileModified": "2026-11-24T13:05:12Z" }
   ```

   También lo ves en `https://deuda-vacaciones-production.up.railway.app/api/status` o pulsando **Ver estado** en `/admin.html`.

## Variante 2: enviar las filas con "Obtener filas de una tabla"

### Preparar el Excel (una sola vez, lo hace RRHH)

1. Abre `Control_Deuda_Vacaciones.xlsb` y guárdalo como **Libro de Excel (.xlsx)** en la misma carpeta de SharePoint
   (Archivo → Guardar como). A partir de ahora RRHH trabaja sobre el `.xlsx`.
2. Hoja **Control**: selecciona desde la fila de cabecera (`Código`, `Nombre`, …, `Estado`) hasta la **última fila de
   empleado**, sin incluir el título de arriba ni la fila de totales de abajo. Insertar → **Tabla**, marca *La tabla tiene
   encabezados*. En *Diseño de tabla* → *Nombre de tabla*: `TablaControl`.
3. Hoja **Movimientos**: selecciona desde la cabecera (`Fecha`, `Código`, …, `Observación`) hasta la última línea con datos
   (o unas filas más, para dejar sitio). Insertar → **Tabla**, con encabezados. Nombre: `TablaMovimientos`.
4. Guarda. Las tablas crecen solas cuando se escribe justo debajo de la última fila, así que RRHH sigue trabajando igual.
   Las filas vacías o sin código las ignora la app.

### El flujo

1. **Desencadenador**: igual que en la variante 1 (SharePoint, *Cuando se modifica un archivo (solo propiedades)*).
2. **Condición**: *Nombre de archivo con extensión* es igual a `Control_Deuda_Vacaciones.xlsx`.
3. **Obtener filas de una tabla** (Excel Online (Business)) → renómbrala `Filas Control`.
   - *Ubicación*: `https://airlinkdistribution.sharepoint.com/sites/AirlinkDR` · *Biblioteca de documentos*: la del Excel ·
     *Archivo*: `Control_Deuda_Vacaciones.xlsx` · *Tabla*: `TablaControl`.
   - *Opciones avanzadas* → *Formato de fecha y hora*: **ISO 8601**.
   - ⋯ → *Configuración* → **Paginación**: activada, umbral `5000` (sin esto la acción devuelve como máximo 256 filas).
4. Otra **Obtener filas de una tabla** → renómbrala `Filas Movimientos`, mismos datos con *Tabla*: `TablaMovimientos`
   y la misma paginación.
5. **HTTP**
   - *Método*: `POST` · *URI*: `https://deuda-vacaciones-production.up.railway.app/api/admin/rows`
   - *Encabezados*: `x-admin-key` = la `ADMIN_KEY` · `Content-Type` = `application/json` ·
     `x-file-modified` = contenido dinámico **Modificado** (del desencadenador) · `x-source` = `power automate (filas)`
   - *Cuerpo*: escribe esto y, donde pone `value`, inserta el contenido dinámico **value** de cada acción
     (aparecen como *Filas Control → value* y *Filas Movimientos → value*):

     ```json
     {
       "control": [value de Filas Control],
       "movimientos": [value de Filas Movimientos]
     }
     ```

     Si el diseñador lo pone entre comillas, bórralas: `"control": value` tiene que quedar sin comillas alrededor del token.
   - *Configuración* de la acción HTTP → **Entradas seguras** y **Salidas seguras**.
6. **Guardar** y **Probar**. La respuesta `200` es igual que en la variante 1 pero con `"via": "power automate (filas)"`.

**Solo con la tabla Control**: si no quieres mandar los movimientos, deja una sola acción *Obtener filas* y en el cuerpo
pon `{ "control": value }` (o directamente el contenido dinámico **body** de la acción). La consulta funciona igual; lo único
que no aparece en pantalla es el historial de horas extras, porque los importes pagados salen de las columnas de Control.

La app acepta las claves tal como vienen de Excel (`Código`, `Días adeudados iniciales`, `Deuda restante (RD$)`…),
los números como texto, y las fechas en ISO 8601 o como número de serie de Excel. Ignora `@odata.etag` e `ItemInternalId`.

## Qué pasa en cada caso

- **El Excel está abierto en Excel Online con autoguardado**: el desencadenador puede dispararse varias veces
  durante la edición. No pasa nada: cada envío reemplaza al anterior en un segundo.
- **Lo que llega no es válido** (un guardado a medias, o faltan columnas en la tabla): la app responde `400`, el flujo
  queda en rojo y se **conserva la versión anterior**. El siguiente guardado lo arregla.
- **Railway está caído en el momento del guardado**: la ejecución falla; vuelve a ejecutarla desde el historial
  del flujo (*Volver a enviar*) o guarda el Excel otra vez.
- **Se reinicia el servicio**: arranca con la última copia del volumen; no hace falta reenviar nada.
- **Primera carga**: si no quieres esperar a que RRHH guarde, ejecuta el flujo con *Probar* o sube el archivo una
  vez desde `/admin.html`.

## Probar el endpoint sin Power Automate

Desde PowerShell en Windows, con el archivo en el escritorio:

```powershell
Invoke-RestMethod -Method Post -Uri "https://deuda-vacaciones-production.up.railway.app/api/admin/upload" `
  -Headers @{ "x-admin-key" = "TU_ADMIN_KEY"; "x-file-name" = "Control_Deuda_Vacaciones.xlsb"; "x-source" = "prueba manual" } `
  -ContentType "application/octet-stream" -InFile "$HOME\Desktop\Control_Deuda_Vacaciones.xlsb"
```
