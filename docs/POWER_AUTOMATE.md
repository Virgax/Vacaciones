# Sincronizar el Excel con Power Automate

Con este flujo no hace falta registrar ninguna aplicación en Entra ID. Cada vez que RRHH guarda
`Control_Deuda_Vacaciones.xlsb` en SharePoint, Power Automate lo envía a la app, que lo valida, lo guarda
en su volumen y lo usa desde ese momento. La pantalla muestra "Sincronizado desde SharePoint" y la fecha
de modificación del archivo en SharePoint.

```
 RRHH guarda el Excel ──► SharePoint ──► Power Automate (cuando se modifica un archivo)
                                               │  Obtener contenido del archivo
                                               ▼
                                   HTTP POST /api/admin/upload ──► app en Railway (volumen /data)
```

> **Licencia**: la acción **HTTP** de Power Automate es un conector *premium*. Hace falta que la cuenta que
> crea el flujo tenga Power Automate Premium (o un plan por flujo). Si no la tienes, las alternativas son
> la lectura directa con Microsoft Graph (README, sección 2) o subir el Excel a mano desde `/admin.html`.

## Datos que necesitas

| Dato | Dónde está |
|---|---|
| URL de la app | `https://deuda-vacaciones-production.up.railway.app` |
| `ADMIN_KEY` | Railway → proyecto *AIZ Deuda Vacaciones* → servicio `deuda-vacaciones` → Variables |
| Sitio de SharePoint | `https://airlinkdistribution.sharepoint.com/sites/AirlinkDR` |
| Biblioteca y carpeta del Excel | donde RRHH guarda `Control_Deuda_Vacaciones.xlsb` |

## Crear el flujo (10 minutos)

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

## Qué pasa en cada caso

- **El Excel está abierto en Excel Online con autoguardado**: el desencadenador puede dispararse varias veces
  durante la edición. No pasa nada: cada envío reemplaza al anterior en un segundo.
- **El archivo que llega no es un Excel válido** (por ejemplo un guardado a medias): la app responde `400`, el flujo
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
