# Bienes reposeídos · RapiPréstamos

Versión del módulo **Bienes reposeídos** que funciona sin servidor propio:

- **Página:** `index.html`, publicada con GitHub Pages. No trae ningún dato: todo se pide al servidor después de entrar con usuario y contraseña.
- **Servidor y base de datos:** Google Apps Script (`apps-script/Code.gs`) pegado en una hoja de Google de la cuenta **proyectosrapiprestamos**. Cada tabla es una pestaña de la hoja (`BR_Finca`, `BR_Carga`, `BR_Avaluo`…).
- **Archivos:** los Excel originales que se suben, los modelos de plantilla, los adjuntos y los PDF de avalúos quedan en la carpeta de Drive **"Bienes reposeídos – archivos"**.

Tiene las mismas 6 secciones del sistema anterior (Cuadro y ventas, Gestión de campo, Documentos, Dashboard, Gráficas gerenciales, Alquileres), más la sección nueva **Avalúos**.

> **Datos privados:** las plantillas de minutas y el listado inicial de avalúos tienen nombres y cédulas, así que **no están en este repositorio**, porque es público. Van en el archivo `Privado.gs`, que se entrega aparte y se pega solo en Apps Script.

## Puesta en marcha (una sola vez)

1. **Hoja de Google.** Entra a Google Drive con la cuenta **proyectosrapiprestamos** y crea una hoja nueva, por ejemplo "Bienes reposeídos – base". No la compartas: la página entra por el script.
2. **Pegar el código.** En la hoja: **Extensiones → Apps Script**.
   - Borra lo que trae `Código.gs`, pega **todo** `apps-script/Code.gs` y guarda.
   - Crea un archivo nuevo con **+ → Secuencia de comandos**, ponle de nombre `Privado`, pega todo `Privado.gs` (el archivo aparte) y guarda.
3. **Instalar.** Recarga la hoja. Aparece el menú **Bienes reposeídos → Instalar**. Ejecútalo y acepta los permisos (hoja y Drive).
   - Crea las pestañas, la carpeta de archivos, las 14 plantillas y los 82 avalúos.
   - Te muestra la **clave de instalación**. Si la pierdes, está en el menú **Ver la clave de instalación**.
4. **Publicar el script.** En Apps Script: **Implementar → Nueva implementación → ⚙ → Aplicación web**.
   - *Ejecutar como:* **Yo** · *Quién tiene acceso:* **Cualquier usuario**.
   - Copia la URL que termina en **`/exec`**.
5. **Conectar la página.** En `index.html`, busca `PEGAR_AQUI_LA_URL_DEL_APPS_SCRIPT` y pon esa URL. Si la página todavía no la tiene, también se puede pegar en el recuadro que aparece al abrirla.
6. **GitHub Pages.** En el repositorio: **Settings → Pages → Branch: `main` / `(root)` → Save**. En un minuto queda en `https://rapiprestamos-prog.github.io/bienes-reposeidos-rapi/`.
7. **Primer ingreso.** La página pide la clave de instalación para crear al **administrador**: tu nombre, tu usuario y tu contraseña.
8. **Tu equipo.** En **Usuarios** crea a cada persona con su rol. La contraseña que pones es temporal y la cambian al entrar.

> Cada vez que cambies `Code.gs`: **Implementar → Gestionar implementaciones → ✏ → Versión: Nueva → Implementar**. Así la URL no cambia.

## Roles

| Rol | Puede |
|---|---|
| Administrador | Todo, más usuarios, bitácora, eliminar cargas y borrar todo |
| Gestor | Subir reportes, editar fincas, gestiones, documentos, alquileres y avalúos |
| Consulta | Ver y descargar |

## Seguridad

- Las contraseñas se guardan **cifradas** (sal + SHA-256 repetido). Nunca se guardan en texto.
- 5 intentos fallidos → bloqueo de 15 minutos.
- La sesión dura 8 horas. Cambiar la contraseña, el rol o desactivar a alguien cierra sus sesiones abiertas.
- Cada ingreso, carga, cambio y descarga queda en la pestaña **BR_Bitacora** (en la página: Usuarios → Bitácora).

## Qué entiende al subir un Excel

- **Activos disponibles para la venta** (el de contabilidad, hoja "2026" y "datos de ventas"). Si no trae la columna *Estado*, las fincas nuevas entran como disponibles y las que ya están conservan su estatus.
- Reporte de ventas (ventas en proceso, finalizadas y no concretadas, visitas, resumen ejecutivo).
- Análisis de trazabilidad.
- Cuadro de lotes con mejoras (casas).
- Control de alquileres.

Cada Excel queda en el historial de cargas con su archivo original, que se puede volver a descargar.

## Avalúos

- Cada avalúo se une solo a su finca:
  - por el número en el nombre del archivo: `FINCA 56575`, `F339402`, `30357148-4B12`, `placa 936979`;
  - si el nombre no trae número, por el nombre del deudor. Estos salen marcados para revisar.
- La finca también se puede asignar a mano con **Cambiar**.
- **Subir listado:** un CSV o Excel con las columnas *Archivo* y *Link* (como el de la carpeta AVALUOS de OneDrive). Se agregan los nuevos y se conservan los que ya estaban.
- **Agregar avalúo:** un enlace o un PDF suelto, que queda en Drive.
- El Dashboard y la ficha de cada finca muestran su PDF de avalúo.

## Archivos

| Archivo | Qué es |
|---|---|
| `index.html` | La página (GitHub Pages) |
| `apps-script/Code.gs` | El servidor (Apps Script) |
| `Privado.gs` | **No está en el repositorio.** Plantillas de documentos y listado inicial de avalúos |
