# Operaciones de archivo: cambiar nombre, eliminar, exportar

Este capítulo cubre las operaciones de archivo comunes a todos los editores; las opciones de exportación propias de cada editor están en su capítulo.

El **menú ⋯** de la fila (pase el cursor por una fila de archivo) reúne estas acciones:

![El menú ⋯ de una fila de archivo](img/file-ops.png)

## Cambiar nombre

Dos puntos de entrada, un mismo conjunto de comprobaciones:

- **⋯ ▸ Cambiar nombre** en una fila de la pantalla de inicio.
- **Haga doble clic en una pestaña de archivo** para cambiar el nombre en el sitio (véase [Pestañas y gestión de ventanas](help://tabs-and-windows)).

Reglas: la extensión se conserva automáticamente; los caracteres no permitidos, los puntos finales y los nombres reservados (CON/NUL y similares) se rechazan con un mensaje; también se bloquea un conflicto de nombres en la misma carpeta. El archivo real del disco se renombra, y la lista de recientes y los destacados se actualizan.

## Eliminar

- **⋯ ▸ Eliminar** en Inicio: mueve el archivo a la **papelera del sistema**, desde donde el sistema operativo permite restaurarlo.
- Tras eliminar aparece unos segundos una notificación con deshacer; deshacer devuelve el archivo a su sitio.

## Duplicar

**⋯ ▸ Duplicar** crea una copia llamada <name> copy en la misma carpeta y la abre en una pestaña nueva; si el nombre ya existe, se añade un contador automáticamente.

## Guardar y Guardar como

- **⌘S / ctrl+S** guarda el archivo actual; si no tiene título, primero pregunta por la ubicación y el nombre.
- **Guardar como** escribe un archivo nuevo y deja el original intacto; las ediciones posteriores se dirigen al archivo nuevo.
- Cada guardado es atómico (archivo temporal y renombrado): salir a mitad de la escritura no puede corromper el archivo.
- El autoguardado solo entra en juego tras el primer guardado manual (véase [Inicio rápido](help://getting-started)).

## Exportar a PDF

- **Docs**: Archivo ▸ Exportar como PDF (o el botón de la cinta de opciones), paginado según el diseño actual.
- **Slides**: la exportación rasteriza página a página, con progreso para presentaciones grandes.
- **Sheets**: la exportación sigue la paginación de impresión.
- Las exportaciones se renderizan en una ventana oculta y terminan donde usted elija.

## Exportar a Word / imágenes

- **PDF ▸ PDF a Word**: convierte el PDF en .docx (conversión local; los diseños complejos se conservan en la medida de lo posible).
- **Docs** puede exportar las páginas como imágenes (un PNG por página).

## Imprimir

Archivo ▸ Imprimir en cada editor (⌘P/ctrl+P) abre el cuadro de diálogo de impresión del sistema; los PDF se imprimen con el orden de páginas y las rotaciones actuales.

## Dónde viven los archivos sin título

La ubicación que elija en el primer guardado es su hogar; antes de eso el documento solo existe en la memoria. El autoguardado toma el relevo solo después de ese primer guardado.
