# El editor HTML

El editor HTML abre .html / .htm con dos modos: **Vista previa** (la página renderizada) y **Código**.

- **Vista previa**: renderizado real; las hojas de estilo y las imágenes relativas se cargan junto al archivo.
- **Inspector de la vista previa**: haga clic para seleccionar un elemento, doble clic para editar su texto en el sitio, borrarlo con la barra de herramientas y Preguntar a la IA sobre la selección.
- **Modo de código**: edite el HTML; ctrl+F para buscar, Reemplazar todo guarda el marcado reescrito.
- **Guardar**: fiel byte a byte (se conservan el BOM, los CRLF y el salto de línea final); los guardados sin cambios no reescriben.
- **Zoom**: ctrl+rueda o pellizco escala la vista previa; ctrl+Z dentro de la vista previa deshace la última edición.

## La barra de herramientas

Haga clic en cualquier elemento de la vista previa y una barra de herramientas flota por encima:

![La barra de herramientas flotante sobre un elemento seleccionado](img/html-toolbar.png)

- **Archivo e historial**: Guardar, Guardar como, Deshacer, Rehacer, Buscar; el interruptor de **Autoguardado** escribe los cambios a intervalos.
- Conmutador **Vista previa / Código**; **Pantalla completa** muestra la página a pantalla completa.
- **Formato**: negrita, cursiva, aumentar/reducir el tamaño de fuente; el **panel de estilo** del elemento seleccionado (colores y más).
- **Insertar**: título, párrafo, tabla, imagen (por enlace), más.
- **Acciones de imagen** (con una imagen seleccionada): recortar, **Quitar fondo**, reemplazar, bloquear proporción.
- **Acciones de elemento** (con un elemento seleccionado en el inspector de la vista previa): eliminar, duplicar, subir/bajar.
- **Botón de IA**: abre el panel de IA; pregunte directamente sobre el elemento seleccionado.
