# PDF: lectura, anotación y censura

El editor de PDF tiene cinco pestañas en la cinta de opciones: **Inicio / Anotar / Editar / Páginas / Vista**. Lee y escribe: el texto es editable, el contenido se puede censurar y firmar, y los formularios se pueden rellenar.

## Lectura y navegación

- Barra lateral izquierda: **Miniaturas** (haga clic para saltar, el intervalo visible aparece resaltado) o **Esquema** (los marcadores, cuando existen).
- Zoom: el control de proporción de la esquina inferior derecha; ctrl+rueda ajusta el zoom por pasos.
- Rotación: por página o de todas las páginas desde el menú Páginas; las rotaciones se vuelven a escribir al guardar.
- Búsqueda: ctrl+F en texto completo, con todas las coincidencias resaltadas.
- PDF cifrados: una solicitud de contraseña (en su propia ventana pequeña) los abre; la contraseña solo se usa en esta sesión.

## Selección de texto y marcado (Anotar)

Pruébelo en cualquier párrafo:

1. **Arrastre el ratón por encima de una frase**: al soltar, flota una barra de anotación sobre ella:

![La barra de anotación tras seleccionar texto](img/pdf-highlight.png)

2. Elija **Resaltar** (la muestra amarilla abre una paleta de colores), **Subrayar** o **Tachar**; **Preguntar a la IA** envía la selección junto con su pregunta al panel de IA.
3. Para deshacer una anotación, vuelva a seleccionar el mismo fragmento arrastrándolo y haga clic en el botón activo de la barra (un conmutador al estilo de Word), o selecciónelo y pulse Supr.

Referencia:

- Al arrastrar sobre texto aparece una barra emergente: **resaltar / subrayar / tachar / copiar / Preguntar a la IA**.
- Los colores salen de la paleta; **aplicar el mismo marcado a un intervalo ya marcado lo quita** (conmutador al estilo de Word).
- Los marcados ya guardados en el archivo se pueden seleccionar y borrar (menú ⋯ o Supr).
- **Nota**: mientras hay una herramienta de dibujo activada, la capa de texto no se puede seleccionar; la herramienta se desactiva sola tras cada colocación, así que vuelve al modo selección para la siguiente acción.

## Herramientas de dibujo (Anotar)

Seis herramientas: **Dibujar, Rectángulo, Elipse, Flecha, Nota**, más **Censurar área** en la pestaña Editar.

- Cada herramienta es un conmutador: pulse para activarla; **se desactiva sola en cuanto se coloca una forma** (vuelva a pulsar la herramienta para seguir); pulsar la herramienta activa también la desactiva.
- El dibujo libre sigue el grosor del trazo; rectángulo/elipse/flecha se trazan arrastrando; los colores salen de la paleta de dibujo.
- Las formas colocadas se pueden seleccionar, borrar, arrastrar y (rectángulo/elipse) cambiar de tamaño.
- **La censura, el procedimiento completo** (para ocultar una línea de texto):

  1. Pestaña Anotar ▸ haga clic en **Censurar área** (la herramienta se activa).
  2. **Arrastre un rectángulo sobre el contenido**: queda cubierto por una marca rayada y la barra de herramientas gana los botones **borrar marcas / aplicar censura**:

  ![La página después de marcar una censura](img/pdf-redact.png)

  3. Haga clic en **Aplicar censura** y confirme: el resultado es una copia de trabajo en la que el texto y las imágenes cubiertos se han eliminado físicamente (no solo ocultos) y no se puede deshacer; el documento original queda intacto.

  ¿Se equivocó? El botón de borrar marcas elimina las marcas actuales para que pueda volver a dibujar.

## Notas adhesivas y cadenas de comentarios

- La **herramienta Nota** clava un chinchete y abre una tarjeta en el margen para ese texto (el nombre del autor es configurable); al confirmar, se guarda como una anotación de texto PDF estándar.
- Haga clic en un chinchete para abrir la cadena: **Responder** (cadenas planas al estilo de WPS/Acrobat), **Editar** su comentario, **Eliminar** un comentario o una cadena entera.
- Las ediciones en curso sobreviven hasta que el guardado escribe el texto nuevo en la misma anotación del archivo, con lo que se mantienen intactas las cadenas de respuestas.

## Editar el contenido del PDF (Editar)

- **Editar texto**: pulse el texto para editarlo bloque por bloque (motor pdfium; las fuentes se ajustan lo mejor posible).
- **Insertar texto**: coloque texto localizable con la fuente, el tamaño y el color que elija.
- **Insertar imagen / sello**.
- Formularios: los campos AcroForm se rellenan directamente; los valores se escriben al guardar.

## Firmas

- **Firma manuscrita**: dibújela; puede vincularse a un campo de firma de formulario.
- **Firma de imagen**: coloque una imagen como firma.
- Las firmas guardadas se pueden reutilizar.

## Operaciones de página (Páginas)

- **Girar / eliminar / reordenar**: arrastre las miniaturas para reordenar; la eliminación pide confirmación.
- **Extraer página**: exporte las páginas seleccionadas a un PDF nuevo.
- **Dividir**: por intervalos, en varios archivos.
- **Combinar**: añade otros PDF. Los tamaños se suman **antes** de leer nada y **todo lo que supere 1 GiB en total se rechaza** con un error legible (así la memoria queda acotada).
- Los cambios a nivel de página se vuelven a escribir en el siguiente guardado; Guardar como deja el original intacto.

## Exportación e impresión

- **PDF a Word**: conversión local a .docx.
- **Imprimir**: el orden y las rotaciones actuales a través del cuadro de diálogo del sistema; admite intervalos de páginas.

## Guardar

- El guardado normal o el autoguardado vuelven a escribir las anotaciones y las ediciones en el archivo (de forma atómica).
- **La censura pasa por su procedimiento de aplicación**, que produce una copia y deja el original intacto, para que el contenido sensible no quede en él.
