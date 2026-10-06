# Docs: procesamiento de texto

Docs es el procesador de texto similar a Word: lee y escribe .docx reales con una paginación WYSIWYG fiel.

## La cinta de opciones

Pestañas: **Inicio / Insertar / Dibujar / Disposición / Diseño / Referencias / Revisar / Vista**, más pestañas contextuales para el objeto seleccionado (Diseño y disposición de tabla, imágenes, encabezado y pie de página).

- **Inicio**: portapapeles; fuente (incluidos tamaños y signos de énfasis de Asia oriental) con **Borrar todo el formato** y un conmutador para **Mostrar u ocultar marcas de formato**; párrafo (alineación/sangría/interlineado/listas) más **Definir nueva viñeta / Definir nuevo formato de número / Lista multinivel**, que guardan sus propios estilos de lista en el documento; estilos (Título 1-6/Normal/Cita, modificables) con un **Panel de estilos** para la lista completa.
- **Insertar**: saltos de página y de sección; tablas (una cuadrícula de filas × columnas, o **Insertar tabla…** para un tamaño exacto); imágenes, formas, cuadros de texto; **Portada** y **Página en blanco** de una galería predefinida; **Gráfico**; **Letra capital**; **WordArt**; campos (fecha, hora, número de página, total de páginas, nombre de archivo); hipervínculos, **Marcador** y referencias cruzadas; comentarios; encabezado y pie de página y números de página; símbolos y ecuaciones.
  - **Gráfico** inserta un objeto de gráfico real con sus propios datos — de barras, de líneas o circular — y no una imagen. _Editar datos_ de Word abre los números que hay detrás.
- **Dibujar**: tinta sobre la página, en el grupo **Herramientas de dibujo** — **Seleccionar** vuelve a la edición de texto, y luego **Lápiz**, **Marcador de resaltado** y **Borrador** (un clic o un barrido elimina todo el trazo). Junto a ellos, **Estilo de lápiz** / **Estilo de marcador de resaltado** es un solo control con muestras de color y una fila de grosores; su etiqueta sigue a la herramienta que esté activa. La tinta se guarda en el documento como una anotación que flota sobre el texto, así que sobrevive a guardar y reabrir, y **Borrar todo** en el grupo siguiente la elimina por completo.
- **Disposición**: márgenes, orientación y tamaño del papel, columnas, sangrías y espaciado de párrafo.
- **Diseño**: temas, conjuntos de colores, marca de agua, bordes de página.
- **Referencias**: tabla de contenido (actualizable), notas al pie y notas al final, títulos, referencias cruzadas.

  ![La pestaña Referencias](img/docs-references.png)

- **Revisar**: **Editor** revisa la ortografía, la gramática y la puntuación de todo el documento; **Traducir**; corrección ortográfica; comentarios (**Comentarios con IA** trabaja los que están abiertos); control de cambios con las vistas Todas las revisiones / Revisiones simples, aceptar/rechazar y **Resumen IA de revisiones**; contar palabras; **Comparar** con otro archivo; **Proteger documento**.
- **Vista**: cinco maneras de mirar el archivo — **Diseño de impresión**, **Diseño web**, **Esquema**, **Modo de lectura** y **Vista previa de páginas**; alejar/acercar/100 %/ancho de página/una página; **Panel de IA**; **Modo oscuro**; regla, líneas de la cuadrícula y el panel de navegación; **Nueva pestaña**, **Dividir** y **Cambiar de pestaña**; el cuadro de diálogo de **atajos de teclado** en el que se puede buscar.
  - **Modo oscuro** oscurece la página y el lienzo que la rodea, nunca la cinta — la división de Word entre una superficie de edición oscura y una ventana oscura. La elección se recuerda y gana sobre el tema de la aplicación en cualquier caso.
  - **Dividir** abre un segundo panel debajo que se desplaza de forma independiente y refleja el primero; ciérralo con la × de su borde.

## El panel de navegación

**Vista ▸ panel de navegación** abre un panel lateral con la estructura de títulos del documento, un cuadro de búsqueda sobre todo el documento y una miniatura por página. Si está abierto se recuerda entre sesiones, de modo que un documento por el que se navega por la estructura sigue siendo navegable.

**La estructura** es el árbol de títulos. Haga clic derecho en un título de ella para plegarlo o reestructurarlo, y no solo para navegar:

- **Contraer / Expandir** en un título pliega todo su subárbol: el capítulo desaparece, su texto sigue en el documento.
- **Contraer todo / Expandir todo** pliega o despliega todo de una vez. En un informe largo esa es la diferencia entre una estructura legible y un muro de texto.
- **Mostrar niveles de título** filtra el árbol a las profundidades que le interesan, así que _Mostrar título 1_ le deja una tabla de contenido que puede recorrer de verdad.
- **Aumentar nivel / Disminuir nivel** cambian el nivel del título, y con él el nivel que heredan todos los títulos que están debajo: así un capítulo se convierte en una sección.
- **Nuevo título antes / después** inserta uno en la posición del cursor, sin salir del panel.
- **Eliminar** quita el título _y todo lo que hay debajo_, que es el único con el que hay que tener cuidado: borra un subárbol, no una línea.
- **Seleccionar título y contenido** selecciona desde el título hasta el final de su subárbol, listo para editar una sección entera.

## Menú contextual

Haga clic derecho en cualquier parte del cuerpo: el menú se ajusta a lo que ha pulsado. Principales grupos:

- **Portapapeles**: cortar / copiar / pegar / **Texto Unicode sin formato**.
- **Fuente, párrafo**: cambiar familia y tamaño, negrita/cursiva/subrayado, alineación/sangría/interlineado sin pasar por la cinta de opciones.
- **Sinónimos**: muestra sinónimos de la palabra seleccionada; haga clic en uno para sustituirla.
- **Traducir** (IA): traduce la selección al idioma de destino (inglés, chino simplificado, japonés, coreano, francés, alemán, español, …) a través del panel de IA.
- **Nuevo comentario**: adjunta un comentario a la selección.
- **Ortografía** (sobre una palabra mal escrita): sustituciones sugeridas, ignorar todo, agregar al diccionario, definir el idioma de corrección.
- **Hipervínculo**: abrir / editar / copiar vínculo / quitar hipervínculo.
- **Imagen**: ver imagen, guardar imagen como, **ajustar texto** (en línea / cuadrado a la izquierda y derecha / arriba y abajo / detrás del texto / delante del texto), orden de disposición.
- **Campos** (tabla de contenido, números de página): actualizar campo / mostrar códigos de campo / editar campo.
- **Numeración** (dentro de una lista): reiniciar numeración / continuar numeración / cambiar nivel de lista / establecer valor de numeración.
- **Tabla** (cursor dentro de una tabla): insertar filas/columnas, combinar / dividir celdas, dividir tabla, ajustar automáticamente, alineación de celdas, distribuir filas/columnas, propiedades de la tabla, menú de eliminación, seleccionar.

## Edición

- Buscar y reemplazar (ctrl+F / ctrl+H): distinguir mayúsculas y minúsculas, palabras completas, expresiones regulares.
- Copiar formato; deshacer/rehacer en varios niveles; opciones de pegado.
- Tablas: combinar/dividir celdas, operaciones de filas y columnas, bordes y sombreado, ordenar, fórmulas.
- Imágenes: ajustar texto, recortar, comprimir; lienzo de dibujo.

## Tipografía de Asia oriental

- La compresión de signos de puntuación y las reglas de salto de línea (kinsoku) coinciden con Word; conversión de ancho completo a medio.
- Las fuentes candidatas cubren los nombres de familia habituales de Asia oriental en Windows y macOS.

## IA

- Botón de IA en la cinta y panel lateral: reescribir, ampliar, traducir, resumir, insertar tablas predeterminadas, además de instrucciones libres.
- Cada turno de la IA guarda primero una instantánea; puede volver atrás desde la lista de versiones, y ese propio retroceso se puede deshacer.

## Guardar y exportar

- Guarda .docx reescribiendo solo los párrafos modificados: el contenido intacto queda idéntico byte a byte.
- Exporta a PDF (con la paginación actual) y a imágenes página por página.

## Imprimir

ctrl+P a través del cuadro de diálogo del sistema, con páginas WYSIWYG.
