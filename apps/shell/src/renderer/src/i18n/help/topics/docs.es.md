# Docs: procesamiento de texto

Docs es el procesador de texto similar a Word: lee y escribe .docx reales con una paginación WYSIWYG fiel.

## La cinta de opciones

Pestañas: **Inicio / Insertar / Disposición / Diseño / Referencias / Revisar / Vista**, más pestañas contextuales para el objeto seleccionado (Diseño de tabla, imágenes).

- **Inicio**: portapapeles; fuente (incluidos tamaños y signos de énfasis de Asia oriental); párrafo (alineación/sangría/interlineado/listas); estilos (Título 1-6/Normal/Cita, modificables).
- **Insertar**: saltos de página y de sección, tablas (incluidas las dibujadas y las rápidas), imágenes, formas, hipervínculos, encabezado y pie de página, número de página, fecha, cuadros de texto.
- **Disposición**: márgenes, orientación y tamaño del papel, columnas, sangrías y espaciado de párrafo.
- **Diseño**: temas, conjuntos de colores, marca de agua, bordes de página.
- **Referencias**: tabla de contenido (actualizable), notas al pie y notas al final, títulos, referencias cruzadas.

  ![La pestaña Referencias](img/docs-references.png)

- **Revisar**: corrección ortográfica, comentarios, control de cambios (vistas Todas las revisiones / Revisiones simples), contar palabras.
- **Vista**: regla, líneas de la cuadrícula, panel de navegación, zoom y el cuadro de diálogo de **atajos de teclado** en el que se puede buscar.

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
