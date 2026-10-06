# El editor Markdown

El editor Markdown abre .md / .markdown con la experiencia de código fuente y vista previa renderizada.

- **Abrir**: desde Inicio o Archivo ▸ Abrir; la línea de comandos también funciona.
- **Editar**: edición de texto sin formato; las extensiones GFM (tablas, listas de tareas, tachado, enlaces automáticos) se renderizan en la vista previa.
- **Vista previa**: en vivo; los recursos relativos como las imágenes se resuelven junto al documento.
- **Guardar**: fiel byte a byte: el BOM, los CRLF y la presencia de un salto de línea final se conservan; un guardado sin cambios no reescribe el archivo.
- **Buscar y reemplazar**: ctrl+F busca en el código fuente; Reemplazar todo lo reescribe.
- **IA**: unos botones predefinidos permiten que el asistente reescriba, amplíe o traduzca el documento.

## La barra de herramientas

Una fila de botones encima del editor (pase el cursor para ver las descripciones):

![La barra de herramientas de Markdown](img/md-toolbar.png)

- **Archivo e historial**: Guardar, Guardar como, Deshacer, Rehacer, Buscar; el interruptor de **Autoguardado** de la derecha escribe los cambios en el disco a intervalos.
- **Botón de IA**: abre el panel de IA; a su lado están los preajustes de reescritura, ampliación y traducción.
- **Estilo de párrafo** (desplegable): cambiar entre el cuerpo del texto y los niveles de título.
- **Formato en línea**: **Negrita**, _Cursiva_, ~~Tachado~~, `Código en línea`, enlace.
- **Listas**: con viñetas, numerada, lista de tareas.
- **Insertar**: tabla, imagen, línea separadora.
- **Propiedades**: inserta el bloque YAML front matter al principio del archivo o salta hasta él.
- **Esquema**: saltar por la jerarquía de títulos.
- **Ortografía**: activa o desactiva la corrección ortográfica de este documento.

Tres ejemplos rápidos:

- **Título**: ponga el cursor en la línea ▸ desplegable de estilo de párrafo ▸ «Título 1».
- **Tabla**: haga clic en **Insertar tabla** ▸ arrastre el número de filas y columnas ▸ escriba en las celdas; la vista previa lo renderiza de inmediato.
- **Lista de tareas**: seleccione unas líneas ▸ haga clic en **Lista de tareas** ▸ cada línea pasa a ser `- [ ]` y se representa como casillas de verificación en la vista previa.

## Exportar

Menú Archivo, todo local y todo pregunta dónde poner el resultado:

- **Exportar como Word…** y **Exportar como PDF…** escriben un .docx o .pdf de verdad.
- **Exportar como imágenes…** escribe un PNG por página en la carpeta que elija.
- **Convertir y abrir en Docs** convierte a .docx y lo abre en la pestaña Docs integrada aquí en la app: no es un traspaso a nada en la nube, y la copia convertida vive en una carpeta de caché que se limpia al cabo de una semana aproximadamente.

## Vista de código fuente

La cinta de opciones lleva un conmutador **Código fuente** (localizado junto con la aplicación). Al activarlo, el editor se sustituye por el Markdown en crudo: exactamente el texto que escribe un guardado, nada embellecido, nada normalizado por debajo de usted.

- **La edición es fiel byte a byte.** Un guardado desde la vista de código fuente produce los mismos bytes que un guardado desde el editor: el BOM, los CRLF y la presencia de un salto de línea final se conservan.
- **Es el mismo documento.** Alterne entre el editor y la fuente tantas veces como quiera; la fuente es el propio texto del editor, no una copia que haya que combinar.
- **La barra de herramientas de formato no está disponible** mientras esta vista está abierta, porque la mayoría de esos botones insertan construcciones del editor que solo tienen sentido en la parte renderizada. Vuelve a aparecer cuando cierra la vista.
- **Los archivos JSON y otros archivos en modo fuente** se abren aquí directamente: no hay nada que renderizar, así que la fuente _es_ el documento.
