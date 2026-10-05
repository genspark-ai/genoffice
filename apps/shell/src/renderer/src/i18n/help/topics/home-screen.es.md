# La pantalla de inicio: dónde están sus archivos

Inicio es la página de arranque de GenOffice: una barra de navegación a la izquierda, listas de archivos y tarjetas de creación rápida a la derecha.

![La pantalla de inicio](img/home-screen.png)

## Navegación en la barra lateral

- **Recientes**: los archivos que abrió recientemente, agrupados por fecha (esta semana / este mes / anteriores).
- **Destacados**: los archivos que marcó como destacados. Pase el cursor por una fila y haga clic en la estrella para añadirlos o quitarlos.
- **Genspark Projects**: tras iniciar sesión en su cuenta de Genspark, muestra los proyectos que creó con Genspark AI en la web; haga clic en uno para seguir editando en el navegador. Admite búsqueda, ordenación por fecha, actualización y cargar más.
- **Carpetas**: ancle directorios frecuentes en la barra lateral (Añadir carpeta…) y salté a ellos como si fueran marcadores. Las raíces no disponibles se muestran como no disponibles y se pueden quitar de la lista.
- **Papelera**: apunta a la papelera del sistema; los archivos borrados acaban allí y el sistema operativo permite restaurarlos.

## La lista de archivos

Cada fila muestra un icono, el nombre del archivo, la fecha de modificación y más. El **menú ⋯** de la fila ofrece:

- **Cambiar nombre**: en el sitio, la extensión se conserva automáticamente.
- **Destacar / Quitar de destacados**
- **Duplicar**: crea una copia en la misma carpeta.
- **Eliminar**: mueve el archivo a la papelera del sistema; no es una eliminación definitiva.
- **Mostrar en la carpeta**: localiza el archivo en el explorador de archivos.

## Búsqueda

El cuadro de búsqueda de arriba cubre dos cosas a la vez:

- **Nombres de archivo**: filtrado rápido por nombre.
- **Contenido de los archivos**: GenOffice indexa sus archivos en segundo plano (el texto de los .docx/xlsx/pptx/pdf/md/html, con OCR como respaldo para los PDF escaneados), así que buscar en el cuerpo del texto también encuentra archivos. El ámbito y los interruptores se ajustan en los ajustes de búsqueda.

## Tarjetas de inicio rápido

Las tarjetas que están encima de las listas crean un documento en un solo paso. Al hacer clic en una tarjeta se crea un archivo de ese tipo y se abre su editor: puede escribir de inmediato o dejar que la IA lo redacte (cada editor tiene un **botón de IA** en la cinta de opciones y **Pedir a la IA** en el menú contextual de la selección).

El archivo nuevo va a la carpeta seleccionada en la barra lateral; si no hay ninguna, va a la carpeta predeterminada.

Qué hace cada tarjeta:

- **AI Docs** (.docx): un documento de texto en blanco en el editor Docs. El archivo se escribe en el disco solo al **guardar por primera vez**; los documentos nuevos se abren con el panel de IA desplegado (se desactiva en Configuración → «Abrir el panel de IA en documentos nuevos»).
- **AI Sheets** (.xlsx): una hoja de cálculo en blanco en el editor Sheets. Hasta que no guarde, no existe ningún archivo en el disco; el nombre queda reservado para el primer guardado. Tras la primera generación con IA, el archivo también se puede renombrar automáticamente según su contenido.
- **AI Slides** (.pptx): una presentación en blanco en el editor Slides.
- **AI Markdown** (.md): un documento Markdown en blanco en el editor Markdown.
- **AI HTML** (.html): una página web en blanco en el editor HTML.
- **AI PDF** (.pdf): distinto de los demás: **crea de inmediato** un PDF real en blanco de una sola página en la carpeta de destino y lo abre como un archivo normal (el editor PDF trabaja con archivos reales). Útil para anotar, censurar o añadir texto; el archivo se puede renombrar automáticamente según su contenido al guardarlo por primera vez.
- **Abrir archivo local**: un selector de archivos del sistema para Word (.docx/.doc), Excel (.xlsx/.xlsm/.xls/.csv/.tsv), PowerPoint (.pptx/.ppt), PDF, Markdown (.md/.markdown) y páginas web (.html/.htm). Permite seleccionar varios; cada archivo recibe su propia pestaña.

> Sugerencia: Archivo ▸ Nuevo en la barra de menús crea los mismos tipos de documento (⌘N/Ctrl+N crea un documento de texto de forma predeterminada); arrastrar un archivo a la ventana lo abre.

## Proyectos en la nube (Genspark Projects)

- El primer uso requiere iniciar sesión en su cuenta de Genspark (flujo por código de dispositivo: GenOffice muestra un código y usted completa el inicio de sesión en el navegador).
- La lista de proyectos se sincroniza con la web; Abrir en el navegador lleva allí para continuar.
- No iniciar sesión no afecta a ninguna función local.
