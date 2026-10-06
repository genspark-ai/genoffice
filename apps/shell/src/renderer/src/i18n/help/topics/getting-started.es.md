# Inicio rápido: la interfaz y lo esencial

GenOffice es una suite ofimática que funciona por completo en su equipo: una ventana, una fila de pestañas, con seis editores: Docs (procesamiento de texto), Sheets (hojas de cálculo), Slides (presentaciones), PDF, Markdown y HTML. Los archivos son .docx / .xlsx / .pptx / .pdf reales, totalmente intercambiables con Word, Excel y PowerPoint. No necesita red.

## Descripción de la interfaz

![La pantalla de inicio](img/home-screen.png)

La ventana tiene tres partes:

- **Barra de pestañas (arriba)**: cada archivo abierto es una pestaña. La pestaña Inicio de la izquierda siempre está presente y no se puede cerrar; las demás pestañas son sus documentos. Haga doble clic en una pestaña para cambiar el nombre del archivo en el sitio.
- **Área de contenido**: el editor (o la pantalla de inicio) que corresponde a la pestaña activa.
- **Menú**: en la barra de menús del sistema en macOS, en la parte superior de la ventana en Windows/Linux. Los menús Archivo/Edición/Vista se adaptan al editor activo.

## Crear un documento

Sirve cualquiera de estas vías:

- Haga clic en una tarjeta de creación rápida de la sección [Inicio rápido](help://getting-started) de **Inicio** (AI Docs, AI Sheets, AI Slides, …).
- Menú **Archivo ▸ Nuevo**: Docs (⌘N/ctrl+N), Sheets, Slides, Markdown, HTML o PDF.
- Arrastre un archivo a la ventana, o haga doble clic en él en el explorador de archivos (si GenOffice es la aplicación predeterminada).

Un documento nuevo se abre sin título; el archivo en el disco solo se crea al guardarlo por primera vez.

## Abrir archivos

- El menú **Archivo ▸ Abrir** (⌘O/ctrl+O) abre el selector del sistema: .docx / .doc / .xlsx / .xls / .csv / .tsv / .pptx / .ppt / .pdf / .md / .html.
- Haga clic en cualquier elemento de la lista **Recientes** de la pantalla de inicio.
- `genoffice <file>` en un terminal también abre archivos.

## El modelo de guardado

- **Guardado manual**: ⌘S/ctrl+S, o Archivo ▸ Guardar / Guardar como. El primer guardado pregunta por la ubicación y el nombre.
- **El autoguardado** se activa solo después de haber guardado el archivo a mano al menos una vez: un PDF que solo se lee nunca se reescribe a escondidas. El autoguardado se dispara poco después de que cambie el contenido.
- Al cerrar una pestaña con cambios sin guardar, primero se pregunta si se quiere Guardar / Descartar / Cancelar.
- Cada escritura es atómica (archivo temporal y renombrado), así que un corte de luz no puede dejar un archivo a medias.

## Atajos de teclado frecuentes

| Acción            | macOS | Windows / Linux |
| ----------------- | ----- | --------------- |
| Nuevo documento   | ⌘N    | ctrl+N          |
| Abrir             | ⌘O    | ctrl+O          |
| Guardar           | ⌘S    | ctrl+S          |
| Cerrar pestaña    | ⌘W    | ctrl+W          |
| Abrir este manual | F1    | F1              |
| Contraer la cinta | ⌥⌘R   | Ctrl+F1         |

**Contraer la cinta de opciones** funciona en todos los editores. La fila de pestañas permanece y la banda de comandos que hay debajo se oculta; la pestaña seleccionada hace doble función como control de contracción, de modo que mientras la cinta está contraída no hay ninguna pestaña seleccionada y pulsar cualquier pestaña vuelve a mostrar la banda. Hacer doble clic en una pestaña hace lo mismo. La forma en que la dejó se recuerda editor por editor.

Los atajos dentro de cada editor (copiar formato, buscar y reemplazar, operaciones de tabla, …) están en sus capítulos; Docs incluye además un cuadro de diálogo de atajos de teclado en el que se puede buscar (**⌘/**) (véase su capítulo).

## Los atajos Option+Comando

Option+Comando es la capa que Word reserva para los saltos estructurados, y GenOffice la rellena igual. Todo lo siguiente es propio de Docs:

| Atajo (macOS)   | Qué hace               | Windows / Linux    |
| --------------- | ---------------------- | ------------------ |
| ⌥⌘1 / ⌥⌘2 / ⌥⌘3 | Título 1 / 2 / 3       | Ctrl+Alt+1 / 2 / 3 |
| ⌥⌘0             | Normal                 | Ctrl+Alt+0         |
| ⌥⌘M             | Párrafo                | Ctrl+Alt+M         |
| ⌥⌘A             | Nuevo comentario       | Ctrl+Alt+A         |
| ⌥⌘F             | Insertar nota al pie   | Ctrl+Alt+F         |
| ⌥⌘E             | Insertar nota al final | Ctrl+Alt+D         |
| ⌥⌘G             | Ir a                   | Ctrl+G             |

Dos de ellos cambian en Windows, por la misma razón que hace que Word los separe. **macOS es dueño de ⌥⌘D** — muestra y oculta el Dock —, así que la nota al final es ⌥⌘E en el Mac y Ctrl+Alt+D en cualquier otro sitio. Y **Ir a** se quita el Alt: Ctrl+G, donde la combinación del Mac sí lo lleva.

Eso deja ⌥⌘D libre para que GenOffice lo use en macOS si algún comando futuro lo necesita.

## Adónde ir después

- Dónde viven los archivos: [La pantalla de inicio](help://home-screen).
- Gestionar muchos archivos abiertos: [Pestañas y gestión de ventanas](help://tabs-and-windows).
- Dejar que la IA haga el trabajo: [El panel de IA](help://ai-panel).
- Idioma, tema, aplicaciones predeterminadas: [Configuración, idioma, tema e integraciones MCP](help://settings-integrations).
