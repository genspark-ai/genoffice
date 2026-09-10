<p align="center">
  <a href="https://genoffice.ai/">
    <picture>
      <source srcset="../assets/readme/hero-dark.webp" media="(prefers-color-scheme: dark)">
      <img src="../assets/readme/hero.webp" alt="GenOffice: la suite ofimática de IA de código abierto con Docs, Sheets, Slides, PDF, Markdown y HTML, con un panel de IA integrado" width="100%">
    </picture>
  </a>
</p>

<h1 align="center">GenOffice</h1>

<p align="center"><b>La primera suite ofimática de IA de código abierto completa del mundo.</b><br>
Archivos Word, Excel, PowerPoint y PDF, editados por ti y tu IA, guardados de nuevo en los formatos reales.</p>

<p align="center">
  <a href="../../LICENSE"><img src="https://img.shields.io/github/license/genspark-ai/genoffice" alt="Licencia: Apache-2.0"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases/latest"><img src="https://img.shields.io/github/v/release/genspark-ai/genoffice" alt="Última versión"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases"><img src="https://img.shields.io/github/downloads/genspark-ai/genoffice/total" alt="Descargas"></a>
  <a href="https://github.com/genspark-ai/genoffice/stargazers"><img src="https://img.shields.io/github/stars/genspark-ai/genoffice?style=flat" alt="Estrellas de GitHub"></a>
</p>

<p align="center"><a href="../../README.md">English</a> · <b>Español</b> · <a href="README.pt-BR.md">Português (Brasil)</a> · <a href="README.de.md">Deutsch</a> · <a href="README.fr.md">Français</a> · <a href="README.zh-CN.md">简体中文</a> · <a href="README.zh-TW.md">繁體中文</a> · <a href="README.ko.md">한국어</a> · <a href="README.ja.md">日本語</a> · <a href="README.ar.md">العربية</a> · <a href="README.ru.md">Русский</a> · <a href="README.it.md">Italiano</a> · <a href="README.nl.md">Nederlands</a> · <a href="README.pl.md">Polski</a> · <a href="README.cs.md">Čeština</a> · <a href="README.id.md">Bahasa Indonesia</a> · <a href="README.ms.md">Bahasa Melayu</a> · <a href="README.th.md">ไทย</a> · <a href="README.hi.md">हिन्दी</a> · <a href="README.he.md">עברית</a></p>

<p align="center">
  <a href="#download"><b>Descargar</b></a> ·
  <a href="https://genoffice.ai/"><b>Sitio web</b></a> ·
  <a href="https://genoffice.ai/join"><b>Comunidad</b></a> ·
  <a href="../../PRIVACY.md"><b>Privacidad</b></a>
</p>

GenOffice es una alternativa gratuita y de código abierto a Microsoft Office
para macOS, Windows y Linux. Abre y guarda archivos nativos `.docx`, `.xlsx`
y `.pptx`, edita PDF, Markdown y HTML, y coloca un agente de IA junto a cada
documento: no un chat pegado a un lado, sino un editor que lee el archivo,
hace el cambio y te muestra exactamente qué tocó.

- **Formatos reales, con preservación de bytes.** Solo se reescribe lo que
  editas. Todo lo demás en el archivo se conserva byte por byte, así que los
  documentos siguen funcionando en Word, Excel y PowerPoint.
- **IA que puedes revisar.** Los cambios llegan como cambios con seguimiento
  y diffs con reversión de un clic. Las hojas de cálculo obtienen fórmulas en
  vivo, no números pegados. Las presentaciones y páginas se generan en el
  lienzo y siguen siendo totalmente editables.
- **Local por diseño.** Los archivos se abren, editan, guardan y convierten
  en tu equipo. PDF → Word / Excel / PowerPoint, Markdown → Word y HTML →
  Word se ejecutan todos en el dispositivo. Solo las llamadas de IA salen de
  tu equipo, hacia el proveedor que elijas.
- **Tus claves o ninguna.** Inicia sesión con Genspark y olvídate de las
  claves, o usa tu propia clave para Claude, OpenAI, Gemini, DeepSeek, Kimi,
  GLM, Qwen, Doubao, MiniMax, Grok, Mistral, OpenRouter, o cualquier endpoint
  compatible con OpenAI, servidores locales incluidos.

**Consíguelo:** [macOS](https://github.com/genspark-ai/genoffice/releases/latest) (Apple Silicon e Intel) ·
[Windows](https://github.com/genspark-ai/genoffice/releases/latest) (x64 y Arm) ·
[Linux](https://github.com/genspark-ai/genoffice/releases/latest) (deb, rpm, AppImage) —
detalles y requisitos en [Download](#download).

## Demo

Seis apps, un solo panel de IA. Cada captura es la app real en macOS, con la
IA impulsada por el prompt que puedes leer en el panel.

### 1 · Docs — abre y edita `.docx` con una IA que puedes revisar

<table>
<tr>
<td width="50%"><img src="../assets/readme/docs-report.webp" alt="GenOffice Docs mostrando una página de informe anual a dos columnas con una imagen de portada a todo lo ancho, una tabla de KPI con sombreado, encabezado y pie de página, con un zoom del 80% y el panel de IA contraído"></td>
<td width="50%"><img src="../assets/readme/docs-ai.webp" alt="GenOffice Docs: una descripción general de la empresa con una imagen de banner; la IA ajustó la sección Overview e insertó una nueva sección con viñetas, y el panel ofrece deshacer el cambio con un solo clic"></td>
</tr>
<tr>
<td><b>Abre el archivo tal como lo compone Word</b> — secciones a dos columnas, imágenes a sangre completa, tablas con sombreado, encabezados y pies de página, con la paginación calculada según las métricas de línea de Word. Los estilos, comentarios, cambios con seguimiento, ecuaciones y trazos de tinta se conservan intactos en el ida y vuelta.</td>
<td><b>Pide el cambio</b> — la IA lee los bloques que necesita, reescribe la sección Overview e inserta una nueva sección con viñetas. Cada turno de la IA es una instantánea que puedes revertir; con <b>Track changes</b> activado, los cambios llegan como revisiones al estilo Word.</td>
</tr>
</table>

### 2 · Sheets — `.xlsx` con fórmulas y gráficos en vivo, no números pegados

<table>
<tr>
<td width="50%"><img src="../assets/readme/sheets-ai.webp" alt="GenOffice Sheets: la IA añadió una hoja Summary con ingresos por región y categoría usando fórmulas SUMIF, además de un gráfico de columnas, e informa de 43 cambios aplicados con un botón para deshacer"></td>
<td width="50%"><img src="../assets/readme/sheets-qa.webp" alt="GenOffice Sheets: al preguntar qué región lideró los ingresos del segundo trimestre, la IA responde Europa con el desglose por categoría y cita las celdas que usó como enlaces, junto a la hoja Orders"></td>
</tr>
<tr>
<td><b>Créala</b> — a partir de una sola frase, el agente añade una hoja Summary con auténticas fórmulas <code>SUMIF</code> por región y categoría, inserta un gráfico de columnas y aplica los 43 cambios como un único lote que se puede deshacer.</td>
<td><b>Pregúntale</b> — las preguntas sobre el libro se responden con el razonamiento y las celdas exactas usadas, presentadas como citas en las que se puede hacer clic. Por debajo: un motor <code>.xlsx</code> en Rust desarrollado internamente, tablas dinámicas, segmentaciones de datos, formato condicional y trazado de fórmulas.</td>
</tr>
</table>

### 3 · Slides — de un prompt a una presentación `.pptx`

<img src="../assets/readme/slides-generate.webp" alt="Time-lapse de GenOffice Slides generando el deck para inversores de Aurora Home: la IA planifica la narrativa en el panel, las diapositivas aparecen una tras otra en el lienzo, y el deck terminado concluye con la propuesta final" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/slides-cover.webp" alt="GenOffice Slides: la diapositiva de portada de una presentación para inversores de Aurora Home generada por IA en el lienzo, con la instrucción original de una línea y el resumen de la IA de lo que construyó en el panel"></td>
<td width="50%"><img src="../assets/readme/slides-ai.webp" alt="GenOffice Slides: la diapositiva de cierre diseñada de la misma presentación de 11 diapositivas, con la tira de miniaturas a la izquierda y el panel de IA resumiendo el hilo narrativo"></td>
</tr>
<tr>
<td><b>Con una sola línea</b> — «Crea una presentación de 10 diapositivas para inversores de Aurora Home…». GenOffice planifica el hilo narrativo, investiga las cifras y redacta cada diapositiva en el lienzo como un auténtico <code>.pptx</code>.</td>
<td><b>Una presentación terminada</b> — once diapositivas diseñadas con tipografía e imágenes consistentes y una llamada a la acción de cierre; sigue editando con patrones, diseños, guías inteligentes y recorte no destructivo, o pide al panel que reestilice, reescriba y reordene.</td>
</tr>
</table>

### 4 · PDF — edita texto de PDF directamente, convierte PDF a Word en el dispositivo

<table>
<tr>
<td width="50%"><img src="../assets/readme/pdf-edit.webp" alt="GenOffice PDF: el modo Editar texto resalta el contorno de cada bloque de texto de la página para editarlo directamente, mientras el panel de IA responde una pregunta sobre el informe con citas de página"></td>
<td width="50%"><img src="../assets/readme/pdf-convert.webp" alt="GenOffice Docs mostrando un documento de Word convertido localmente a partir del PDF de la revisión trimestral de Helios, abierto en una segunda pestaña junto al PDF original"></td>
</tr>
<tr>
<td><b>Edita dentro de la página</b> — el modo Edit text resalta el contorno de cada bloque de texto para volver a escribirlo directamente; el flujo de contenido se reescribe mediante PDFium con las fuentes originales, sin recurrir a una anotación que tapa el texto. Pregunta a la IA sobre un informe extenso y obtén respuestas con citas de página.</td>
<td><b>Convierte en el dispositivo</b> — <b>PDF Converter → PDF to Word</b> produce un <code>.docx</code> editable que se abre en Docs junto al original, con los encabezados, filas de estadísticas y párrafos intactos. Los destinos Excel y PowerPoint funcionan igual; las páginas escaneadas pasan por el OCR del sistema.</td>
</tr>
</table>

### 5 · HTML — un generador de páginas e interfaces con IA, primero el brief de diseño

Indica para qué es la página y para quién es. La IA propone primero un
**brief de diseño** — el gancho, la paleta, la tipografía y las direcciones
de estilo — y luego construye un único archivo `.html` autocontenido
siguiendo esos tokens.

<img src="../assets/readme/html-restyle-motion.webp" alt="Time-lapse de GenOffice HTML rediseñando la landing page de Lumen: una sola solicitud de Restyle en el panel convierte la página oscura de Midnight Studio en la cálida versión Solar Daybreak, mientras cada sección y todo el contenido permanecen intactos" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/html-ai.webp" alt="GenOffice HTML: una landing page generada para una lámpara de escritorio solar en la dirección de estilo oscuro Midnight Studio, mostrada en la vista previa en vivo con el panel de IA resumiendo la página que acaba de construir"></td>
<td width="50%"><img src="../assets/readme/html-restyle.webp" alt="La misma landing page de Lumen reestilizada por la IA en la cálida dirección Solar Daybreak: fondo tipo papel, titulares con tipografía serif y un acento naranja, conservando todas las secciones y todo el texto"></td>
</tr>
<tr>
<td><b>Generada a partir de un prompt</b> — un hero llamativo, tarjetas de funciones, precios y un formulario de lista de espera para Lumen, construido en la dirección Midnight Studio. Haz clic en cualquier elemento para reestilizarlo, doble clic para editar el texto, o cambia a la vista de código de CodeMirror.</td>
<td><b>El mismo diseño, otra dirección</b> — una sola solicitud de <b>Restyle</b> cambia los tokens del brief y la página se adapta: papel cálido, serif editorial, acento naranja sol, sin reescribir nada. Preséntala en pantalla completa, o expórtala como PDF o como un documento de Word nativo y editable.</td>
</tr>
</table>
<table>
<tr>
<td width="50%"><img src="../assets/readme/html-dashboard.webp" alt="GenOffice HTML: una interfaz de panel personal generada para una diseñadora freelance en un estilo lino cálido, con una barra lateral izquierda, un saludo en tipografía serif y cuatro tarjetas de métricas"></td>
<td width="50%"><img src="../assets/readme/html-report.webp" alt="GenOffice HTML: un informe de datos generado sobre el mercado de vehículos eléctricos en estilo periódico (broadsheet), con una cabecera serif, una cifra destacada de 17,3 millones y una fila de estadísticas"></td>
</tr>
<tr>
<td><b>Maquetas de interfaz</b> — la plantilla inicial «personal dashboard» convierte una persona en un diseño funcional: barra lateral izquierda, saludo, minigráfico de horas facturables, tarjetas de facturas y utilización, todo en HTML real que puedes entregar a un desarrollador.</td>
<td><b>Historias de datos</b> — la plantilla inicial «data report» construye un periódico editorial: cabecera serif, una cifra destacada, una fila de estadísticas separada por una línea, gráficos SVG en línea y una nota metodológica.</td>
</tr>
</table>

### 6 · Markdown — un editor de bloques sobre `.md` plano, con Ask AI

<table>
<tr>
<td width="50%"><img src="../assets/readme/markdown-ai.webp" alt="GenOffice Markdown: un párrafo seleccionado muestra un popover de Ask AI con una instrucción escrita y chips de sugerencia como Polish, Make more concise, Expand y Fix grammar, además de los botones Send now y Add to queue"></td>
<td width="50%"><img src="../assets/readme/markdown-render.webp" alt="GenOffice Markdown mostrando un documento de notas de lanzamiento con una tabla, un diagrama de flujo de Mermaid y una lista de tareas, con las sugerencias iniciales del panel de IA a la izquierda"></td>
</tr>
<tr>
<td><b>Ask AI sobre una selección</b> — selecciona cualquier pasaje y aparece un chip de <b>Ask AI</b>: escribe una instrucción o elige una sugerencia, envíala al momento, o encola varias ediciones ancladas y ejecútalas de una sola vez. El mismo punto de entrada existe en todas las apps.</td>
<td><b>Se renderiza, se guarda como Markdown plano</b> — encabezados, listas, tablas, imágenes, bloques de código y diagramas de Mermaid en un editor de bloques Tiptap, que se escriben de nuevo como <code>.md</code> plano, con una exportación <b>Markdown → Word</b> totalmente local.</td>
</tr>
</table>

## Por qué GenOffice

- **Código abierto**, Apache-2.0, desarrollado abiertamente en GitHub.
- **Se ejecuta en tu equipo.** Apps nativas para macOS, Windows y Linux; los
  archivos se quedan en tu disco y cada edición, guardado y conversión ocurre
  en tu máquina.
- **Archivos Office reales.** `.docx`, `.xlsx` y `.pptx` nativos, con
  preservación de bytes: las partes del archivo que no tocas se copian
  exactamente igual.
- **Una IA que edita el propio documento.** Cambios con seguimiento en Docs,
  fórmulas y gráficos en vivo en Sheets, diapositivas dibujadas directamente
  en el lienzo, y cada turno de la IA deja una instantánea a la que puedes
  volver.
- **Tu modelo, tu clave.** Inicia sesión con Genspark, o usa tu propia clave
  para Claude, OpenAI, Gemini, DeepSeek y más; también incluye servidores
  locales y cualquier endpoint compatible con OpenAI.
- **PDF hecho como debe ser.** Edita el texto directamente en la página y
  convierte PDF a Word, Excel o PowerPoint en el dispositivo, con OCR del
  sistema para escaneos.
- **También Markdown y HTML**, con el mismo panel de IA y exportación local a
  Word.
- **Gratis**, tanto para personas individuales como para equipos.

## Backends de IA

**Inicia sesión con Genspark** y no hay nada que configurar: las llamadas al
modelo se enrutan a través del proxy de Genspark (familias Claude, GPT y
Gemini) y los agentes obtienen búsqueda web y de imágenes, generación de
imágenes, y análisis de imagen/audio/video.

**O usa tu propia clave.** Settings → AI incluye Claude, OpenAI, Gemini,
DeepSeek, Kimi, GLM, Qwen, Doubao, MiniMax, Grok, Mistral, OpenRouter y
OpenCode Zen/Go, además de una ranura personalizada para cualquier endpoint
compatible con OpenAI (URL base + clave), incluidos servidores de modelos
locales. La búsqueda y los medios tienen sus propios proveedores por
capacidad en **AI Media & Search**: Serper o Tavily para búsqueda web, y
OpenAI, Gemini, Doubao/Seedream, GLM, Grok, Qwen, MiniMax o cualquier
endpoint de imágenes compatible con OpenAI para generación de imágenes y
análisis de imagen/video.

Toda la suite incluye temas claro, oscuro y del sistema. Los temas solo
cambian lo que se ve en pantalla: las exportaciones, las impresiones y los
archivos guardados siempre conservan los colores propios del documento.

<a id="download"></a>

## Descargar

| Plataforma                              | Requisitos                                            | Descarga                                                                                         |
| --------------------------------------- | ----------------------------------------------------- | ------------------------------------------------------------------------------------------------ |
| **macOS** — Apple Silicon (arm64)       | macOS 11+                                             | [Última versión `.dmg` (arm64)](https://github.com/genspark-ai/genoffice/releases/latest)        |
| **macOS** — Intel (x64)                 | macOS 11+                                             | [Última versión `.dmg` (x64)](https://github.com/genspark-ai/genoffice/releases/latest)          |
| **Windows** (x64, la mayoría de los PC) | Windows 10+, Intel/AMD                                | [Instalador `-x64.exe` más reciente](https://github.com/genspark-ai/genoffice/releases/latest)   |
| **Windows** en Arm (ARM64)              | Windows 11 en Arm (Snapdragon X y similares)          | [Instalador `-arm64.exe` más reciente](https://github.com/genspark-ai/genoffice/releases/latest) |
| **Linux** — Debian / Ubuntu             | x86_64, glibc 2.34+ (Ubuntu 22.04 o posterior)        | [Última versión `.deb`](https://github.com/genspark-ai/genoffice/releases/latest)                |
| **Linux** — Fedora / RHEL / openSUSE    | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+) | [Última versión `.rpm`](https://github.com/genspark-ai/genoffice/releases/latest)                |
| **Linux** — otras distribuciones        | x86_64, glibc 2.34+, FUSE 2                           | [Última versión `.AppImage`](https://github.com/genspark-ai/genoffice/releases/latest)           |

Todas las compilaciones provienen de `main`; los instaladores de macOS y
Windows están firmados. Las versiones anteriores están en la página de
[Releases](https://github.com/genspark-ai/genoffice/releases).

<details>
<summary><b>Instalación en Linux</b></summary>

El paquete deb se instala con apt: descarga las dependencias y añade
GenOffice al menú de aplicaciones:

```bash
sudo apt install ./genoffice_<version>_amd64.deb
```

En Fedora, la familia RHEL u openSUSE, instala el rpm en su lugar:

```bash
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

El AppImage se ejecuta directamente: instala el runtime de FUSE 2
(`sudo apt install libfuse2`; en Ubuntu 24.04 el paquete es `libfuse2t64`),
haz que el archivo sea ejecutable y luego ejecútalo:

```bash
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

</details>

## Cómo funciona

Siete apps de Electron — Docs, Sheets, Slides, PDF, Markdown, HTML y el
shell con pestañas — comparten una única capa de motor formada por paquetes
puros de TypeScript, más un sidecar en Rust para `.xlsx`. El archivo original
siempre es la fuente de verdad: los cambios se aplican como parches
acotados, y todo lo que el editor no tocó sobrevive intacto el ida y vuelta.

```
open docx ─► archive original by hash (never touched)
          ─► parse word/document.xml into a block tree, each block anchored to its original XML
          ─► Tiptap editor (manual + AI editing, dirty tracking)
save      ─► dirty blocks → OOXML fragments (referencing existing styles only)
          ─► splice into the original document.xml; untouched blocks keep their bytes
          ─► repack the zip; every other entry is copied byte-for-byte
```

El recorrido paquete por paquete (motores docx/pptx, `pdf2docx`,
`html2docx`, el núcleo del agente y los proveedores) está en
[CONTRIBUTING.md](../../CONTRIBUTING.md#engine-packages).

## Desarrollo

```bash
npm install
npm run fixtures     # generate test .docx fixtures
npm test             # engine + app unit tests (docs/sheets/slides need no display)
npm run typecheck    # tsc --noEmit across every workspace
npm run dev          # all six editors + shell against Vite dev servers
npm run dev:docs     # a single app (same pattern works per workspace)
npm run dist:mac     # package macOS dmg (regenerates third-party notices)
npm run dist:win     # package Windows nsis installer
npm run dist:linux   # package Linux AppImage + deb + rpm
```

La app de sheets además necesita un toolchain de Rust para su sidecar de
xlsx (`cargo` en el PATH); `npm run build -w @genoffice/sheets` lo compila
automáticamente. Consulta [CONTRIBUTING.md](../../CONTRIBUTING.md) para ver
las comprobaciones que debe pasar cada cambio y cómo se integran las pull
requests.

## Comunidad

GenOffice está en desarrollo activo y tu opinión ayuda a darle forma.

- **Reporta un error o solicita una función** en
  [GitHub Issues](https://github.com/genspark-ai/genoffice/issues).
- **Únete al chat de grupo de GenOffice** en
  [GenTeam](https://genoffice.ai/join) para hablar con el equipo y otros
  usuarios.
- **Dale una estrella al repositorio** si GenOffice te resulta útil: es la
  mejor forma de apoyar el proyecto.

## Preguntas frecuentes

<details>
<summary><b>¿Es GenOffice gratuito?</b></summary>

Sí. GenOffice es gratuito y de código abierto bajo la licencia Apache-2.0:
sin prueba gratuita, sin nivel de pago para las apps en sí.

</details>

<details>
<summary><b>¿Puede GenOffice abrir archivos de Microsoft Word, Excel y PowerPoint?</b></summary>

Sí. GenOffice abre y guarda archivos nativos `.docx`, `.xlsx` y `.pptx`. Al
guardar se preservan los bytes: las partes del archivo que no tocaste se
escriben de nuevo byte por byte, así que los documentos siguen funcionando
en Microsoft Office.

</details>

<details>
<summary><b>¿Funciona GenOffice sin conexión?</b></summary>

La edición de documentos es completamente local: los archivos nunca salen
de tu equipo para abrirse, editarse, guardarse o convertirse. Las funciones
de IA (agentes, búsqueda, herramientas de imagen) necesitan conexión a la
red, ya sea con un inicio de sesión de Genspark o con tu propia clave de API
de modelo.

</details>

<details>
<summary><b>¿Puede GenOffice editar archivos PDF?</b></summary>

Sí: una edición real de texto e imágenes en PDF que reescribe el flujo de
contenido de la página conservando las fuentes originales, sin recurrir a
anotaciones que tapan el contenido.

</details>

<details>
<summary><b>¿Puede GenOffice convertir PDF a Word, Excel o PowerPoint?</b></summary>

Sí, íntegramente en el dispositivo: extracción a nivel de carácter con
PDFium más análisis de diseño basado en geometría, sin servicios en la nube
ni subidas. Las páginas escaneadas también quedan cubiertas: en macOS y
Windows el OCR del sistema las lee, así que se convierten en texto editable
en lugar de una imagen de página.

</details>

<details>
<summary><b>¿Puedo usar mi propio modelo de IA o clave de API?</b></summary>

Sí. Además del inicio de sesión sin clave con Genspark, GenOffice admite tu
propia clave para Claude, OpenAI, Gemini, DeepSeek, Kimi, GLM, Qwen, Doubao,
MiniMax, Grok, Mistral, OpenRouter y OpenCode Zen/Go, además de cualquier
endpoint compatible con OpenAI, incluidos servidores de modelos locales. La
búsqueda, la generación de imágenes y el análisis de imagen/video usan sus
propias claves en Settings → AI Media & Search.

</details>

<details>
<summary><b>¿Puede GenOffice convertir HTML a Word?</b></summary>

Sí: Export as Word en la app de HTML produce un `.docx` nativo y editable,
íntegramente en el dispositivo. La página se renderiza en el Chromium
integrado y se reduce a estructuras reales de Word: encabezados, párrafos,
listas, tablas, tarjetas, filas de KPI, campos de formulario y fondos de
página; solo los elementos visuales sin equivalente en Word (gráficos,
iconos, cajas decoradas) se incrustan como imágenes.

</details>

<details>
<summary><b>¿GenOffice recopila algún dato?</b></summary>

Las compilaciones oficiales empaquetadas envían por defecto un análisis de
uso limitado, y puedes desactivar este reporte en cualquier momento en
Settings → General. El análisis nunca envía el contenido de los documentos,
nombres de archivo, rutas de archivo, identidad de la cuenta ni direcciones
de correo electrónico. Consulta [GenOffice Privacy](../../PRIVACY.md) para
ver la lista completa de eventos y datos.

</details>

## Seguridad

Consulta [SECURITY.md](../../SECURITY.md) para conocer la postura de
seguridad del proceso (sandboxing del renderer, validación de IPC, control
de enlaces externos) y los modelos de amenaza para el contenido generado por
IA.

## Agradecimientos

GenOffice no sería posible sin estos proyectos de código abierto:

- [Electron](https://www.electronjs.org/): el runtime de escritorio de
  todas las apps.
- [Univer](https://github.com/dream-num/univer) (Apache-2.0): el núcleo de
  interfaz de hoja de cálculo que Sheets extiende.
- [PDFium](https://pdfium.googlesource.com/pdfium/) (BSD-3-Clause, incluido
  mediante [@embedpdf/pdfium](https://github.com/embedpdf/embed-pdf-viewer)):
  el motor de flujo de contenido detrás de la edición real de texto e
  imágenes en PDF.
- [pdf.js](https://github.com/mozilla/pdf.js) (Apache-2.0) y
  [pdf-lib](https://github.com/Hopding/pdf-lib) (MIT): renderizado de PDF y
  ensamblaje de documentos.
- [Tiptap](https://tiptap.dev/) / [ProseMirror](https://prosemirror.net/):
  los editores de bloques de Docs y Markdown.
- [CodeMirror](https://codemirror.net/) (MIT): el editor de código fuente en
  HTML.
- [Konva](https://konvajs.org/): renderizado en canvas para Slides y los
  gráficos de Sheets.
- [HarfBuzz](https://github.com/harfbuzz/harfbuzz) (wasm): métricas de
  composición de texto (text-shaping) para escrituras complejas.
- [calamine](https://github.com/tafia/calamine) e
  [IronCalc](https://github.com/ironcalc/IronCalc): las capas de lectura y
  cálculo del sidecar de xlsx en Rust.
- [libeot](https://github.com/umanwizard/libeot) (MPL-2.0): el decodificador
  MicroType Express para las fuentes de PowerPoint incrustadas, portado a
  TypeScript.
- [React](https://react.dev/) (MIT): la capa de interfaz de usuario de cada
  app.
- [Mermaid](https://mermaid.js.org/) (MIT) y [KaTeX](https://katex.org/)
  (MIT): diagramas y matemáticas en Markdown y Docs.
- [opentype.js](https://opentype.js.org/) (MIT): análisis de fuentes para
  métricas y búsqueda de glifos.
- [JSZip](https://stuk.github.io/jszip/) (MIT) y
  [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser)
  (MIT): las capas de contenedor OOXML y XML.
- [Fluent UI System Icons](https://github.com/microsoft/fluentui-system-icons)
  (MIT): el conjunto de iconos de las cintas de opciones.
- [electron-updater](https://www.electron.build/) (MIT): las actualizaciones
  dentro de la app.
- Las fuentes Liberation, Carlito, Caladea y Noto CJK (OFL/Apache-2.0): las
  fuentes de documento incluidas.

`npm run notices` regenera el resumen de licencias de terceros incluido
(`tools/gen-third-party-notices.mjs`); todas las dependencias en tiempo de
ejecución son MIT/Apache-2.0/BSD-3-Clause/OFL.

## Licencia

GenOffice tiene licencia [Apache License 2.0](../../LICENSE), con una
excepción: el directorio `ee/` está reservado para futuros módulos
empresariales y está cubierto por la
[GenOffice Enterprise License](../../ee/LICENSE).

Los nombres y logotipos de GenOffice y Genspark son marcas registradas de
Mainfunc, Inc. La licencia Apache-2.0 no otorga permiso para usarlos
(consulta la sección 6); los forks deben usar su propia identidad de marca.
