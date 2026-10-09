# Conectar un agente de codificación

GenOffice habla el Model Context Protocol, de modo que un agente de codificación puede leer, escribir y renderizar sus documentos con los mismos motores que usa la aplicación. El agente no está adivinando un formato de archivo: recibe los esquemas tipados de las op a partir de las mismas definiciones contra las que valida el executor.

## Registrarlo desde la aplicación

**Configuración ▸ Integraciones** es el sitio donde se hace esto. El panel tiene dos mitades, y puede usar una o las dos.

**El skill.** Una fila por cada agente de codificación que se encuentre en este equipo —Claude Code, Codex, Cursor, Gemini CLI, Copilot CLI, OpenCode, Windsurf—, cada una con **Instalar**, **Actualizar** y **Desinstalar**, más **Instalar en otra carpeta…**, **Descargar skill (zip)** y **Copiar ruta**. Si su asistente no aparece en la lista, señale a GenOffice una carpeta desde la que lea `SKILL.md`, o guarde el zip y deje que el asistente lo instale. El skill y el MCP pueden convivir: el asistente elige uno, y hacen exactamente lo mismo.

**MCP.** Se ofrecen dos vías: **Iniciado por el asistente (recomendado)**, donde añade la configuración que se muestra a su cliente y el asistente arranca el servidor por su cuenta, y el **Servidor HTTP local**, que la aplicación ejecuta por usted. De cualquier manera, el asistente acaba hablando con GenOffice y usted nunca escribe una orden.

## Registrarlo desde la línea de comandos

Lo mismo desde una terminal: esta es la vía avanzada, y la que conviene usar cuando el agente está en un sitio donde el panel no puede encontrarlo:

```sh
genoffice mcp install all
```

Encuentra los agentes de codificación de este equipo y escribe la entrada del servidor stdio en la configuración de cada uno, dejando el resto del archivo tal como lo encontró.

```sh
genoffice mcp list             # where each agent stands
genoffice mcp install cursor    # just one
genoffice mcp uninstall cursor  # take it back out
```

Un agente instalado en un sitio poco habitual necesita `--dir <path>`; `--force` reescribe una entrada que ya está ahí.

Todo lo que acepta el servidor cabe en una pantalla: las formas de registro, eliminación y listado, el servicio por HTTP y las dos opciones de esquema:

![La salida real de genoffice mcp --help: las formas install, uninstall y list, con las opciones --http, --host, --token, --compact-schemas, --dir y --force](img/mcp.png)

## Ejecutarlo sin asistente

Para un cliente en otro equipo, sírvalo por HTTP en su lugar:

```sh
genoffice mcp --http 8765 --token <secret>
```

`--host <addr>` cambia dónde escucha. Pásele al cliente el mismo token.

Por HTTP también viajan los archivos: `PUT /files/<name>` sube uno, cada herramienta acepta una URL `http(s)` en lugar de una ruta, y las salidas vuelven como URL de descarga y, cuando son lo bastante pequeñas, como recursos incrustados. Las ops, las especificaciones y el Markdown se pasan en línea en ambos casos.

## Qué recibe el agente

Cada orden es una herramienta. Las interesantes:

- **`docs`, `sheet`, `slides`** — leer y editar un archivo por la ruta de escritura de la propia aplicación, una **op** cada vez. Una presentación nueva sigue `deck_start`, `deck_page`, `deck_build`.
- **`render`** — un PNG por página, maquetado por el renderizador de la aplicación, para que el agente mire una diapositiva en vez de adivinarla.
- **`pdf`** — la capa de texto de un PDF, página a página, sin proceso de la aplicación.
- **`info`** — metadatos y un resumen de la estructura, normalmente la primera llamada más barata sobre un archivo desconocido.
- **`search`, `image`, `media`** — los proveedores configurados en la aplicación, así que el agente no necesita sus propias claves.
- **`merge`** — rellenar una plantilla `{{key}}`.

## Esquemas y un presupuesto más pequeño

`apply` y `create` anuncian sus parámetros `ops`, `cells` y `data` con el esquema tipado de cada op, generado a partir de `genoffice guide <domain> --json`. Es preciso y no es pequeño. Un cliente con la ventana de contexto apretada puede pedir matrices sencillas en su lugar:

```sh
genoffice mcp --compact-schemas
GENOFFICE_MCP_COMPACT_SCHEMAS=1 genoffice mcp
```

## Por qué el skill

Un agente que no conoce el vocabulario de las ops adivinará. El skill lleva la referencia y las guías de diseño —el mismo material que imprime `genoffice guide`—, de modo que el asistente escribe ops cuya especificación realmente ha leído. Instálelo desde el panel de arriba, o con `genoffice skill` desde una terminal.

## Qué alcanza en la aplicación

El servidor no se limita a los archivos del disco. Mientras GenOffice está en marcha, el agente también puede trabajar a través de la ventana:

- **`open_in_genoffice`** abre un archivo en una pestaña y le da el foco.
- **`open_documents`** enumera todos los documentos que tiene abiertos —id, tipo, ruta y si tienen cambios sin guardar— y después lee el contenido actual de uno o lo cierra, guardando antes salvo que le pida que lo descarte.
- **Las herramientas de contenido** toman ese id (o la ruta) como argumento `document`, así que una edición cae en la pestaña que ya tenía abierta y la ventana cambia para mostrarla.

Dos cosas quedan fuera de alcance: no hay panel de IA, y el diálogo de actualización dentro de la aplicación no se aplica.

## El panel Servidor HTTP local

En **Servidor HTTP local**, la aplicación ejecuta el servidor por su cuenta en lugar de dejarlo en manos del asistente: un interruptor de activación, un campo **Puerto**, un indicador **En ejecución / Detenido** y la **Generación en segundo plano** (escribir los documentos directamente en una ruta sin abrir la interfaz) y el **Ejemplo de configuración del cliente** para copiar. Al abrir **Avanzado** se añaden las dos URL de conexión —Streamable HTTP y la antigua de SSE—, una URL de **Comprobación de estado** y un interruptor de **Registro** que anota la actividad del servidor y de las herramientas en un archivo local que usted puede **Abrir**, **Actualizar** o **Borrar** desde ahí. Solo escucha en localhost.
