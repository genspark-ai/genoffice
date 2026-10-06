# Conectar un agente de programación

GenOffice habla el Model Context Protocol, así que un agente de programación puede leer, escribir y renderizar sus documentos con los mismos motores que usa la aplicación. El agente no está adivinando un formato de archivo: recibe los esquemas tipados de las ops a partir de las mismas definiciones que el ejecutor usa para validar.

## Registrarlo

El caso habitual es un solo comando:

```sh
genoffice mcp install all
```

Encuentra los agentes de programación de esta máquina —Claude Code, Codex, Cursor, Gemini CLI, Copilot CLI, OpenCode, Windsurf— y escribe la entrada del servidor stdio en la configuración propia de cada uno, dejando el resto de ese archivo tal como lo encontró.

```sh
genoffice mcp list             # where each agent stands
genoffice mcp install cursor    # just one
genoffice mcp uninstall cursor  # take it back out
```

Un agente que haya instalado en un sitio poco habitual acepta `--dir <path>`; `--force` reescribe una entrada que ya está ahí.

## Ejecutarlo usted mismo

Para un cliente en otra máquina, publíquelo por HTTP en su lugar:

```sh
genoffice mcp --http 8765 --token <secret>
```

`--host <addr>` cambia dónde escucha. Pase el mismo token al cliente.

Por HTTP también viajan archivos: `PUT /files/<name>` sube uno, todas las herramientas aceptan una URL `http(s)` en lugar de una ruta, y los resultados vuelven como URL de descarga —y, cuando son lo bastante pequeños, como recursos incrustados. Las ops, las specs y el Markdown se pasan en línea en ambos casos.

## Lo que recibe el agente

Todos los comandos son herramientas. Las interesantes:

- **`docs`, `sheet`, `slides`** — leer y editar un archivo por la vía de escritura de la propia aplicación, una **op** cada vez. Una presentación nueva sigue `deck_start`, `deck_page`, `deck_build`.
- **`render`** — un PNG por página, maquetado por el renderizador de la aplicación, para que el agente pueda mirar una diapositiva en vez de imaginársela.
- **`pdf`** — la capa de texto de un PDF, página a página, sin ningún proceso de la aplicación.
- **`info`** — metadatos y un resumen de la estructura, que suele ser la primera llamada más barata para un archivo desconocido.
- **`search`, `image`, `media`** — los proveedores configurados en la aplicación, así que el agente no necesita sus propias claves.
- **`merge`** — rellenar una plantilla `{{key}}`.

## Esquemas y un presupuesto más ajustado

`apply` y `create` anuncian sus parámetros `ops`, `cells` y `data` con el esquema tipado de cada op, generado a partir de `genoffice guide <domain> --json`. Es preciso, y no es pequeño. Un cliente con una ventana de contexto ajustada puede pedir matrices simples en su lugar:

```sh
genoffice mcp --compact-schemas
GENOFFICE_MCP_COMPACT_SCHEMAS=1 genoffice mcp
```

## La skill

Un agente que no conozca el vocabulario de las ops tendrá que adivinar. `genoffice skill` instala una skill de GenOffice en los agentes que encuentra, con la referencia y las guías de diseño: el mismo material que imprime `genoffice guide`.

## Lo que no es

El servidor MCP lee y escribe archivos. No es la ventana: no hay panel de IA, y el diálogo de actualización de la aplicación no se aplica. Si un paso necesita la ventana, abra el archivo.
