# Línea de comandos y agentes

Cada instalación trae un comando `genoffice` que maneja los mismos motores que la ventana: los mismos analizadores, el mismo escritor, el mismo renderizador. Un archivo que guarda la aplicación y un archivo que escribe el comando son el mismo archivo, y una comprobación que supera el panel de IA de la aplicación la supera también el comando.

```sh
genoffice --help          # every command
genoffice guide slides    # the op reference, for writing your own
```

Esa es toda la superficie en una sola pantalla: cada comando con una línea que dice lo que hace, y luego las opciones globales y los códigos de salida:

![La salida real de genoffice --help: el banner de versión, cada comando con una descripción de una línea, y las opciones globales y los códigos de salida](img/cli.png)

## Cómo obtener el comando

macOS y Windows lo incluyen dentro del paquete de la aplicación. Para usarlo por su nombre, ejecute una vez `genoffice install-cli`: crea un enlace simbólico del binario incluido en `/usr/local/bin`, o en el `PATH` de su usuario en Windows.

## Los comandos que conviene conocer

| Comando           | Qué hace                                                                                                                                                                                                                          |
| ----------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `open`            | Abre un documento en la aplicación; arranca la aplicación si no está en marcha.                                                                                                                                                   |
| `convert`         | Convierte entre formatos usando los motores de la propia aplicación.                                                                                                                                                              |
| `create`          | Crea un documento a partir de contenido estructurado.                                                                                                                                                                             |
| `render`          | Un PNG por página, tal como los maqueta el renderizador.                                                                                                                                                                          |
| `pdf`             | Lee la capa de texto de un PDF página a página, sin ningún proceso de la aplicación.                                                                                                                                              |
| `info`            | Metadatos y un resumen de la estructura de un documento.                                                                                                                                                                          |
| `search`          | Búsqueda web o de imágenes a través del proveedor configurado en la aplicación.                                                                                                                                                   |
| `image` / `media` | Genera una imagen, o describe y hace preguntas sobre un archivo de imagen, vídeo o audio.                                                                                                                                         |
| `merge`           | Rellena los marcadores de posición `{{key}}` de una plantilla `.docx`, `.pptx` o `.xlsx`.                                                                                                                                         |
| `capabilities`    | Informa de qué funciones en la nube están configuradas en esta máquina.                                                                                                                                                           |
| `guide`           | La referencia de ops y las guías de diseño, generadas a partir de las mismas definiciones que el ejecutor usa para validar, así que no pueden apartarse de lo que `apply` acepta. `--json` la devuelve con el esquema de cada op. |
| `install-cli`     | Pone `genoffice` en el `PATH`.                                                                                                                                                                                                    |
| `skill`           | Lista los agentes de programación encontrados en esta máquina e instala o actualiza en ellos la skill de GenOffice.                                                                                                               |
| `mcp`             | Publica todos los comandos como una herramienta del Model Context Protocol. Vea **Conectar un agente de programación**.                                                                                                           |

## Edición: documentos, hojas, presentaciones

`genoffice docs`, `genoffice sheet` y `genoffice slides` leen y escriben por la misma vía de escritura que usa la aplicación, y comparten un mismo vocabulario: una **op** es una edición suelta, y una **spec** es una lista de ops que se aplican en orden.

```sh
genoffice sheet read report.xlsx
genoffice slides apply deck.pptx --ops ops.json --dry-run
genoffice docs apply brief.docx --ops ops.json
```

`--dry-run` informa de lo que haría un lote y no escribe nada, que es la forma barata de revisar una spec antes de aplicarla. El panel de IA de la aplicación funciona exactamente con estas ops, así que todo lo que puede pedirle, usted puede programarlo.

## Model Context Protocol

`genoffice mcp` publica todos los comandos como herramientas de MCP, y `genoffice mcp install <agent|all>` la registra en la configuración propia de un agente de programación. Vea **Conectar un agente de programación** para ese lado.

## Lo que el comando no hace

Lee y escribe el archivo. No es la aplicación: no hay ventana, y el diálogo de actualización de la aplicación no se aplica. Todo lo que necesite la ventana —el panel de IA, la revisión de calidad de una diapositiva renderizada— tiene que esperar a que abra el archivo. `genoffice render` le da los píxeles sin ella, y `genoffice slides` audita por su cuenta la disposición de una presentación.
