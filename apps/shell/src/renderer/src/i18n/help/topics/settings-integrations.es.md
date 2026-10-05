# Configuración, idioma, tema e integraciones MCP

## Abrir la configuración

La fila de cuenta de la parte inferior izquierda de Inicio abre el panel de configuración (muestra «Iniciar sesión» cuando está desconectado); las opciones relacionadas con la IA están en su sección Modelo de IA

![La ventana de configuración](img/settings-integrations.png): la configuración de modelos se trata en Modelos de IA y ajustes.

## Idioma

- La configuración ofrece **21 idiomas de interfaz**: inglés, chino simplificado, japonés, coreano, francés, alemán, español, tailandés, indonesio, ruso, árabe, portugués, italiano, polaco, checo, neerlandés, malayo, hebreo, hindi, chino tradicional, vietnamita.
- El cambio se aplica de inmediato y se conserva; la barra de menús nativa se reconstruye con el idioma.

## Tema

Claro / Oscuro / Seguir el sistema. «Seguir el sistema» sigue la apariencia del sistema operativo, y los editores cambian de aspecto en sincronía sin parpadeos.

## Aplicaciones predeterminadas

La configuración puede registrar GenOffice como programa predeterminado para .docx / .xlsx / .pptx / .pdf y formatos similares (registro de aplicación predeterminada a nivel de plataforma; confirme cuando se le pida).

## Avisos de software de terceros y actualizaciones

- Ayuda ▸ Avisos de software de terceros: el inventario completo de licencias de código abierto que se distribuye con la aplicación.
- Ayuda ▸ Buscar actualizaciones: lanza una comprobación manual; si hay una versión más reciente, ofrece instalarla.

## Iniciar sesión en Genspark

- El acceso a la sesión (en la configuración o en la lista de proyectos de la nube) usa un flujo de **código de dispositivo**: GenOffice muestra un código y abre el inicio de sesión en el navegador; continúa solo cuando ha terminado.
- El inicio de sesión solo se usa para: la lista de proyectos de la nube y los modelos alojados de Genspark. Sin él, todas las funciones locales y los modelos personalizados siguen funcionando.
- Cerrar la sesión es un clic en la configuración.

## Integración MCP (para usuarios avanzados / clientes de IA)

GenOffice incorpora un **servidor MCP** local para que los clientes de IA externos (Claude Desktop, Cursor, …) puedan leer y escribir directamente en sus documentos:

- Inicio: `genoffice mcp` en la línea de comandos (puerto y token de autenticación configurables; solo loopback de forma predeterminada).
- Capacidades: crear/abrir/editar docx, xlsx y pptx, leer contenidos, convertir formatos, exportar a PDF y más: el mismo conjunto de herramientas que usan las aplicaciones de escritorio.
- Seguridad: la autenticación por token es opcional pero recomendable; el escucha se queda en la máquina local de forma predeterminada; véase `genoffice mcp --help`.

## Chuleta de la línea de comandos

| Comando            | Qué hace                      |
| ------------------ | ----------------------------- |
| `genoffice <file>` | abrir un archivo              |
| `genoffice mcp`    | iniciar el servidor MCP local |
| `genoffice --help` | todos los comandos y opciones |
