# Configuración, idioma, tema e integraciones MCP

## Abrir la configuración

La fila de cuenta de la parte inferior izquierda de Inicio abre el panel de configuración (muestra «Iniciar sesión» cuando está desconectado). Tiene seis secciones: Cuenta, Modelo de IA, Medios de IA y búsqueda, General, Integraciones y Acerca de.

![Configuración ▸ General, donde están el idioma, el tema, el guardado automático y el interruptor de estadísticas de uso](img/settings-general.png)

La configuración de modelos tiene su propio artículo; en **Medios de IA y búsqueda** se activan, por proveedor, la generación de imágenes, el análisis de imágenes, el análisis de vídeo, la búsqueda web y la búsqueda de archivos locales.

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

**Integraciones** es el panel que conecta GenOffice con un agente de programación, y tiene su propio artículo: Conectar un agente de programación. La versión corta: elija una vía (la línea de comandos o MCP), siga esa sección y después abra una conversación nueva y pregunte.

![Configuración ▸ Integraciones: los tres pasos, luego las filas de habilidades y las opciones de MCP](img/settings-integrations.png)

En **Servidor HTTP local**, la aplicación también puede ejecutar el servidor por su cuenta —un interruptor de activación y un puerto— y **Avanzado** añade la URL de comprobación de estado y el archivo de registro, en lugar de dejarlo en manos del asistente. Solo escucha en localhost.

## Chuleta de la línea de comandos

| Comando            | Qué hace                      |
| ------------------ | ----------------------------- |
| `genoffice <file>` | abrir un archivo              |
| `genoffice mcp`    | iniciar el servidor MCP local |
| `genoffice --help` | todos los comandos y opciones |
