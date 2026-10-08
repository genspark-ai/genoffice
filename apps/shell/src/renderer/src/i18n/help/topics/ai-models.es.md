# Modelos de IA y ajustes

## Proveedores y modelos

Los modelos y las claves se configuran en Configuración (la fila de cuenta de la parte inferior izquierda de Inicio):

![La ventana de configuración](img/settings-general.png)

- **Genspark alojado**: inicie sesión (flujo por código de dispositivo) y úselo, sin configuración alguna.
- **Endpoints personalizados (BYOK)**: Configuración ▸ IA toma una URL base y una clave API por protocolo: compatible con OpenAI, Anthropic, Gemini, DeepSeek, DashScope (qwen) y más. Las claves se guardan en el archivo de ajustes de la aplicación en este equipo y se envían solo en las cabeceras de la petición.
- Se puede elegir un modelo distinto por capacidad: chat/generación, generación de imágenes, análisis de imágenes.
- **Probar conexión**: comprueba que el endpoint sea accesible y que el modelo esté visible antes de guardar.
- Las URL base pueden llevar una ruta y una cadena de consulta (estilo pasarela); las rutas del endpoint se concatenan correctamente.

## Integración con la CLI (tipo Codex)

- Los ajustes aceptan la ruta de un programa CLI local (los directorios personales no ASCII y un prefijo ~ funcionan; ~ se expande automáticamente); Detectar modelos sondea los modelos disponibles de esa CLI.
- La validación solo comprueba la existencia: no impone restricciones de juego de caracteres.

## Cuándo se aplican los cambios

- Los cambios de modelo y de endpoint se aplican de inmediato; una conversación en curso mantiene la configuración anterior hasta su siguiente vuelta.
