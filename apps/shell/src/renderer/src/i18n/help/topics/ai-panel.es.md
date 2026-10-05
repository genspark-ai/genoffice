# El panel de IA

Todos los editores pueden invocar el panel de IA: seleccione algo, dé una instrucción y vea el resultado en streaming.

## Abrirlo y usarlo

![El panel de IA en Docs](img/ai-panel.png)

- Entradas: el **botón de IA** en la cinta de opciones de cada editor, **Pedir a la IA** en los menús contextuales, o Preguntar a la IA en la barra de marcado.
- Describa la tarea en lenguaje natural (reescribir esto / convertir esta columna en porcentajes / rehacer el diseño de esta página…) y pulse Intro.
- Las respuestas se renderizan **en streaming**; cuando la IA necesita herramientas (leer el documento, editarlo, ejecutar un script), las ejecuta y sigue hasta terminar.
- **Detener**: interrumpa la vuelta en curso en cualquier momento.

## Qué sabe hacer

- **Docs**: reescribir/ampliar/traducir/resumir, insertar tablas e imágenes, ajustar el formato; cada vuelta guarda primero una instantánea.
- **Sheets**: fórmulas, relleno de datos, transformaciones masivas, formato.
- **Slides**: generación de una presentación completa, ajuste del diseño, reescritura de los textos.
- **PDF**: preguntas y resúmenes sobre el texto o las páginas seleccionadas.
- **Markdown / HTML**: reescribir, ampliar, traducir.

## Retroceso y seguridad

- El panel de Docs conserva una **lista de versiones**: una instantánea por vuelta, puede volver a cualquiera de ellas y ese propio retroceso se puede deshacer con Ctrl+Z. Las instantáneas sobreviven a reabrir el documento.
- Las ediciones de la IA pasan por la misma canalización de edición que las manuales (se pueden deshacer y requieren guardar): nada omite su confirmación de guardado.

## Privacidad

- Las instrucciones y el contenido pertinente del documento van al **servicio de modelo que haya configurado** (Genspark alojado o un endpoint personalizado, capítulo siguiente); sin configuración, no se envía nada.
- Los archivos locales no se suben a ningún otro sitio; las claves BYOK viven solo en las cabeceras de la petición: nunca en el disco ni en los registros.
