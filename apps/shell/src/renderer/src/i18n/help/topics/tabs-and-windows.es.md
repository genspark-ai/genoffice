# Pestañas y gestión de ventanas

Todos los archivos abiertos comparten una sola ventana; la barra de pestañas de arriba permite pasar de uno a otro al estilo de un navegador.

![La barra de pestañas: Inicio más dos documentos](img/tabs.png)

## Fundamentos

- **Cambiar**: haga clic en una pestaña o gire la rueda sobre la barra para recorrerlas.
- **Cerrar**: la × de la pestaña, o ⌘W/ctrl+W. Si hay cambios sin guardar, pregunta primero; el menú de todas las pestañas permite cerrar de una vez las demás / las de la derecha.
- **Nuevo**: el + al extremo derecho de la barra.
- **La activación es inmediata**: al pulsar una pestaña ya se cambia, sin esperar a que termine el clic.

## Cambiar nombre: haga doble clic en una pestaña

**Haga doble clic en cualquier pestaña de archivo** y el título se convierte en un campo de texto en el sitio: Intro confirma, Esc cancela, la pérdida de foco confirma; la Intro que confirma un candidato del IME no se confunde con una confirmación. El cambio de nombre pasa por las mismas comprobaciones de seguridad que el de una fila de la pantalla de inicio (caracteres no permitidos, conflictos de nombre), renombra el archivo en el disco y actualiza la lista de recientes.

## Arrastrar para reordenar

Mantenga pulsada una pestaña y arrástrela de lado a lado para reordenarla; las vecinas se apartan en vivo y el orden queda fijo al soltar. Una zona muerta de 4 píxeles evita que un clic normal desplace las pestañas.

## Menú contextual

Con el botón derecho en una pestaña: nueva pestaña, cerrar, cerrar las demás, cerrar las de la derecha, duplicar el archivo y más (traducido al idioma de la aplicación).

## Separar en una ventana

**Arrastre una pestaña fuera de la barra** (o use Abrir en una ventana nueva) y se convierte en su propia ventana con el documento vivo; arrástrela de vuelta para acoplarla de nuevo a la barra. Una ventana separada es una ventana equivalente: puede seguir editando y guardando allí.

## La lista de todas las pestañas

Cuando las pestañas desbordan, el ▾ del extremo derecho de la barra abre la lista completa (un menú nativo, nunca tapado por el área de contenido); elija con las flechas.

## La barra de herramientas de cada editor

Cada pestaña de editor tiene una barra de herramientas arriba (la disposición exacta varía algo según el editor):

- **Guardar** (⌘S/ctrl+S) y **Guardar como**.
- **Deshacer / Rehacer**: en varios niveles; cada editor mantiene su propio historial.
- **Buscar** (ctrl+F): abre el panel de búsqueda de ese editor.
- Interruptor de **Autoguardado**: activado, los cambios se escriben en el disco a intervalos; desactivado, solo se escribe al guardar a mano (el distintivo de cambios sin guardar en la pestaña o el título se lo recuerda).

## La pestaña Inicio

La pestaña Inicio de la izquierda no se puede cerrar; para volver desde cualquier editor, haga clic en ella o use Archivo ▸ Inicio.
