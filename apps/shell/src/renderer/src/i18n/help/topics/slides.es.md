# Slides: presentaciones

Slides es el editor similar a PowerPoint: lee y escribe .pptx reales.

## La interfaz

- **Cinta de opciones**: la pestaña Inicio reúne insertar y dar formato (cuadro de texto/forma/imagen/tabla/gráfico), fuente y párrafo, alinear y organizar (orden de superposición, alinear, distribuir).
- **Carril de miniaturas** (izquierda): hacer clic para cambiar, arrastrar para reordenar, menú contextual para nuevo/duplicar/eliminar.
- **Lienzo**: edición WYSIWYG; arrastre, tiradores de tamaño, guías.
- **Notas**: notas del orador por diapositiva, se conservan en los flujos de exportación.

## Pestañas de la cinta de opciones

La barra de pestañas (en macOS empieza por Inicio; en Windows se añade una pestaña Archivo):

- **Inicio**: insertar y dar formato: cuadro de texto, formas, imágenes, tablas, gráficos; fuente y párrafo; alinear y organizar (superposición/alinear/distribuir); diseño.
- **Insertar**: cuadro de texto, tabla (filas y columnas a elegir), imágenes, número de página, un botón de salto (púlselo durante la presentación para ir a una diapositiva) y más.
- **Dibujar**: el **lápiz** (dibujar a mano en la diapositiva, guardado como tinta en la página) y el **marcador de resaltado** (trazos translúcidos y más gruesos), con grosor de trazo; pulsar otra vez la herramienta la cancela.
- **Diseño**: temas, esquema de color y fondo; patrones y diseños.
- **Transiciones**: elija una transición para la diapositiva actual (surte efecto en el modo Presentador de PowerPoint), con aplicación a todas; «Ninguna» la quita.

  ![La pestaña Transiciones](img/slides-transitions.png)

- **Animaciones**: efectos de entrada y énfasis para la forma seleccionada, **Trayectorias de la animación** (moverse a lo largo de una ruta); **Vista previa** reproduce las animaciones de la diapositiva en el lienzo; «Ninguna» las quita.

  ![La pestaña Animaciones](img/slides-animations.png)

  Pruébelo: seleccione el cuadro de texto del título ▸ pestaña Animaciones ▸ elija un efecto de entrada ▸ **Vista previa** lo reproduce en el lienzo.

- **Presentación con diapositivas**: presentar desde el principio o desde la diapositiva actual, más los ajustes de presentación.
- **Revisar**: **nuevo comentario** en la diapositiva actual (se escribe en el pptx y es visible en PowerPoint).
- **Vista**: **Normal** (miniaturas y lienzo), **Vista Esquema** (recorrer y saltar por el texto), **Clasificador de diapositivas** (vista general en cuadrícula, doble clic para editar), **Vista de lectura** (pantalla completa, diapositiva a diapositiva; Esc para salir).

## Menús contextuales

El menú depende de dónde haga clic derecho:

- **Sobre el lienzo vacío**: diapositiva nueva (diseño a elegir), cortar/copiar/pegar la diapositiva (también **Pegar como imagen** y **Pegar conservando el formato de origen**), duplicar/quitar/ocultar la diapositiva, **agregar sección** (antes), renombrar la sección/mover arriba/mover abajo/quitar/contraer todo/expandir todo, **formato del fondo** / **cambiar la imagen del fondo**, restablecer el diseño de la diapositiva.
- **Sobre un elemento seleccionado**: editar texto, hipervínculo, tamaño y posición, **alinear** (izquierda/centrar horizontal/derecha/arriba/centrar vertical/abajo) y distribuir horizontal y verticalmente, traer al frente / enviar al fondo, agrupar/desagrupar/reagrupar, voltear horizontal/verticalmente, recortar imagen / reemplazar imagen / guardar como imagen, cambiar la forma, **establecer como forma predeterminada**, editar puntos; los elementos de texto reciben además negrita/cursiva/subrayado/viñetas/numeración.
- **Sobre una tabla seleccionada**: insertar filas/columnas (arriba/abajo/izquierda/derecha), quitar filas/columnas, combinar / combinar a la derecha / combinar hacia abajo, dividir celdas, **sombreado de celda**, anclar el contenido de la celda (arriba/centro/abajo).

## Contenido

- Cuadros de texto, formas (relleno/contorno/sombra), imágenes (recortar/reemplazar), tablas, gráficos (columnas/barras/líneas/circulos, …, con datos editables).
- Patrones y diseños: unifican fuentes, marcadores de posición y fondos; las diapositivas nuevas heredan el diseño elegido.
- Los colores y las fuentes del tema siguen al tema.

## Generación con IA

- La tarjeta AI Slides de la pantalla de inicio: dé un tema o un esquema y la IA construye la presentación; si falla la generación en la nube, recurre a la generación local.
- Sigua ajustando después con el panel de IA (la IA reorganiza el formato mediante un entorno de pruebas de scripts controlado: el mismo mecanismo que en [Sheets](help://sheets)).

## Presentar y exportar

- **Exportar como PDF**: rasterizado diapositiva a diapositiva en una ventana oculta, con progreso y vigilancia de tiempo de espera para presentaciones grandes.
- Exportación de imágenes: un PNG por diapositiva.
- Presentación a pantalla completa donde la versión lo permita.

## Guardar

Ida y vuelta completa en .pptx: las páginas sin modificar quedan idénticas byte a byte; las notas, los patrones y las anotaciones se conservan.
