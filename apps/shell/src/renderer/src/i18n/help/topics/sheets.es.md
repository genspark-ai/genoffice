# Sheets: hojas de cálculo

Sheets es el editor de hojas de cálculo similar a Excel; el cálculo se ejecuta en un proceso separado, un motor Rust (un fallo ahí nunca tumba la aplicación). Abre y guarda .xlsx reales; .csv y .tsv se abren como tablas.

## La interfaz

- **Cinta de opciones**: ocho pestañas, que se detallan una a una más abajo.
- **Barra de fórmulas**: muestra y edita la fórmula de la celda activa; se admiten las funciones habituales.
- **Pestañas de hoja** (abajo): agregar / cambiar el nombre / eliminar / mover hojas.
- **Edición de celdas**: doble clic o simplemente escribir; Intro confirma y baja, Tab pasa a la derecha, Esc cancela (costumbres de Excel).
- **Atajos**: alineados con la familia de Excel (ctrl+C/V/X, ctrl+Z/Y, ctrl+F, …).

## Pestañas de la cinta de opciones

- **Inicio**: fuente, relleno, bordes, formatos de número (moneda/porcentual/miles, aumentar o disminuir decimales), alineación, combinar, insertar filas/columnas y su tamaño, formato condicional, dar formato como tabla, estilos de celda, portapapeles y copiar formato, ordenar y filtrar.
- **Insertar**: formas, iconos, símbolos, ecuación, captura de pantalla y más.
- **Disposición de página**: colores y fuentes del tema, interruptores para imprimir líneas de la cuadrícula y encabezados, vista previa de salto de página.
- **Fórmulas**: Autosuma e inserción de funciones, definir nombres (también desde la selección), rastrear precedentes y dependientes, ventana Inspección, recalcular hoja o libro.
- **Datos**: ordenar y filtrar (incluido el filtro avanzado y borrar el filtro), texto en columnas, combinar libros, actualizar todo.
- **Revisar**: explorar comentarios (mostrar, anterior/siguiente), traducir.
- **Vista**: interruptores de cuadrícula y encabezados, zoom, Normal / vista previa de salto de página.
- **Diseño de gráfico**: aparece cuando hay un gráfico seleccionado: tipo de gráfico, estilos y colores, edición del rango de datos.

La pestaña Datos, botón por botón (de izquierda a derecha en la imagen):

![La pestaña Datos](img/sheets-data.png)

- **Tabla dinámica**: construye una tabla dinámica a partir del rango actual; arrastre los campos para agregar.
- **Actualizar**: recalcula los datos de la tabla dinámica actual.
- **Desde texto/CSV**: importa un .csv/.txt como hoja nueva, separándolo por un delimitador.
- **Combinar libros**: trae las hojas de otros archivos .xlsx a este libro.
- **Actualizar todo**: recalcula todas las tablas dinámicas y los conjuntos de datos externos.
- **Ordenar** (desplegable): ascendente / descendente / orden personalizado (reglas de varias columnas).
- **Filtro**: añade menús desplegables ▼ a la fila de encabezado; marque los valores que quiere conservar.
- Los pequeños botones apilados junto a él: **Borrar** (recuperar todas las filas), **Volver a aplicar** (ejecutar otra vez el filtro actual), **Avanzadas** (filtrar con un rango de criterios).
- **Texto en columnas** (desplegable): divide una columna en varias por delimitador o por ancho fijo.
- **Relleno rápido**: dé un ejemplo y el resto de la columna se rellena siguiendo el patrón (ctrl+E).
- **Quitar duplicados**: quita las filas duplicadas según las columnas seleccionadas.
- **Validación de datos** (desplegable): reglas de entrada para la selección (listas desplegables, rangos numéricos, …).
- **Consolidar**: agrega varios rangos en un mismo lugar por categoría.
- **Análisis de hipótesis** (desplegable): buscar objetivo / tablas de datos.
- **Agrupar / Desagrupar** (desplegable): grupos de filas o columnas con contraer y expandir.
- **Subtotal**: inserta filas de subtotal por categoría.

La pestaña Fórmulas, botón por botón:

![La pestaña Fórmulas](img/sheets-formulas.png)

- **Insertar función** (fx): busca funciones con un asistente de argumentos.
- **Autosuma** (desplegable): SUMA con un clic, además de promedio/recuento/máximo/mínimo.
- **Recientes / Financieras / Lógicas / Texto / Fecha y hora / Búsqueda y referencia / Matemáticas y trigonometría / Más**: explore e inserte funciones por categoría.
- **Administrador de nombres**: ver, crear y quitar rangos con nombre.
- **Definir nombre** (desplegable): da nombre a la selección; **Utilizar en la fórmula** inserta un nombre existente; **Crear desde la selección** nombra rangos a partir de su fila o columna de encabezado.
- **Rastrear precedentes / Rastrear dependientes**: las flechas azules muestran de dónde vienen los datos de una fórmula y a qué alimentan; **Quitar flechas** las borra.
- **Mostrar fórmulas**: las celdas muestran la propia fórmula en lugar del resultado.
- **Comprobación de errores**: localiza y explica los errores de fórmula.
- **Ventana Inspección**: fije las celdas que le interesan y siga sus valores en vivo.
- **Opciones para el cálculo** (desplegable): recálculo automático o manual; en el modo manual, **Calcular ahora** y **Calcular hoja** lo ejecutan a mano.

## Números y formato

- Formatos de número: Normal, Número, Moneda, Porcentual, Fecha/hora, Fracción, Científico y más.
- Alineación, ajuste de línea, celdas combinadas, bordes y rellenos.
- Alturas de fila y anchos de columna arrastrando; doble clic en un borde ajusta automáticamente.

## Datos

**Ordenar y filtrar** (por ejemplo, descendente por una columna):

1. Haga clic en **cualquier celda de esa columna** (no hace falta seleccionar la columna entera).
2. Pestaña Inicio ▸ **Ordenar y filtrar** ▸ **Descendente**; las filas enteras se reordenan juntas (el área se ordena como un todo).
3. Para reglas propias (varias columnas, por color): la misma ruta, **Orden personalizado**.
4. Filtrado: seleccione la fila de encabezado y haga clic en **Ordenar y filtrar ▸ Filtro**; cada encabezado recibe un menú ▼ donde marca los valores que quiere conservar; borrar el filtro lo devuelve todo.

- Ordenar y filtrar.
- Inmovilizar paneles.
- .csv / .tsv: se abren directamente como tabla (un tsv separado por tabulaciones se analiza como una sola tabla); al guardar se vuelve a escribir el formato original.

## Menús contextuales

- **En la cuadrícula**: el menú propio del editor (Univer): cortar/copiar/pegar, insertar y quitar filas/columnas, ocultar, combinar celdas, inmovilizar paneles y otros elementos habituales.
- **En la barra de estado inferior**: elija qué estadísticas muestra la barra de estado (promedio / recuento / suma, …); la elección se conserva.
- **En una pestaña de hoja abajo**: agregar / cambiar el nombre / quitar / colorear / ocultar hojas (menú de pestañas de Univer).
- El menú contextual de la barra de pestañas superior se trata en [Pestañas y gestión de ventanas](help://tabs-and-windows).

## IA

- El panel de IA lateral: seleccione un rango e indíquelo en lenguaje natural (reformatear, generar datos, escribir fórmulas).
- Las ediciones de la IA se pueden revertir desde el panel.

## Guardar y exportar

- Guarda .xlsx (fórmulas y formatos se conservan); Guardar como; la exportación a PDF sigue la paginación de impresión.
- El autoguardado sigue la regla global (se activa tras el primer guardado manual).

## Estabilidad

- El motor de cálculo está aislado por proceso de la interfaz: si unos datos extremos lo matan, recibe un mensaje y un intento de recuperación de la sesión, no un fallo de la aplicación.
