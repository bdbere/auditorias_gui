# Analizador de Auditorías v2.0

## Descripción
El Analizador de Auditorías v2.0 es una aplicación de escritorio desarrollada en Python que automatiza el análisis, procesamiento y visualización de resultados de encuestas de satisfacción o auditorías de servicios institucionales. A través de una interfaz gráfica intuitiva, la herramienta permite cargar datos brutos, calcular promedios de respuestas basados en una escala Likert, exportar reportes detallados en Excel y generar gráficos estadísticos para la toma de decisiones.

## Características Principales
* **Carga de Datos:** Compatible con archivos `.xlsx` y `.csv`. Detecta y asigna automáticamente la columna "Carrera" (ubicada en la columna G / índice 6).
* **Análisis Cuantitativo:** Mapea respuestas cualitativas ("Excelente", "Bueno", "Regular", "Malo", "Muy Malo") a una escala numérica (5 a 1) para calcular promedios exactos por pregunta y por categoría de servicio.
* **Reportes en Excel Automatizados:** Genera un archivo Excel de salida con las siguientes características:
  * Pestañas (hojas) independientes por cada "Carrera".
  * Tablas de frecuencias que muestran el recuento de respuestas y totales.
  * Alertas visuales con formato condicional (resalta en rojo promedios críticos por debajo de 3.0).
  * Estilos de celda predefinidos (colores, bordes y alineaciones) para un aspecto profesional listo para entregar.
* **Visualización Dinámica:** Interfaz integrada que permite generar gráficos de barras personalizados, filtrando la información por Carrera y por el tipo de servicio evaluado (ej. Centro de información, Titulación, Cafetería, etc.).

## Tecnologías Utilizadas
* **Python 3.x**
* **Pandas:** Manipulación, limpieza y estructuración de los datos.
* **Tkinter:** Construcción de la interfaz gráfica de usuario (GUI).
* **Openpyxl:** Creación y formateo avanzado de los reportes en Excel.
* **Matplotlib:** Renderizado de gráficos de barras para el análisis visual.

## Instalación y Requisitos

1. Clona este repositorio en tu entorno local:
   ```bash
   git clone [https://github.com/tu-usuario/analizador-auditorias.git](https://github.com/tu-usuario/analizador-auditorias.git)

2. Asegúrate de tener Python instalado en tu sistema.

3. Instala las dependencias requeridas ejecutando:
   ```bash
   pip install pandas openpyxl matplotlib
## Guía de uso

1. Ejecuta el script principal desde tu terminal o entorno de desarrollo (como PyCharm):
   ```bash
   python main.py
2. En la ventana principal de la aplicación, haz clic en 1. Cargar Archivo y selecciona tu base de datos con los resultados de las auditorías.
3. Tras una carga exitosa, se habilitarán las siguientes opciones:
  * Generar Reporte Excel: Te pedirá elegir la ruta para guardar el documento procesado con todas las hojas, cálculos y estilos aplicados.
  * Generar Gráficos: Abrirá una ventana secundaria donde podrás seleccionar una Carrera y un Servicio específico en los menús desplegables para visualizar su respectivo gráfico de rendimiento.

## Estructura de Datos Esperada
Para el correcto funcionamiento del análisis, el archivo de origen debe cumplir con:

* Columna G: Estar reservada para la información de la Carrera del usuario encuestado. El script la identificará y renombrará si es necesario.

* Escala de Respuestas: El texto de las celdas a evaluar debe coincidir con la escala definida (Excelente, Bueno, Regular, Malo, Muy Malo) sin importar el uso de mayúsculas.

* Mapeo de Columnas: Las preguntas deben estar distribuidas en los rangos de columnas establecidos en el diccionario servicios_columnas dentro del código (por ejemplo, columnas de la H a la O para "Centro de información").

## Autora
* Berenice Hernández Juárez - Ingeniería en Sistemas Computacionales - Tecnm Querétaro
