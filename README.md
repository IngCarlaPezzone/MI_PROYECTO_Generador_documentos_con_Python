# Generador de documentos con Python

Con este script se puede generar una serie de documentos Word a partir de información tabulada en un Excel, donde contiene todo el seguimiento de los documentos que se van generando para formar parte del Sistema de Gestión Documentos de la organización.

Esta pensado para el siguiente uso:
- Tiene una plantilla en formato `.docx` donde se determina el formato del documento: encabezado, logo, títulos, pie de página, textos que se repitan en todos los documentos, etc. Es importante que se decidan estos elementos desde el primer momento ya que este script no agrega modificaciones de formato.
- Tiene una planilla de cálculos `.xlsx` donde se colocan los títulos en columnas y en las filas la información de cada documento. La primer columna corresponda al código del documento con el que se realizarán los distintos procesos.
- Se crea una carpeta donde se guardan los documentos generados.
- El script `utils.py` contiene todas las funciones de utilidad para la ejecución del `generador.py`.
- El `generador.py` lee la información del excel y verifica si todos los códigos se encuentra en la carpeta. Pueden ocurrir dos cosas:
  - Todos los códigos del Excel ya tienen generado su correspondiente documento, entonces lo que hace el script es actualizar solo algunos campos que pudieron haberse modificado o no en el Excel. Por ejemplo: para este caso, el apartado "Descripción" va personalizado uno por uno, por lo que ese texto el script no lo modifica nunca. En cambio, otros apartados como "Alcance", "Responsables" y "Registros" pueden ir teniendo cambios a medida que se construye el sistema de documentos. Estos últimos sí se actualizan a la última versión del Excel.
  - En el Excel hay mas códigos que no tienen su correspondiente documento en la carpeta. En este caso el script genera el documento según la plantilla y la información del Excel.

## 20220830 Actualización

Agregué funcionalidad. Agrega al final del documento en el apartado de Documentos Asociados, todos los documentos que tenga asociados con sus correspondientes link hacie ese documento.
Cambié las palabras Alcance por Objetivos y Responsable por Responsabilidades.
