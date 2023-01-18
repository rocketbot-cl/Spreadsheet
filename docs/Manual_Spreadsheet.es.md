# GSpreadsheet
  
Módulo para manejar Google Spreadsheet desde Rocketbot.  

*Read this in other languages: [English](Manual_Spreadsheets.md), [Español](Manual_Spreadsheets.es.md).*
  
![banner](imgs/Banner_spreadsheet.png)
## Como instalar este módulo
  
__Descarga__ e __instala__ el contenido en la carpeta 'modules' en la ruta de Rocketbot.  



## Descripción de los comandos

### Configurar credenciales
  
Obtiene los permisos para manejar Google Spreadsheet con Rocketbot
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta archivo de credenciales|Ruta del archivo de credenciales de Google|Ruta|

### Crear Hoja
  
Crear una nueva hoja en la Hoja de cálculo seleccionada
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ingrese Spreadsheet|Ingrese el ID de la hoja de calculo|Hoja de calculo|
|Buscar por|Seleccione el campo por el cual buscar|Name|
|Nombre de la hoja|Nombre que tendrá la hoja de calculo|Nombre: |
|Variable donde se guardará el ID|Variable donde se guardará el ID de la hoja de calculo creada|{Resultado}|

### Borrar Hoja
  
Elimina una hoja de la Hoja de cálculo seleccionada
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ingrese Spreadsheet|Ingrese el nombre de la hoja de calculo a borrar|Hoja de calculo|
|Buscar por|Seleccione el campo por el cual buscar|Name|
|Id de la hoja a borrar|Ingrese el id de la hoja a borrar|624974891 |

### Escribir en celdas
  
Escribe en una celda o rango de celdas desde una hoja de cálculo seleccionada
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ingrese Spreadsheet|Ingrese el nombre de la hoja de calculo|Hoja de calculo|
|Buscar por|Seleccione el campo por el cual buscar|Name|
|Ingrese Hoja|Ingrese el nombre de la hoja|Hoja 1|
|Ingrese la celda a escribir |Celda donde se escribirá el texto|A1|
|Ingrese texto |Texto a escribir en la celda|[["data","data"],["data","data"]]|

### Leer celdas
  
Lee celdas o rango de celdas desde una hoja de cálculo seleccionada, ejemplo A1 o A1:B5
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ingrese Spreadsheet|Ingrese el nombre de la hoja de calculo|Hoja de calculo|
|Buscar por|Seleccione el campo por el cual buscar|Name|
|Ingrese Hoja|Ingrese el nombre de la hoja|Hoja 1|
|Ingrese celda o rango de celdas |Ingrese la celda o rango de celdas a leer|A1|
|Resultado |Variable donde se almacenará el resultado|resultado|

### Obtener hojas
  
Obtener lista de hojas con su id desde una hoja de cálculo seleccionada
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ingrese Spreadsheet|Ingrese el id de la hoja de calculo|Hoja de calculo|
|Buscar por|Seleccione el campo por el cual buscar|Name|
|Resultado |Variable donde se almacenará el resultado|resultado|

### Contar Celdas
  
Cuenta filas y columnas de una hoja seleccionada
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ingrese Spreadsheet|Ingrese el nombre de la hoja de calculo|Hoja de calculo|
|Buscar por|Seleccione el campo por el cual buscar|Name|
|Ingrese Hoja|Ingrese el nombre de la hoja|Hoja 1|
|Resultado |Variable donde se almacenará el resultado|resultado|
