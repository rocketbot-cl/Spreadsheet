# GSpreadsheet
  
Module to manage Google Spreadsheet from Rocketbot.  

*Read this in other languages: [English](Manual_Spreadsheets.md), [Español](Manual_Spreadsheets.es.md).*
  
![banner](imgs/Banner_spreadsheet.png)
## How to install this module
  
__Download__ and __install__ the content in 'modules' folder in Rocketbot path  



## Description of the commands

### Setup credentials
  
Get permissions to handle Google Spreadsheet with Rocketbot
|Parameters|Description|example|
| --- | --- | --- |
|Credentials file path|Path of the Google credentials file|Path|

### Create Sheet
  
Create a new sheet in Spreadsheet selected
|Parameters|Description|example|
| --- | --- | --- |
|Spreadsheet|Enter the ID of the spreadsheet|Spreadsheet|
|Search by|Select the field to search by|Name|
|Sheet name|Name that the spreadsheet will have|Name: |
|Variable where the ID will be saved|Variable where the ID of the created spreadsheet will be saved|{Result}|

### Delete Sheet
  
Delete a sheet from Spreadsheet selected
|Parameters|Description|example|
| --- | --- | --- |
|Spreadsheet|Enter the name of the spreadsheet to delete|Spreadsheet|
|Search by|Select the field to search by|Name|
|Sheet Id for delete|Enter the id of the sheet to delete|624974831 |

### Set cells
  
Set a cell or range of cells from Spreadsheet selected
|Parameters|Description|example|
| --- | --- | --- |
|Spreadsheet|Enter the spreadsheet name|Spreadsheet|
|Search by|Select the field to search by|Name|
|Sheet|Enter the sheet name|Sheet1|
|Enter cell |Cell where the text will be written|A1|
|Text |Text to write in the cell|[["data","data"],["data","data"]]|

### Read cells
  
Read a cell or range of cells from Spreadsheet selected, eg. A1 or A1:B5
|Parameters|Description|example|
| --- | --- | --- |
|Spreadsheet|Enter the name of the spreadsheet|Spreadsheet|
|Search by|Select the field to search by|Name|
|Sheet|Enter the name of the sheet|Sheet1|
|Enter cell or range of cells |Enter the cell or range of cells to read|A1|
|Result |Variable where the result will be stored|result|

### Get sheets
  
Get list of sheets with your id from a Spreadsheet selected
|Parameters|Description|example|
| --- | --- | --- |
|Spreadsheet|Enter the id of the spreadsheet|Spreadsheet|
|Search by|Select the field to search by|Name|
|Result |Variable where the result will be stored|result|

### Count Cells
  
Count rows and columns on a sheet
|Parameters|Description|example|
| --- | --- | --- |
|Spreadsheet|Enter the name of the spreadsheet|Spreadsheet|
|Search by|Select the field to search by|Name|
|Sheet|Enter the name of the sheet|Sheet1|
|Result |Variable where the result will be stored|result|
