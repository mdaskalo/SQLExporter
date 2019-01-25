# SQLExcelExporter
A simple java application to import SQL Data into Excel Spreadsheet


Usage :

```
java -jar SQLExcelExporter.jar -config <file>
```


A JSON Configuration file can be used with the following content -
Options:
* Option large=true|false - use streaming for large files
* File extensions can be xls or xlsx
* Placeholder ##Date## in filename will be replaced with date in reverse format yyyyMMdd

```json
{
  "datasource": {
    "className": "oracle.jdbc.driver.OracleDriver",
    "jdbcUrl": "jdbc:oracle:thin:@//localhost:1521/orcl",
    "username": "user",
    "password": "password"
  },
  "excelFile": [
    {
      "id": "1",
      "large": true,
      "worksheet": [
        {
          "id": "1",
          "sqlQuery": "Select * from Person",
          "workSheetName": "Person"
        },
        {
          "id": "2",
          "sqlQuery": "Select * from Stock",
          "workSheetName": "Stock"
        }
      ],
      "fileName": "C:/PathToFile/##Date## Filename123.xlsx",
      "preparationProcedureStatement": "{ call DataPreparationProcedure }"
    }
  ]
}

```

The tool is single threaded and multiple worksheets with different dataset can be created using the tool.
