package com.sachinhandiekar.sqltools.excel.model;

import com.google.gson.annotations.SerializedName;

import java.util.List;

/**
 * A class to denote the SQLExcelExporterConfig JSON data model.
 * <p>
 * E.g.
 * <p>
 * {
 * "datasource": {
 * "className": "oracle.jdbc.driver.OracleDriver",
 * "jdbcUrl": "jdbc:oracle:thin:@//localhost:1521/orcl",
 * "username": "user1",
 * "password": "pass1"
 * },
 * "excelFile": [
 * {
 * "id": "1",
 * "large": true,
 * "worksheet": [
 * {
 * "id": "1",
 * "sqlQuery": "Select * from Table1",
 * "workSheetName": "Q1"
 * },
 * {
 * "id": "2",
 * "sqlQuery": "Select * from Table2",
 * "workSheetName": "Q2"
 * }
 * ],
 * "fileName": "C:/temp/##Date##-excelFile1.xls",
 * "preparationProcedureStatement": "{ call proc123 }"
 * }
 * ]
 * }
 */
public class SQLExcelExporterConfig {

    @SerializedName("datasource")
    private Datasource datasource;

    @SerializedName("excelFile")
    private List<ExcelFile> excelFiles;

    public Datasource getDatasource() {
        return datasource;
    }

    public void setDatasource(Datasource datasource) {
        this.datasource = datasource;
    }


    public List<ExcelFile> getExcelFiles() {
        return excelFiles;
    }

    public void setExcelFiles(List<ExcelFile> excelFiles) {
        this.excelFiles = excelFiles;
    }
}
