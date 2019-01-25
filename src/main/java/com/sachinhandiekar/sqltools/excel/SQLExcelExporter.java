package com.sachinhandiekar.sqltools.excel;

import com.google.gson.Gson;
import com.sachinhandiekar.sqltools.excel.model.SQLExcelExporterConfig;
import com.sachinhandiekar.sqltools.excel.model.Worksheet;
import com.sachinhandiekar.sqltools.excel.model.ExcelFile;
import org.apache.commons.cli.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.sql.CallableStatement;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Locale;

/**
 * A main class to run the SQLExcelReporter
 *
 * Usage : java -jar SQLExcelExporter.jar -config <file>
 *
 */
public class SQLExcelExporter {

    private static final int HEADER_ROW = 0;

    private static final int DATA_ROW = 1;

    private static final Logger logger = LoggerFactory.getLogger(SQLExcelExporter.class);

    public static void main(String[] args) {
    	Connection connection = null;
        try {

            // Parse the CLI arguments to get the location of the config
            String jsonConfigFilePath = parseCLIArgs(args);

            Gson gson = new Gson();

            // Load the JSON Config file
            logger.info("Reading JSON Config from: " + jsonConfigFilePath);

            BufferedReader br = new BufferedReader(new FileReader(jsonConfigFilePath));

            //convert the json string back to object
            SQLExcelExporterConfig sqlExcelImporterConfig = gson.fromJson(br, SQLExcelExporterConfig.class);

            // Load the driver class
            logger.debug("Loading driver class : " + sqlExcelImporterConfig.getDatasource().getClassName());
            Class.forName(sqlExcelImporterConfig.getDatasource().getClassName()).newInstance();

            // Create a connection to the database
            logger.debug("Creating a connection to the database...");
            connection = DriverManager.getConnection(sqlExcelImporterConfig.getDatasource().getJdbcUrl(),
                    sqlExcelImporterConfig.getDatasource().getUserName(),
                    sqlExcelImporterConfig.getDatasource().getPassword());


            //Iterate through the list of excelFile
            List<ExcelFile> excelFileList = sqlExcelImporterConfig.getExcelFiles();

            for (ExcelFile excelFile : excelFileList) {
        		logger.info("*ExcelFile " + excelFile.getId() + " Large: " + excelFile.isLarge());
            	Workbook workBook = null;
            	if (excelFile.isLarge()) 
            		workBook = new SXSSFWorkbook(1000);
            	else if (excelFile.getFileName().toLowerCase().endsWith(".xlsx")) 
            		workBook = new XSSFWorkbook();
            	else if (excelFile.getFileName().toLowerCase().endsWith(".xls")) 
            		workBook = new HSSFWorkbook();
            	else
            	{
            		logger.error("File name can have extensions xls or xlsx only.");
            		System.exit(1);
            	}
            	
            	if (excelFile.getPreparationProcedureStatement() != null && 
            			excelFile.getPreparationProcedureStatement().trim() != "" )
            	{
            		logger.info("**Stored procedure " + excelFile.getPreparationProcedureStatement());
            		executeStroedProcedure(excelFile.getPreparationProcedureStatement(), connection);
            	}

                // Iterate through the list of worksheet for each excelFile
                List<Worksheet> worksheets = excelFile.getWorksheets();

                for (Worksheet workSheet : worksheets) {
            		logger.info("**Worksheet " + workSheet.getId());
                    ResultSet resultSet = getResultSetForQuery(workSheet.getSqlQuery(), connection);
                    generateWorksheet(workSheet.getWorkSheetName(), workBook, resultSet);
                    resultSet.close();
                }

                String fullFilePath = excelFile.getFileName();
                
                LocalDateTime ldt = LocalDateTime.now();
                DateTimeFormatter formmat1 = DateTimeFormatter.ofPattern("yyyyMMdd", Locale.ENGLISH);
                String fileNamePrefix = formmat1.format(ldt);
                fullFilePath = fullFilePath.replace("##Date##", fileNamePrefix);
                
                FileOutputStream fileOut = new FileOutputStream(fullFilePath);
                workBook.write(fileOut);
                fileOut.close();
            }
            logger.info("Data has been successfully exported to excel files.");
        } catch (Exception e) {
            logger.error("An error occurred while exporting data to excel. " + e);
            e.printStackTrace();
            System.exit(1);
        }
        finally {
            try {
				if (connection != null && !connection.isClosed())
				{
					connection.close();
				}
			} catch (SQLException e) {
			}
        }
    }

    private static String parseCLIArgs(String[] args) {
        CommandLineParser parser = new DefaultParser();
        String configFilePath = null;
        try {

            Options options = new Options();

            // add t option
            options.addOption("config", true, "Configuration file for SQLExcelExporter");
            // parse the command line arguments
            CommandLine line = parser.parse(options, args);

            if (line.hasOption("config")) {
                configFilePath = line.getOptionValue("config");
            } else {
                HelpFormatter formatter = new HelpFormatter();
                formatter.printHelp("SQLExcelExporter -config C:/pathto/sqlExcelExporter.json", options);
                System.exit(0);
            }
        } catch (Exception exp) {
            // oops, something went wrong
            logger.error("Parsing failed.  Reason: " + exp.getMessage());
        }

        return configFilePath;

    }

    
    private static void generateHeaderRow(Sheet sheet, ResultSet rs) throws SQLException {

        Row headerRow = sheet.createRow(HEADER_ROW);

        ResultSetMetaData resultsetMetadata = rs.getMetaData();
        int columnCount = resultsetMetadata.getColumnCount();

        for (int i = 0; i < columnCount; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(resultsetMetadata.getColumnName(i + 1));
        }
    }


    private static void populateRows(Sheet sheet, ResultSet rs) throws SQLException {
        int rowCounter = DATA_ROW;
        
//        for (int i=1;i<=rs.getMetaData().getColumnCount();i++) {
//            String colName = rs.getMetaData().getColumnName(i);
//            String colType = rs.getMetaData().getColumnTypeName(i);
//            System.out.println(colName+" of type "+colType);
//	    }
        
        while (rs.next()) {

        	Row row = sheet.createRow(rowCounter);
            int columnCount = rs.getMetaData().getColumnCount();

            for (int i = 0; i < columnCount; i++) {
            	if (rs.getMetaData().getColumnTypeName(i+1).equals("NUMBER"))
            		row.createCell(i).setCellValue(rs.getDouble(i + 1));
            	else
            		row.createCell(i).setCellValue(rs.getString(i + 1));
            }
            rowCounter++;
        }
    }


    private static ResultSet getResultSetForQuery(String query, Connection connection) throws SQLException {
        Statement statement = connection.createStatement();
        return statement.executeQuery(query);
    }
    
    private static void executeStroedProcedure(String query, Connection connection) throws SQLException {
    	// Example "{ call proc3 }";
    	CallableStatement cs = connection.prepareCall(query);
    	cs.execute();
    }

    /**
     * Generate a worksheet and populate it with a header row and data rows
     *
     * @param workSheetName name of the worksheet
     * @param workbook a reference to the HFFSWorkbook (Apache POI)
     * @param resultSet a JDBC resultset containing the data
     * @throws SQLException if any error occurs
     */
    private static void generateWorksheet(String workSheetName, Workbook workbook, ResultSet resultSet) throws SQLException {
    	Sheet workSheet = workbook.createSheet(workSheetName);

        // Create the first Header row
        // Get all the column names from the ResultSet
        generateHeaderRow(workSheet, resultSet);
        
        workSheet.createFreezePane(0, 1);

        // Populate the data in the rows
        populateRows(workSheet, resultSet);
    }
}