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
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.math.BigDecimal;
import java.sql.CallableStatement;
import java.sql.Connection;
import java.sql.Date;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.SQLType;
import java.sql.Statement;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.TimeZone;
import java.util.regex.Pattern;
import java.util.Map;

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
    	String resultingFileName=null;
        try {

            // Parse the CLI arguments to get the location of the config
            String jsonConfigFilePath = parseCLIArgs(args);

            Gson gson = new Gson();

            // Load the JSON Config file
            logger.info("Reading JSON Config from: " + jsonConfigFilePath);

            BufferedReader br = new BufferedReader(new FileReader(jsonConfigFilePath));

            //convert the json string back to object
            SQLExcelExporterConfig sqlExcelImporterConfig = gson.fromJson(br, SQLExcelExporterConfig.class);

            resultingFileName = performExport(sqlExcelImporterConfig);
            logger.info("Successfully exported data in excel format in file " + resultingFileName);
        } catch (Exception e) {
            logger.error("An error occurred while exporting data to excel. " + e);
            e.printStackTrace();
            System.exit(1);
        }

    }

	/**
	 * @param sqlExcelImporterConfig
	 * @return The full filename of the resulting file or null
	 * @throws InstantiationException
	 * @throws IllegalAccessException
	 * @throws ClassNotFoundException
	 * @throws SQLException
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public static String performExport(SQLExcelExporterConfig sqlExcelImporterConfig) throws InstantiationException,
			IllegalAccessException, ClassNotFoundException, FileNotFoundException, IOException {
		Connection connection=null;
		String fullFilePath = null;
		// Load the driver class
		logger.debug("Loading driver class : " + sqlExcelImporterConfig.getDatasource().getClassName());
		Class.forName(sqlExcelImporterConfig.getDatasource().getClassName()).newInstance();

		// Create a connection to the database
		logger.debug("Creating a connection to the database...");
		try {
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
	
			    fullFilePath = excelFile.getFileName();
			    
			    LocalDateTime ldt = LocalDateTime.now();
			    DateTimeFormatter formmat1 = DateTimeFormatter.ofPattern("yyyyMMdd", Locale.ENGLISH);
			    String fileNamePrefix = formmat1.format(ldt);
			    fullFilePath = fullFilePath.replace("##Date##", fileNamePrefix);
			    
			    FileOutputStream fileOut = new FileOutputStream(fullFilePath);
			    workBook.write(fileOut);
			    fileOut.close();
			    
			    return fullFilePath;
			}
		
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			logger.error("SQLException encountered while reading data : " + e.toString(),e);
			e.printStackTrace();
		}
		finally {
			if (connection != null) {
				try {
					connection.close();
				} 
				catch (SQLException e) {
					logger.error("SQLException encountered while closing the connection : " + e.toString(),e);
				}
			}
		}
		
		logger.info("Data has been successfully exported to excel files.");
		return fullFilePath;
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

    
    private static void mshGenerateHeaderRow(Sheet sheet, ResultSet resultSet) throws SQLException {
        Workbook workbook = sheet.getWorkbook();
        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        
        Row titleRow = sheet.createRow(0);
        ResultSetMetaData metaData = resultSet.getMetaData();
        int columnCount = metaData.getColumnCount();
        for (int colIndex = 0; colIndex < columnCount; colIndex++) {
            String title = metaData.getColumnLabel(colIndex + 1);
            Cell cell = titleRow.createCell(colIndex);
            cell.setCellValue(title);
            CellStyle style = workbook.createCellStyle();
            style.setFont(boldFont);
            cell.setCellStyle(style);
        }
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

    private static final LinkedHashMap<String, String> mapSqlTypeExcelFormat=new LinkedHashMap<String, String>();;
    static {
    	/*
    	<util:map map-class="java.util.LinkedHashMap">
        <entry key="NUMBER\(\d+,2\)" value="0.00" />
        <entry key="NUMBER\(\d+,0\)" value="0" />
        <entry key="NUMBER.*" value="0.###" />
        <entry key="DECIMAL\(\d+,2\)" value="0.00" />
        <entry key="DECIMAL\(\d+,0\)" value="0" />
        <entry key="DECIMAL.*" value="0.###" />
        <entry key=".*CHAR.*" value="text" />
        <entry key="DATE.*" value="dd/MM/yyyy" />
        <entry key="TIMESTAMP.*" value="dd/MM/yyyy" />
    </util:map>
    */
    	
    	mapSqlTypeExcelFormat.put("NUMBER\\(\\d+,2\\)" ,"0.00" );
    	mapSqlTypeExcelFormat.put("NUMBER\\(\\d+,0\\)" ,"0" );
    	mapSqlTypeExcelFormat.put("NUMBER.*" ,"0.###" );
    	mapSqlTypeExcelFormat.put("INT\\(\\d+,0\\)" ,"0" );
    	mapSqlTypeExcelFormat.put("BIGINT\\(\\d+,0\\)" ,"0" );
    	mapSqlTypeExcelFormat.put("BIT\\(1,0\\)" ,"0" );
    	mapSqlTypeExcelFormat.put("DECIMAL\\(\\d+,2\\)" ,"0.00" );
    	mapSqlTypeExcelFormat.put("DECIMAL\\(\\d+,0\\)" ,"0" );
    	mapSqlTypeExcelFormat.put("DECIMAL.*" ,"0.####" );
    	mapSqlTypeExcelFormat.put("NUMERIC.*" ,"0.####" );
    	mapSqlTypeExcelFormat.put(".*CHAR.*" ,"text" );
    	mapSqlTypeExcelFormat.put("DATETIME\\(\\d+,3\\)" ,"dd.MM.yyyy h:mm:ss.000" );
    	mapSqlTypeExcelFormat.put("DATETIME\\(\\d+,0\\)" ,"dd.MM.yyyy h:mm:ss" );
    	mapSqlTypeExcelFormat.put("DATE.*" ,"dd.MM.yyyy" );
    	mapSqlTypeExcelFormat.put("TIMESTAMP.*\"" ,"dd.MM.yyyy h:mm:ss.000" );
    }
    
    private static CellStyle getDataStyle(Workbook workbook, ResultSetMetaData metaData, int colIndex, DataFormat dataFormat) throws SQLException {
        CellStyle dataStyle = workbook.createCellStyle();
        String columnType = metaData.getColumnTypeName(colIndex + 1).toUpperCase();
        columnType += "(" + metaData.getPrecision(colIndex + 1);
        columnType += "," + metaData.getScale(colIndex + 1) + ")";
        String excelFormat = getExcelFormat(columnType);
        final short format = dataFormat.getFormat(excelFormat);
        logger.info("Column "+colIndex+" columnType "+columnType + " excelFormat="+excelFormat);
        dataStyle.setDataFormat(format);
        return dataStyle;
    }
     
    private static String getExcelFormat(String columnType) {
        for (Map.Entry<String, String> entry : mapSqlTypeExcelFormat.entrySet()) {
            if (Pattern.matches(entry.getKey(), columnType)) {
                return entry.getValue();
            }
        }
        return "text";
    }

    private static void populateRows(Sheet sheet, ResultSet rs) throws SQLException {
        int rowCounter = DATA_ROW;
        
//        for (int i=1;i<=rs.getMetaData().getColumnCount();i++) {
//            String colName = rs.getMetaData().getColumnName(i);
//            String colType = rs.getMetaData().getColumnTypeName(i);
//            System.out.println(colName+" of type "+colType);
//	    }
        int columnCount = rs.getMetaData().getColumnCount();
        CellStyle[] dataStyles=new CellStyle[columnCount];
        int rowsOffset=0;
        Workbook workbook=sheet.getWorkbook();
        DataFormat dataFormat =workbook.createDataFormat();
        while (rs.next()) {


        	if (rowCounter==DATA_ROW) {
        		//Row rowd = sheet.createRow(rowCounter);
        		
        		
        		for (int i = 0; i < columnCount; i++) {
                	String typeName=rs.getMetaData().getColumnTypeName(i+1);
                	int type=rs.getMetaData().getColumnType(i+1); 
                	int displaySize=rs.getMetaData().getColumnDisplaySize(i+1);
                	int precisionSize=rs.getMetaData().getPrecision(i+1);
                	int scaleSize=rs.getMetaData().getScale(i+1);
                	String columnClassName=rs.getMetaData().getColumnClassName(i+1);
                	String columnLabel=rs.getMetaData().getColumnLabel(i+1);
                	dataStyles[i]=getDataStyle(sheet.getWorkbook(),rs.getMetaData(),i,dataFormat);
//                	rowd.createCell(i).setCellValue(
//                			"typeName:"+typeName+
//                			"; typeInt="+type+
//                			"; DisplaySize="+displaySize+
//                			"; precisionSize"+precisionSize+
//                			"; scaleSize"+scaleSize+
//                			"; columnLabel="+columnLabel+
//                			"; columnClassName="+columnClassName
//                			);
        		}
        		
//        		rowsOffset=1;
        	}

        	Row row = sheet.createRow(rowCounter+rowsOffset);

            for (int i = 0; i < columnCount; i++) {
            	String typeName=rs.getMetaData().getColumnTypeName(i+1);
            	int type=rs.getMetaData().getColumnType(i+1); 
//            	rs.getMetaData().getColumnDisplaySize(column)
//            	if (type == java.sql.Types.)

            	/*
            	if (rs.getMetaData().getColumnTypeName(i+1).equals("NUMBER"))
            		row.createCell(i).setCellValue(rs.getDouble(i + 1));
            	else
            		row.createCell(i).setCellValue(rs.getString(i + 1));
            
*/            	
                Object value = rs.getObject(i + 1);
                final Cell cell = row.createCell(i);
                if (value == null) {
                    cell.setCellValue("");
                } else {
                    if (value instanceof Calendar) {
                        cell.setCellValue((Calendar) value);
                    } else if (value instanceof Date) {
                        cell.setCellValue((Date) value);
                    } else if (value instanceof java.sql.Timestamp) {
                    	java.sql.Timestamp t=(java.sql.Timestamp) value;
                    	java.util.GregorianCalendar gcal=java.util.GregorianCalendar.from(t.toLocalDateTime().atZone(TimeZone.getDefault().toZoneId()));
                        cell.setCellValue(gcal);
                        //cell.setCellType(CellType.NUMERIC);
                    } else if (value instanceof String) {
                        cell.setCellValue((String) value);
                    } else if (value instanceof Boolean) {
                        cell.setCellValue(((Boolean) value).booleanValue());
                    } else if (value instanceof Double) {
                        cell.setCellValue(((Double) value).doubleValue());
                    } else if (value instanceof Integer) {
                        cell.setCellValue(((Integer) value).doubleValue());
                    } else if (value instanceof Long) {
                        cell.setCellValue(((Long) value).doubleValue());
                    } else if (value instanceof BigDecimal) {
                    	BigDecimal bd=(BigDecimal)value;
                    	double dbd = bd.doubleValue();
                    	if ( !Double.isNaN(dbd) && Double.isFinite(dbd) && 
                    				(bd.equals(new BigDecimal(dbd))
                    						|| bd.precision() < 16
                    						)) 
                    	{
                    		cell.setCellValue(dbd);
                    	}
                    	else 
                    	{
                    		cell.setCellValue(bd.toPlainString());
                    	}
                    } else {
                    	cell.setCellValue(rs.getString(i+1));
                    }
                    cell.setCellStyle(dataStyles[i]);
                }
            	
            }
            rowCounter++;           	

        }
        
        for (int i = 0; i < columnCount; i++) {
            sheet.autoSizeColumn(i);
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