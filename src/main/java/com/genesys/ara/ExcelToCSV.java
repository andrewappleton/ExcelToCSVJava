package com.genesys.ara;

import com.codoid.products.exception.FilloException;
import com.codoid.products.fillo.Connection;
import com.codoid.products.fillo.Fillo;
import com.codoid.products.fillo.Recordset;
import org.pmw.tinylog.Logger;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.Properties;

public class ExcelToCSV {

    static String PROPS_FILE = "./ExcelToCSVJava.properties";
    static String TEXT_QUALIFIER = "\"";

    boolean showExcelOutput = false;
    boolean overwriteCsv = true;
    boolean useTextQualifier = false;
    String delimiter = ",";

    public ExcelToCSV(String excelFilename) {
        checkProperties();
        Logger.info(String.format("Delimiter: %s",delimiter));
        Logger.info(String.format("Using Text Qualifier: %s",useTextQualifier));
        Logger.info(String.format("Overwrite CSV file: %s",overwriteCsv));
        convert(excelFilename);
    }

    private void checkProperties() {
        try {
            Properties prop = new Properties();
            prop.load(new FileReader(new File(".").getCanonicalPath() + File.separator + PROPS_FILE));
            Logger.info(String.format("Found properties file: %s",PROPS_FILE));
            delimiter = prop.getProperty("delimiter",delimiter);
            overwriteCsv = "true".equals(prop.getProperty("overwrite-csv",
                    String.valueOf(overwriteCsv)));
            useTextQualifier = "true".equals(prop.getProperty("use-text-qualifier",
                    String.valueOf(useTextQualifier)));
            showExcelOutput = "true".equals(prop.getProperty("show-excel-output",
                    String.valueOf(showExcelOutput)));
        } catch (IOException e) {
            Logger.warn(String.format("No properties file named %s found, using defaults",PROPS_FILE));
        }
    }

    private boolean fileExists(String filename) {
        File f = new File(filename);
        Logger.info(String.format("Checking for existence of file %s",filename));
        return f.exists();
    }

    private void convert(String excelFilename) {
        Fillo fillo = new Fillo();
        Recordset recordset = null;
        Connection connection = null;
        OutputStream os = null;
        OutputStreamWriter osWriter = null;
        String csvFilename = excelFilename.substring(0,excelFilename.lastIndexOf('.'))+".csv";
        String query = "";
        String csvRow = "";
        String currRecord = "";
        int csvWriteCount = 0;
        try {
            if (!overwriteCsv && fileExists(csvFilename)) {
                Logger.error("ERROR: CSV file exists and should not be overwritten!");
                return;
            }
            os = new FileOutputStream(csvFilename);
            osWriter = new OutputStreamWriter(os, StandardCharsets.UTF_8);
            connection = fillo.getConnection(excelFilename);
            Logger.info(String.format("Found Excel file: %s",excelFilename));
            for (String table : connection.getMetaData().getTableNames()) {
                query = String.format("select * from %s", table);
                recordset = connection.executeQuery(query);
                Logger.info(String.format("Received %d records!", recordset.getCount()));
                while (recordset.next()) {
                    for (String fieldName : recordset.getFieldNames()) {
                        currRecord = useTextQualifier ?
                                wrapField(recordset.getField(fieldName)) :
                                recordset.getField(fieldName);
                        csvRow += String.format("%s%s", currRecord, delimiter);
                    }
                    csvRow = csvRow.substring(0, csvRow.length() - 1);
                    if (showExcelOutput)
                        System.out.println(csvRow);
                    csvRow += "\n";
                    osWriter.append(csvRow);
                    csvWriteCount ++;
                    csvRow = "";
                }
            }
            Logger.info(String.format("Wrote %d records to %s file",csvWriteCount,csvFilename));
            recordset.close();
            connection.close();
        } catch (FilloException e) {
            Logger.error(e.getMessage());
            e.printStackTrace();
        } catch (FileNotFoundException e) {
            Logger.error("ERROR: Could not write output CSV file.");
        } catch (IOException e) {
            Logger.error("ERROR: Could not append to CSV file");
        }
        finally {
            try {
                if (osWriter != null) osWriter.close();
            } catch (IOException e) {
                Logger.error("ERROR: " + e.getMessage());
            }
            try {
                os.close();
            } catch (IOException e) {
                Logger.error("ERROR: " + e.getMessage());
            }
        }
    }

    private String wrapField(String field) {
        return TEXT_QUALIFIER + field + TEXT_QUALIFIER;
    }

    public static void main(String [] args) {
        Logger.info("Java Excel CSV Converter 1.0.0");
        if (args.length != 1) {
            Logger.info("Usage: java -jar ExcelToCSV.jar <filename.xlsx>");
        } else {
            new ExcelToCSV(args[0]);
        }
    }
}