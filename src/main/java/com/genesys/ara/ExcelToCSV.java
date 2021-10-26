package com.genesys.ara;

import com.codoid.products.exception.FilloException;
import com.codoid.products.fillo.Connection;
import com.codoid.products.fillo.Fillo;
import com.codoid.products.fillo.Recordset;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.Properties;

public class ExcelToCSV {

    static String PROPS_FILE = "./ExcelToCSVJava.properties";
    boolean overwriteCsv = true;
    boolean useTextQualifier = false;
    String textQualifier = "\"";
    String delimiter = ",";

    public ExcelToCSV(String excelFilename) {
        checkProperties();
        System.out.println(String.format("Delimiter: %s",delimiter));
        System.out.println(String.format("Using Text Qualifier: %s",useTextQualifier));
        System.out.println(String.format("Overwrite CSV file: %s",overwriteCsv));
        convert(excelFilename);
    }

    private void checkProperties() {
        try {
            Properties prop = new Properties();
            prop.load(new FileReader(new File(".").getCanonicalPath() + File.separator + PROPS_FILE));
            System.out.println(String.format("Found properties file: %s",PROPS_FILE));
            delimiter = prop.getProperty("delimiter",delimiter);
            overwriteCsv = "true".equals(prop.getProperty("overwrite-csv",
                    String.valueOf(overwriteCsv)));
            useTextQualifier = "true".equals(prop.getProperty("use-text-qualifier",
                    String.valueOf(useTextQualifier)));
        } catch (IOException e) {
            System.err.println(String.format("No properties file named %s found, using defaults",PROPS_FILE));
        }
    }

    private boolean fileExists(String filename) {
        File f = new File(filename);
        System.out.println(String.format("Checking for existence of file %s",filename));
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
        try {
            if (!overwriteCsv && fileExists(csvFilename)) {
                System.err.println("ERROR: CSV file exists and should not be overwriiten!");
                return;
            }
            os = new FileOutputStream(csvFilename);
            osWriter = new OutputStreamWriter(os, StandardCharsets.UTF_8);
            connection = fillo.getConnection(excelFilename);
            System.out.println(String.format("Found Excel file: %s",excelFilename));
            for (String table : connection.getMetaData().getTableNames()) {
                query = String.format("select * from %s", table);
                recordset = connection.executeQuery(query);
                System.out.println(String.format("Received %d records!", recordset.getCount()));
                while (recordset.next()) {
                    for (String fieldName : recordset.getFieldNames()) {
                        currRecord = useTextQualifier ?
                                wrapField(recordset.getField(fieldName)) :
                                recordset.getField(fieldName);
                        csvRow += String.format("%s%s", currRecord, delimiter);
                    }
                    csvRow = csvRow.substring(0, csvRow.length() - 1);
                    System.out.println(csvRow);
                    csvRow += "\n";
                    osWriter.append(csvRow);
                    csvRow = "";
                }
            }
            recordset.close();
            connection.close();
        } catch (FilloException e) {
            System.err.println(e.getMessage());
            e.printStackTrace();
        } catch (FileNotFoundException e) {
            System.err.println("ERROR: Could not write output CSV file.");
        } catch (IOException e) {
            System.err.println("ERROR: Could not append to CSV file");
        }
        finally {
            try {
                if (osWriter != null) osWriter.close();
            } catch (IOException e) {
                System.err.println("ERROR: " + e.getMessage());
            }
            try {
                os.close();
            } catch (IOException e) {
                System.err.println("ERROR: " + e.getMessage());
            }
        }
    }

    private String wrapField(String field) {
        return textQualifier + field + textQualifier;
    }

    public static void main(String [] args) {
        System.out.println("Java Excel CSV Converter 1.0.0");
        if (args.length != 1) {
            System.out.println("Usage: java -jar ExcelToCSV.jar <filename.xlsx>");
        } else {
            new ExcelToCSV(args[0]);
        }
    }
}


///TODO: Add logging framework... log4j?