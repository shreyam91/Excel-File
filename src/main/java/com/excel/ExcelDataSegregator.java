package com.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDataSegregator {

    // Method to process the Excel file and segregate data by status
    public void processExcel(String inputFilePath) {
        try (FileInputStream inputStream = new FileInputStream(new File(inputFilePath))) {
            Workbook workbook = new XSSFWorkbook(inputStream);
            Sheet sheet = workbook.getSheetAt(0); // Data is in the first sheet

            // Map to hold data segregated by status
            HashMap<String, List<Row>> statusData = new HashMap<>();

            // Loop through each row in the sheet (skip header row)
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);

                int statusColumnIndex = 5;  // Assuming status is in column 6 (index 5)
                // Get the status value from the specified column
                String status = row.getCell(statusColumnIndex).getStringCellValue();

                // Add the row to the corresponding status list
                List<Row> statusRows = statusData.get(status);
                if (statusRows == null) {
                    statusRows = new ArrayList<>();
                    statusData.put(status, statusRows);
                }
                statusRows.add(row);
            }

            // Create a new workbook for the split data
            XSSFWorkbook newWorkbook = new XSSFWorkbook();

            // Loop through each status and create a new sheet with its data
            for (HashMap.Entry<String, List<Row>> entry : statusData.entrySet()) {
                // Generate a filename using the status and current date/time
                String status = entry.getKey();
                String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
                String outputFileName = "data_" + status + "_" + timestamp + ".xlsx";

                ExcelUtils.createSheetWithData(newWorkbook, status, entry.getValue());

                // Write the split data to a new Excel file
                try (FileOutputStream outputStream = new FileOutputStream(outputFileName)) {
                    newWorkbook.write(outputStream);
                }

                System.out.println("Excel file for status '" + status + "' saved as: " + outputFileName);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
