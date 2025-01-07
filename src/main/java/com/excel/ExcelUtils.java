package com.excel;

import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {

    // Method to create a new sheet in the workbook with the specified name and data
    public static void createSheetWithData(XSSFWorkbook workbook, String sheetName, List<Row> rows) {
        XSSFSheet newSheet = workbook.createSheet(sheetName);

        // Copy data into the new sheet, including the header row
        int targetRow = 0;

        // Assuming the first row contains the header
        Row headerRow = rows.get(0);
        Row newHeaderRow = newSheet.createRow(targetRow++);
        copyRow(headerRow, newHeaderRow);

        // Copy remaining rows
        for (Row row : rows) {
            Row newRow = newSheet.createRow(targetRow++);
            copyRow(row, newRow);
        }
    }

    // Helper method to copy data from one row to another
    public static void copyRow(Row sourceRow, Row targetRow) {
        for (int i = 0; i < sourceRow.getPhysicalNumberOfCells(); i++) {
            Cell sourceCell = sourceRow.getCell(i);
            Cell targetCell = targetRow.createCell(i);

            // Copy cell value based on its type
            if (sourceCell.getCellType() == CellType.STRING) {
                targetCell.setCellValue(sourceCell.getStringCellValue());
            } else if (sourceCell.getCellType() == CellType.NUMERIC) {
                targetCell.setCellValue(sourceCell.getNumericCellValue());
            } else if (sourceCell.getCellType() == CellType.BOOLEAN) {
                targetCell.setCellValue(sourceCell.getBooleanCellValue());
            }
            // Add more cell types if needed (Date, Formula, etc.)
        }
    }
}
