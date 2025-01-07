package com.excel;

public class ExcelController {
    public static void main(String[] args) {
        String inputFilePath = "resources/mock_data.xlsx"; // input file path or name 

        // Create an instance of ExcelDataSegregator and process the Excel file
        // ExcelDataSegregator segregator = new ExcelDataSegregator();
        ExcelDataSegregator segregator = new ExcelDataSegregator();
        segregator.processExcel(inputFilePath);
    }
}
