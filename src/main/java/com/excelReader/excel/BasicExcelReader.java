package com.excelReader.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import java.io.FileInputStream;
import java.util.*;

public class BasicExcelReader {

    public static void main(String[] args) throws Exception {

        // File path to your Excel file
        String filePath = "C:/Users/amitz/Downloads/source_invoices.xlsx";

        // ===== BASIC EXCEL READING =====

        // Step 1: Open the Excel file
        FileInputStream fis = new FileInputStream(filePath);
        Workbook workbook = WorkbookFactory.create(fis);

        // Step 2: Get the first sheet (or specify sheet name)
        Sheet sheet = workbook.getSheetAt(0); // First sheet
        // Or use: Sheet sheet = workbook.getSheet("Sheet1");

        // Step 3: Create variables to store data
        List<Map<String, Object>> allData = new ArrayList<>();
        Map<String, List<Object>> columnData = new HashMap<>();
        List<String> headers = new ArrayList<>();

        // Step 4: Read header row (usually first row)
        Row headerRow = sheet.getRow(0); // Row 0 is the first row
        if (headerRow != null) {
            for (Cell cell : headerRow) {
                String headerValue = getCellValueAsString(cell);
                if (headerValue != null && !headerValue.trim().isEmpty()) {
                    headers.add(headerValue.trim());
                    columnData.put(headerValue.trim(), new ArrayList<>());
                }
            }
        }

        System.out.println("Headers found: " + headers);

        // Step 5: Iterate through data rows
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue; // Skip empty rows

            // Store entire row as a Map
            Map<String, Object> rowData = new HashMap<>();

            // Iterate through cells in this row
            for (int colIndex = 0; colIndex < headers.size(); colIndex++) {
                Cell cell = row.getCell(colIndex);
                Object cellValue = getCellValue(cell);

                String header = headers.get(colIndex);
                rowData.put(header, cellValue);

                // Also store in column-wise data
                columnData.get(header).add(cellValue);
            }

            // Only add row if it has some data
            if (!rowData.isEmpty()) {
                allData.add(rowData);
            }
        }

        // ===== DISPLAY STORED DATA =====

        System.out.println("\n=== ROW-WISE DATA ===");
        for (int i = 0; i < allData.size(); i++) {
            System.out.println("Row " + (i + 1) + ": " + allData.get(i));
        }

        System.out.println("\n=== COLUMN-WISE DATA ===");
        for (String header : headers) {
            System.out.println(header + ": " + columnData.get(header));
        }

        // ===== ACCESS SPECIFIC VALUES =====

        System.out.println("\n=== ACCESS SPECIFIC VALUES ===");
        if (!allData.isEmpty()) {
            // Get first row data
            Map<String, Object> firstRow = allData.get(0);

            // Access by column name
            Object invoiceNumber = firstRow.get("Invoice Number");
            Object resourceName = firstRow.get("Resource Name");
            Object perMonthRate = firstRow.get("Per Month Rate");

            System.out.println("First row - Invoice Number: " + invoiceNumber);
            System.out.println("First row - Resource Name: " + resourceName);
            System.out.println("First row - Per Month Rate: " + perMonthRate);
        }

        // Close resources
        workbook.close();
        fis.close();
    }

    // Helper method to get cell value as Object
    private static Object getCellValue(Cell cell) {
        if (cell == null) return null;

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                // Check if it's a date
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                } else {
                    return cell.getNumericCellValue();
                }
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case FORMULA:
                // For simplicity, return the formula result
                try {
                    return cell.getNumericCellValue();
                } catch (Exception e) {
                    return cell.getStringCellValue();
                }
            default:
                return null;
        }
    }

    // Helper method to get cell value as String
    private static String getCellValueAsString(Cell cell) {
        if (cell == null) return null;

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return String.valueOf(cell.getNumericCellValue());
                } catch (Exception e) {
                    return cell.getStringCellValue();
                }
            default:
                return null;
        }
    }
}