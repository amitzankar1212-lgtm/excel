package com.excelReader.excel;

import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;

public class SourceExcelReader {

    private static final Logger logger = LoggerFactory.getLogger(SourceExcelReader.class);

    public static Map<String, InvoiceData> readInvoices(
            String headerFilePath,
            String itemsFilePath) throws Exception {

        logger.info("📂 Reading invoice data from files");
        logger.debug("📄 Header file: {}", headerFilePath);
        logger.debug("📄 Items file: {}", itemsFilePath);

        // ===== READ HEADER FILE =====
        logger.debug("🔍 Reading header file data");
        List<Map<String, Object>> headers;
        try (Workbook headerWb = WorkbookFactory.create(new FileInputStream(headerFilePath))) {
            FormulaEvaluator headerEval = headerWb.getCreationHelper().createFormulaEvaluator();
            Sheet headerSheet = headerWb.getSheetAt(0);
            headers = readHeaderRecords(headerSheet, headerEval);
        }
        logger.info("✅ Successfully read {} header records", headers.size());

        // ===== READ ITEMS FILE =====
        logger.debug("🔍 Reading items file data");
        List<Map<String, Object>> items;
        try (Workbook itemsWb = WorkbookFactory.create(new FileInputStream(itemsFilePath))) {
            FormulaEvaluator itemsEval = itemsWb.getCreationHelper().createFormulaEvaluator();
            Sheet itemsSheet = itemsWb.getSheetAt(0);
            items = readItemsSheet(itemsSheet, itemsEval);
        }
        logger.info("✅ Successfully read {} item records", items.size());

        // ===== HEADER DRIVEN MATCHING =====
        logger.debug("🔗 Matching header records with item records");
        Map<String, InvoiceData> result = new LinkedHashMap<>();

        for (Map<String, Object> header : headers) {

            String invoiceNo =
                    String.valueOf(header.get("Invoice Number")).trim();
            logger.debug("🎯 Processing invoice: {}", invoiceNo);

            List<Map<String, Object>> matchedItems = new ArrayList<>();

            for (Map<String, Object> item : items) {

                String itemInv =
                        String.valueOf(item.get("Invoice Number")).trim();

                if (invoiceNo.equals(itemInv)) {
                    matchedItems.add(item);
                }
            }

            if (!matchedItems.isEmpty()) {
                result.put(
                        invoiceNo,
                        new InvoiceData(header, matchedItems)
                );
                logger.debug("✅ Matched {} items for invoice: {}", matchedItems.size(), invoiceNo);
            } else {
                logger.warn("⚠️ No items found for invoice: {}", invoiceNo);
            }
        }

        logger.info("🎉 Successfully created {} invoice records", result.size());
        return result;
    }

    // ================= HELPERS =================

    private static List<Map<String, Object>> readHeaderRecords(
            Sheet sheet, FormulaEvaluator evaluator) {

        List<Map<String, Object>> list = new ArrayList<>();

        Row headerRow = sheet.getRow(0);
        if (headerRow == null) return list;

        for (int r = 1; r <= sheet.getLastRowNum(); r++) {

            Row dataRow = sheet.getRow(r);
            if (dataRow == null) continue;

            Map<String, Object> map = new HashMap<>();

            for (int c = 0; c < headerRow.getLastCellNum(); c++) {

                Cell keyCell = headerRow.getCell(c);
                if (keyCell == null ||
                        keyCell.getCellType() != CellType.STRING)
                    continue;

                String key = keyCell.getStringCellValue().trim();
                if (key.isEmpty()) continue;

                Object value = getValue(dataRow.getCell(c), evaluator);
                map.put(key, value);
            }

            if (!map.isEmpty()) list.add(map);
        }
        return list;
    }

    private static List<Map<String, Object>> readItemsSheet(
            Sheet sheet, FormulaEvaluator evaluator) {

        List<Map<String, Object>> list = new ArrayList<>();

        Row headerRow = sheet.getRow(0);
        if (headerRow == null) return list;

        for (int r = 1; r <= sheet.getLastRowNum(); r++) {

            Row row = sheet.getRow(r);
            if (row == null) continue;

            Map<String, Object> map = new HashMap<>();

            for (int c = 0; c < headerRow.getLastCellNum(); c++) {

                Cell keyCell = headerRow.getCell(c);
                if (keyCell == null ||
                        keyCell.getCellType() != CellType.STRING)
                    continue;

                String key = keyCell.getStringCellValue().trim();
                if (key.isEmpty()) continue;

                Object value = getValue(row.getCell(c), evaluator);
                map.put(key, value);
            }
            list.add(map);
        }
        return list;
    }

    private static Object getValue(Cell cell, FormulaEvaluator evaluator) {

        if (cell == null) return "";

        if (cell.getCellType() == CellType.FORMULA) {
            CellValue cv = evaluator.evaluate(cell);
            return cv.getCellType() == CellType.NUMERIC
                    ? cv.getNumberValue()
                    : cv.getStringValue();
        }

        return cell.getCellType() == CellType.NUMERIC
                ? cell.getNumericCellValue()
                : cell.getStringCellValue();
    }

    /**
     * Get the current status of a specific invoice from the items Excel file
     * @param itemsFilePath Path to the items Excel file
     * @param invoiceNo Invoice number to check
     * @return Current status string, or null if not found
     */
    public static String getInvoiceStatus(String itemsFilePath, String invoiceNo) {
        Workbook workbook = null;
        FileInputStream fis = null;

        try {
            fis = new FileInputStream(itemsFilePath);
            workbook = WorkbookFactory.create(fis);
            Sheet sheet = workbook.getSheetAt(0);

            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                return null;
            }

            // Find the column index for "Status"
            int statusColumnIndex = -1;
            for (int c = 0; c < headerRow.getLastCellNum(); c++) {
                Cell keyCell = headerRow.getCell(c);
                if (keyCell != null && keyCell.getCellType() == CellType.STRING) {
                    String key = keyCell.getStringCellValue().trim();
                    if ("Status".equalsIgnoreCase(key)) {
                        statusColumnIndex = c;
                        break;
                    }
                }
            }

            if (statusColumnIndex == -1) {
                return null;
            }

            // Find the row with the matching invoice number
            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                // Find invoice number column
                for (int c = 0; c < headerRow.getLastCellNum(); c++) {
                    Cell keyCell = headerRow.getCell(c);
                    if (keyCell != null && keyCell.getCellType() == CellType.STRING) {
                        String key = keyCell.getStringCellValue().trim();
                        if ("Invoice Number".equalsIgnoreCase(key)) {
                            Cell invoiceCell = row.getCell(c);
                            if (invoiceCell != null) {
                                String cellInvoiceNo = getValue(invoiceCell, null).toString().trim();
                                if (invoiceNo.trim().equals(cellInvoiceNo)) {
                                    // Found the matching invoice, get status
                                    Cell statusCell = row.getCell(statusColumnIndex);
                                    if (statusCell != null) {
                                        return getValue(statusCell, null).toString().trim();
                                    } else {
                                        return ""; // Status column exists but cell is empty
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return null; // Invoice number not found

        } catch (Exception e) {
            return null; // Error reading file
        } finally {
            // Ensure resources are properly closed
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (Exception e) {
                    // Log but don't throw
                    System.err.println("Warning: Could not close workbook: " + e.getMessage());
                }
            }
            if (fis != null) {
                try {
                    fis.close();
                } catch (Exception e) {
                    // Log but don't throw
                    System.err.println("Warning: Could not close file input stream: " + e.getMessage());
                }
            }
        }
    }

    /**
     * Update the status of a specific invoice in the items Excel file
     * @param itemsFilePath Path to the items Excel file
     * @param invoiceNo Invoice number to update
     * @param newStatus New status to set (pending, Inprocess, completed)
     */
    public static void updateInvoiceStatus(String itemsFilePath, String invoiceNo, String newStatus) throws Exception {
        logger.debug("📝 Updating status for invoice {} to '{}'", invoiceNo, newStatus);

        Workbook workbook = null;
        FileInputStream fis = null;

        try {
            fis = new FileInputStream(itemsFilePath);
            workbook = WorkbookFactory.create(fis);
            Sheet sheet = workbook.getSheetAt(0);

            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                logger.error("❌ Header row not found in items file: {}", itemsFilePath);
                throw new Exception("Header row not found in items file");
            }

            // Find the column index for "Status"
            int statusColumnIndex = -1;
            for (int c = 0; c < headerRow.getLastCellNum(); c++) {
                Cell keyCell = headerRow.getCell(c);
                if (keyCell != null && keyCell.getCellType() == CellType.STRING) {
                    String key = keyCell.getStringCellValue().trim();
                    if ("Status".equalsIgnoreCase(key)) {
                        statusColumnIndex = c;
                        logger.debug("📍 Found Status column at index: {}", c);
                        break;
                    }
                }
            }

            if (statusColumnIndex == -1) {
                logger.error("❌ Status column not found in items file: {}", itemsFilePath);
                throw new Exception("Status column not found in items file");
            }

            // Find ALL rows with the matching invoice number and update their status
            int updatedCount = 0;
            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                // Find invoice number column
                for (int c = 0; c < headerRow.getLastCellNum(); c++) {
                    Cell keyCell = headerRow.getCell(c);
                    if (keyCell != null && keyCell.getCellType() == CellType.STRING) {
                        String key = keyCell.getStringCellValue().trim();
                        if ("Invoice Number".equalsIgnoreCase(key)) {
                            Cell invoiceCell = row.getCell(c);
                            if (invoiceCell != null) {
                                String cellInvoiceNo = getValue(invoiceCell, null).toString().trim();
                                if (invoiceNo.trim().equals(cellInvoiceNo)) {
                                    // Found the matching invoice, update status
                                    Cell statusCell = row.getCell(statusColumnIndex);
                                    if (statusCell == null) {
                                        statusCell = row.createCell(statusColumnIndex);
                                    }
                                    statusCell.setCellValue(newStatus);
                                    updatedCount++;
                                    logger.debug("✅ Updated status in row {} for invoice: {}", r, invoiceNo);
                                    break; // Break inner loop, continue to next row
                                }
                            }
                        }
                    }
                }
            }

            if (updatedCount == 0) {
                logger.error("❌ Invoice number {} not found in items file: {}", invoiceNo, itemsFilePath);
                throw new Exception("Invoice number " + invoiceNo + " not found in items file");
            }

            logger.debug("💾 Saving status updates to file: {}", itemsFilePath);
            // Save the changes
            try (FileOutputStream fos = new FileOutputStream(itemsFilePath)) {
                workbook.write(fos);
            }
            logger.info("✅ Successfully updated status for {} record(s) of invoice: {}", updatedCount, invoiceNo);

        } finally {
            // Ensure resources are properly closed
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (Exception e) {
                    // Log but don't throw
                    System.err.println("Warning: Could not close workbook: " + e.getMessage());
                }
            }
            if (fis != null) {
                try {
                    fis.close();
                } catch (Exception e) {
                    // Log but don't throw
                    System.err.println("Warning: Could not close file input stream: " + e.getMessage());
                }
            }
        }
    }
}
