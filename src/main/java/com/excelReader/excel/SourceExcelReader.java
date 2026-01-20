package com.excelReader.excel;

import org.apache.poi.ss.usermodel.*;
import java.io.FileInputStream;
import java.util.*;

public class SourceExcelReader {

    public static Map<String, InvoiceData> readInvoices(
            String headerFilePath,
            String itemsFilePath) throws Exception {

        // ===== READ HEADER FILE =====
        Workbook headerWb =
                WorkbookFactory.create(new FileInputStream(headerFilePath));
        FormulaEvaluator headerEval =
                headerWb.getCreationHelper().createFormulaEvaluator();

        Sheet headerSheet = headerWb.getSheetAt(0);
        List<Map<String, Object>> headers =
                readHeaderRecords(headerSheet, headerEval);

        headerWb.close();

        // ===== READ ITEMS FILE =====
        Workbook itemsWb =
                WorkbookFactory.create(new FileInputStream(itemsFilePath));
        FormulaEvaluator itemsEval =
                itemsWb.getCreationHelper().createFormulaEvaluator();

        Sheet itemsSheet = itemsWb.getSheetAt(0);
        List<Map<String, Object>> items =
                readItemsSheet(itemsSheet, itemsEval);

        itemsWb.close();

        // ===== HEADER DRIVEN MATCHING =====
        Map<String, InvoiceData> result = new LinkedHashMap<>();

        for (Map<String, Object> header : headers) {

            String invoiceNo =
                    String.valueOf(header.get("Invoice Number")).trim();

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
            }
        }

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
}
