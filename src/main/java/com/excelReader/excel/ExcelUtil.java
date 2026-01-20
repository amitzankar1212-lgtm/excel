package com.excelReader.excel;


import org.apache.poi.ss.usermodel.*;

import java.util.Map;

public class ExcelUtil {

    public static Cell getCell(Sheet sheet, int row, int col) {

        Row r = sheet.getRow(row);
        if (r == null) r = sheet.createRow(row);

        Cell c = r.getCell(col);
        if (c == null) c = r.createCell(col);

        return c;
    }

    // ===== LABEL FINDER (ROBUST) =====
    public static Cell findCell(Sheet sheet, String label) {

        String target = normalize(label);

        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING) {
                    String cellText = normalize(cell.getStringCellValue());
                    if (cellText.contains(target)) {
                        return cell;
                    }
                }
            }
        }
        return null;
    }

    public static int findRow(Sheet sheet, String label) {

        Cell cell = findCell(sheet, label);
        return cell != null ? cell.getRowIndex() : -1;
    }

    private static String normalize(String s) {
        return s.toLowerCase()
                .replace(":", "")
                .replace("-", "")
                .replaceAll("\\s+", "");
    }

    // ===== NUMBER HELPERS =====
    public static double extractNumber(Object value) {

        if (value == null) return 0;

        if (value instanceof Number) {
            return ((Number) value).doubleValue();
        }

        String s = value.toString().trim();

        if (s.contains("-")) {
            s = s.split("-")[0];
        }

        return Double.parseDouble(s);
    }

    public static double extractNumber(Object value, int defaultValue) {
        if (value == null) return defaultValue;

        try {
            if (value instanceof Number) {
                return ((Number) value).intValue();
            }

            String s = value.toString().trim();
            if (s.isEmpty()) return defaultValue;

            if (s.contains("-")) {
                s = s.split("-")[0];
            }

            return Integer.parseInt(s);
        } catch (NumberFormatException e) {
            return defaultValue;
        }
    }

    public static String extractString(Object value, String defaultValue) {
        if (value == null) return defaultValue;

        String s = value.toString().trim();
        return s.isEmpty() ? defaultValue : s;
    }


    public static String buildParticulars(Map<String, Object> src) {

        double totalDays   = extractNumber(src.get("Total Days"));
        double workingDays = extractNumber(src.get("Working Days"));
        double rate        = extractNumber(src.get("Per Month Rate"));

        double displayAmount =
                (rate / totalDays) * workingDays;

        return "Software Development Consultancy\n" +
                "Charges-" + src.get("Resource Name") + "\n" +
                "Actual working days in month-" + (int) totalDays + "\n" +
                "Present working days in month-" + (int) workingDays + "\n" +
                "Calculation:(" +
                (long) rate + "/" +
                (int) totalDays + "*" +
                (int) workingDays + ")=" +
                Math.round(displayAmount);
    }


}
