package com.excelReader.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;

public class TemplateWriter {

    // Base row positions - these can be adjusted if template changes
    private static final int INVOICE_TITLE_ROW = 8;
    private static final int INVOICE_DATE_ROW = 9;
    private static final int CLIENT_NAME_ROW = 10;
    private static final int INVOICE_NO_ROW = 13; // Moved down by 2 rows (PAN + GSTN, no extra gap)

    // Dynamic positions calculated from base rows
    private static final int PARTICULARS_HEADER_ROW = INVOICE_NO_ROW + 2; // 2 rows after invoice no
    private static final int ITEMS_START_ROW = PARTICULARS_HEADER_ROW + 1;

    public static void generateInvoice(
            String invoiceNo,
            Map<String, Object> header,
            List<Map<String, Object>> items,
            String templatePath,
            String outputDir) throws Exception {

        // Skip if no items in this invoice
        if (items == null || items.isEmpty()) {
            System.out.println("⚠️ Skipping invoice " + invoiceNo + " - no items found");
            return;
        }

        Workbook wb;
        Sheet sheet;

        try {
            wb = WorkbookFactory.create(new FileInputStream(templatePath));
            sheet = wb.getSheetAt(0);
        } catch (Exception e) {
            // Create a basic template if file doesn't exist
            wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook();
            sheet = wb.createSheet("Invoice");

            // Add basic headers
            setCell(sheet, INVOICE_TITLE_ROW, 1, "INVOICE");
            setCell(sheet, INVOICE_DATE_ROW, 1, "Invoice Date:");
            setCell(sheet, CLIENT_NAME_ROW, 1, "Client Name:");
            setCell(sheet, CLIENT_NAME_ROW + 1, 1, "Client PAN:");
            setCell(sheet, CLIENT_NAME_ROW + 2, 1, "Client GSTN:");
            setCell(sheet, INVOICE_NO_ROW, 1, "Invoice No:");
            setCell(sheet, PARTICULARS_HEADER_ROW, 1, "Particulars");
            setCell(sheet, PARTICULARS_HEADER_ROW, 4, "QTY");
            setCell(sheet, PARTICULARS_HEADER_ROW, 5, "Rate");
            setCell(sheet, PARTICULARS_HEADER_ROW, 6, "Amount");
            setCell(sheet, ITEMS_START_ROW, 5, "Subtotal:");
            setCell(sheet, ITEMS_START_ROW + 1, 5, "SGST 9%:");
            setCell(sheet, ITEMS_START_ROW + 2, 5, "CGST 9%:");
            setCell(sheet, ITEMS_START_ROW + 3, 5, "Total:");
            setCell(sheet, ITEMS_START_ROW + 4, 1, "Amount in Words:");
        }

        // ===== HEADER =====
        // Use header data if available, otherwise try to get from first item
        Map<String, Object> headerSource = header;
        if (headerSource == null && !items.isEmpty()) {
            headerSource = items.get(0); // Use first item as fallback for header data
        }

        if (headerSource != null) {
            // Map the fields from clientDetails file
            String invoiceDate = formatDate(headerSource.get("Invoice Date"));
            String clientName = headerSource.get("Client Name") != null ? headerSource.get("Client Name").toString() : "";
            String clientAddress = headerSource.get("Client Address") != null ? headerSource.get("Client Address").toString() : "";
            String clientPan = headerSource.get("Client PAN") != null ? headerSource.get("Client PAN").toString() : "";
            String clientGstn = headerSource.get("Client GSTIN") != null ? headerSource.get("Client GSTIN").toString() : "";

            setCell(sheet, INVOICE_DATE_ROW, 3, invoiceDate);
            setCell(sheet, CLIENT_NAME_ROW, 3, clientName + "\n" + clientAddress);
            setCell(sheet, CLIENT_NAME_ROW + 1, 3, clientPan); // Client PAN row
            setCell(sheet, CLIENT_NAME_ROW + 2, 3, clientGstn); // Client GSTN row
        }
        setCell(sheet, INVOICE_NO_ROW, 3, invoiceNo);

        // Calculate total rows needed: 1 item row for each item
        int numItems = items.size();
        int totalRowsToInsert = numItems; // For each item: 1 content row

        // Shift rows down to make space for items with blank rows in between
        sheet.shiftRows(ITEMS_START_ROW, sheet.getLastRowNum(), totalRowsToInsert);

        // ===== PARTICULARS =====
        double subTotal = 0;
        int currentRowIndex = ITEMS_START_ROW;

        // Create rows for each item
        for (int i = 0; i < numItems; i++) {
            // ===== ITEM ROW =====
            int itemRowNum = currentRowIndex;

            // Set row height for item row
            Row itemRow = sheet.getRow(itemRowNum);
            if (itemRow == null) {
                itemRow = sheet.createRow(itemRowNum);
            }
            itemRow.setHeightInPoints(100); // Height for 5-line Particulars content

            // Get values from the current item in the source invoice
            Map<String, Object> currentItem = items.get(i);
            String particulars = ExcelUtil.buildParticulars(currentItem);
            double rate = ExcelUtil.extractNumber(currentItem.get("Per Month Rate"));
            double workingDays = ExcelUtil.extractNumber(currentItem.get("Working Days"));
            double totalDays = ExcelUtil.extractNumber(currentItem.get("Total Days"));
            double amount = round2((rate / totalDays) * workingDays);

            subTotal += amount;

            // Get other values from source invoice
            String hsnSac = ExcelUtil.extractString(currentItem.get("HSN/SAC"), "998314");
            double qty = ExcelUtil.extractNumber(currentItem.get("QTY"), 1);

            // Set cell values with borders
            setCellWithBordersAndCenter(sheet, itemRowNum, 1, particulars);
            setCellWithBordersAndCenter(sheet, itemRowNum, 3, hsnSac);
            setCellWithBordersAndCenter(sheet, itemRowNum, 4, qty);
            setCellWithBordersAndCenter(sheet, itemRowNum, 5, rate);
            setCellWithBordersAndCenter(sheet, itemRowNum, 6, amount);

            // Merge Particulars columns for item row
            sheet.addMergedRegion(new CellRangeAddress(itemRowNum, itemRowNum, 1, 2));

            currentRowIndex += 1; // Move to next item row
        }

        // Calculate totals after accumulating all amounts
        subTotal = round2(subTotal);
        double sgst = round2(subTotal * 0.09);
        double cgst = round2(subTotal * 0.09);
        double total = round2(subTotal + sgst + cgst);

        // Set the totals in their new positions (shifted down by totalRowsToInsert)
        int totalsRowOffset = ITEMS_START_ROW + totalRowsToInsert;
        setCell(sheet, totalsRowOffset, 5, "Sub Total:");
        setCell(sheet, totalsRowOffset, 6, subTotal);
        setCell(sheet, totalsRowOffset + 1, 5, "SGST:");
        setCell(sheet, totalsRowOffset + 1, 6, sgst);
        setCell(sheet, totalsRowOffset + 2, 5, "CGST:");
        setCell(sheet, totalsRowOffset + 2, 6, cgst);
        setCell(sheet, totalsRowOffset + 3, 5, "Total:");
        setCell(sheet, totalsRowOffset + 3, 6, total);
        setCellWithBordersCenterItalicBold(sheet,totalsRowOffset+ 4,1, amountInWords(total));

        // Auto-size columns with limits
        autoSizeColumnsWithLimit(sheet);

        String outFile =
                outputDir + "invoice_" +
                        invoiceNo.replace("/", "_") +
                        ".xlsm";

        try (FileOutputStream fos = new FileOutputStream(outFile)) {
            wb.write(fos);
        }
        wb.close();

        PdfGeneratorLibreOffice.generatePdfFromExcel(outFile);
    }

    private static void autoSizeColumnsWithLimit(Sheet sheet) {
        // Auto-size all columns to fit content
        for (int col = 1; col <= 6; col++) {
            sheet.autoSizeColumn(col);
        }

        // For merged cells (Particulars column), ensure minimum width for proper display
        if (sheet.getColumnWidth(2) < 5 * 256) {
            sheet.setColumnWidth(2, 5 * 256); // Minimum width for merged area
        }

        // Ensure HSN/SAC column (column 3) has adequate width
        if (sheet.getColumnWidth(3) < 10 * 256) {
            sheet.setColumnWidth(3, 10 * 256); // Minimum width for HSN/SAC column
        }

        // Ensure QTY column (column 4) has adequate width
        if (sheet.getColumnWidth(4) < 8 * 256) {
            sheet.setColumnWidth(4, 8 * 256); // Minimum width for QTY column
        }

        // Ensure Rate column (column 5) has adequate width for rate values
        if (sheet.getColumnWidth(5) < 12 * 256) {
            sheet.setColumnWidth(5, 12 * 256); // Minimum width for Rate column
        }

        // Ensure Amount column (column 6) has adequate width for currency values
        if (sheet.getColumnWidth(6) < 15 * 256) {
            sheet.setColumnWidth(6, 18 * 256); // Minimum width for Amount column
        }
    }

    private static String amountInWords(double amount) {
        long rupees = (long) amount;
        long paise = Math.round((amount - rupees) * 100);

        return convert(rupees) + " Rupees"
                + (paise > 0 ? " and " + convert(paise) + " Paise" : "");
    }



    private static String convert(long number) {
        if (number == 0) return "Zero";

        String[] units = {
                "", "One", "Two", "Three", "Four", "Five",
                "Six", "Seven", "Eight", "Nine", "Ten",
                "Eleven", "Twelve", "Thirteen", "Fourteen",
                "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"
        };

        String[] tens = {
                "", "", "Twenty", "Thirty", "Forty",
                "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"
        };

        if (number < 20)
            return units[(int) number];

        if (number < 100)
            return tens[(int) number / 10] + " " + units[(int) number % 10];

        if (number < 1000)
            return units[(int) number / 100] + " Hundred " + convert(number % 100);

        if (number < 100000)
            return convert(number / 1000) + " Thousand " + convert(number % 1000);

        if (number < 10000000)
            return convert(number / 100000) + " Lakh " + convert(number % 100000);

        return convert(number / 10000000) + " Crore " + convert(number % 10000000);
    }

    private static void setCell(
            Sheet sheet,
            int row,
            int col,
            Object value) {

        if (value == null) return;

        Row r = sheet.getRow(row);
        if (r == null) r = sheet.createRow(row);

        Cell c = r.getCell(col);
        if (c == null) c = r.createCell(col);

        if (value instanceof Number) {
            c.setCellValue(((Number) value).doubleValue());
        } else {
            c.setCellValue(value.toString());
        }
    }

    private static Cell setCellWithBordersAndCenter(
            Sheet sheet,
            int row,
            int col,
            Object value) {

        Workbook wb = sheet.getWorkbook();
        Row r = sheet.getRow(row);
        if (r == null) r = sheet.createRow(row);

        Cell c = r.getCell(col);
        if (c == null) c = r.createCell(col);

        // Set cell value based on type
        if (value instanceof Number) {
            c.setCellValue(((Number) value).doubleValue());
        } else if (value != null) {
            c.setCellValue(value.toString());
        } else {
            c.setCellValue("");
        }

        // Create cell style with borders and top alignment
        CellStyle style = wb.createCellStyle();
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setVerticalAlignment(VerticalAlignment.TOP);
        style.setWrapText(true); // Enable text wrapping

        // If it's a number, add number formatting
        if (value instanceof Number) {
            style.setDataFormat(wb.createDataFormat().getFormat("#,##0.00"));
        }

        c.setCellStyle(style);
        return c;
    }

    private static Cell setCellWithBordersCenterItalicBold(
            Sheet sheet,
            int row,
            int col,
            Object value) {

        Workbook wb = sheet.getWorkbook();
        Row r = sheet.getRow(row);
        if (r == null) r = sheet.createRow(row);

        Cell c = r.getCell(col);
        if (c == null) c = r.createCell(col);

        // Set cell value based on type
        if (value instanceof Number) {
            c.setCellValue(((Number) value).doubleValue());
        } else if (value != null) {
            c.setCellValue(value.toString());
        } else {
            c.setCellValue("");
        }

        // Create cell style with borders, center alignment, italic and bold
        CellStyle style = wb.createCellStyle();
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setVerticalAlignment(VerticalAlignment.TOP);
        style.setWrapText(true); // Enable text wrapping

        // Create font with italic and bold styling
        Font font = wb.createFont();
        font.setItalic(true);
        font.setBold(true);
        style.setFont(font);

        // If it's a number, add number formatting
        if (value instanceof Number) {
            style.setDataFormat(wb.createDataFormat().getFormat("#,##0.00"));
        }

        c.setCellStyle(style);
        return c;
    }

    private static double round2(double value) {
        return Math.round(value * 100.0) / 100.0;
    }

    private static String formatDate(Object dateValue) {
        if (dateValue == null) return "";

        try {
            double excelDate;
            if (dateValue instanceof Number) {
                excelDate = ((Number) dateValue).doubleValue();
            } else {
                // Try to parse as string
                excelDate = Double.parseDouble(dateValue.toString().trim());
            }


            long excelEpoch = (long) ((excelDate - 25569) * 24 * 60 * 60 * 1000); // 25569 is the serial number for 1970-01-01
            Date javaDate = new Date(excelEpoch);

            // Format as dd/MM/yyyy
            SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
            return sdf.format(javaDate);

        } catch (Exception e) {
            // If conversion fails, return the original value as string
            return dateValue.toString();
        }
    }
}