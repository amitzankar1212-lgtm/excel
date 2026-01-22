package com.excelReader.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;

public class TemplateWriter {

    private static final Logger logger = LoggerFactory.getLogger(TemplateWriter.class);

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

        // Check currency from header data or first item to determine tax calculations and layout
        boolean isDollarCurrency = false;
        Map<String, Object> currencySource = header;
        if (currencySource == null && !items.isEmpty()) {
            currencySource = items.get(0);
        }

        if (currencySource != null) {
            String currency = currencySource.get("Currency") != null ? currencySource.get("Currency").toString().toLowerCase() : "";
            if ("dollar".equals(currency)) {
                isDollarCurrency = true;
            }
        }

        // Calculate dynamic row positions based on currency
        int invoiceNoRow = isDollarCurrency ? CLIENT_NAME_ROW + 1 : INVOICE_NO_ROW;
        int particularsHeaderRow = invoiceNoRow + 2;
        int itemsStartRow = particularsHeaderRow + 1;

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
            setCellWithBordersLeftTop(sheet, INVOICE_DATE_ROW, 1, "Invoice Date:");
            setCellWithBordersLeftTop(sheet, CLIENT_NAME_ROW, 1, "Client Name:");

            // Calculate positions based on currency
            int templateInvoiceNoRow = CLIENT_NAME_ROW + 1; // Default after client name
            int templateParticularsRow = PARTICULARS_HEADER_ROW; // Default position

            if (!isDollarCurrency) {
                // For non-dollar currency, show PAN and GSTN
                setCell(sheet, CLIENT_NAME_ROW + 1, 1, "Client PAN:");
                setCell(sheet, CLIENT_NAME_ROW + 2, 1, "Client GSTN:");
                templateInvoiceNoRow = CLIENT_NAME_ROW + 3; // Invoice No after PAN and GSTN
            }

            setCell(sheet, templateInvoiceNoRow, 1, "Invoice No:");
            setCell(sheet, templateParticularsRow, 1, "Particulars");
            setCell(sheet, templateParticularsRow, 4, "QTY");
            setCell(sheet, templateParticularsRow, 5, "Rate");
            setCell(sheet, templateParticularsRow, 6, "Amount");

            // Add totals section based on currency
            setCell(sheet, ITEMS_START_ROW, 5, "Subtotal:");
            if (isDollarCurrency) {
                // For dollar currency, only show Subtotal and Total
                setCell(sheet, ITEMS_START_ROW + 1, 5, "Total:");
                setCell(sheet, ITEMS_START_ROW + 2, 1, "Amount in Words:");
            } else {
                // For regular currency, show CGST/SGST
                setCell(sheet, ITEMS_START_ROW + 1, 5, "SGST 9%:");
                setCell(sheet, ITEMS_START_ROW + 2, 5, "CGST 9%:");
                setCell(sheet, ITEMS_START_ROW + 3, 5, "Total:");
                setCell(sheet, ITEMS_START_ROW + 4, 1, "Amount in Words:");
            }
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

            setCellWithBordersLeftAlign(sheet, INVOICE_DATE_ROW, 3, invoiceDate);
            setCellWithBordersLeftTop(sheet, CLIENT_NAME_ROW, 3, clientName + "\n" + clientAddress);

            // Set fixed height for Client Name & Address row to match invoice_template2
            Row addressRow = sheet.getRow(CLIENT_NAME_ROW);
            if (addressRow != null) {
                addressRow.setHeightInPoints(50f); // Fixed height to match invoice_template2
            }

            // Only show PAN and GSTN for non-dollar currency
            if (!isDollarCurrency) {
                String clientPan = headerSource.get("Client PAN") != null ? headerSource.get("Client PAN").toString() : "";
                String clientGstn = headerSource.get("Client GSTIN") != null ? headerSource.get("Client GSTIN").toString() : "";
                setCellWithBordersLeftAlign(sheet, CLIENT_NAME_ROW + 1, 3, clientPan); // Client PAN row
                setCellWithBordersLeftAlign(sheet, CLIENT_NAME_ROW + 2, 3, clientGstn); // Client GSTN row
            }

            setCellWithBordersLeftAlign(sheet, invoiceNoRow, 3, invoiceNo);
        } else {
            // Fallback if no header data
            setCellWithBordersLeftAlign(sheet, invoiceNoRow, 3, invoiceNo);
        }

        // Calculate total rows needed: 1 item row for each item
        int numItems = items.size();
        int totalRowsToInsert = numItems; // For each item: 1 content row

        // Shift rows down to make space for items with blank rows in between
        sheet.shiftRows(itemsStartRow, sheet.getLastRowNum(), totalRowsToInsert);

        // ===== PARTICULARS =====
        double subTotal = 0;
        int currentRowIndex = itemsStartRow;

        // Create rows for each item
        for (int i = 0; i < numItems; i++) {
            // ===== ITEM ROW =====
            int itemRowNum = currentRowIndex;

            // Set row height for item row
            Row itemRow = sheet.getRow(itemRowNum);
            if (itemRow == null) {
                itemRow = sheet.createRow(itemRowNum);
            }
            itemRow.setHeightInPoints(35); // Further reduced height for Particulars content

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
            setCellWithBordersLeftTop(sheet, itemRowNum, 1, particulars);
            setCellWithBordersCenterTop(sheet, itemRowNum, 3, hsnSac);
            setCellWithBordersCenterTop(sheet, itemRowNum, 4, qty);
            setCellWithBordersRightTop(sheet, itemRowNum, 5, rate);
            setCellWithBordersRightTop(sheet, itemRowNum, 6, amount);

            // Merge Particulars columns for item row
            sheet.addMergedRegion(new CellRangeAddress(itemRowNum, itemRowNum, 1, 2));

            currentRowIndex += 1; // Move to next item row
        }

        // Calculate totals after accumulating all amounts
        subTotal = round2(subTotal);

        // Set the totals in their new positions (shifted down by totalRowsToInsert)
        int totalsRowOffset = itemsStartRow + totalRowsToInsert;
        setCellWithBordersLeftTop(sheet, totalsRowOffset, 5, "Sub Total:");
        setCellWithBordersRightTop(sheet, totalsRowOffset, 6, subTotal);

        if (isDollarCurrency) {
            // For dollar currency, only show Subtotal and Total
            double total = subTotal; // No taxes for dollar currency
            setCellWithBordersLeftTop(sheet, totalsRowOffset + 1, 5, "Total:");
            setCellWithBordersRightTop(sheet, totalsRowOffset + 1, 6, total);

            // Set height for Amount in Words row to exactly two lines
            Row amountInWordsRow = sheet.getRow(totalsRowOffset + 1);
            if (amountInWordsRow == null) {
                amountInWordsRow = sheet.createRow(totalsRowOffset +  1);
            }
            amountInWordsRow.setHeightInPoints(18.7f); // Height for exactly two lines (45% reduction from 34f)
            setCellWithBordersCenterItalicBoldNoWrap(sheet, totalsRowOffset + 2, 1, "Amount in words: " + amountInWords(total, true));
        } else {
            // For regular currency, show CGST/SGST
            double sgst = round2(subTotal * 0.09);
            double cgst = round2(subTotal * 0.09);
            double total = round2(subTotal + sgst + cgst);

            setCellWithBordersLeftAlign(sheet, totalsRowOffset + 1, 5, "SGST:");
            setCellWithBordersRightTop(sheet, totalsRowOffset + 1, 6, sgst);
            setCellWithBordersLeftAlign(sheet, totalsRowOffset + 2, 5, "CGST:");
            setCellWithBordersRightTop(sheet, totalsRowOffset + 2, 6, cgst);
            setCellWithBordersLeftAlign(sheet, totalsRowOffset + 3, 5, "Total:");
            setCellWithBordersRightTop(sheet, totalsRowOffset + 3, 6, total);

            // Set height for Amount in Words row to exactly two lines
            Row amountInWordsRow = sheet.getRow(totalsRowOffset + 1);
            if (amountInWordsRow == null) {
                amountInWordsRow = sheet.createRow(totalsRowOffset + 1);
            }
            amountInWordsRow.setHeightInPoints(18.7f); // Height for exactly two lines (45% reduction from 34f)
            setCellWithBordersCenterItalicBoldNoWrap(sheet, totalsRowOffset + 4, 1, "Amount in words: " + amountInWords(total, false));
        }

        // Auto-size columns with limits
        autoSizeColumnsWithLimit(sheet, isDollarCurrency);

        // Construct filename: Invoice_NUM_MST_InvoiceName
        String invoiceName = "";
        if (header != null && header.get("Invoice Name") != null) {
            invoiceName = header.get("Invoice Name").toString().trim();
            // Sanitize filename by replacing invalid characters
            invoiceName = invoiceName.replaceAll("[\\\\/:*?\"<>|]", "_");
        }

        String fileName = "Invoice_" + invoiceNo.replace("/", "_") + "_MST_" + invoiceName;
        String outFile = outputDir + fileName + ".xlsx";

        try {
            logger.debug("💾 Saving Excel file: {}", outFile);
            try (FileOutputStream fos = new FileOutputStream(outFile)) {
                wb.write(fos);
            }
            wb.close();
            logger.info("📄 Excel file generated successfully: {}", outFile);

            logger.debug("🖨️ Generating PDF from Excel: {}", outFile);
            PdfGeneratorLibreOffice.generatePdfFromExcel(outFile);
            String pdfFile = outFile.replace(".xlsm", ".pdf");
            logger.info("📋 PDF file generated successfully: {}", pdfFile);

        } catch (Exception ex) {
            logger.error("❌ Failed to generate files for invoice {}: {}", invoiceNo, ex.getMessage(), ex);
            throw ex;
        }
    }

    private static void autoSizeColumnsWithLimit(Sheet sheet, boolean isDollarCurrency) {
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
        int rateColumnWidth = isDollarCurrency ? 18 * 256 : 12 * 256; // Wider for template2
        if (sheet.getColumnWidth(5) < rateColumnWidth) {
            sheet.setColumnWidth(5, rateColumnWidth); // Wider minimum width for Rate column in template2
        }

        // Ensure Amount column (column 6) has adequate width for currency values
        if (sheet.getColumnWidth(6) < 15 * 256) {
            sheet.setColumnWidth(6, 18 * 256); // Minimum width for Amount column
        }
    }

    private static String amountInWords(double amount, boolean isDollar) {
        long wholeUnits = (long) amount;
        long fractionalUnits = Math.round((amount - wholeUnits) * 100);

        String currencyName = isDollar ? "Dollars" : "Rupees";
        String fractionalName = isDollar ? "Cents" : "Paise";

        return convert(wholeUnits) + " " + currencyName
                + (fractionalUnits > 0 ? " and " + convert(fractionalUnits) + " " + fractionalName : "");
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

        // Adjust row height if text is long
        adjustRowHeightForText(sheet, row, col, value);
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

        // Create cell style with borders and center alignment
        CellStyle style = wb.createCellStyle();
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true); // Enable text wrapping

        // If it's a number, add number formatting
        if (value instanceof Number) {
            style.setDataFormat(wb.createDataFormat().getFormat("#,##0.00"));
        }

        c.setCellStyle(style);

        // Adjust row height if text is long
        adjustRowHeightForText(sheet, row, col, value);

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
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
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

        // Adjust row height if text is long
        adjustRowHeightForText(sheet, row, col, value);

        return c;
    }

    private static Cell setCellWithBordersLeftAlign(
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

        // Create cell style with borders and left alignment
        CellStyle style = wb.createCellStyle();
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true); // Enable text wrapping

        // If it's a number, add number formatting
        if (value instanceof Number) {
            style.setDataFormat(wb.createDataFormat().getFormat("#,##0.00"));
        }

        c.setCellStyle(style);

        // Adjust row height if text is long
        adjustRowHeightForText(sheet, row, col, value);

        return c;
    }

    private static Cell setCellWithBordersLeftTop(
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

        // Create cell style with borders and left top alignment
        CellStyle style = wb.createCellStyle();
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.TOP);
        style.setWrapText(true); // Enable text wrapping

        // If it's a number, add number formatting
        if (value instanceof Number) {
            style.setDataFormat(wb.createDataFormat().getFormat("#,##0.00"));
        }

        c.setCellStyle(style);

        // Adjust row height if text is long
        adjustRowHeightForText(sheet, row, col, value);

        return c;
    }

    private static Cell setCellWithBordersRightTop(
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

        // Create cell style with borders and right top alignment
        CellStyle style = wb.createCellStyle();
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.RIGHT);
        style.setVerticalAlignment(VerticalAlignment.TOP);
        style.setWrapText(true); // Enable text wrapping

        // If it's a number, add number formatting
        if (value instanceof Number) {
            style.setDataFormat(wb.createDataFormat().getFormat("#,##0.00"));
        }

        c.setCellStyle(style);

        // Adjust row height if text is long
        adjustRowHeightForText(sheet, row, col, value);

        return c;
    }

    private static Cell setCellWithBordersCenterTop(
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

        // Create cell style with borders and center top alignment
        CellStyle style = wb.createCellStyle();
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.TOP);
        style.setWrapText(true); // Enable text wrapping

        // If it's a number, add number formatting
        if (value instanceof Number) {
            style.setDataFormat(wb.createDataFormat().getFormat("#,##0.00"));
        }

        c.setCellStyle(style);

        // Adjust row height if text is long
        adjustRowHeightForText(sheet, row, col, value);

        return c;
    }

    private static Cell setCellWithBordersCenterItalicBoldNoWrap(
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

        // Create cell style with borders, center alignment, italic and bold, no text wrapping
        CellStyle style = wb.createCellStyle();
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true); // Disable text wrapping for single line

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

        // Adjust row height if text is long to fit within 2 lines
        adjustRowHeightForText(sheet, row, col, value);

        return c;
    }

    private static void adjustRowHeightForText(Sheet sheet, int row, int col, Object value) {
        if (value == null || value instanceof Number) return;

        String text = value.toString();
        if (text.isEmpty()) return;

        Row r = sheet.getRow(row);
        if (r == null) return;

        // Get column width in characters (approximate)
        double columnWidthInChars = sheet.getColumnWidth(col) / 256.0;

        // Estimate characters per line (rough approximation)
        int charsPerLine = (int) Math.max(1, columnWidthInChars * 0.8);

        // Count approximate lines needed
        int linesNeeded = 1;
        if (text.length() > charsPerLine) {
            // Simple line counting based on spaces and length
            String[] words = text.split("\\s+");
            int currentLineLength = 0;

            for (String word : words) {
                if (currentLineLength + word.length() + 1 > charsPerLine) {
                    linesNeeded++;
                    currentLineLength = word.length();
                } else {
                    currentLineLength += word.length() + 1;
                }
            }
        }

        // Set minimum height per line (in points) - further reduced for compact display
        float baseHeight = 8f; // Further reduced base height
        float heightPerLine = 6f; // Further reduced additional height per line

        float newHeight = baseHeight + (heightPerLine * (linesNeeded - 1));
        newHeight = Math.max(newHeight, r.getHeightInPoints()); // Don't reduce height

        r.setHeightInPoints(newHeight);
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