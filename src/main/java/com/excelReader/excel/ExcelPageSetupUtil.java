package com.excelReader.excel;

import org.apache.poi.ss.usermodel.*;

import java.io.*;

public class ExcelPageSetupUtil {

    public static void applyA4NarrowLayout(String excelPath) throws Exception {

        FileInputStream fis = new FileInputStream(excelPath);
        Workbook workbook = WorkbookFactory.create(fis);
        fis.close();

        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {

            Sheet sheet = workbook.getSheetAt(i);
            PrintSetup ps = sheet.getPrintSetup();

            // A4 size
            ps.setPaperSize(PrintSetup.A4_PAPERSIZE);

            // Portrait mode (false = portrait, true = landscape)
            ps.setLandscape(false);

            // Fit content to page width
            sheet.setFitToPage(true);
            sheet.setAutobreaks(true);

            // Narrow margins (in inches)
            sheet.setMargin(Sheet.LeftMargin, 0.25);
            sheet.setMargin(Sheet.RightMargin, 0.25);
            sheet.setMargin(Sheet.TopMargin, 0.75);
            sheet.setMargin(Sheet.BottomMargin, 0.75);

            // Center horizontally (optional)
            sheet.setHorizontallyCenter(true);
        }

        FileOutputStream fos = new FileOutputStream(excelPath);
        workbook.write(fos);
        fos.close();
        workbook.close();
    }
}
