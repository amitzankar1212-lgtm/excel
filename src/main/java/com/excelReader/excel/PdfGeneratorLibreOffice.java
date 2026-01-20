package com.excelReader.excel;

import java.io.File;

public class PdfGeneratorLibreOffice {

    private static final String LIBRE_OFFICE_PATH =
            "C:\\Program Files\\LibreOffice\\program\\soffice.exe";

    public static void generatePdfFromExcel(String excelPath)
            throws Exception {

        File excelFile = new File(excelPath);
        if (!excelFile.exists()) {
            throw new RuntimeException(
                    "Excel file not found: " + excelPath);
        }

        File outputDir = excelFile.getParentFile();

        ProcessBuilder pb = new ProcessBuilder(
                LIBRE_OFFICE_PATH,
                "--headless",
                "--nologo",
                "--nolockcheck",
                "--nodefault",
                "--convert-to", "pdf:calc_pdf_Export",
                excelFile.getAbsolutePath(),
                "--outdir", outputDir.getAbsolutePath()
        );

        pb.redirectErrorStream(true);
        Process process = pb.start();

        int exitCode = process.waitFor();
        if (exitCode != 0) {
            throw new RuntimeException(
                    "Excel to PDF conversion failed");
        }
    }
}
