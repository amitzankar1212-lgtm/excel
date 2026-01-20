package com.excelReader.excel;

import java.util.Map;

public class ExcelApplication {

	public static void main(String[] args) throws Exception {

		String headerFile =
				"C:/Users/amitz/Downloads/Invoices/clientDetails.xlsx";

		String itemsFile =
				"C:/Users/amitz/Downloads/Invoices/invoiceDetails.xlsx";

		String templatePath =
				"C:/Users/amitz/Downloads/Invoices/invoice_template.xlsx";

		String outputDir =
				"C:/Users/amitz/Downloads/Invoices/op/";

		Map<String, InvoiceData> invoices =
				SourceExcelReader.readInvoices(
						headerFile, itemsFile);

		for (Map.Entry<String, InvoiceData> e : invoices.entrySet()) {

			TemplateWriter.generateInvoice(
					e.getKey(),
					e.getValue().header,
					e.getValue().items,
					templatePath,
					outputDir
			);
		}

		System.out.println("✅ All invoices generated correctly");
	}


}
