package com.excelReader.excel;

import java.util.Map;

public class ExcelApplication {

	public static void main(String[] args) throws Exception {

		String headerFile =
				"C:/Users/amitz/Downloads/Invoices/clientDetails.xlsx";
//				"C:/Users/amitz/Downloads/Invoices/temp/temp2.xlsx";

		String itemsFile =
				"C:/Users/amitz/Downloads/Invoices/invoiceDetails.xlsx";
//				"C:/Users/amitz/Downloads/Invoices/temp/temp1.xlsx";

		String templatePath =
				"C:/Users/amitz/Downloads/Invoices/invoice_template.xlsx";
		String templatePath2 =
				"C:/Users/amitz/Downloads/Invoices/invoice_template_2.xlsx";

		String outputDir =
				"C:/Users/amitz/Downloads/Invoices/op/";

		Map<String, InvoiceData> invoices =
				SourceExcelReader.readInvoices(
						headerFile, itemsFile);

		for (Map.Entry<String, InvoiceData> e : invoices.entrySet()) {

			// Check currency to determine which template to use
			String currency = "";
			String selectedTemplatePath = templatePath; // Default template

			if (e.getValue().header != null) {
				Object currencyObj = e.getValue().header.get("Currency");
				if (currencyObj != null) {
					currency = currencyObj.toString().toLowerCase();
					if ("dollar".equals(currency)) {
						selectedTemplatePath = templatePath2; // Use template 2 for dollar currency
					}
				}
			}

			TemplateWriter.generateInvoice(
					e.getKey(),
					e.getValue().header,
					e.getValue().items,
					selectedTemplatePath,
					outputDir
			);
		}

		System.out.println("✅ All invoices generated correctly");
	}


}
