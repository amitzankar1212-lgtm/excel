package com.excelReader.excel;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.util.Map;

@SpringBootApplication
public class ExcelApplication implements CommandLineRunner {

	private static final Logger logger = LoggerFactory.getLogger(ExcelApplication.class);

	@Value("${invoice.header.file}")
	private String headerFile;

	@Value("${invoice.items.file}")
	private String itemsFile;

	@Value("${invoice.template.path}")
	private String templatePath;

	@Value("${invoice.template2.path}")
	private String templatePath2;

	@Value("${invoice.output.directory}")
	private String outputDir;

	public static void main(String[] args) {
		SpringApplication.run(ExcelApplication.class, args);
	}

	@Override
	public void run(String... args) throws Exception {

		logger.info("🚀 Starting Invoice Generation Application");
		logger.warn("⚠️  IMPORTANT: Make sure invoiceDetails.xlsx is NOT open in Excel before running this application");

		try {
			// Resolve relative paths to absolute paths based on current working directory
			String resolvedHeaderFile = resolvePath(headerFile);
			String resolvedItemsFile = resolvePath(itemsFile);
			String resolvedTemplatePath = resolvePath(templatePath);
			String resolvedTemplatePath2 = resolvePath(templatePath2);
			String resolvedOutputDir = resolvePath(outputDir);

			logger.info("📂 Loading invoice data from: {}", resolvedHeaderFile);
			logger.info("📂 Loading invoice items from: {}", resolvedItemsFile);

			Map<String, InvoiceData> invoices =
					SourceExcelReader.readInvoices(
							resolvedHeaderFile, resolvedItemsFile);

			logger.info("📊 Found {} invoices to process", invoices.size());

			// Force garbage collection and wait for file handles to be released
			System.gc();
			Thread.sleep(2000); // Increased delay

			for (Map.Entry<String, InvoiceData> e : invoices.entrySet()) {

				logger.info("🔄 Processing invoice: {}", e.getKey());

				// Update status to "Inprocess" with retry mechanism
				updateInvoiceStatusWithRetry(resolvedItemsFile, e.getKey(), "INPROGRESS", logger);

				// Check currency to determine which template to use
				String currency = "";
				String selectedTemplatePath = resolvedTemplatePath; // Default template

				if (e.getValue().header != null) {
					Object currencyObj = e.getValue().header.get("Currency");
					if (currencyObj != null) {
						currency = currencyObj.toString().toLowerCase();
						if ("dollar".equals(currency)) {
							selectedTemplatePath = resolvedTemplatePath2; // Use template 2 for dollar currency
							logger.debug("💵 Using dollar currency template for invoice: {}", e.getKey());
						} else {
							logger.debug("💰 Using regular currency template for invoice: {}", e.getKey());
						}
					}
				}

				try {
					TemplateWriter.generateInvoice(
							e.getKey(),
							e.getValue().header,
							e.getValue().items,
							selectedTemplatePath,
							resolvedOutputDir
					);
					logger.info("✅ Successfully generated invoice: {}", e.getKey());

					// Update status to "completed" with retry mechanism
					updateInvoiceStatusWithRetry(resolvedItemsFile, e.getKey(), "COMPLETED", logger);

				} catch (Exception ex) {
					logger.error("❌ Failed to generate invoice: {} - Error: {}", e.getKey(), ex.getMessage(), ex);
					throw ex; // Re-throw to stop processing if one fails
				}
			}

			logger.info("🎉 All {} invoices generated successfully!", invoices.size());
			logger.info("✅ Invoice generation process completed successfully. Exiting...");
			System.exit(0);

		} catch (Exception ex) {
			logger.error("💥 Fatal error during invoice generation: {}", ex.getMessage(), ex);
			logger.error("❌ Invoice generation process failed. Exiting with error code 1...");
			System.exit(1);
		}
	}

	/**
	 * Update invoice status with retry mechanism for file locking issues
	 */
	private void updateInvoiceStatusWithRetry(String itemsFile, String invoiceNo, String status, Logger logger) {
		int maxRetries = 5; // Increased retries
		int retryDelay = 2000; // Increased base delay to 2 seconds

		for (int attempt = 1; attempt <= maxRetries; attempt++) {
			try {
				// Check if file is accessible before attempting update
				if (!isFileAccessible(itemsFile)) {
					throw new Exception("File is not accessible (might be open in Excel)");
				}

				SourceExcelReader.updateInvoiceStatus(itemsFile, invoiceNo, status);
				logger.debug("📝 Updated status to '{}' for invoice: {} (attempt {})", status, invoiceNo, attempt);
				return; // Success, exit retry loop
			} catch (Exception ex) {
				if (attempt == maxRetries) {
					logger.error("❌ CRITICAL: Could not update status to '{}' for invoice {} after {} attempts", status, invoiceNo, maxRetries);
					logger.error("❌ Last error: {}", ex.getMessage());
					logger.error("💡 SOLUTION: 1) Close Excel completely, 2) Check Task Manager for Excel processes, 3) Restart if needed");
					logger.error("💡 The invoice was generated successfully, but status tracking failed");
				} else {
					logger.warn("Retrying status update for invoice {} (attempt {}/{}): {}", invoiceNo, attempt, maxRetries, ex.getMessage());
					try {
						// Force garbage collection before retry
						System.gc();
						Thread.sleep(retryDelay * attempt); // Exponential backoff
					} catch (InterruptedException ie) {
						Thread.currentThread().interrupt();
						break;
					}
				}
			}
		}
	}

	/**
	 * Check if file is accessible for writing
	 */
	private boolean isFileAccessible(String filePath) {
		try {
			java.io.File file = new java.io.File(filePath);
			if (!file.exists()) return false;

			// Try to open file for reading to check accessibility
			try (java.io.FileInputStream fis = new java.io.FileInputStream(file)) {
				return true;
			}
		} catch (Exception e) {
			return false;
		}
	}

	/**
	 * Resolve relative paths to absolute paths based on the current working directory
	 */
	private String resolvePath(String path) {
		if (path.startsWith("/")) {
			// Absolute Unix-style path
			return path;
		} else if (path.contains(":")) {
			// Already an absolute Windows path (contains drive letter)
			return path;
		} else {
			// Relative path - resolve relative to current working directory
			return System.getProperty("user.dir") + "/" + path;
		}
	}
}
