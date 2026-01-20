package com.excelReader.excel;

import java.util.List;
import java.util.Map;

public class InvoiceData {

    public final Map<String, Object> header;
    public final List<Map<String, Object>> items;

    public InvoiceData(
            Map<String, Object> header,
            List<Map<String, Object>> items) {

        this.header = header;
        this.items = items;
    }
}
