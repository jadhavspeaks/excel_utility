package com.excelcomparator;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

public class StreamingExcelProcessor {

    private final List<String> headers = new ArrayList<>();
    private List<String> currentRow = new ArrayList<>();
    private SXSSFWorkbook workbook;
    private Sheet sheet;
    private int outputRowNum = 0;
    private long processedRowCount = 0;
    private long ticksDetected = 0;
    private long rowsWritten = 0;

    public String normalize(File inputFile, File outputFile) throws Exception {
        OPCPackage pkg = OPCPackage.open(inputFile);
        XSSFReader reader = new XSSFReader(pkg);
        SharedStrings sharedStrings = reader.getSharedStringsTable();
        XMLReader parser = XMLReaderFactory.createXMLReader();

        workbook = new SXSSFWorkbook(100);
        sheet = workbook.createSheet("Normalized Data");

        ContentHandler handler = new SheetHandler(sharedStrings);
        parser.setContentHandler(handler);

        try (InputStream sheetStream = reader.getSheetsData().next()) {
            InputSource sheetSource = new InputSource(sheetStream);
            parser.parse(sheetSource);
        }

        try (FileOutputStream out = new FileOutputStream(outputFile)) {
            workbook.write(out);
        }
        workbook.dispose();

        return String.format(
            "--- Normalization Complete ---\nRows Processed: %d\nNon-Empty Cells Found: %d\nOutput Rows Written: %d",
            processedRowCount, ticksDetected, rowsWritten
        );
    }

    private class SheetHandler extends DefaultHandler {
        private final SharedStrings sharedStrings;
        private String cellValue;
        private boolean isFirstRow = true;
        private boolean isSharedString;

        private SheetHandler(SharedStrings sst) {
            this.sharedStrings = sst;
        }

        public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
            if (name.equals("row")) {
                currentRow = new ArrayList<>();
            } else if (name.equals("c")) {
                cellValue = "";
                String cellType = attributes.getValue("t");
                isSharedString = "s".equals(cellType);
            }
        }

        public void endElement(String uri, String localName, String name) throws SAXException {
            if (name.equals("v")) {
                if (isSharedString) {
                    int idx = Integer.parseInt(cellValue);
                    cellValue = new XSSFRichTextString(sharedStrings.getItemAt(idx).getString()).toString();
                }
            } else if (name.equals("c")) {
                currentRow.add(cellValue);
            } else if (name.equals("row")) {
                if (isFirstRow) {
                    processedRowCount++;
                    headers.addAll(currentRow);
                    Row headerRow = sheet.createRow(outputRowNum++);
                    headerRow.createCell(0).setCellValue("Key");
                    headerRow.createCell(1).setCellValue("ColumnName");
                    headerRow.createCell(2).setCellValue("Value");
                    rowsWritten++;
                    isFirstRow = false;
                } else {
                    if (!currentRow.isEmpty()) {
                        processedRowCount++;
                        String key = currentRow.get(0);
                        for (int i = 1; i < currentRow.size(); i++) {
                            String value = currentRow.get(i);
                            if (value != null && !value.trim().isEmpty()) {
                                ticksDetected++;
                                Row outputRow = sheet.createRow(outputRowNum++);
                                outputRow.createCell(0).setCellValue(key);
                                if (i < headers.size()) {
                                    outputRow.createCell(1).setCellValue(headers.get(i));
                                }
                                outputRow.createCell(2).setCellValue(value);
                                rowsWritten++;
                            }
                        }
                    }
                }
            }
        }

        public void characters(char[] ch, int start, int length) throws SAXException {
            cellValue = new String(ch, start, length);
        }
    }
}
