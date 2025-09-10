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
import java.util.function.Consumer;

import org.xml.sax.helpers.XMLReaderFactory;

public class StreamingExcelProcessor {

    private final List<String> headers = new ArrayList<>();
    private List<String> currentRow = new ArrayList<>();
    private SXSSFWorkbook workbook;
    private Sheet sheet;
    private int outputRowNum = 0;
    private long processedRowCount = 0;
    private long ticksDetected = 0;
    private long rowsWritten = 0;
    private Consumer<String> logger;

    public void processFile(File inputFile, File outputFile, String sheetName, int headerRows, Consumer<String> logger) throws Exception {
        this.logger = logger;
        this.workbook = new SXSSFWorkbook(100);
        this.sheet = workbook.createSheet("Normalized Data");

        OPCPackage pkg = OPCPackage.open(inputFile);
        XSSFReader reader = new XSSFReader(pkg);
        SharedStrings sharedStrings = reader.getSharedStringsTable();
        XMLReader parser = XMLReaderFactory.createXMLReader();

        ContentHandler handler = new SheetHandler(sharedStrings, headerRows);
        parser.setContentHandler(handler);

        boolean sheetFound = false;
        XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) reader.getSheetsData();
        while (sheets.hasNext()) {
            try (InputStream sheetStream = sheets.next()) {
                if (sheets.getSheetName().equals(sheetName)) {
                    sheetFound = true;
                    log("Processing sheet: " + sheetName);
                    InputSource sheetSource = new InputSource(sheetStream);
                    parser.parse(sheetSource);
                    break;
                }
            }
        }

        if (!sheetFound) {
            workbook.close();
            pkg.close();
            throw new IllegalArgumentException("Sheet '" + sheetName + "' not found in the workbook.");
        }

        log(String.format("Processed %d rows, found %d non-empty cells, wrote %d rows to output.",
                processedRowCount, ticksDetected, rowsWritten));

        try (FileOutputStream out = new FileOutputStream(outputFile)) {
            workbook.write(out);
        }
        workbook.dispose();
        pkg.close();
    }

    private void log(String message) {
        if (logger != null) {
            logger.accept(message);
        }
    }

    private class SheetHandler extends DefaultHandler {
        private final SharedStrings sharedStrings;
        private final int headerRows;
        private int currentRowNumber = 0;
        private String cellValue;
        private boolean isSharedString;
        private java.util.Map<Integer, StringBuilder> headerBuilders = new java.util.TreeMap<>();

        private SheetHandler(SharedStrings sst, int headerRows) {
            this.sharedStrings = sst;
            this.headerRows = headerRows;
        }

        public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
            if (name.equals("row")) {
                currentRow = new ArrayList<>();
                currentRowNumber++;
            } else if (name.equals("c")) {
                cellValue = "";
                String cellType = attributes.getValue("t");
                isSharedString = "s".equals(cellType);
            }
        }

        public void endElement(String uri, String localName, String name) throws SAXException {
            if (name.equals("v")) {
                if (isSharedString) {
                    try {
                        int idx = Integer.parseInt(cellValue);
                        cellValue = new XSSFRichTextString(sharedStrings.getItemAt(idx).getString()).toString();
                    } catch (NumberFormatException e) {
                        // Handle cases where the value is not a valid integer index
                    }
                }
            } else if (name.equals("c")) {
                currentRow.add(cellValue);
            } else if (name.equals("row")) {
                processedRowCount++;
                if (currentRowNumber <= headerRows) {
                    // This is a header row
                    for (int i = 0; i < currentRow.size(); i++) {
                        StringBuilder sb = headerBuilders.computeIfAbsent(i, k -> new StringBuilder());
                        if (sb.length() > 0) sb.append(" | ");
                        sb.append(currentRow.get(i));
                    }
                    if (currentRowNumber == headerRows) {
                        // Last header row, finalize headers
                        int maxCols = headerBuilders.keySet().stream().max(Integer::compareTo).orElse(-1) + 1;
                        for (int i = 0; i < maxCols; i++) {
                            headers.add(headerBuilders.getOrDefault(i, new StringBuilder()).toString());
                        }
                        Row headerRow = sheet.createRow(outputRowNum++);
                        headerRow.createCell(0).setCellValue("Key");
                        headerRow.createCell(1).setCellValue("ColumnName");
                        headerRow.createCell(2).setCellValue("Value");
                        rowsWritten++;
                    }
                } else {
                    // This is a data row
                    if (!currentRow.isEmpty()) {
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
