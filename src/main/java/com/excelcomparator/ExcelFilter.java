package com.excelcomparator;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.function.Consumer;

public class ExcelFilter {

    public void processFile(File inputFile, File outputFile, String sheetName, int keyColumnIndex, List<String> keyWords, Consumer<String> logger) throws Exception {
        try (InputStream inputStream = new FileInputStream(inputFile);
             Workbook workbook = new XSSFWorkbook(inputStream);
             SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(100)) {

            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                throw new IllegalArgumentException("Sheet not found: " + sheetName);
            }

            Sheet newSheet = sxssfWorkbook.createSheet(sheetName);
            List<String> headerRow = new ArrayList<>();
            Row header = sheet.getRow(0);
            if (header != null) {
                Row newHeader = newSheet.createRow(0);
                for (int i = 0; i < header.getLastCellNum(); i++) {
                    Cell cell = header.getCell(i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    String headerValue = (cell == null) ? "" : new DataFormatter().formatCellValue(cell);
                    headerRow.add(headerValue);
                    newHeader.createCell(i).setCellValue(headerValue);
                }
            } else {
                logger.accept("Warning: Header row is empty or not found.");
            }


            int outputRowCount = 1;
            int processedRowCount = 0;

            Iterator<Row> rowIterator = sheet.iterator();
            if (rowIterator.hasNext()) {
                rowIterator.next(); // Skip header
            }

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                processedRowCount++;

                Cell keyCell = row.getCell(keyColumnIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (keyCell != null) {
                    String keyValue = new DataFormatter().formatCellValue(keyCell);
                    if (keyWords.contains(keyValue)) {
                        Row newRow = newSheet.createRow(outputRowCount++);
                        for (int i = 0; i < headerRow.size(); i++) {
                            Cell oldCell = row.getCell(i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                            Cell newCell = newRow.createCell(i);
                            if (oldCell != null) {
                                switch (oldCell.getCellType()) {
                                    case STRING:
                                        newCell.setCellValue(oldCell.getStringCellValue());
                                        break;
                                    case NUMERIC:
                                        if (DateUtil.isCellDateFormatted(oldCell)) {
                                            newCell.setCellValue(oldCell.getDateCellValue());
                                        } else {
                                            newCell.setCellValue(oldCell.getNumericCellValue());
                                        }
                                        break;
                                    case BOOLEAN:
                                        newCell.setCellValue(oldCell.getBooleanCellValue());
                                        break;
                                    case FORMULA:
                                        newCell.setCellValue(oldCell.getCellFormula());
                                        break;
                                    case BLANK:
                                        // Do nothing
                                        break;
                                    default:
                                        newCell.setCellValue(new DataFormatter().formatCellValue(oldCell));
                                }
                            }
                        }
                    }
                }

                if (processedRowCount % 1000 == 0) {
                    logger.accept("Processed " + processedRowCount + " rows...");
                }
            }

            try (FileOutputStream out = new FileOutputStream(outputFile)) {
                sxssfWorkbook.write(out);
            }

            logger.accept("Filtering complete. Processed " + processedRowCount + " rows.");
            logger.accept("Output file created at: " + outputFile.getAbsolutePath());

        } finally {
            // sxssfWorkbook.dispose(); // Important for cleaning up temporary files
        }
    }
}
