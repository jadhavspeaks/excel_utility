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

    public void processFile(File inputFile, File outputFile, String sheetName, int headerRows, List<FilterCondition> conditions, Consumer<String> logger) throws Exception {
        try (InputStream inputStream = new FileInputStream(inputFile);
             Workbook workbook = new XSSFWorkbook(inputStream);
             SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(100)) {

            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                throw new IllegalArgumentException("Sheet not found: " + sheetName);
            }

            Sheet newSheet = sxssfWorkbook.createSheet(sheetName);
            List<String> headers = ExcelUtil.getHeaders(inputFile, sheetName, headerRows);
            Row newHeader = newSheet.createRow(0);
            for (int i = 0; i < headers.size(); i++) {
                newHeader.createCell(i).setCellValue(headers.get(i));
            }

            int outputRowCount = 1;
            int processedRowCount = 0;

            Iterator<Row> rowIterator = sheet.iterator();
            for (int i = 0; i < headerRows && rowIterator.hasNext(); i++) {
                rowIterator.next();
            }

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                processedRowCount++;

                boolean allConditionsMet = true;
                if (conditions.isEmpty()) {
                    allConditionsMet = false; // Don't include any rows if there are no filters
                }

                for (FilterCondition condition : conditions) {
                    int columnIndex = headers.indexOf(condition.getColumnName());
                    if (columnIndex == -1) {
                        allConditionsMet = false;
                        logger.accept("Warning: Column '" + condition.getColumnName() + "' not found. Skipping row.");
                        break;
                    }
                    Cell cell = row.getCell(columnIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    String cellValue = (cell == null) ? "" : new DataFormatter().formatCellValue(cell);
                    if (!condition.matches(cellValue)) {
                        allConditionsMet = false;
                        break;
                    }
                }

                if (allConditionsMet) {
                    Row newRow = newSheet.createRow(outputRowCount++);
                    for (int i = 0; i < headers.size(); i++) {
                        Cell oldCell = row.getCell(i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                        Cell newCell = newRow.createCell(i);
                        if (oldCell != null) {
                             switch (oldCell.getCellType()) {
                                case STRING: newCell.setCellValue(oldCell.getStringCellValue()); break;
                                case NUMERIC:
                                    if (DateUtil.isCellDateFormatted(oldCell)) {
                                        newCell.setCellValue(oldCell.getDateCellValue());
                                    } else {
                                        newCell.setCellValue(oldCell.getNumericCellValue());
                                    }
                                    break;
                                case BOOLEAN: newCell.setCellValue(oldCell.getBooleanCellValue()); break;
                                case FORMULA: newCell.setCellValue(oldCell.getCellFormula()); break;
                                case BLANK: break;
                                default: newCell.setCellValue(new DataFormatter().formatCellValue(oldCell));
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
