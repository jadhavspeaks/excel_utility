package com.excelcomparator;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class ExcelUtil {

    public static class ExcelData {
        private final List<String> headers;
        private final List<List<Object>> data;

        public ExcelData(List<String> headers, List<List<Object>> data) {
            this.headers = headers;
            this.data = data;
        }

        public List<String> getHeaders() { return headers; }
        public List<List<Object>> getData() { return data; }
    }

    public static List<String> getSheetNames(File file) throws IOException {
        List<String> sheetNames = new ArrayList<>();
        try (Workbook workbook = WorkbookFactory.create(file)) {
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                sheetNames.add(workbook.getSheetName(i));
            }
        }
        return sheetNames;
    }

    public static ExcelData readExcel(File file, String sheetName, int headerRows, int numRows) throws IOException {
        List<String> headers = new ArrayList<>();
        List<List<Object>> data = new ArrayList<>();

        try (Workbook workbook = WorkbookFactory.create(file)) {
            Sheet sheet = (sheetName == null || sheetName.isEmpty()) ? workbook.getSheetAt(0) : workbook.getSheet(sheetName);
            if (sheet == null) throw new IOException("Sheet not found");

            Row headerRow = sheet.getRow(headerRows - 1);
            if (headerRow == null) throw new IOException("Header row not found");

            for (Cell cell : headerRow) {
                headers.add(getCellValueAsString(cell));
            }

            int rowsToRead = (numRows == -1) ? sheet.getLastRowNum() : Math.min(headerRows + numRows -1, sheet.getLastRowNum());
            for (int i = headerRows; i <= rowsToRead; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                List<Object> rowData = new ArrayList<>();
                for (int j = 0; j < headers.size(); j++) {
                    rowData.add(getCellValueAsObject(row.getCell(j)));
                }
                data.add(rowData);
            }
        }
        return new ExcelData(headers, data);
    }

    public static void writeFilteredDataToExcel(ExcelData excelData, File file) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Filtered Data");

            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < excelData.getHeaders().size(); i++) {
                headerRow.createCell(i).setCellValue(excelData.getHeaders().get(i));
            }

            int rowNum = 1;
            for (List<Object> dataRow : excelData.getData()) {
                Row row = sheet.createRow(rowNum++);
                for (int i = 0; i < dataRow.size(); i++) {
                    setCellValue(row.createCell(i), dataRow.get(i));
                }
            }

            for (int i = 0; i < excelData.getHeaders().size(); i++) {
                sheet.autoSizeColumn(i);
            }

            try (FileOutputStream fileOut = new FileOutputStream(file)) {
                workbook.write(fileOut);
            }
        }
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        DataFormatter formatter = new DataFormatter();
        return formatter.formatCellValue(cell);
    }

    private static Object getCellValueAsObject(Cell cell) {
        if (cell == null) return null;
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) return cell.getDateCellValue();
                return cell.getNumericCellValue();
            case BOOLEAN: return cell.getBooleanCellValue();
            case FORMULA: return cell.getCellFormula();
            default: return null;
        }
    }

    private static void setCellValue(Cell cell, Object value) {
        if (value instanceof String) cell.setCellValue((String) value);
        else if (value instanceof Number) cell.setCellValue(((Number) value).doubleValue());
        else if (value instanceof Date) cell.setCellValue((Date) value);
        else if (value instanceof Boolean) cell.setCellValue((Boolean) value);
        else if (value != null) cell.setCellValue(value.toString());
    }
}
