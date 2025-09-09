package com.excelcomparator;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.stream.Collectors;

public class ExcelUtil {

    public static class ExcelData {
        private final List<String> headers;
        private final List<List<Object>> data;
        private final List<Integer> originalRowNumbers;

        public ExcelData(List<String> headers, List<List<Object>> data, List<Integer> originalRowNumbers) {
            this.headers = headers;
            this.data = data;
            this.originalRowNumbers = originalRowNumbers;
        }

        public List<String> getHeaders() { return headers; }
        public List<List<Object>> getData() { return data; }
        public List<Integer> getOriginalRowNumbers() { return originalRowNumbers; }
    }

    public static List<String> getSheetNames(File file) throws IOException {
        List<String> sheetNames = new ArrayList<>();
        try (Workbook workbook = WorkbookFactory.create(file)) {
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                sheetNames.add(workbook.getSheetName(i));
            }
        } catch (Exception e) {
            throw new IOException("Error reading Excel file: " + e.getMessage(), e);
        }
        return sheetNames;
    }

    public static ExcelData readExcel(File file, String sheetName, int headerRows, int numRows) throws IOException {
        List<String> headers = new ArrayList<>();
        List<List<Object>> data = new ArrayList<>();
        List<Integer> originalRowNumbers = new ArrayList<>();

        try (Workbook workbook = WorkbookFactory.create(file)) {
            Sheet sheet = (sheetName == null || sheetName.isEmpty()) ? workbook.getSheetAt(0) : workbook.getSheet(sheetName);
            if (sheet == null) throw new IOException("Sheet '" + sheetName + "' not found.");

            int maxCols = 0;
            Map<Integer, StringBuilder> headerBuilders = new TreeMap<>();
            for (int i = 0; i < headerRows; i++) {
                Row headerRow = sheet.getRow(i);
                if (headerRow != null) maxCols = Math.max(maxCols, headerRow.getLastCellNum());
            }

            for (int i = 0; i < headerRows; i++) {
                Row headerRow = sheet.getRow(i);
                for (int j = 0; j < maxCols; j++) {
                    String cellValue = getCellValueFromMergedRegion(sheet, i, j).trim();
                    StringBuilder sb = headerBuilders.computeIfAbsent(j, k -> new StringBuilder());
                    if (!cellValue.isEmpty()) {
                        if (sb.length() > 0) sb.append(" | ");
                        sb.append(cellValue);
                    }
                }
            }
            for (int i = 0; i < maxCols; i++) {
                headers.add(headerBuilders.getOrDefault(i, new StringBuilder()).toString());
            }

            int rowsToRead = (numRows == -1) ? sheet.getLastRowNum() : Math.min(headerRows + numRows - 1, sheet.getLastRowNum());
            for (int i = headerRows; i <= rowsToRead; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                List<Object> rowData = new ArrayList<>();
                originalRowNumbers.add(row.getRowNum() + 1);
                for (int j = 0; j < headers.size(); j++) {
                    rowData.add(getCellValueAsObject(row.getCell(j)));
                }
                data.add(rowData);
            }
        } catch (Exception e) {
            throw new IOException("Error reading Excel file: " + e.getMessage(), e);
        }
        return new ExcelData(headers, data, originalRowNumbers);
    }

    public static ExcelData readAndAutoNormalize(File file, String sheetName, int headerRows) throws IOException {
        ExcelData rawData = readExcel(file, sheetName, headerRows, -1);
        List<String> originalHeaders = rawData.getHeaders();
        List<List<Object>> originalData = rawData.getData();
        List<Integer> originalRowNums = rawData.getOriginalRowNumbers();

        List<String> identifierColumns = new ArrayList<>();
        List<String> valueColumns = new ArrayList<>();

        try (Workbook workbook = WorkbookFactory.create(file)) {
            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) throw new IOException("Sheet '" + sheetName + "' not found.");

            boolean[] isValueColumn = new boolean[originalHeaders.size()];
            if (headerRows > 0) {
                for (CellRangeAddress region : sheet.getMergedRegions()) {
                    if (region.getLastRow() < headerRows && region.getFirstColumn() != region.getLastColumn()) {
                        for (int i = region.getFirstColumn(); i <= region.getLastColumn(); i++) {
                            if (i < isValueColumn.length) isValueColumn[i] = true;
                        }
                    }
                }
            }
            for (int i = 0; i < originalHeaders.size(); i++) {
                if (isValueColumn[i]) valueColumns.add(originalHeaders.get(i));
                else identifierColumns.add(originalHeaders.get(i));
            }
        }

        if (valueColumns.isEmpty()) return rawData;

        List<String> newHeaders = new ArrayList<>(identifierColumns);
        newHeaders.add("Dimension");
        newHeaders.add("Applicable");

        List<List<Object>> newData = new ArrayList<>();
        List<Integer> newRowNumbers = new ArrayList<>();
        List<Integer> identifierIndices = identifierColumns.stream().map(originalHeaders::indexOf).collect(Collectors.toList());
        List<Integer> valueIndices = valueColumns.stream().map(originalHeaders::indexOf).collect(Collectors.toList());

        for (int i = 0; i < originalData.size(); i++) {
            List<Object> row = originalData.get(i);
            int originalRowNum = originalRowNums.get(i);
            List<Object> identifierValues = new ArrayList<>();
            for (int index : identifierIndices) {
                identifierValues.add(index < row.size() ? row.get(index) : "");
            }
            for (int valueIndex : valueIndices) {
                List<Object> newRow = new ArrayList<>(identifierValues);
                newRow.add(originalHeaders.get(valueIndex));
                Object cellValue = valueIndex < row.size() ? row.get(valueIndex) : null;
                newRow.add(cellValue != null && !cellValue.toString().trim().isEmpty() ? "Yes" : "No");
                newData.add(newRow);
                newRowNumbers.add(originalRowNum);
            }
        }
        return new ExcelData(newHeaders, newData, newRowNumbers);
    }

    private static String getCellValueFromMergedRegion(Sheet sheet, int rowNum, int colNum) {
        for (CellRangeAddress region : sheet.getMergedRegions()) {
            if (region.isInRange(rowNum, colNum)) {
                return getCellValueAsString(sheet.getRow(region.getFirstRow()).getCell(region.getFirstColumn()));
            }
        }
        return getCellValueAsString(sheet.getRow(rowNum) != null ? sheet.getRow(rowNum).getCell(colNum) : null);
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

    public static void writeResultToExcel(ComparisonResult result, ExcelData data1, ExcelData data2, File file) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Comparison Result");
            int rowNum = 0;

            ComparisonSummary summary = result.getSummary();
            Font boldFont = workbook.createFont();
            boldFont.setBold(true);
            CellStyle boldStyle = workbook.createCellStyle();
            boldStyle.setFont(boldFont);

            Row r = sheet.createRow(rowNum++);
            r.createCell(0).setCellValue("Comparison Summary");
            r.getCell(0).setCellStyle(boldStyle);
            sheet.createRow(rowNum++).createCell(0).setCellValue("File 1 Columns: " + summary.getFile1ColumnCount());
            sheet.createRow(rowNum++).createCell(0).setCellValue("File 2 Columns: " + summary.getFile2ColumnCount());
            sheet.createRow(rowNum++).createCell(0).setCellValue("Common Columns: " + summary.getCommonColumns().size());

            if(!summary.getColumnsOnlyInFile1().isEmpty()){
                r = sheet.createRow(rowNum++);
                r.createCell(0).setCellValue("Columns only in File 1:");
                r.getCell(0).setCellStyle(boldStyle);
                for(String col : summary.getColumnsOnlyInFile1()) sheet.createRow(rowNum++).createCell(1).setCellValue(col);
            }
            if(!summary.getColumnsOnlyInFile2().isEmpty()){
                r = sheet.createRow(rowNum++);
                r.createCell(0).setCellValue("Columns only in File 2:");
                r.getCell(0).setCellStyle(boldStyle);
                for(String col : summary.getColumnsOnlyInFile2()) sheet.createRow(rowNum++).createCell(1).setCellValue(col);
            }
            rowNum++;

            Row headerRow = sheet.createRow(rowNum++);
            List<String> headers = new ArrayList<>();
            headers.add("File 1 Row");
            headers.add("File 2 Row");
            headers.add("Status");
            headers.add("Mismatched Columns");
            headers.addAll(result.getHeaders());
            for (int i = 0; i < headers.size(); i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers.get(i));
                cell.setCellStyle(boldStyle);
            }

            for (ComparisonResult.ResultRow resultRow : result.getResultRows()) {
                Row row = sheet.createRow(rowNum++);
                int cellIdx = 0;

                Cell r1Cell = row.createCell(cellIdx++);
                if (resultRow.getOriginalRowNum1() > 0) r1Cell.setCellValue(resultRow.getOriginalRowNum1());
                Cell r2Cell = row.createCell(cellIdx++);
                if (resultRow.getOriginalRowNum2() > 0) r2Cell.setCellValue(resultRow.getOriginalRowNum2());
                row.createCell(cellIdx++).setCellValue(resultRow.getStatus().toString());
                row.createCell(cellIdx++).setCellValue(String.join(", ", resultRow.getMismatchedColumns()));

                for (String header : result.getHeaders()) {
                    Object val1 = getCombinedValue(resultRow.getData1(), getAllIndices(data1.getHeaders(), header));
                    Object val2 = getCombinedValue(resultRow.getData2(), getAllIndices(data2.getHeaders(), header));
                    Cell cell = row.createCell(cellIdx++);
                    if (resultRow.getStatus() == ComparisonResult.RowStatus.MISMATCH && !Objects.equals(val1, val2)) {
                        setCellValue(cell, String.format("%s -> %s", val1, val2));
                    } else if (resultRow.getStatus() == ComparisonResult.RowStatus.MISSING_IN_FILE_1) {
                        setCellValue(cell, val2);
                    } else {
                        setCellValue(cell, val1);
                    }
                }
            }

            for (int i = 0; i < headers.size(); i++) sheet.autoSizeColumn(i);

            try (FileOutputStream fileOut = new FileOutputStream(file)) {
                workbook.write(fileOut);
            }
        }
    }

    public static List<Integer> getAllIndices(List<String> headers, String header) {
        List<Integer> indices = new ArrayList<>();
        if (headers == null) return indices;
        for (int i = 0; i < headers.size(); i++) {
            if (header.equals(headers.get(i))) indices.add(i);
        }
        return indices;
    }

    public static Object getCombinedValue(List<Object> data, List<Integer> indices) {
        if (indices.isEmpty() || data == null) return "";
        if (indices.size() == 1) {
            int idx = indices.get(0);
            return (idx < data.size()) ? data.get(idx) : "";
        }
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < indices.size(); i++) {
            int idx = indices.get(i);
            if (idx < data.size()) {
                Object val = data.get(idx);
                if (val != null) sb.append(val);
            }
            if (i < indices.size() - 1) sb.append(" | ");
        }
        return sb.toString();
    }

    private static void setCellValue(Cell cell, Object value) {
        if (value instanceof String) cell.setCellValue((String) value);
        else if (value instanceof Number) cell.setCellValue(((Number) value).doubleValue());
        else if (value instanceof Date) cell.setCellValue((Date) value);
        else if (value instanceof Boolean) cell.setCellValue((Boolean) value);
        else if (value != null) cell.setCellValue(value.toString());
    }
}
