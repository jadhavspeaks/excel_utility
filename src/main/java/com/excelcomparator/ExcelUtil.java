package com.excelcomparator;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Objects;
import java.util.TreeMap;
import java.util.Map;
import java.util.stream.Collectors;

public class ExcelUtil {

    public static class ExcelData {
        private final List<String> headers;
        private final List<List<Object>> data;

        public ExcelData(List<String> headers, List<List<Object>> data) {
            this.headers = headers;
            this.data = data;
        }

        public List<String> getHeaders() {
            return headers;
        }

        public List<List<Object>> getData() {
            return data;
        }
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

        try (Workbook workbook = WorkbookFactory.create(file)) {
            Sheet sheet;
            if (sheetName == null || sheetName.isEmpty()) {
                sheet = workbook.getSheetAt(0);
            } else {
                sheet = workbook.getSheet(sheetName);
            }

            if (sheet == null) {
                throw new IOException("Sheet '" + sheetName + "' not found in the workbook.");
            }

            // Read headers, accounting for merged regions
            int maxCols = 0;
            List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
            Map<Integer, StringBuilder> headerBuilders = new TreeMap<>();

            for (int i = 0; i < headerRows; i++) {
                Row headerRow = sheet.getRow(i);
                if (headerRow == null) continue;
                maxCols = Math.max(maxCols, headerRow.getLastCellNum());
            }

            for (int i = 0; i < headerRows; i++) {
                Row headerRow = sheet.getRow(i);
                for (int j = 0; j < maxCols; j++) {
                    Cell cell = headerRow != null ? headerRow.getCell(j) : null;
                    String cellValue = getCellValueFromMergedRegion(sheet, i, j, mergedRegions).trim();
                    StringBuilder sb = headerBuilders.computeIfAbsent(j, k -> new StringBuilder());
                    if (!cellValue.isEmpty()) {
                        if (sb.length() > 0) {
                            sb.append(" | ");
                        }
                        sb.append(cellValue);
                    }
                }
            }

            for (int i = 0; i < maxCols; i++) {
                headers.add(headerBuilders.getOrDefault(i, new StringBuilder()).toString());
            }

            int rowsToRead = (numRows == -1) ? sheet.getLastRowNum() : Math.min(headerRows + numRows - 1, sheet.getLastRowNum());

            // Read data rows
            for (int i = headerRows; i <= rowsToRead; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                List<Object> rowData = new ArrayList<>();
                for (int j = 0; j < headers.size(); j++) {
                    Cell cell = row.getCell(j, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    rowData.add(getCellValueAsObject(cell));
                }
                data.add(rowData);
            }
        } catch (Exception e) {
            throw new IOException("Error reading Excel file: " + e.getMessage(), e);
        }

        return new ExcelData(headers, data);
    }

    public static ExcelData readAndAutoNormalize(File file, String sheetName, int headerRows) throws IOException {
        // First, get the raw data and headers
        ExcelData rawData = readExcel(file, sheetName, headerRows, -1);
        List<String> originalHeaders = rawData.getHeaders();
        List<List<Object>> originalData = rawData.getData();

        // --- Automated Header Analysis ---
        List<String> identifierColumns = new ArrayList<>();
        List<String> valueColumns = new ArrayList<>();

        try (Workbook workbook = WorkbookFactory.create(file)) {
            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                throw new IOException("Sheet '" + sheetName + "' not found.");
            }

            boolean[] isValueColumn = new boolean[originalHeaders.size()];

            if (headerRows > 0) {
                for (CellRangeAddress region : sheet.getMergedRegions()) {
                    if (region.getLastRow() < headerRows) {
                        if (region.getFirstColumn() != region.getLastColumn()) {
                            for (int i = region.getFirstColumn(); i <= region.getLastColumn(); i++) {
                                if (i < isValueColumn.length) {
                                    isValueColumn[i] = true;
                                }
                            }
                        }
                    }
                }
            }

            for (int i = 0; i < originalHeaders.size(); i++) {
                if (isValueColumn[i]) {
                    valueColumns.add(originalHeaders.get(i));
                } else {
                    identifierColumns.add(originalHeaders.get(i));
                }
            }
        } catch (Exception e) {
            throw new IOException("Failed during header analysis: " + e.getMessage(), e);
        }

        if (valueColumns.isEmpty()) {
            return rawData; // Nothing to normalize, return as-is.
        }

        // --- Automated Un-Pivoting Transformation ---
        List<String> newHeaders = new ArrayList<>(identifierColumns);
        newHeaders.add("Dimension");
        newHeaders.add("Applicable");

        List<List<Object>> newData = new ArrayList<>();

        // Get indices for faster lookup
        List<Integer> identifierIndices = identifierColumns.stream()
            .map(originalHeaders::indexOf).collect(Collectors.toList());
        List<Integer> valueIndices = valueColumns.stream()
            .map(originalHeaders::indexOf).collect(Collectors.toList());

        for (List<Object> row : originalData) {
            // Extract identifier values for the current row
            List<Object> identifierValues = new ArrayList<>();
            for (int index : identifierIndices) {
                identifierValues.add(index < row.size() ? row.get(index) : "");
            }

            // Un-pivot the value columns
            for (int valueIndex : valueIndices) {
                List<Object> newRow = new ArrayList<>(identifierValues);

                // Add Dimension (the header of the value column)
                newRow.add(originalHeaders.get(valueIndex));

                // Add Applicable (Yes/No based on cell content)
                Object cellValue = valueIndex < row.size() ? row.get(valueIndex) : null;
                String cellContent = (cellValue != null) ? cellValue.toString().trim() : "";

                if (!cellContent.isEmpty()) {
                    newRow.add("Yes");
                } else {
                    newRow.add("No");
                }

                newData.add(newRow);
            }
        }

        return new ExcelData(newHeaders, newData);
    }


    private static String getCellValueFromMergedRegion(Sheet sheet, int rowNum, int colNum, List<CellRangeAddress> mergedRegions) {
        for (CellRangeAddress region : mergedRegions) {
            if (region.isInRange(rowNum, colNum)) {
                Row firstRow = sheet.getRow(region.getFirstRow());
                if (firstRow != null) {
                    Cell firstCell = firstRow.getCell(region.getFirstColumn());
                    return getCellValueAsString(firstCell);
                }
                return "";
            }
        }
        Row row = sheet.getRow(rowNum);
        return (row != null) ? getCellValueAsString(row.getCell(colNum)) : "";
    }


    public static void writeResultToExcel(ComparisonResult result, ExcelData data1, ExcelData data2, File file) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Comparison Result");

            // Create cell styles
            CellStyle matchStyle = createStyle(workbook, IndexedColors.WHITE);
            CellStyle mismatchRowStyle = createStyle(workbook, IndexedColors.LIGHT_YELLOW);
            CellStyle missingStyle = createStyle(workbook, IndexedColors.ROSE);
            CellStyle mismatchCellStyle = createStyle(workbook, IndexedColors.GOLD);

            // Header
            Row headerRow = sheet.createRow(0);
            List<String> headers = new ArrayList<>();
            headers.add("Status");
            headers.addAll(result.getHeaders());
            for (int i = 0; i < headers.size(); i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers.get(i));
            }

            List<String> headers1 = data1.getHeaders();
            List<String> headers2 = data2.getHeaders();

            // Data rows
            int rowNum = 1;
            for (ComparisonResult.ResultRow resultRow : result.getResultRows()) {
                Row row = sheet.createRow(rowNum++);
                CellStyle rowStyle = matchStyle;
                switch (resultRow.getStatus()) {
                    case MISMATCH: rowStyle = mismatchRowStyle; break;
                    case MISSING_IN_FILE_1: case MISSING_IN_FILE_2: rowStyle = missingStyle; break;
                }

                // Status cell
                Cell statusCell = row.createCell(0);
                statusCell.setCellValue(resultRow.getStatus().toString());
                statusCell.setCellStyle(rowStyle);

                // Data cells
                for (int i = 0; i < result.getHeaders().size(); i++) {
                    Cell cell = row.createCell(i + 1);
                    String header = result.getHeaders().get(i);

                    int idx1 = headers1.indexOf(header);
                    int idx2 = headers2.indexOf(header);
                    Object value = null;
                    boolean isMismatch = false;

                    switch (resultRow.getStatus()) {
                         case MISSING_IN_FILE_1:
                            value = (idx2 != -1 && resultRow.getData2() != null && idx2 < resultRow.getData2().size()) ? resultRow.getData2().get(idx2) : "";
                            break;
                        case MISSING_IN_FILE_2:
                            value = (idx1 != -1 && resultRow.getData1() != null && idx1 < resultRow.getData1().size()) ? resultRow.getData1().get(idx1) : "";
                            break;
                        case MATCH:
                            value = (idx1 != -1 && resultRow.getData1() != null && idx1 < resultRow.getData1().size()) ? resultRow.getData1().get(idx1) : "";
                            break;
                        case MISMATCH:
                            Object val1 = (idx1 != -1 && resultRow.getData1() != null && idx1 < resultRow.getData1().size()) ? resultRow.getData1().get(idx1) : null;
                            Object val2 = (idx2 != -1 && resultRow.getData2() != null && idx2 < resultRow.getData2().size()) ? resultRow.getData2().get(idx2) : null;
                            if (!Objects.equals(val1, val2)) {
                                value = String.format("%s -> %s", val1, val2);
                                isMismatch = true;
                            } else {
                                value = val1;
                            }
                            break;
                    }

                    setCellValue(cell, value);
                    cell.setCellStyle(isMismatch ? mismatchCellStyle : rowStyle);
                }
            }

            // Autosize columns
            for (int i = 0; i < headers.size(); i++) {
                sheet.autoSizeColumn(i);
            }

            // Write to file
            try (FileOutputStream fileOut = new FileOutputStream(file)) {
                workbook.write(fileOut);
            }
        }
    }

    private static CellStyle createStyle(Workbook workbook, IndexedColors color) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(color.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    private static void setCellValue(Cell cell, Object value) {
        if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue());
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value != null) {
            cell.setCellValue(value.toString());
        }
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        DataFormatter formatter = new DataFormatter();
        return formatter.formatCellValue(cell);
    }

    private static Object getCellValueAsObject(Cell cell) {
        if (cell == null) {
            return null;
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                } else {
                    return cell.getNumericCellValue();
                }
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return "";
            default:
                return null;
        }
    }
}
