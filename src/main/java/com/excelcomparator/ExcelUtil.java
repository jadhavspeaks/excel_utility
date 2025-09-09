package com.excelcomparator;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Objects;
import java.util.TreeMap;
import java.util.Map;
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

        public List<String> getHeaders() {
            return headers;
        }

        public List<List<Object>> getData() {
            return data;
        }

        public List<Integer> getOriginalRowNumbers() {
            return originalRowNumbers;
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
        List<Integer> originalRowNumbers = new ArrayList<>();

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
                originalRowNumbers.add(row.getRowNum() + 1); // Capture row number
                for (int j = 0; j < headers.size(); j++) {
                    Cell cell = row.getCell(j, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    rowData.add(getCellValueAsObject(cell));
                }
                data.add(rowData);
            }
        } catch (Exception e) {
            throw new IOException("Error reading Excel file: " + e.getMessage(), e);
        }

        return new ExcelData(headers, data, originalRowNumbers);
    }

    private static class HeaderAnalysisResult {
        final List<String> identifierColumns = new ArrayList<>();
        final Map<String, List<String>> groupedColumns = new LinkedHashMap<>();
        final List<String> allOriginalHeaders;

        HeaderAnalysisResult(List<String> allOriginalHeaders) {
            this.allOriginalHeaders = allOriginalHeaders;
        }
    }

    private static HeaderAnalysisResult analyzeHeaders(Sheet sheet, int headerRows, int maxCols) {
        List<String> tempHeaders = new ArrayList<>();
         Map<Integer, StringBuilder> headerBuilders = new TreeMap<>();

        for (int i = 0; i < headerRows; i++) {
            Row headerRow = sheet.getRow(i);
            for (int j = 0; j < maxCols; j++) {
                Cell cell = headerRow != null ? headerRow.getCell(j) : null;
                String cellValue = getCellValueFromMergedRegion(sheet, i, j, sheet.getMergedRegions()).trim();
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
            tempHeaders.add(headerBuilders.getOrDefault(i, new StringBuilder()).toString());
        }

        HeaderAnalysisResult result = new HeaderAnalysisResult(tempHeaders);
        boolean[] isGrouped = new boolean[maxCols];

        if (headerRows > 1) {
            for (int r = 0; r < headerRows - 1; r++) {
                for (CellRangeAddress region : sheet.getMergedRegions()) {
                    if (region.getFirstRow() == r && region.getLastRow() == r && region.getFirstColumn() != region.getLastColumn()) {
                        String groupName = getCellValueAsString(sheet.getRow(r).getCell(region.getFirstColumn()));
                        List<String> children = new ArrayList<>();
                        for (int c = region.getFirstColumn(); c <= region.getLastColumn(); c++) {
                            children.add(result.allOriginalHeaders.get(c));
                            isGrouped[c] = true;
                        }
                        result.groupedColumns.put(groupName, children);
                    }
                }
            }
        }

        for (int i = 0; i < maxCols; i++) {
            if (!isGrouped[i]) {
                result.identifierColumns.add(result.allOriginalHeaders.get(i));
            }
        }
        return result;
    }

    public static ExcelData readAndAutoNormalize(File file, String sheetName, int headerRows) throws IOException {
        try (Workbook workbook = WorkbookFactory.create(file)) {
            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) throw new IOException("Sheet '" + sheetName + "' not found.");

            int maxCols = 0;
            for (int i = 0; i < headerRows; i++) {
                Row headerRow = sheet.getRow(i);
                if (headerRow != null) maxCols = Math.max(maxCols, headerRow.getLastCellNum());
            }

            HeaderAnalysisResult analysis = analyzeHeaders(sheet, headerRows, maxCols);

            if (analysis.groupedColumns.isEmpty()) {
                return readExcel(file, sheetName, headerRows, -1); // No groups, return raw data
            }

            List<String> newHeaders = new ArrayList<>(analysis.identifierColumns);
            newHeaders.add("Group Name");
            newHeaders.add("Child Dimension");

            List<List<Object>> newData = new ArrayList<>();
            List<Integer> newRowNumbers = new ArrayList<>();
            ExcelData rawExcelData = readExcel(file, sheetName, headerRows, -1);
            List<List<Object>> originalData = rawExcelData.getData();
            List<Integer> originalRowNums = rawExcelData.getOriginalRowNumbers();

            List<Integer> identifierIndices = analysis.identifierColumns.stream()
                .map(analysis.allOriginalHeaders::indexOf)
                .collect(Collectors.toList());

            for (int i = 0; i < originalData.size(); i++) {
                List<Object> row = originalData.get(i);
                int originalRowNum = originalRowNums.get(i);

                List<Object> identifierValues = new ArrayList<>();
                for (int index : identifierIndices) {
                    identifierValues.add(index < row.size() ? row.get(index) : "");
                }

                for (Map.Entry<String, List<String>> groupEntry : analysis.groupedColumns.entrySet()) {
                    String groupName = groupEntry.getKey();
                    for (String childHeader : groupEntry.getValue()) {
                        int childIndex = analysis.allOriginalHeaders.indexOf(childHeader);
                        Object cellValue = (childIndex != -1 && childIndex < row.size()) ? row.get(childIndex) : null;

                        String cellContent = (cellValue != null) ? cellValue.toString().trim() : "";
                        if (!cellContent.isEmpty()) {
                            List<Object> newRow = new ArrayList<>(identifierValues);
                            newRow.add(groupName);
                            newRow.add(childHeader);
                            newData.add(newRow);
                            newRowNumbers.add(originalRowNum);
                        }
                    }
                }
            }

            return new ExcelData(newHeaders, newData, newRowNumbers);
        } catch (Exception e) {
            throw new IOException("Failed to read and normalize Excel file: " + e.getMessage(), e);
        }
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
            int rowNum = 0;

            // --- Write Summary ---
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
                for(String col : summary.getColumnsOnlyInFile1()){
                     sheet.createRow(rowNum++).createCell(1).setCellValue(col);
                }
            }
            if(!summary.getColumnsOnlyInFile2().isEmpty()){
                r = sheet.createRow(rowNum++);
                r.createCell(0).setCellValue("Columns only in File 2:");
                r.getCell(0).setCellStyle(boldStyle);
                for(String col : summary.getColumnsOnlyInFile2()){
                     sheet.createRow(rowNum++).createCell(1).setCellValue(col);
                }
            }

            rowNum++; // Add a blank line after summary

            // --- Write Main Results Table ---
            CellStyle matchStyle = createStyle(workbook, IndexedColors.WHITE);
            CellStyle mismatchRowStyle = createStyle(workbook, IndexedColors.LIGHT_YELLOW);
            CellStyle missingStyle = createStyle(workbook, IndexedColors.ROSE);
            CellStyle mismatchCellStyle = createStyle(workbook, IndexedColors.GOLD);

            Row headerRow = sheet.createRow(rowNum++);
            List<String> headers = new ArrayList<>();
            headers.add("File 1 Row");
            headers.add("File 2 Row");
            headers.add("Status");
            headers.addAll(result.getHeaders());
            for (int i = 0; i < headers.size(); i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers.get(i));
                cell.setCellStyle(boldStyle);
            }

            List<String> headers1 = data1.getHeaders();
            List<String> headers2 = data2.getHeaders();

            for (ComparisonResult.ResultRow resultRow : result.getResultRows()) {
                Row row = sheet.createRow(rowNum++);
                int cellIdx = 0;

                // Row numbers
                Cell r1Cell = row.createCell(cellIdx++);
                if (resultRow.getOriginalRowNum1() > 0) r1Cell.setCellValue(resultRow.getOriginalRowNum1());
                Cell r2Cell = row.createCell(cellIdx++);
                if (resultRow.getOriginalRowNum2() > 0) r2Cell.setCellValue(resultRow.getOriginalRowNum2());

                // Status cell
                Cell statusCell = row.createCell(cellIdx++);
                statusCell.setCellValue(resultRow.getStatus().toString());

                // Data cells
                for (String header : result.getHeaders()) {
                    Cell cell = row.createCell(cellIdx++);
                    Object val1 = getCombinedValue(resultRow.getData1(), getAllIndices(headers1, header));
                    Object val2 = getCombinedValue(resultRow.getData2(), getAllIndices(headers2, header));

                    if (resultRow.getStatus() == ComparisonResult.RowStatus.MISMATCH && !Objects.equals(val1, val2)) {
                        setCellValue(cell, String.format("%s -> %s", val1, val2));
                        cell.setCellStyle(mismatchCellStyle);
                    } else if (resultRow.getStatus() == ComparisonResult.RowStatus.MISSING_IN_FILE_1) {
                         setCellValue(cell, val2);
                    } else {
                        setCellValue(cell, val1);
                    }
                }

                // Set row style
                CellStyle rowStyle = matchStyle;
                switch (resultRow.getStatus()) {
                    case MISMATCH: rowStyle = mismatchRowStyle; break;
                    case MISSING_IN_FILE_1: case MISSING_IN_FILE_2: rowStyle = missingStyle; break;
                }
                for(int i=0; i < cellIdx; i++){
                    if(row.getCell(i).getCellStyle() == null){
                         row.getCell(i).setCellStyle(rowStyle);
                    }
                }
            }

            for (int i = 0; i < headers.size(); i++) {
                sheet.autoSizeColumn(i);
            }

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
    public static List<Integer> getAllIndices(List<String> headers, String header) {
        List<Integer> indices = new ArrayList<>();
        if (headers == null) return indices;
        for (int i = 0; i < headers.size(); i++) {
            if (header.equals(headers.get(i))) {
                indices.add(i);
            }
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
                if (val != null) {
                    sb.append(val);
                }
            }
            if (i < indices.size() - 1) {
                sb.append(" | ");
            }
        }
        return sb.toString();
    }
}
