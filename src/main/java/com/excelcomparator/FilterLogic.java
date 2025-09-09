package com.excelcomparator;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.stream.Collectors;

public class FilterLogic {

    public static ExcelUtil.ExcelData filter(ExcelUtil.ExcelData data1, ExcelUtil.ExcelData data2, FilterMappingConfig config) {
        if (config.getDirection() == FilterMappingConfig.FilterDirection.FILE1_FILTERS_FILE2) {
            return applyFile1ToFile2Filter(data1, data2, config);
        } else {
            return applyFile2ToFile1Filter(data1, data2, config);
        }
    }

    private static ExcelUtil.ExcelData applyFile1ToFile2Filter(ExcelUtil.ExcelData data1, ExcelUtil.ExcelData data2, FilterMappingConfig config) {
        // Step 1: Build a set of all valid rules from File 1
        Set<String> validAttributeReportRules = new HashSet<>();
        Set<String> validAttributeProductRules = new HashSet<>();

        int attr1Idx = data1.getHeaders().indexOf(config.getFile1_AttributeColumn());
        if (attr1Idx == -1) return new ExcelUtil.ExcelData(data2.getHeaders(), new ArrayList<>());

        List<Integer> report1Indices = config.getFile1_ReportColumns().stream().map(h -> data1.getHeaders().indexOf(h)).collect(Collectors.toList());
        List<Integer> product1Indices = config.getFile1_ProductColumns().stream().map(h -> data1.getHeaders().indexOf(h)).collect(Collectors.toList());

        for (List<Object> row : data1.getData()) {
            String attribute = row.get(attr1Idx).toString();
            for (int rIdx : report1Indices) {
                if (rIdx != -1 && rIdx < row.size() && isTicked(row.get(rIdx))) {
                    String reportHeader = data1.getHeaders().get(rIdx);
                    validAttributeReportRules.add(attribute + "||" + reportHeader);
                }
            }
            for (int pIdx : product1Indices) {
                 if (pIdx != -1 && pIdx < row.size() && isTicked(row.get(pIdx))) {
                    String productHeader = data1.getHeaders().get(pIdx);
                    validAttributeProductRules.add(attribute + "||" + productHeader);
                }
            }
        }

        // Step 2: Filter File 2 based on the rules
        List<List<Object>> filteredData = new ArrayList<>();
        int report2Idx = data2.getHeaders().indexOf(config.getFile2_ReportColumn());
        int product2Idx = data2.getHeaders().indexOf(config.getFile2_ProductColumn());
        List<Integer> attr2Indices = config.getFile2_AttributeColumns().stream().map(h -> data2.getHeaders().indexOf(h)).collect(Collectors.toList());

        for (int i = 0; i < data2.getData().size(); i++) {
            List<Object> row = data2.getData().get(i);
            String report = (report2Idx != -1 && report2Idx < row.size()) ? row.get(report2Idx).toString() : "";
            String product = (product2Idx != -1 && product2Idx < row.size()) ? row.get(product2Idx).toString() : "";

            for (int attrIdx : attr2Indices) {
                if (attrIdx != -1 && attrIdx < row.size()) {
                    String attribute = row.get(attrIdx).toString();
                    if (validAttributeReportRules.contains(attribute + "||" + report) || validAttributeProductRules.contains(attribute + "||" + product)) {
                        filteredData.add(row);
                        break;
                    }
                }
            }
        }
        return new ExcelUtil.ExcelData(data2.getHeaders(), filteredData);
    }

    private static ExcelUtil.ExcelData applyFile2ToFile1Filter(ExcelUtil.ExcelData data1, ExcelUtil.ExcelData data2, FilterMappingConfig config) {
        // Step 1: Extract distinct combinations from File 2
        Set<String> validReportAttributes = new HashSet<>();
        Set<String> validProductAttributes = new HashSet<>();
        int report2Idx = data2.getHeaders().indexOf(config.getFile2_ReportColumn());
        int product2Idx = data2.getHeaders().indexOf(config.getFile2_ProductColumn());
        List<Integer> attr2Indices = config.getFile2_AttributeColumns().stream().map(h -> data2.getHeaders().indexOf(h)).collect(Collectors.toList());

        for (List<Object> row : data2.getData()) {
            String report = (report2Idx != -1 && report2Idx < row.size()) ? row.get(report2Idx).toString() : "";
            String product = (product2Idx != -1 && product2Idx < row.size()) ? row.get(product2Idx).toString() : "";
            for (int attrIdx : attr2Indices) {
                if (attrIdx != -1 && attrIdx < row.size()) {
                    String attribute = row.get(attrIdx).toString();
                    if (!report.isEmpty()) validReportAttributes.add(attribute + "||" + report);
                    if (!product.isEmpty()) validProductAttributes.add(attribute + "||" + product);
                }
            }
        }

        // Step 2: Filter File 1 based on these combinations
        List<List<Object>> filteredData = new ArrayList<>();
        int attr1Idx = data1.getHeaders().indexOf(config.getFile1_AttributeColumn());
        if (attr1Idx == -1) return new ExcelUtil.ExcelData(data1.getHeaders(), new ArrayList<>());

        for (List<Object> row : data1.getData()) {
            String attribute = row.get(attr1Idx).toString();
            boolean keepRow = false;
            for (String reportHeader : config.getFile1_ReportColumns()) {
                int rIdx = data1.getHeaders().indexOf(reportHeader);
                if (rIdx != -1 && rIdx < row.size() && isTicked(row.get(rIdx))) {
                    if (validReportAttributes.contains(attribute + "||" + reportHeader)) {
                        keepRow = true;
                        break;
                    }
                }
            }
            if (keepRow) {
                filteredData.add(row);
                continue;
            }
            for (String productHeader : config.getFile1_ProductColumns()) {
                int pIdx = data1.getHeaders().indexOf(productHeader);
                if (pIdx != -1 && pIdx < row.size() && isTicked(row.get(pIdx))) {
                    if (validProductAttributes.contains(attribute + "||" + productHeader)) {
                        keepRow = true;
                        break;
                    }
                }
            }
            if (keepRow) {
                filteredData.add(row);
            }
        }
        return new ExcelUtil.ExcelData(data1.getHeaders(), filteredData);
    }

    private static boolean isTicked(Object obj) {
        if (obj == null) return false;
        String s = obj.toString().trim();
        return !s.isEmpty();
    }
}
