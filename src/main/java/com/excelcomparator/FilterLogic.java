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
        Set<String> validRules = new HashSet<>();
        int attr1Idx = data1.getHeaders().indexOf(config.getFile1_AttributeColumn());
        if (attr1Idx == -1) return new ExcelUtil.ExcelData(data2.getHeaders(), new ArrayList<>(), new ArrayList<>());

        List<Integer> report1Indices = config.getFile1_ReportColumns().stream().map(data1.getHeaders()::indexOf).collect(Collectors.toList());
        List<Integer> product1Indices = config.getFile1_ProductColumns().stream().map(data1.getHeaders()::indexOf).collect(Collectors.toList());

        for (List<Object> row : data1.getData()) {
            String attribute = row.get(attr1Idx).toString();
            for (int rIdx : report1Indices) {
                if (rIdx != -1 && rIdx < row.size() && isTicked(row.get(rIdx))) {
                    validRules.add(attribute + "||" + data1.getHeaders().get(rIdx));
                }
            }
            for (int pIdx : product1Indices) {
                 if (pIdx != -1 && pIdx < row.size() && isTicked(row.get(pIdx))) {
                    validRules.add(attribute + "||" + data1.getHeaders().get(pIdx));
                }
            }
        }

        List<List<Object>> filteredData = new ArrayList<>();
        List<Integer> filteredRowNumbers = new ArrayList<>();
        int report2Idx = data2.getHeaders().indexOf(config.getFile2_ReportColumn());
        int product2Idx = data2.getHeaders().indexOf(config.getFile2_ProductColumn());
        List<Integer> attr2Indices = config.getFile2_AttributeColumns().stream().map(data2.getHeaders()::indexOf).collect(Collectors.toList());

        for (int i = 0; i < data2.getData().size(); i++) {
            List<Object> row = data2.getData().get(i);
            String report = (report2Idx != -1 && report2Idx < row.size()) ? row.get(report2Idx).toString() : "";
            String product = (product2Idx != -1 && product2Idx < row.size()) ? row.get(product2Idx).toString() : "";
            for (int attrIdx : attr2Indices) {
                if (attrIdx != -1 && attrIdx < row.size()) {
                    String attribute = row.get(attrIdx).toString();
                    if (validRules.contains(attribute + "||" + report) || validRules.contains(attribute + "||" + product)) {
                        filteredData.add(row);
                        filteredRowNumbers.add(data2.getOriginalRowNumbers().get(i));
                        break;
                    }
                }
            }
        }
        return new ExcelUtil.ExcelData(data2.getHeaders(), filteredData, filteredRowNumbers);
    }

    private static ExcelUtil.ExcelData applyFile2ToFile1Filter(ExcelUtil.ExcelData data1, ExcelUtil.ExcelData data2, FilterMappingConfig config) {
        Set<String> validCombos = new HashSet<>();
        int report2Idx = data2.getHeaders().indexOf(config.getFile2_ReportColumn());
        int product2Idx = data2.getHeaders().indexOf(config.getFile2_ProductColumn());
        List<Integer> attr2Indices = config.getFile2_AttributeColumns().stream().map(data2.getHeaders()::indexOf).collect(Collectors.toList());

        for (List<Object> row : data2.getData()) {
            String report = (report2Idx != -1 && report2Idx < row.size()) ? row.get(report2Idx).toString() : "";
            String product = (product2Idx != -1 && product2Idx < row.size()) ? row.get(product2Idx).toString() : "";
            for (int attrIdx : attr2Indices) {
                if (attrIdx != -1 && attrIdx < row.size()) {
                    String attribute = row.get(attrIdx).toString();
                    if (!report.isEmpty()) validCombos.add(attribute + "||" + report);
                    if (!product.isEmpty()) validCombos.add(attribute + "||" + product);
                }
            }
        }

        List<List<Object>> filteredData = new ArrayList<>();
        List<Integer> filteredRowNumbers = new ArrayList<>();
        int attr1Idx = data1.getHeaders().indexOf(config.getFile1_AttributeColumn());
        if (attr1Idx == -1) return new ExcelUtil.ExcelData(data1.getHeaders(), new ArrayList<>(), new ArrayList<>());

        for (int i = 0; i < data1.getData().size(); i++) {
            List<Object> row = data1.getData().get(i);
            String attribute = row.get(attr1Idx).toString();
            boolean keepRow = false;
            for (String reportHeader : config.getFile1_ReportColumns()) {
                int rIdx = data1.getHeaders().indexOf(reportHeader);
                if (rIdx != -1 && rIdx < row.size() && isTicked(row.get(rIdx)) && validCombos.contains(attribute + "||" + reportHeader)) {
                    keepRow = true;
                    break;
                }
            }
            if (keepRow) {
                filteredData.add(row);
                filteredRowNumbers.add(data1.getOriginalRowNumbers().get(i));
                continue;
            }
            for (String productHeader : config.getFile1_ProductColumns()) {
                int pIdx = data1.getHeaders().indexOf(productHeader);
                if (pIdx != -1 && pIdx < row.size() && isTicked(row.get(pIdx)) && validCombos.contains(attribute + "||" + productHeader)) {
                    keepRow = true;
                    break;
                }
            }
            if (keepRow) {
                filteredData.add(row);
                filteredRowNumbers.add(data1.getOriginalRowNumbers().get(i));
            }
        }
        return new ExcelUtil.ExcelData(data1.getHeaders(), filteredData, filteredRowNumbers);
    }

    private static boolean isTicked(Object obj) {
        if (obj == null) return false;
        String s = obj.toString().trim();
        return !s.isEmpty();
    }
}
