package com.excelcomparator;

import java.util.AbstractMap;
import java.util.Collections;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class ComparatorLogic {

    public static ComparisonResult compare(
            ExcelUtil.ExcelData data1,
            ExcelUtil.ExcelData data2,
            Map<String, String> keyColumnMap,
            boolean compareAllColumns,
            boolean findMissingRows) {

        List<ComparisonResult.ResultRow> resultRows = new ArrayList<>();
        List<String> headers1 = data1.getHeaders();
        List<String> headers2 = data2.getHeaders();
        List<Integer> rowNums1 = data1.getOriginalRowNumbers();
        List<Integer> rowNums2 = data2.getOriginalRowNumbers();

        Set<String> h1Set = new LinkedHashSet<>(headers1);
        Set<String> h2Set = new LinkedHashSet<>(headers2);
        List<String> onlyIn1 = h1Set.stream().filter(h -> !h2Set.contains(h)).collect(Collectors.toList());
        List<String> onlyIn2 = h2Set.stream().filter(h -> !h1Set.contains(h)).collect(Collectors.toList());
        Set<String> common = h1Set.stream().filter(h2Set::contains).collect(Collectors.toSet());
        ComparisonSummary summary = new ComparisonSummary(h1Set.size(), h2Set.size(), onlyIn1, onlyIn2, common);

        List<String> unifiedHeaders = new ArrayList<>(new LinkedHashSet<>(Stream.concat(headers1.stream(), headers2.stream()).collect(Collectors.toList())));

        Map<String, List<Map.Entry<Integer, List<Object>>>> mapOfData2 = new HashMap<>();
        if (findMissingRows) {
            for (int i = 0; i < data2.getData().size(); i++) {
                String key = buildKey(data2.getData().get(i), keyColumnMap.values(), headers2);
                mapOfData2.computeIfAbsent(key, k -> new ArrayList<>()).add(new AbstractMap.SimpleEntry<>(rowNums2.get(i), data2.getData().get(i)));
            }
        }

        for (int i = 0; i < data1.getData().size(); i++) {
            List<Object> row1 = data1.getData().get(i);
            int rowNum1 = rowNums1.get(i);
            String key1 = buildKey(row1, keyColumnMap.keySet(), headers1);
            List<Map.Entry<Integer, List<Object>>> matchingEntries = mapOfData2.get(key1);

            if (matchingEntries != null && !matchingEntries.isEmpty()) {
                Map.Entry<Integer, List<Object>> entry2 = matchingEntries.remove(0);
                if (matchingEntries.isEmpty()) mapOfData2.remove(key1);

                if (compareAllColumns) {
                    List<String> mismatchedCols = new ArrayList<>();
                    for (String header : unifiedHeaders) {
                        Object val1 = ExcelUtil.getCombinedValue(row1, ExcelUtil.getAllIndices(headers1, header));
                        Object val2 = ExcelUtil.getCombinedValue(entry2.getValue(), ExcelUtil.getAllIndices(headers2, header));
                        if (!Objects.equals(val1, val2)) {
                            mismatchedCols.add(header);
                        }
                    }
                    resultRows.add(new ComparisonResult.ResultRow(mismatchedCols.isEmpty() ? ComparisonResult.RowStatus.MATCH : ComparisonResult.RowStatus.MISMATCH, row1, entry2.getValue(), rowNum1, entry2.getKey(), mismatchedCols));
                } else {
                    resultRows.add(new ComparisonResult.ResultRow(ComparisonResult.RowStatus.MATCH, row1, entry2.getValue(), rowNum1, entry2.getKey(), Collections.emptyList()));
                }
            } else if (findMissingRows) {
                resultRows.add(new ComparisonResult.ResultRow(ComparisonResult.RowStatus.MISSING_IN_FILE_2, row1, null, rowNum1, 0, Collections.emptyList()));
            }
        }

        if (findMissingRows) {
            for (List<Map.Entry<Integer, List<Object>>> remainingEntries : mapOfData2.values()) {
                for (Map.Entry<Integer, List<Object>> entry : remainingEntries) {
                    resultRows.add(new ComparisonResult.ResultRow(ComparisonResult.RowStatus.MISSING_IN_FILE_1, null, entry.getValue(), 0, entry.getKey(), Collections.emptyList()));
                }
            }
        }
        return new ComparisonResult(resultRows, unifiedHeaders, summary);
    }

    private static String buildKey(List<Object> row, Collection<String> keyHeaders, List<String> allHeaders) {
        StringJoiner joiner = new StringJoiner("||");
        for (String keyHeader : keyHeaders) {
            int index = allHeaders.indexOf(keyHeader);
            if (index != -1 && index < row.size()) {
                Object value = row.get(index);
                joiner.add(value != null ? value.toString() : "");
            } else {
                joiner.add("");
            }
        }
        return joiner.toString();
    }
}
