package com.excelcomparator;

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

        // Create a unified header list
        List<String> unifiedHeaders = Stream.concat(headers1.stream(), headers2.stream())
                .distinct().collect(Collectors.toList());

        Map<String, Integer> keyColumnIndexes1 = keyColumnMap.keySet().stream()
                .collect(Collectors.toMap(h -> h, headers1::indexOf));
        Map<String, Integer> keyColumnIndexes2 = keyColumnMap.values().stream()
                .collect(Collectors.toMap(h -> h, headers2::indexOf));

        Map<String, List<List<Object>>> mapOfData2 = new HashMap<>();
        if (findMissingRows) {
            for (List<Object> row : data2.getData()) {
                String key = buildKey(row, keyColumnMap.values(), headers2);
                mapOfData2.computeIfAbsent(key, k -> new ArrayList<>()).add(row);
            }
        }

        for (List<Object> row1 : data1.getData()) {
            String key1 = buildKey(row1, keyColumnMap.keySet(), headers1);
            List<List<Object>> matchingRows2 = mapOfData2.get(key1);

            if (matchingRows2 != null && !matchingRows2.isEmpty()) {
                List<Object> row2 = matchingRows2.remove(0); // Take the first match
                if (matchingRows2.isEmpty()) {
                    mapOfData2.remove(key1);
                }

                if (compareAllColumns) {
                    List<Boolean> mismatches = new ArrayList<>();
                    boolean hasMismatch = false;
                    for (String header : unifiedHeaders) {
                        int idx1 = headers1.indexOf(header);
                        int idx2 = headers2.indexOf(header);
                        Object val1 = (idx1 != -1 && idx1 < row1.size()) ? row1.get(idx1) : null;
                        Object val2 = (idx2 != -1 && idx2 < row2.size()) ? row2.get(idx2) : null;
                        if (!Objects.equals(val1, val2)) {
                            mismatches.add(true);
                            hasMismatch = true;
                        } else {
                            mismatches.add(false);
                        }
                    }
                    if (hasMismatch) {
                        resultRows.add(new ComparisonResult.ResultRow(ComparisonResult.RowStatus.MISMATCH, row1, row2, mismatches));
                    } else {
                        resultRows.add(new ComparisonResult.ResultRow(ComparisonResult.RowStatus.MATCH, row1, row2, Collections.nCopies(unifiedHeaders.size(), false)));
                    }
                } else {
                    resultRows.add(new ComparisonResult.ResultRow(ComparisonResult.RowStatus.MATCH, row1, row2, Collections.nCopies(unifiedHeaders.size(), false)));
                }

            } else if (findMissingRows) {
                resultRows.add(new ComparisonResult.ResultRow(ComparisonResult.RowStatus.MISSING_IN_FILE_2, row1, null, null));
            }
        }

        if (findMissingRows) {
            for (List<List<Object>> remainingRows : mapOfData2.values()) {
                for (List<Object> row2 : remainingRows) {
                    resultRows.add(new ComparisonResult.ResultRow(ComparisonResult.RowStatus.MISSING_IN_FILE_1, null, row2, null));
                }
            }
        }

        return new ComparisonResult(resultRows, unifiedHeaders);
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
