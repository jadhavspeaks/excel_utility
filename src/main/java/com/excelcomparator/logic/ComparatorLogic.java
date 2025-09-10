package com.excelcomparator.logic;

import com.excelcomparator.model.ComparisonResult;
import com.excelcomparator.model.ComparisonSummary;
import com.excelcomparator.util.ExcelUtil;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

/**
 * Contains the core logic for comparing two sets of Excel data.
 */
public class ComparatorLogic {

    /**
     * Compares two ExcelData objects based on specified key columns.
     *
     * @param data1       The first Excel data set.
     * @param data2       The second Excel data set.
     * @param keyColumns  A map where keys are key-column names from file 1 and values are corresponding key-column names from file 2.
     * @param fullRow     If true, compares all columns; if false, only compares key columns.
     * @param findMissing If true, identifies rows that are unique to each file.
     * @return A ComparisonResult object containing the detailed results of the comparison.
     */
    public static ComparisonResult compare(ExcelUtil.ExcelData data1, ExcelUtil.ExcelData data2, Map<String, String> keyColumns, boolean fullRow, boolean findMissing) {
        List<String> headers1 = data1.getHeaders();
        List<String> headers2 = data2.getHeaders();
        List<List<Object>> rows1 = data1.getData();
        List<List<Object>> rows2 = data2.getData();

        List<Integer> keyIndices1 = keyColumns.keySet().stream().map(headers1::indexOf).collect(Collectors.toList());
        List<Integer> keyIndices2 = keyColumns.values().stream().map(headers2::indexOf).collect(Collectors.toList());

        Map<String, List<Object>> map2 = new HashMap<>();
        Map<String, Integer> rowNumMap2 = new HashMap<>();
        if (findMissing) {
            for (int i = 0; i < rows2.size(); i++) {
                String key = buildKey(rows2.get(i), keyIndices2);
                map2.put(key, rows2.get(i));
                rowNumMap2.put(key, data2.getOriginalRowNumbers().get(i));
            }
        }

        List<ComparisonResult.ResultRow> resultRows = new ArrayList<>();
        List<String> commonHeaders = headers1.stream().filter(headers2::contains).collect(Collectors.toList());
        if(fullRow) {
             commonHeaders = new ArrayList<>(data1.getHeaders());
             commonHeaders.addAll(data2.getHeaders());
             commonHeaders = commonHeaders.stream().distinct().collect(Collectors.toList());
        }

        boolean[] row1Matched = new boolean[rows1.size()];

        for (int i = 0; i < rows1.size(); i++) {
            List<Object> row1 = rows1.get(i);
            String key1 = buildKey(row1, keyIndices1);
            int originalRowNum1 = data1.getOriginalRowNumbers().get(i);

            if (map2.containsKey(key1)) {
                List<Object> row2 = map2.get(key1);
                int originalRowNum2 = rowNumMap2.get(key1);
                List<String> mismatchedColumns = new ArrayList<>();
                boolean mismatch = false;

                if (fullRow) {
                    for (String header : commonHeaders) {
                        int idx1 = headers1.indexOf(header);
                        int idx2 = headers2.indexOf(header);
                        Object val1 = (idx1 != -1) ? row1.get(idx1) : null;
                        Object val2 = (idx2 != -1) ? row2.get(idx2) : null;
                        if (!Objects.equals(val1, val2)) {
                            mismatch = true;
                            mismatchedColumns.add(header);
                        }
                    }
                }

                resultRows.add(new ComparisonResult.ResultRow(
                        mismatch ? ComparisonResult.RowStatus.MISMATCH : ComparisonResult.RowStatus.MATCH,
                        row1, row2, mismatchedColumns, originalRowNum1, originalRowNum2));

                row1Matched[i] = true;
                map2.remove(key1); // Remove from map to find extras in file 2 later
            } else if (findMissing) {
                resultRows.add(new ComparisonResult.ResultRow(ComparisonResult.RowStatus.MISSING_IN_FILE_2, row1, null, new ArrayList<>(), originalRowNum1, -1));
            }
        }

        if (findMissing) {
            for (String key2 : map2.keySet()) {
                List<Object> row2 = map2.get(key2);
                int originalRowNum2 = rowNumMap2.get(key2);
                resultRows.add(new ComparisonResult.ResultRow(ComparisonResult.RowStatus.MISSING_IN_FILE_1, null, row2, new ArrayList<>(), -1, originalRowNum2));
            }
        }

        ComparisonSummary summary = new ComparisonSummary(headers1, headers2);
        return new ComparisonResult(resultRows, commonHeaders, summary);
    }

    /**
     * Builds a composite key string from a row's data based on the specified key column indices.
     *
     * @param row        The list of objects representing a row.
     * @param keyIndices The list of integer indices for the key columns.
     * @return A single string representing the composite key.
     */
    private static String buildKey(List<Object> row, List<Integer> keyIndices) {
        StringBuilder key = new StringBuilder();
        for (int index : keyIndices) {
            if (index < row.size() && row.get(index) != null) {
                key.append(row.get(index).toString());
            }
            key.append("||");
        }
        return key.toString();
    }
}
