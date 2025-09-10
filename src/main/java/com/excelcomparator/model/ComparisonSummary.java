package com.excelcomparator.model;

import java.util.List;
import java.util.Set;
import java.util.TreeSet;
import java.util.stream.Collectors;

/**
 * Holds summary information about the comparison of two sets of headers.
 * It calculates the total column counts, common columns, and columns unique to each file.
 */
public class ComparisonSummary {
    private final int file1ColumnCount;
    private final int file2ColumnCount;
    private final Set<String> commonColumns;
    private final Set<String> columnsOnlyInFile1;
    private final Set<String> columnsOnlyInFile2;

    /**
     * Constructs a ComparisonSummary by analyzing two lists of headers.
     *
     * @param headers1 The list of headers from the first file.
     * @param headers2 The list of headers from the second file.
     */
    public ComparisonSummary(List<String> headers1, List<String> headers2) {
        this.file1ColumnCount = headers1.size();
        this.file2ColumnCount = headers2.size();

        Set<String> set1 = new TreeSet<>(String.CASE_INSENSITIVE_ORDER);
        set1.addAll(headers1);

        Set<String> set2 = new TreeSet<>(String.CASE_INSENSITIVE_ORDER);
        set2.addAll(headers2);

        this.commonColumns = set1.stream()
                                .filter(set2::contains)
                                .collect(Collectors.toCollection(() -> new TreeSet<>(String.CASE_INSENSITIVE_ORDER)));

        this.columnsOnlyInFile1 = set1.stream()
                                     .filter(h -> !set2.contains(h))
                                     .collect(Collectors.toCollection(() -> new TreeSet<>(String.CASE_INSENSITIVE_ORDER)));

        this.columnsOnlyInFile2 = set2.stream()
                                     .filter(h -> !set1.contains(h))
                                     .collect(Collectors.toCollection(() -> new TreeSet<>(String.CASE_INSENSITIVE_ORDER)));
    }

    public int getFile1ColumnCount() { return file1ColumnCount; }
    public int getFile2ColumnCount() { return file2ColumnCount; }
    public Set<String> getCommonColumns() { return commonColumns; }
    public Set<String> getColumnsOnlyInFile1() { return columnsOnlyInFile1; }
    public Set<String> getColumnsOnlyInFile2() { return columnsOnlyInFile2; }

    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        sb.append("--- Comparison Summary ---\n");
        sb.append("File 1 Total Columns: ").append(file1ColumnCount).append("\n");
        sb.append("File 2 Total Columns: ").append(file2ColumnCount).append("\n");
        sb.append("Common Columns (").append(commonColumns.size()).append("): ").append(String.join(", ", commonColumns)).append("\n");
        if (!columnsOnlyInFile1.isEmpty()) {
            sb.append("Columns Only in File 1 (").append(columnsOnlyInFile1.size()).append("): ").append(String.join(", ", columnsOnlyInFile1)).append("\n");
        }
        if (!columnsOnlyInFile2.isEmpty()) {
            sb.append("Columns Only in File 2 (").append(columnsOnlyInFile2.size()).append("): ").append(String.join(", ", columnsOnlyInFile2)).append("\n");
        }
        return sb.toString();
    }
}
