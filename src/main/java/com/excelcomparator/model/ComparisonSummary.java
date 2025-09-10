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
    private long matchedRecords = 0;
    private long mismatchedRecords = 0;
    private long missingInFile1 = 0;
    private long missingInFile2 = 0;


    /**
     * Constructs a ComparisonSummary by analyzing two lists of headers.
     *
     * @param headers1 The list of headers from the first file.
     * @param headers2 The list of headers from the second file.
     */
    public ComparisonSummary(int f1Cols, int f2Cols, Set<String> common, Set<String> only1, Set<String> only2) {
        this.file1ColumnCount = f1Cols;
        this.file2ColumnCount = f2Cols;
        this.commonColumns = common;
        this.columnsOnlyInFile1 = only1;
        this.columnsOnlyInFile2 = only2;
    }

    public void setMatchedRecords(long matchedRecords) { this.matchedRecords = matchedRecords; }
    public void setMismatchedRecords(long mismatchedRecords) { this.mismatchedRecords = mismatchedRecords; }
    public void setMissingInFile1(long missingInFile1) { this.missingInFile1 = missingInFile1; }
    public void setMissingInFile2(long missingInFile2) { this.missingInFile2 = missingInFile2; }

    public int getFile1ColumnCount() { return file1ColumnCount; }
    public int getFile2ColumnCount() { return file2ColumnCount; }
    public Set<String> getCommonColumns() { return commonColumns; }
    public Set<String> getColumnsOnlyInFile1() { return columnsOnlyInFile1; }
    public Set<String> getColumnsOnlyInFile2() { return columnsOnlyInFile2; }

    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        sb.append("--- Record Summary ---\n");
        sb.append("Matched Records: ").append(matchedRecords).append("\n");
        sb.append("Mismatched Records: ").append(mismatchedRecords).append("\n");
        sb.append("Records Only in File 1: ").append(missingInFile2).append("\n");
        sb.append("Records Only in File 2: ").append(missingInFile1).append("\n\n");
        sb.append("--- Column Summary ---\n");
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
