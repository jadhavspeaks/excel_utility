package com.excelcomparator;

import java.util.List;
import java.util.Set;

public class ComparisonSummary {
    private final int file1ColumnCount;
    private final int file2ColumnCount;
    private final List<String> columnsOnlyInFile1;
    private final List<String> columnsOnlyInFile2;
    private final Set<String> commonColumns;

    public ComparisonSummary(int file1ColumnCount, int file2ColumnCount, List<String> columnsOnlyInFile1, List<String> columnsOnlyInFile2, Set<String> commonColumns) {
        this.file1ColumnCount = file1ColumnCount;
        this.file2ColumnCount = file2ColumnCount;
        this.columnsOnlyInFile1 = columnsOnlyInFile1;
        this.columnsOnlyInFile2 = columnsOnlyInFile2;
        this.commonColumns = commonColumns;
    }

    public int getFile1ColumnCount() { return file1ColumnCount; }
    public int getFile2ColumnCount() { return file2ColumnCount; }
    public List<String> getColumnsOnlyInFile1() { return columnsOnlyInFile1; }
    public List<String> getColumnsOnlyInFile2() { return columnsOnlyInFile2; }
    public Set<String> getCommonColumns() { return commonColumns; }

    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        sb.append("--- Comparison Summary ---\n");
        sb.append("File 1 Distinct Columns: ").append(file1ColumnCount).append("\n");
        sb.append("File 2 Distinct Columns: ").append(file2ColumnCount).append("\n");
        sb.append("Common Columns: ").append(commonColumns.size()).append("\n\n");

        if (!columnsOnlyInFile1.isEmpty()) {
            sb.append("Columns only in File 1: \n");
            columnsOnlyInFile1.forEach(c -> sb.append("  - ").append(c).append("\n"));
            sb.append("\n");
        }
        if (!columnsOnlyInFile2.isEmpty()) {
            sb.append("Columns only in File 2: \n");
            columnsOnlyInFile2.forEach(c -> sb.append("  - ").append(c).append("\n"));
            sb.append("\n");
        }
        sb.append("--------------------------\n");
        return sb.toString();
    }
}
