package com.excelcomparator.model;

import java.util.List;

/**
 * Represents the complete result of a comparison between two Excel sheets.
 * This class holds the list of result rows, the combined headers, and a summary of the comparison.
 */
public class ComparisonResult {

    /**
     * Enum representing the status of a compared row.
     */
    public enum RowStatus {
        /** Indicates the rows match perfectly based on the comparison criteria. */
        MATCH,
        /** Indicates the rows match on key columns but differ in other columns. */
        MISMATCH,
        /** Indicates a row with a specific key exists in File 2 but not in File 1. */
        MISSING_IN_FILE_1,
        /** Indicates a row with a specific key exists in File 1 but not in File 2. */
        MISSING_IN_FILE_2
    }

    /**
     * Represents a single row in the comparison result table.
     * It holds the data from both files for a given key, the comparison status, and metadata.
     */
    public static class ResultRow {
        private final RowStatus status;
        private final List<Object> data1;
        private final List<Object> data2;
        private final List<String> mismatchedColumns;
        private final int originalRowNum1;
        private final int originalRowNum2;

        /**
         * Constructs a new ResultRow.
         *
         * @param status            The comparison status of the row.
         * @param data1             The data from the row in the first file.
         * @param data2             The data from the row in the second file.
         * @param mismatchedColumns A list of column names that have different values.
         * @param r1                The original row number from the first file.
         * @param r2                The original row number from the second file.
         */
        public ResultRow(RowStatus status, List<Object> data1, List<Object> data2, List<String> mismatchedColumns, int r1, int r2) {
            this.status = status;
            this.data1 = data1;
            this.data2 = data2;
            this.mismatchedColumns = mismatchedColumns;
            this.originalRowNum1 = r1;
            this.originalRowNum2 = r2;
        }

        public RowStatus getStatus() { return status; }
        public List<Object> getData1() { return data1; }
        public List<Object> getData2() { return data2; }
        public List<String> getMismatchedColumns() { return mismatchedColumns; }
        public int getOriginalRowNum1() { return originalRowNum1; }
        public int getOriginalRowNum2() { return originalRowNum2; }
    }

    private final List<ResultRow> resultRows;
    private final List<String> headers;
    private final ComparisonSummary summary;

    /**
     * Constructs a new ComparisonResult.
     *
     * @param resultRows The list of all processed rows for the result.
     * @param headers    The combined list of headers for the result table.
     * @param summary    The summary of the comparison.
     */
    public ComparisonResult(List<ResultRow> resultRows, List<String> headers, ComparisonSummary summary) {
        this.resultRows = resultRows;
        this.headers = headers;
        this.summary = summary;
    }

    public List<ResultRow> getResultRows() { return resultRows; }
    public List<String> getHeaders() { return headers; }
    public ComparisonSummary getSummary() { return summary; }
}
