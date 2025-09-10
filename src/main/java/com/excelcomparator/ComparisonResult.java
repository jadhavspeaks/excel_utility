package com.excelcomparator;

import java.util.List;

public class ComparisonResult {

    public enum RowStatus {
        MATCH,
        MISMATCH,
        MISSING_IN_FILE_2,
        MISSING_IN_FILE_1
    }

    public static class ResultRow {
        private final RowStatus status;
        private final List<Object> data1;
        private final List<Object> data2;
        private final int originalRowNum1;
        private final int originalRowNum2;
        private final List<String> mismatchedColumns;

        public ResultRow(RowStatus status, List<Object> data1, List<Object> data2, int originalRowNum1, int originalRowNum2, List<String> mismatchedColumns) {
            this.status = status;
            this.data1 = data1;
            this.data2 = data2;
            this.originalRowNum1 = originalRowNum1;
            this.originalRowNum2 = originalRowNum2;
            this.mismatchedColumns = mismatchedColumns;
        }

        public RowStatus getStatus() { return status; }
        public List<Object> getData1() { return data1; }
        public List<Object> getData2() { return data2; }
        public int getOriginalRowNum1() { return originalRowNum1; }
        public int getOriginalRowNum2() { return originalRowNum2; }
        public List<String> getMismatchedColumns() { return mismatchedColumns; }
    }

    private final List<ResultRow> resultRows;
    private final List<String> headers;
    private final ComparisonSummary summary;

    public ComparisonResult(List<ResultRow> resultRows, List<String> headers, ComparisonSummary summary) {
        this.resultRows = resultRows;
        this.headers = headers;
        this.summary = summary;
    }

    public List<ResultRow> getResultRows() { return resultRows; }
    public List<String> getHeaders() { return headers; }
    public ComparisonSummary getSummary() { return summary; }
}
