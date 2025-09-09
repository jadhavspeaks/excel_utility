package com.excelcomparator;

import java.util.List;

public class ComparisonResult {

    public enum RowStatus {
        MATCH,
        MISMATCH,
        MISSING_IN_FILE_2, // Present in File 1, missing in File 2
        MISSING_IN_FILE_1  // Present in File 2, missing in File 1
    }

    public static class ResultRow {
        private final RowStatus status;
        private final List<Object> data1;
        private final List<Object> data2;
        private final List<Boolean> mismatches; // for column-level mismatch highlighting
        private final int originalRowNum1;
        private final int originalRowNum2;

        public ResultRow(RowStatus status, List<Object> data1, List<Object> data2, List<Boolean> mismatches, int originalRowNum1, int originalRowNum2) {
            this.status = status;
            this.data1 = data1;
            this.data2 = data2;
            this.mismatches = mismatches;
            this.originalRowNum1 = originalRowNum1;
            this.originalRowNum2 = originalRowNum2;
        }

        public RowStatus getStatus() {
            return status;
        }

        public List<Object> getData1() {
            return data1;
        }

        public List<Object> getData2() {
            return data2;
        }

        public List<Boolean> getMismatches() {
            return mismatches;
        }

        public int getOriginalRowNum1() {
            return originalRowNum1;
        }

        public int getOriginalRowNum2() {
            return originalRowNum2;
        }
    }

    private final List<ResultRow> resultRows;
    private final List<String> headers;
    private final ComparisonSummary summary;

    public ComparisonResult(List<ResultRow> resultRows, List<String> headers, ComparisonSummary summary) {
        this.resultRows = resultRows;
        this.headers = headers;
        this.summary = summary;
    }

    public List<ResultRow> getResultRows() {
        return resultRows;
    }

    public List<String> getHeaders() {
        return headers;
    }

    public ComparisonSummary getSummary() {
        return summary;
    }
}
