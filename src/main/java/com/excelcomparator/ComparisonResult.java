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

        public ResultRow(RowStatus status, List<Object> data1, List<Object> data2, List<Boolean> mismatches) {
            this.status = status;
            this.data1 = data1;
            this.data2 = data2;
            this.mismatches = mismatches;
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
    }

    private final List<ResultRow> resultRows;
    private final List<String> headers;

    public ComparisonResult(List<ResultRow> resultRows, List<String> headers) {
        this.resultRows = resultRows;
        this.headers = headers;
    }

    public List<ResultRow> getResultRows() {
        return resultRows;
    }

    public List<String> getHeaders() {
        return headers;
    }
}
