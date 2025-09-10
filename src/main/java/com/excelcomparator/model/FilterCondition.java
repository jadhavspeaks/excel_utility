package com.excelcomparator.model;

public class FilterCondition {
    private final String columnName;
    private final String operator;
    private final String value;

    public FilterCondition(String columnName, String operator, String value) {
        this.columnName = columnName;
        this.operator = operator;
        this.value = value;
    }

    public String getColumnName() {
        return columnName;
    }

    public String getOperator() {
        return operator;
    }

    public String getValue() {
        return value;
    }

    public boolean matches(String cellValue) {
        if (cellValue == null) {
            return false;
        }
        switch (operator.toUpperCase()) {
            case "EQUALS":
                return cellValue.equalsIgnoreCase(value);
            case "CONTAINS":
                return cellValue.toLowerCase().contains(value.toLowerCase());
            case "NOT_CONTAINS":
                return !cellValue.toLowerCase().contains(value.toLowerCase());
            case "STARTS_WITH":
                return cellValue.toLowerCase().startsWith(value.toLowerCase());
            case "ENDS_WITH":
                return cellValue.toLowerCase().endsWith(value.toLowerCase());
            default:
                return false;
        }
    }
}
