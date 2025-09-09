package com.excelcomparator;

import java.util.List;

public class NormalizationConfig {
    private final List<String> identifierColumns;
    private final List<String> valueColumns;
    private final String newCategoryColumnName;
    private final String newValueColumnName;

    public NormalizationConfig(List<String> identifierColumns, List<String> valueColumns, String newCategoryColumnName, String newValueColumnName) {
        this.identifierColumns = identifierColumns;
        this.valueColumns = valueColumns;
        this.newCategoryColumnName = newCategoryColumnName;
        this.newValueColumnName = newValueColumnName;
    }

    public List<String> getIdentifierColumns() {
        return identifierColumns;
    }

    public List<String> getValueColumns() {
        return valueColumns;
    }

    public String getNewCategoryColumnName() {
        return newCategoryColumnName;
    }

    public String getNewValueColumnName() {
        return newValueColumnName;
    }
}
