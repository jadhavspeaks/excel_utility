package com.excelcomparator;

import java.util.List;

public class FilterMappingConfig {

    public enum FilterDirection {
        FILE1_FILTERS_FILE2,
        FILE2_FILTERS_FILE1
    }

    private FilterDirection direction;
    private String file1_AttributeColumn;
    private List<String> file2_AttributeColumns;
    private List<String> file1_ReportColumns;
    private String file2_ReportColumn;
    private List<String> file1_ProductColumns;
    private String file2_ProductColumn;

    // Getters and Setters
    public FilterDirection getDirection() { return direction; }
    public void setDirection(FilterDirection direction) { this.direction = direction; }
    public String getFile1_AttributeColumn() { return file1_AttributeColumn; }
    public void setFile1_AttributeColumn(String file1_AttributeColumn) { this.file1_AttributeColumn = file1_AttributeColumn; }
    public List<String> getFile2_AttributeColumns() { return file2_AttributeColumns; }
    public void setFile2_AttributeColumns(List<String> file2_AttributeColumns) { this.file2_AttributeColumns = file2_AttributeColumns; }
    public List<String> getFile1_ReportColumns() { return file1_ReportColumns; }
    public void setFile1_ReportColumns(List<String> file1_ReportColumns) { this.file1_ReportColumns = file1_ReportColumns; }
    public String getFile2_ReportColumn() { return file2_ReportColumn; }
    public void setFile2_ReportColumn(String file2_ReportColumn) { this.file2_ReportColumn = file2_ReportColumn; }
    public List<String> getFile1_ProductColumns() { return file1_ProductColumns; }
    public void setFile1_ProductColumns(List<String> file1_ProductColumns) { this.file1_ProductColumns = file1_ProductColumns; }
    public String getFile2_ProductColumn() { return file2_ProductColumn; }
    public void setFile2_ProductColumn(String file2_ProductColumn) { this.file2_ProductColumn = file2_ProductColumn; }
}
