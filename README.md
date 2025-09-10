# Excel Comparator Suite

This is a Java Swing-based desktop application for comparing and manipulating Excel files.

## Features

*   **Advanced Excel Comparison:**
    *   Compare two Excel sheets row by row based on user-defined key columns.
    *   Supports multi-row headers.
    *   Provides a detailed summary of differences, including mismatched columns and rows that are unique to each file.
    *   **Automated Normalization:** An optional feature to "un-pivot" wide-format data into a long format before comparison, making it possible to compare complex layouts.
    *   Export detailed comparison results to an Excel file.

*   **Simple Normalizer Tool:**
    *   A standalone tool (accessible from the `Tools` menu) for converting a wide-format sheet into a long-format (Key, ColumnName, Value) table.
    *   Optimized for very large files using a streaming parser to keep memory usage low.

*   **Multi-Condition Filter Tool:**
    *   A standalone tool (accessible from the `Tools` menu) for filtering an Excel sheet based on multiple criteria.
    *   Users can add multiple conditions (e.g., `Column A = 'Value1'` AND `Column B CONTAINS 'Value2'`).
    *   Uses a streaming writer for efficient handling of large output files.

## Requirements

*   Java 11 or higher.
*   Apache Maven (for building from source).

## How to Build

This project is built using Apache Maven.

1.  **Clone the repository:**
    ```bash
    git clone <repository-url>
    ```

2.  **Navigate to the project directory:**
    ```bash
    cd excel-comparator
    ```

3.  **Build the project using Maven:**
    ```bash
    mvn package
    ```
    This will compile the source code, run tests, and create an executable JAR file with all dependencies in the `target/` directory.

## How to Run

After building the project, you can run the application from the command line:

```bash
java -jar target/excel-comparator-1.0-SNAPSHOT-jar-with-dependencies.jar
```

Alternatively, you can run it directly through Maven:
```bash
mvn exec:java -Dexec.mainClass="com.excelcomparator.ExcelComparator"
```
