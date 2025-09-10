# Java Excel Comparator Utility

## Description

A powerful and easy-to-use desktop application built in Java Swing for comparing two Excel files. This utility allows users to identify differences between two sheets, with features for row and column-level comparison, key-based matching, and exporting results.

## Features

-   **Load Excel Files**: Supports both `.xls` and `.xlsx` file formats.
-   **Sheet Selection**: Choose the specific sheet to compare from a dropdown menu for each file.
-   **Multi-Level Header Support**: Can parse headers that span multiple rows, correctly interpreting merged cells to form descriptive column names.
-   **Automated Normalization**: A powerful feature to automatically detect and transform wide-format data into a standard row-based format. It identifies top-level merged headers and un-pivots the columns beneath them.
-   **Flexible Key Mapping**: Allows users to define one or more columns to use as a composite key for matching rows between the two files.
-   **Customizable Comparison**:
    -   **Key-only view**: A checkbox to switch between a full, detailed comparison report and a simplified view showing only the mapped key columns.
    -   **Detect missing/extra rows**: Identifies rows that exist in one file but not the other.
-   **Enhanced Reporting**:
    -   **Column Summary**: The report includes a summary of column differences (added, deleted, common, counts).
    -   **Precise Mismatch Location**: The report includes the original row numbers and a list of mismatched column names to pinpoint the exact location of any data differences.
-   **Export to Excel**: Exports the full, detailed comparison result to a new Excel file.

## Requirements

-   Java 11 or higher.
-   Apache Maven 3.6.0 or higher.

## How to Run

1.  **Clone the repository** or download the source code.
2.  **Package the application** into an executable JAR:
    ```bash
    mvn package
    ```
3.  **Run the application** from the generated JAR file in the `target` directory:
    ```bash
    java -jar target/excel-comparator-1.0-SNAPSHOT-jar-with-dependencies.jar
    ```

## How to Use

1.  **Load Files**: Click "Choose File..." for each file.
2.  **Set Headers & Sheet**: Use the "Headers" spinner to specify the number of header rows and select the correct sheet from the dropdown.
3.  **Use Normalization (Optional)**: If a file is in a wide format, check the "Enable Automated Normalization" box for that file to transform it before comparison.
4.  **Map Key Columns**: Use the dropdown menus under "Map Key Columns" to select the columns that uniquely identify a row.
5.  **Set Options**:
    -   Leave "Key-only view" unchecked for a full detailed comparison.
    -   Check it to see a simplified view of only the key columns.
6.  **Run Comparison**: Click the "Run Comparison" button.
7.  **View & Export Results**: View the summary and detailed results in the UI, or click "Export to Excel" to save the report.
8.  **Exit**: Click the "Exit" button to close the application.
