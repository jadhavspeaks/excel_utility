# Java Excel Comparator Utility

## Description

A powerful and easy-to-use desktop application built in Java Swing for comparing two Excel files. This utility allows users to identify differences between two sheets, with features for row and column-level comparison, key-based matching, and exporting results.

## Features

-   **Load Excel Files**: Supports both `.xls` and `.xlsx` file formats.
-   **Multi-Level Header Support**: Can parse headers that span multiple rows, correctly interpreting merged cells to form descriptive column names.
-   **Data Preview**: Shows a preview of the first 10 rows of each loaded file.
-   **Normalization (Un-pivot) Mode**: A powerful feature to transform wide-format data into a standard row-based format before comparison. This is ideal for files where columns represent entities (like products or reports) and cells contain checkmarks or values.
-   **Flexible Key Mapping**: Allows users to define one or more columns to use as a composite key for matching rows between the two files.
-   **Customizable Comparison**:
    -   **Column-level comparison**: Highlights specific cells that have different values in matched rows.
    -   **Detect missing/extra rows**: Identifies rows that exist in one file but not the other.
-   **Interactive Results Display**: Shows the comparison results in a table within the application, with color-coding to easily spot matches, mismatches, and missing rows.
-   **Export to Excel**: Exports the full comparison result to a new, color-coded Excel file for analysis and reporting.

## Requirements

-   Java 11 or higher.
-   Apache Maven 3.6.0 or higher.

## How to Run (Development Mode)

1.  **Clone the repository** or download the source code.
2.  **Navigate to the project root directory** in your terminal.
3.  **Compile and run the application** using the following Maven command:
    ```bash
    mvn compile exec:java
    ```
    This command will download the necessary dependencies, compile the source code, and launch the application.

## How to Build and Run from JAR

Alternatively, you can package the application into a single executable JAR file that includes all dependencies.

1.  **Package the application**:
    Run the following Maven command. This will create the JAR file in the `target` directory.
    ```bash
    mvn package
    ```

2.  **Run the JAR file**:
    Once packaged, you can run the application from the JAR file directly:
    ```bash
    java -jar target/excel-comparator-1.0-SNAPSHOT-jar-with-dependencies.jar
    ```

## How to Use

1.  **Load Files**: Click "Choose File 1" and "Choose File 2" to load the Excel files you want to compare.
2.  **Set Header Rows**: Use the "Headers" spinner to specify how many rows make up the column titles. The tool can handle merged cells in the header.
3.  **Select Sheets**: Choose the correct sheet from the dropdown menu for each file.
4.  **Use Normalization Mode (Optional)**:
    *   If one of your files is in a "wide" format (e.g., products listed across columns), check the "Enable Normalization" box for that file.
    *   The Normalization panel will appear. Select the columns that are your main identifiers (e.g., `CDE Name`, `Attribute`).
    *   Select the columns that you want to un-pivot (e.g., `Product A`, `Product B`, ...).
    *   Click "Reload with Normalization Settings" to see a preview of the transformed data.
5.  **Map Key Columns**: Use the dropdown menus under "Map Key Columns" to select the columns that uniquely identify a row. Click "+ Add Key" if you need to use a composite key.
6.  **Set Options**: Check or uncheck the comparison options as needed.
7.  **Run Comparison**: Click the "Run Comparison" button.
8.  **View Results**: The results will be displayed in the table at the bottom.
    -   **White rows**: Matched rows.
    -   **Yellow rows**: Mismatched rows. Mismatched cells within these rows will be highlighted in a different color.
    -   **Pink rows**: Rows that are missing in one of the files.
9.  **Export**: Click "Export to Excel" to save the results to a new Excel file.
