# Java Excel Comparator Utility

## Description

A powerful and easy-to-use desktop application built in Java Swing for comparing two Excel files. This utility allows users to identify differences between two sheets, with features for row and column-level comparison, key-based matching, and exporting results.

## Features

-   **Load Excel Files**: Supports both `.xls` and `.xlsx` file formats.
-   **Data Preview**: Shows a preview of the first 10 rows of each loaded file.
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
2.  **Map Key Columns**: Use the dropdown menus under "Map Key Columns" to select the columns that uniquely identify a row. For example, an `ID` column. Click "+ Add Key" if you need to use a composite key (e.g., `First Name` + `Last Name`).
3.  **Set Options**: Check or uncheck the comparison options as needed.
4.  **Run Comparison**: Click the "Run Comparison" button.
5.  **View Results**: The results will be displayed in the table at the bottom.
    -   **White rows**: Matched rows.
    -   **Yellow rows**: Mismatched rows. Mismatched cells within these rows will be highlighted in a different color.
    -   **Pink rows**: Rows that are missing in one of the files.
6.  **Export**: Click "Export to Excel" to save the results to a new Excel file.
