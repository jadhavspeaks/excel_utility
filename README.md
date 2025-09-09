# Java Excel Filter Utility

## Description

A powerful desktop application built in Java Swing for dynamically filtering and extracting data from one Excel file based on rules defined in another. This tool is designed to handle complex, real-world data structures where column names and layouts may not match, requiring flexible, user-defined mapping.

## Features

-   **Two-Way Filtering**: Filter "File 2" (the data file) using rules from "File 1" (the definition file), or vice-versa.
-   **Advanced Column Mapping**: A step-by-step UI to handle complex mapping scenarios:
    -   **One-to-Many**: Map a single attribute column in one file to multiple attribute columns in the other.
    -   **Many-to-One**: Map multiple tick-mark columns (e.g., for Products or Reports) in one file to a single column containing those values in the other.
-   **Dynamic Rule Creation**: The filtering engine automatically derives rules from the "Filter Definition File" based on which columns have a tick mark (`✓`) or any other non-empty value.
-   **Live Preview**: (To be implemented) The UI will show a preview of the filtered data.
-   **Export to Excel**: Export the final, filtered data to a new, clean Excel file.

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
    java -jar target/excel-filter-1.0-SNAPSHOT-jar-with-dependencies.jar
    ```

## How to Use

The application will guide you through a three-step process:

1.  **Step 1: Load Files & Select Direction**
    -   Use the "Choose..." buttons to select your "Filter Definition File" (Excel 1) and your "Data File" (Excel 2).
    -   Use the radio buttons to select the direction of the filter (e.g., "Use Excel 1 to filter Excel 2").
    -   Click "Next ->".

2.  **Step 2: Map Columns**
    -   The application will display the column headers from both files.
    -   Use the dropdown menus to map the corresponding columns for Attributes, Reports, and Products.
    -   Use the "Add..." buttons to create additional mapping rows for many-to-one or one-to-many relationships.
    -   Click "Run Filter".

3.  **Step 3: Filtered Results**
    -   The table will display the rows from your data file that match the filter criteria.
    -   Click "Export to Excel" to save the results.
    -   Click "<- Back" to adjust your mappings and run the filter again.
