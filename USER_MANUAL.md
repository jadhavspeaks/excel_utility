# Excel Comparator Suite - User Manual

This manual provides detailed instructions on how to use the Excel Comparator Suite, including the main comparison tool and the additional utilities.

## Table of Contents
1.  [Getting Started](#getting-started)
2.  [The Excel Comparator](#the-excel-comparator)
    - [Loading Files](#loading-files)
    - [Automated Normalization](#automated-normalization)
    - [Mapping Key Columns](#mapping-key-columns)
    - [Running the Comparison](#running-the-comparison)
    - [Interpreting the Results](#interpreting-the-results)
    - [Exporting Results](#exporting-results)
3.  [Tool: Simple Normalizer](#tool-simple-normalizer)
4.  [Tool: Simple Keyword Filter](#tool-simple-keyword-filter)

---

## 1. Getting Started

To run the application, you need to have Java installed on your system. You can run the application from the executable JAR file.

**From the command line:**
```bash
java -jar excel-comparator-1.0-SNAPSHOT-jar-with-dependencies.jar
```
This will open the main application window.

---

## 2. The Excel Comparator

The Excel Comparator is the main feature of the application. It allows you to perform a detailed, row-by-row comparison of two Excel sheets.

### Loading Files
1.  **Choose File:** For each file (File 1 and File 2), click the "Choose File..." button to select an `.xls` or `.xlsx` file from your computer.
2.  **Select Header Rows:** Use the "Headers" spinner to specify how many rows at the top of your file make up the header. The tool can handle multi-line headers by merging them.
3.  **Select Sheet:** Once a file is loaded, the "Sheet" dropdown will be populated with all the sheets in that workbook. Select the sheet you wish to compare.
4.  **Preview:** A preview of the first 10 rows of the selected sheet will be displayed in the tables below the file selection panels.

### Automated Normalization
This is a powerful feature for handling "wide" data formats, where data is spread across multiple columns under merged headers (e.g., monthly data).

-   **When to use it:** Use this if your data has a structure where column headers are grouped, and you want to compare the data in a more granular, "un-pivoted" way.
-   **How it works:**
    1.  Check the **"Enable Automated Normalization"** checkbox.
    2.  The tool will automatically detect columns that are under merged headers.
    3.  It will transform the data from a wide format to a long format, creating two new columns: `Dimension` (the original column header) and `Applicable` (with a value of "Yes" if the cell was not empty).
    4.  This new, normalized structure will be used for the comparison.

### Mapping Key Columns
To compare the two sheets, you must tell the tool which columns to use to match rows between the two files. These are your primary keys.

1.  In the "Comparison Options" panel, use the dropdowns to select the column from File 1 and the corresponding column from File 2 that should be used as a key.
2.  You can add multiple key columns by clicking the **"+ Add Key"** button. For a row to be considered a match, the data in **all** mapped key columns must be identical.
3.  Click the **"-"** button to remove a key mapping.

### Running the Comparison
1.  **Detect missing/extra rows:** This checkbox is enabled by default. It ensures the tool will report rows that exist in one file but not the other (based on the key columns).
2.  **Show All Columns:**
    -   If **unchecked** (default), the results table will only show the key columns you mapped. This is useful for quickly verifying that the keys match.
    -   If **checked**, the results table will show all columns from both files, providing a full data comparison.
3.  Click the **"Run Comparison"** button to start the process.

### Interpreting the Results
The results are displayed in two parts:

1.  **Comparison Summary:** A text area at the top shows a summary of the comparison, including:
    -   Total columns in each file.
    -   A list of common columns and columns that are unique to each file.
2.  **Results Table:**
    -   **File 1 Row / File 2 Row:** The original row number from the source file.
    -   **Status:**
        -   `MATCH`: The rows match based on the key columns (and all other columns if "Show All Columns" is checked).
        -   `MISMATCH`: The rows match on key columns, but data in at least one other column is different.
        -   `MISSING_IN_FILE_1`: This row (based on its key) exists in File 2 but not in File 1.
        -   `MISSING_IN_FILE_2`: This row (based on its key) exists in File 1 but not in File 2.
    -   **Mismatched Columns:** For `MISMATCH` rows, this column lists exactly which columns have different values.
    -   **Data Columns:** The remaining columns show the data. For mismatched cells, the format will be `File 1 Value -> File 2 Value`.

### Exporting Results
-   Click the **"Export to Excel"** button to save the full, detailed comparison results (including the summary) to an `.xlsx` file.

---

## 3. Tool: Simple Normalizer

This tool provides a quick way to convert a wide-format Excel sheet into a long-format (un-pivoted) table. It is optimized to handle very large files with low memory usage.

1.  Go to **Tools -> Simple Normalizer** in the menu bar.
2.  **Choose Input File:** Select the Excel file you want to process.
3.  **Select Sheet:** Choose the specific sheet you want to normalize from the dropdown.
4.  Click **"Run Normalization"**.
5.  You will be prompted to specify a location to save the output file.
6.  **Output Format:** The tool takes the first column of your sheet as the **Key**. It then creates a new three-column table with the headers `Key | ColumnName | Value` for every other non-empty cell in the original sheet.
7.  A log will show the progress and a summary of rows processed.

---

## 4. Tool: Simple Keyword Filter

This tool allows you to filter a large Excel sheet based on a list of keywords in a specific column.

1.  Go to **Tools -> Simple Keyword Filter** in the menu bar.
2.  **Choose Input File:** Select the Excel file you want to filter.
3.  **Select Sheet:** Choose the sheet you want to filter.
4.  **Select Key Column:** From the dropdown, select the column that contains the values you want to filter by.
5.  **Enter Keywords:** In the text area, enter the list of keywords you want to find in the key column. Enter one keyword per line.
6.  Click **"Start Filtering"**.
7.  You will be prompted to specify a location to save the filtered output file. The output file will contain only the rows where the value in the key column matched one of your keywords.
