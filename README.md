# Data Mapping for Financial Statements

This Python script is designed to automate the extraction, processing, and structuring of data from multiple accounting (financial) statement files in `.xlsx` format.

## Key Features

- **Automated Batch Processing**: The script processes all `.xlsx` files located in the `balances` directory.
- **Flexible Parsing**: It can analyze files with varying structures, including standard and simplified Russian reporting forms. It is robust against common layout inconsistencies like merged cells and offset data columns.
- **Organization Details Extraction**: Automatically finds and extracts company name, INN (Taxpayer Identification Number), legal address, and unit of measurement.
- **Financial Data Extraction**: Locates and extracts line item codes (`raw_code`) and their corresponding values (`saldo`) for the reporting and previous years.
- **Automatic Code Assignment**: If a line item code is missing in the report, the script assigns it based on the indicator's name using the `raw_code_dataset.txt` lookup file or generates a new composite code.
- **Consolidated Analytical Output**: All collected data is aggregated into a single `for_further_analysis.xlsx` file with multiple sheets:
  - `Organisations`: A directory of all processed companies.
  - Sheets by Year (`2024`, `2023`, `2022`): Pivot tables with financial data, where rows are companies (by INN) and columns are line item codes (`raw_code`).
  - `additional_raw_codes`: A sheet for special-purpose codes that require further interpretation.
- **Detailed Logging**: For traceability and debugging, the script generates a `data_mapping_log.txt` file, which records detailed information about every extracted value.

## Project Structure

For the script to work correctly, your project must have the following folder and file structure:

/your_project_folder/
|-- data_mapping.py         # <-- The main script
|-- raw_code_dataset.txt    # <-- Lookup file for codes and names
|-- /balances/              # <-- Directory for source report files
|   |-- report_1.xlsx
|   |-- report_2.xlsx
|   `-- ...
|
`-- (Files generated after running the script)
    |-- for_further_analysis.xlsx # <-- The final output file with consolidated data
    `-- data_mapping_log.txt      # <-- The detailed processing log

## Installation

1.  Ensure you have Python 3.8 or higher installed.
2.  Install the required libraries using pip:

    ```
    pip install pandas openpyxl xlsxwriter
    ```

## Usage

1.  Set up the project structure as shown above.
2.  Place all your source `.xlsx` report files into the `balances` directory.
3.  Place the `data_mapping.py` script and the `raw_code_dataset.txt` lookup file in the root project folder.
4.  Open a terminal or command prompt in the project's root folder and run the script:

    ```
    python data_mapping.py
    ```
5.  Wait for the process to complete. The console will display the progress, and the output files (`for_further_analysis.xlsx` and `data_mapping_log.txt`) will appear in the project folder.

## Output Files Description

- **`for_further_analysis.xlsx`**:
  - **`Organisations` Sheet**: Contains general information for each unique company.
  - **Yearly Sheets (`2024`, `2023`, ...)**: Contain financial indicators pivoted into a table where each row corresponds to a company (by INN) and columns correspond to line item codes (`raw_code`).
  - **`additional_raw_codes` Sheet**: Contains data for special codes, if any were found in the reports.

- **`data_mapping_log.txt`**:
  - Contains a detailed report of the program's execution.
  - For each value (`saldo`) found, a record is created in the following format:
    ```
    2025-07-22 10:30:00 [INFO] Found: File='report_1.xlsx', Source='Balance Sheet'!J7, Target='2024'!1150, Saldo=245.0
    ```
  - This file allows you to trace the source of every value and simplifies data validation.git