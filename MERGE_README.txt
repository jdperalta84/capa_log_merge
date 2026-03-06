CAPA Log Merge Utility
=====================

**Purpose**
-----------
The script reads multiple CAPA (Corrective Action & Prevention) Excel files (CAR, PTO, PAR, CAF sheets) that follow the 2023+ formatting rules. It consolidates all records into a single master file (`CAPA_Master.xlsx`) that can be uploaded to your CAPA dashboard.

**Key Steps**
-------------
1. **Detect sheet type** – For each input file the script looks for the sheet that contains the doc type (CAR, PTO, PAR, CAF) by matching a short regex. If a sheet isn’t found a warning is logged.
2. **Header handling** – Some files put the actual header on the second row; the script automatically detects this and reads the sheet with the correct header.
3. **Column normalisation** – Each sheet can have slightly different column names. A mapping table per doc type lists possible raw column names and the normalised column name that the script will keep. All other columns are dropped.
4. **Data cleaning**
   * Strip whitespace from string columns.
   * Remove rows with NULL or "VOID" locations and skip a set of predefined locations.
   * Exclude certain QA initials for the PTO sheet.
5. **Date parsing** – Columns that represent dates (`init_date`, `close_date`, etc.) are coerced to `datetime`.
6. **Status and Days to Close** – A `status` column is added (`CLOSED`/`OPEN`). For closed records the number of days between the initialisation date and the close date is computed.
7. **Deduplication** – Across all files the script keeps the most recent record for each unique combination of record number (e.g. `CAR#`, `PTO#`, etc.) and location.
8. **Output** – For each doc type a sheet named `Data – <DocType>s` is written. A `Summary` sheet lists totals, closed/open counts and average days to close per year.

**Summary Format**
-----------------
| Doc Type | Year | Total | Closed | Open | Avg Days to Close |
|----------|------|-------|--------|------|------------------|

The script prints a concise report to the console (or Streamlit UI) and writes the final workbook to `CAPA_Master.xlsx`.

Feel free to copy this text into a `MERGE_README.txt` or add it to your documentation.
