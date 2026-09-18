# Scientific Dataset Processing for Energy Data
Scripts used to process data for Electrical consumption and demand Dataset from the Federal Institute of Espírito Santo, Serra, Brazil

This tool processes raw ISSO time-series energy datasets. It standardizes timestamps, converts timezones to UTC, calculates Active Power (kW) and Energy (kWh), and translates all Portuguese data labels into a standardized English historical format.

It is designed to automatically detecting the correct header rows and parsing CSV, standard Excel (`.xlsx`), and HTML-based legacy Excel (`.xls`) files.

## Prerequisites

Python 3.8+ is required.

### Virtual environment
Create the virtual environment
```
python -m venv .venv
```
Then activate it
```
source .venv/bin/activate  
```
On Windows use: `.venv\Scripts\activate`

### Install dependencies
```
pip install pandas openpyxl lxml
```

## Usage
Run the script from the command line using the process_data.py file. It accepts an input directory of raw files, an output directory, and an optional format flag.
```
python process_data.py --input <path_to_input_directory> --output <path_to_output_directory> --format <output_format>
```

Example:
```
python process_data.py --input ./raw_data --output ./processed_data --format same
```

### Command-Line Arguments

* `-i`, `--input`: **(Required)** Path to the directory containing raw datasets (CSV, XLS, XLSX). The script will recursively search all subfolders.
* `-o`, `--output`: **(Required)** Path to the directory where processed datasets will be saved. The original relative folder structure will be maintained.
* `-f`, `--format`: *(Optional)* Defines the output file format. Options include:
  * `same` (Default): Matches the input format (CSV inputs output as CSV, Excel inputs output as Excel).
  * `csv`: Forces all processed files to be exported as `.csv`.
  * `excel`: Forces all processed files to be exported as `.xlsx` with auto-adjusted column widths.
  * `both`: Generates both `.csv` and `.xlsx` copies for every processed dataset.

### Data Processing Pipeline

1. **Ingestion:** Scans the file for the `Inicio` or `Início` header to automatically bypass junk metadata rows.
2. **Timezone Standardization:** Formats timestamps and generates `UTC-3` and standard `UTC` columns.
3. **Power Calculation:** Calculates Active Power (kW) and Total Energy (kWh) for both direct and reverse flows.
4. **Data Cleaning:** Truncates unnecessary columns (e.g., specific phase kWh, internal analyzer temperatures).
5. **Standardization:** Translates all Portuguese headers into standard English nomenclature (e.g., `Tensão A` -> `Voltage A`, `Frequência` -> `Frequency`) to ensure strict compatibility with historical files.