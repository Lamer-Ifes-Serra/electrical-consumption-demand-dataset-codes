"""
==============================================================================
Scientific Dataset Processing Script for Energy Data
==============================================================================
Transforms raw ISSO time-series datasets by formatting timestamps, converting
timezones to UTC, calculating Active Power (kW) and Energy (kWh), and
translating all data labels to match past processed files exactly.

Execution:
    python process_data.py --input ./raw_data --output ./processed_data --format same
==============================================================================
"""

import os
import re
import argparse
import pandas as pd
from openpyxl import load_workbook

# Exact historical translation dictionary ensuring compatibility with past files
COLUMN_TRANSLATIONS = {
    'UTC-3 Inicio': 'UTC-3 Start',
    'UTC-3 Fim': 'UTC-3 End',
    'UTC Inicio': 'UTC Start',
    'UTC Fim': 'UTC End',
    'Frequencia': 'Frequency',
    'Frequência': 'Frequency',

    # Voltages & Currents
    'Tensão A': 'Voltage A',
    'Tensão B': 'Voltage B',
    'Tensão C': 'Voltage C',
    'Tensão A-B': 'Voltage A-B',
    'Tensão B-C': 'Voltage B-C',
    'Tensão C-A': 'Voltage C-A',
    'Corrente A': 'Current A',
    'Corrente B': 'Current B',
    'Corrente C': 'Current C',
    'Corrente Neutro medido': 'Neutral Current Measured',
    'Corrente Neutro Medido': 'Neutral Current Measured',
    'Corrente Neutro calculado': 'Neutral Current Calculated',
    'Corrente Neutro Calculado': 'Neutral Current Calculated',

    # Apparent & Reactive Power
    'P. Aparente A': 'Apparent Power A',
    'P. Aparente B': 'Apparent Power B',
    'P. Aparente C': 'Apparent Power C',
    'P. Aparente Soma aritmética': 'Apparent Power Arithmetic Sum',
    'P. Aparente Soma Aritmética': 'Apparent Power Arithmetic Sum',
    'P. Aparente Soma vetorial': 'Apparent Power Vector Sum',
    'P. Aparente Soma Vetorial': 'Apparent Power Vector Sum',
    'P. Reativa A': 'Reactive Power A',
    'P. Reativa B': 'Reactive Power B',
    'P. Reativa C': 'Reactive Power C',
    'P. Reativa Total': 'Total Reactive Power',

    # Power Factor & Angles
    'FP Real A': 'Real Power Factor A',
    'FP Real B': 'Real Power Factor B',
    'FP Real C': 'Real Power Factor C',
    'FP Real Soma vetorial': 'Average Real Power Factor', # Historical translation mapping
    'FP Real Soma Vetorial': 'Average Real Power Factor',
    'FP Real Média': 'Average Real Power Factor',
    'Cos(&phi;) A': 'Cos(φ) A',
    'Cos(&phi;) B': 'Cos(φ) B',
    'Cos(&phi;) C': 'Cos(φ) C',
    'Cos(φ) A': 'Cos(φ) A',
    'Cos(φ) B': 'Cos(φ) B',
    'Cos(φ) C': 'Cos(φ) C',
    'Cos(&phi;) Média ponderada': 'Average Cos(φ)',
    'Cos(φ) Média ponderada': 'Average Cos(φ)',
    'Cos(&phi;) Média': 'Average Cos(φ)',
    'Cos(φ) Média': 'Average Cos(φ)',
    'Ind Fase 1': 'Phase Indicator 1',
    'Ind Fase 2': 'Phase Indicator 2',
    'Ind Fase 3': 'Phase Indicator 3',
    'Ind Média': 'Average Indicator',

    # Direct Active Power
    'Direto: P. Ativa Fund+Harm A': 'Direct: Active Power Fund+Harm A',
    'Direto: P. Ativa Fund+Harm B': 'Direct: Active Power Fund+Harm B',
    'Direto: P. Ativa Fund+Harm C': 'Direct: Active Power Fund+Harm C',
    'Direto: P. Ativa Fund+Harm Total': 'Direct: Total Active Power Fund+Harm',
    'Direto: P. Ativa Fund. A': 'Direct: Active Power Fundamental A',
    'Direto: P. Ativa Fund. B': 'Direct: Active Power Fundamental B',
    'Direto: P. Ativa Fund. C': 'Direct: Active Power Fundamental C',
    'Direto: P. Ativa Fund. Total': 'Direct: Total Active Power Fundamental',
    'Direto: P. Ativa Harm. A': 'Direct: Active Power Harmonic A',
    'Direto: P. Ativa Harm. B': 'Direct: Active Power Harmonic B',
    'Direto: P. Ativa Harm. C': 'Direct: Active Power Harmonic C',
    'Direto: P. Ativa Harm. Total': 'Direct: Total Active Power Harmonic',

    # Reverse Active Power
    'Reverso: P. Ativa Fund+Harm A': 'Reverse: Active Power Fund+Harm A',
    'Reverso: P. Ativa Fund+Harm B': 'Reverse: Active Power Fund+Harm B',
    'Reverso: P. Ativa Fund+Harm C': 'Reverse: Active Power Fund+Harm C',
    'Reverso: P. Ativa Fund+Harm Total': 'Reverse: Total Active Power Fund+Harm',
    'Reverso: P. Ativa Fund. A': 'Reverse: Active Power Fundamental A',
    'Reverso: P. Ativa Fund. B': 'Reverse: Active Power Fundamental B',
    'Reverso: P. Ativa Fund. C': 'Reverse: Active Power Fundamental C',
    'Reverso: P. Ativa Fund. Total': 'Reverse: Total Active Power Fundamental',
    'Reverso: P. Ativa Harm. A': 'Reverse: Active Power Harmonic A',
    'Reverso: P. Ativa Harm. B': 'Reverse: Active Power Harmonic B',
    'Reverso: P. Ativa Harm. C': 'Reverse: Active Power Harmonic C',
    'Reverso: P. Ativa Harm. Total': 'Reverse: Total Active Power Harmonic',

    # Unbalances
    'Desequilibrio de tensão (fasorial) Total': 'Voltage Unbalance (Phasor) Total',
    'Desequilibrio de Tensão (Fasorial) Total': 'Voltage Unbalance (Phasor) Total',
    'Desequilibrio de tensão (amplitude) Total': 'Voltage Unbalance (Amplitude) Total',
    'Desequilibrio de Tensão (Amplitude) Total': 'Voltage Unbalance (Amplitude) Total',
    'Desequilibrio de corrente (amplitude) Total': 'Current Unbalance (Amplitude) Total',
    'Desequilibrio de Corrente (Amplitude) Total': 'Current Unbalance (Amplitude) Total',

    # Calculated Metrics
    'Direto: P. Ativa Fund+Harm Total (kW)': 'Direct: Total Active Power Fund+Harm (kW)',
    'Reverso: P. Ativa Fund+Harm Total (kW)': 'Reverse: Total Active Power Fund+Harm (kW)',
    'Direto: Kwh Total': 'Direct: Total kWh',
    'Reverso: Kwh Total': 'Reverse: Total kWh'
}


def adjust_column_width(file_path: str) -> None:
    """
    Dynamically adjusts the column widths of an Excel file to fit the content.
    """
    wb = load_workbook(file_path)

    for ws in wb.worksheets:
        for col in ws.columns:
            max_length = 0

            try:
                column = col[0].column_letter
            except AttributeError:
                continue

            for cell in col:
                try:
                    if cell.value is not None:
                        max_length = max(max_length, len(str(cell.value)))
                except Exception:
                    pass

            ws.column_dimensions[column].width = max_length + 2

    wb.save(file_path)
    wb.close()


def extract_year_month(filename: str) -> str:
    """
    Extracts the year and month (YYYY_MM) from the dataset filename.
    """
    match = re.search(r'^(\d{4})_(\d{2})_', filename)
    if match:
        return f"{match.group(1)}_{match.group(2)}"

    match = re.search(r'ISSO_(\d{4})(\d{2})', filename, re.IGNORECASE)
    if match:
        return f"{match.group(1)}_{match.group(2)}"

    match = re.search(r'(\d{4})_(\d{2})_\d{2}_\d{4}_\d{2}_\d{2}', filename)
    if match:
        return f"{match.group(1)}_{match.group(2)}"

    return "NO_DATE"


def read_dataset_file(filepath: str) -> pd.DataFrame:
    """
    Reads the raw dataset file using robust header detection.
    Handles standard Excel, CSV, and HTML tables masquerading as .xls files.
    """
    extension = os.path.splitext(filepath)[1].lower()
    df = None
    header_idx = None

    def has_inicio(arr):
        return any('inicio' in str(x).lower() or 'início' in str(x).lower() for x in arr)

    if extension == ".csv":
        try:
            temp_df = pd.read_csv(filepath, encoding="latin1", sep=",")
            if has_inicio(temp_df.columns):
                df = temp_df
                header_idx = -1
            else:
                for idx, row in temp_df.iterrows():
                    if has_inicio(row.values):
                        header_idx = idx
                        df = temp_df
                        break
        except Exception:
            pass

    if df is None or header_idx is None:
        try:
            tables = pd.read_html(filepath, encoding='latin1', decimal=',', thousands='.')
            if tables:
                temp_df = tables[0]
                if has_inicio(temp_df.columns):
                    df = temp_df
                    header_idx = -1
                else:
                    for idx, row in temp_df.iterrows():
                        if has_inicio(row.values):
                            header_idx = idx
                            df = temp_df
                            break
        except Exception:
            pass

    if df is None or header_idx is None:
        try:
            temp_df = pd.read_excel(filepath, decimal=',', thousands='.')
            if has_inicio(temp_df.columns):
                df = temp_df
                header_idx = -1
            else:
                for idx, row in temp_df.iterrows():
                    if has_inicio(row.values):
                        header_idx = idx
                        df = temp_df
                        break
        except Exception as e:
            raise ValueError(f"Could not read file {filepath} using CSV, HTML, or Excel parsers: {e}")

    if df is None or header_idx is None:
        raise KeyError("Could not find the 'Inicio' or 'Início' header row in the file.")

    if header_idx != -1:
        df.columns = df.iloc[header_idx]
        df = df.iloc[header_idx + 1:].reset_index(drop=True)

    df.columns = [str(col).replace('""', '').replace('"', '').strip() for col in df.columns]

    for col in df.columns:
        if 'inicio' in col.lower() or 'início' in col.lower():
            df.rename(columns={col: 'Inicio'}, inplace=True)
            break

    for col in df.columns:
        if 'fim' == col.lower().strip():
            df.rename(columns={col: 'Fim'}, inplace=True)
            break

    return df


def process_isso_dataset(filepath: str, output_folder: str, output_format: str, relative_path: str) -> None:
    """
    Core pipeline to process, clean, and export a single ISSO dataset file.
    """
    filename = os.path.basename(filepath)
    input_ext = os.path.splitext(filepath)[1].lower()
    print(f"Processing {os.path.join(relative_path, filename)}...")

    try:
        # =========================
        # INGESTION
        # =========================
        df = read_dataset_file(filepath)

        # =========================
        # TIMESTAMP FORMATTING
        # =========================
        df["Inicio"] = pd.to_datetime(
            df["Inicio"],
            format="mixed",
            dayfirst=True,
            errors="coerce"
        )

        df["Fim"] = pd.to_datetime(
            df["Fim"],
            format="mixed",
            dayfirst=True,
            errors="coerce"
        )

        df.rename(
            columns={
                "Inicio": "UTC-3 Inicio",
                "Fim": "UTC-3 Fim"
            },
            inplace=True
        )

        # Generate standard UTC columns
        df["UTC Inicio"] = df["UTC-3 Inicio"] + pd.Timedelta(hours=3)
        df["UTC Fim"] = df["UTC-3 Fim"] + pd.Timedelta(hours=3)

        # Reorder columns to prioritize temporal data
        cols = list(df.columns)
        cols.remove("UTC Inicio")
        cols.remove("UTC Fim")

        new_order = (
            ["UTC-3 Inicio", "UTC-3 Fim", "UTC Inicio", "UTC Fim"] +
            [c for c in cols if c not in ["UTC-3 Inicio", "UTC-3 Fim"]]
        )
        df = df[new_order]

        # =========================
        # POWER & ENERGY CALCULATIONS
        # =========================
        direct_col = "Direto: P. Ativa Fund+Harm Total"
        reverse_col = "Reverso: P. Ativa Fund+Harm Total"

        df[direct_col] = pd.to_numeric(df[direct_col], errors="coerce")
        df[reverse_col] = pd.to_numeric(df[reverse_col], errors="coerce")

        # Calculate kW
        df["Direto: P. Ativa Fund+Harm Total (kW)"] = df[direct_col] / 1000
        df["Reverso: P. Ativa Fund+Harm Total (kW)"] = df[reverse_col] / 1000

        # Calculate kWh (Assuming 15-minute / 0.25h intervals)
        df["Direto: Kwh Total"] = (df[direct_col] / 1000) * 0.25
        df["Reverso: Kwh Total"] = (df[reverse_col] / 1000) * 0.25

        # =========================
        # DATASET TRUNCATION
        # =========================
        columns_to_remove = [
            "kWh Direto Fase 1",
            "kWh Direto Fase 2",
            "kWh Direto Fase 3",
            "kWh Direto Total",
            "kWh Reverso Fase 1",
            "kWh Reverso Fase 2",
            "kWh Reverso Fase 3",
            "kWh Reverso Total",
            "Demanda Total",
            "Temp. Interna do Analisador Total",
            "Temp. Interna do Analisador"
        ]

        df.drop(
            columns=[c for c in columns_to_remove if c in df.columns],
            inplace=True
        )

        # =========================
        # FINAL STRING FORMATTING
        # =========================
        for col in ["UTC-3 Inicio", "UTC-3 Fim", "UTC Inicio", "UTC Fim"]:
            df[col] = df[col].dt.strftime("%d/%m/%Y %H:%M:%S")

        # =========================
        # TRANSLATION TO HISTORICAL ENGLISH
        # =========================
        df.rename(columns=COLUMN_TRANSLATIONS, inplace=True)

        # =========================
        # DETERMINE OUTPUT FORMATS
        # =========================
        save_csv = False
        save_excel = False

        if output_format == "both":
            save_csv = True
            save_excel = True
        elif output_format == "csv":
            save_csv = True
        elif output_format == "excel":
            save_excel = True
        elif output_format == "same":
            if input_ext == ".csv":
                save_csv = True
            else:
                save_excel = True

        # =========================
        # FILE EXPORT
        # =========================
        year_month = extract_year_month(filename)
        base_filename = f"{year_month}_Processed_ISSO"

        target_dir = os.path.join(output_folder, relative_path)
        os.makedirs(target_dir, exist_ok=True)

        if save_excel:
            excel_filepath = os.path.join(target_dir, f"{base_filename}.xlsx")
            df.to_excel(excel_filepath, index=False, engine="openpyxl")
            adjust_column_width(excel_filepath)

        if save_csv:
            csv_filepath = os.path.join(target_dir, f"{base_filename}.csv")
            df.to_csv(csv_filepath, index=False, encoding="utf-8-sig")

        print(f"✅ Success -> {os.path.join(relative_path, base_filename)}")

    except Exception as e:
        print(f"❌ Error processing {filename}: {e}")


def main():
    parser = argparse.ArgumentParser(
        description="Process and standardize raw ISSO time-series datasets."
    )

    parser.add_argument(
        "-i", "--input",
        required=True,
        help="Path to the input directory containing raw CSV/XLS files."
    )

    parser.add_argument(
        "-o", "--output",
        required=True,
        help="Path to the output directory to save processed datasets."
    )

    parser.add_argument(
        "-f", "--format",
        choices=["same", "csv", "excel", "both"],
        default="same",
        help="Output format to generate. Default is 'same' as input file."
    )

    args = parser.parse_args()

    input_folder = args.input
    output_folder = args.output
    output_format = args.format

    if not os.path.isdir(input_folder):
        print(f"Error: Input directory '{input_folder}' does not exist.")
        return

    dataset_files = []
    for root, dirs, files in os.walk(input_folder):
        for f in files:
            if f.lower().endswith((".csv", ".xls", ".xlsx")) and not f.startswith("~$"):
                filepath = os.path.join(root, f)
                relative_path = os.path.relpath(root, input_folder)
                if relative_path == ".":
                    relative_path = ""
                dataset_files.append((filepath, relative_path))

    dataset_files = sorted(dataset_files, key=lambda x: x[0])

    if not dataset_files:
        print(f"No dataset files found in '{input_folder}'.")
        return

    print(f"Found {len(dataset_files)} files. Starting processing pipeline...\n")

    for filepath, relative_path in dataset_files:
        process_isso_dataset(filepath, output_folder, output_format, relative_path)

    print("\nProcessing complete.")


if __name__ == "__main__":
    main()