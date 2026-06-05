import glob
import os
import pandas as pd


def process_dynamic_sheets(
    input_folder,
    reference_template_path,
    output_folder,
    output_prefix="consolidated",
):
    # 1. Dynamically discover sheets from the Reference Template
    try:
        ref_excel = pd.ExcelFile(reference_template_path)
        target_sheets = ref_excel.sheet_names
        print(f"Found {len(target_sheets)} sheets in reference: {target_sheets}")
    except Exception as e:
        print(f"Error reading reference template: {e}")
        return

    # Load Reference Headers dynamically for whatever sheets were found
    ref_headers = {}
    for sheet in target_sheets:
        ref_df = pd.read_excel(ref_excel, sheet_name=sheet)
        ref_headers[sheet] = list(ref_df.columns)

    # Initialize data storage for each discovered sheet
    # e.g., {'SheetA': [], 'SheetB': [], ...}
    consolidated_data = {sheet: [] for sheet in target_sheets}

    # 2. Get all files in directory
    files = glob.glob(os.path.join(input_folder, "*.xlsx"))

    for file in files:
        # Ignore reference file and previous outputs
        if (
            os.path.abspath(file) == os.path.abspath(reference_template_path)
            or output_prefix in file
        ):
            continue

        print(f"\nProcessing file: {os.path.basename(file)}")

        try:
            input_excel = pd.ExcelFile(file)
            file_sheets = input_excel.sheet_names
            # 3. Process sheets dynamically
            for sheet in target_sheets:
                # Check if the current target sheet exists in this input file
                if sheet not in file_sheets:
                    print(f"  -> Skipping sheet '{sheet}': Not found in file.")
                    continue

                df = pd.read_excel(input_excel, sheet_name=sheet)

                # Robust Header Validation & Aligning for this specific sheet
                if not set(ref_headers[sheet]).issubset(set(df.columns)):
                    print(
                        f"  -> Skipping sheet '{sheet}': Missing required headers."
                    )
                    continue

                # Reorder columns to match the reference template sheet exactly
                df = df[ref_headers[sheet]]
                consolidated_data[sheet].append(df)
                print(f"  -> Successfully validated sheet: '{sheet}'")

        except Exception as e:
            print(f"  -> Could not read file {os.path.basename(file)}. Error: {e}")

    # 4. Data Consolidation and Saving
    os.makedirs(output_folder, exist_ok=True)
    max_rows = 1000000

    print("\n--- Starting Consolidation ---")
    for sheet in target_sheets:
        sheet_dfs = consolidated_data[sheet]

        if not sheet_dfs:
            print(f"No valid data collected for sheet: '{sheet}'")
            continue

        # Combine all dataframes collected for this specific sheet name
        combined_df = pd.concat(sheet_dfs, ignore_index=True)
        print(f"Consolidating '{sheet}'... Total rows: {len(combined_df)}")

        # 5. Split and Save into separate files using the dynamic sheet name
        file_count = 1
        for i in range(0, len(combined_df), max_rows):
            chunk = combined_df.iloc[i : i + max_rows]

            # The output filename now updates using the dynamic sheet variable
            output_filename = os.path.join(
                output_folder, f"{output_prefix}_{sheet}_{file_count}.xlsx"
            )

            chunk.to_excel(output_filename, index=False)
            print(f"  Saved: {output_filename} ({len(chunk)} rows)")
            file_count += 1


# Execution
process_dynamic_sheets(
    input_folder=r"Templates",
    reference_template_path=r"Reference.xlsx",
    output_folder=r"Consolidated_Output",
)