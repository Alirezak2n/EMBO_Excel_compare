import pandas as pd
from tqdm import tqdm
from openpyxl import Workbook
from openpyxl.styles import Font, Color, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import os
import glob

# Define the path to the original Excel file
directory_path = 'C:/Users/Alireza/Downloads/Source data files with issues for Alireza/Source data files with issues for Alireza'
output_directory = os.path.join(directory_path, 'duplicates')

# Ensure the output directory exists
if not os.path.exists(output_directory):
    os.makedirs(output_directory)
# Load the Excel file
excel_files = glob.glob(os.path.join(directory_path, '*.xlsx'))


# Process each file
# Process each file with a progress bar
for file_path in tqdm(excel_files, desc="Processing files"):
    tqdm.write(f"Reading file: {file_path}")
    # Load all sheets from the Excel file
    xls = pd.ExcelFile(file_path)
    workbook_modified = False
    wb = Workbook()
    wb.remove(wb.active)  # Remove the default sheet

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)

        # Flatten the DataFrame to check duplicates across all values
        values = pd.Series(df.values.ravel())
        duplicated_values = values[values.duplicated(keep=False)].unique()  # Get unique duplicated values
        # Continue only if there are duplicated values
        if len(duplicated_values)>1:
            workbook_modified = True
            ws = wb.create_sheet(title=sheet_name)

            # Convert the DataFrame to rows in Excel, including headers and index
            for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
                ws.append(row)
                if r_idx > 1:  # Skip the header row for formatting
                    # Iterate over each cell in the current row
                    for c_idx, cell in enumerate(ws[r_idx]):
                        # If the cell's value is in the list of duplicated values, highlight it
                        if cell.value in duplicated_values:
                            cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # Save the new formatted workbook if modifications were made
    if workbook_modified:
        base_name = os.path.basename(file_path)
        new_base_name = os.path.splitext(base_name)[0] + '_duplicatecheck' + os.path.splitext(base_name)[1]
        new_file_path = os.path.join(output_directory, new_base_name)
        wb.save(new_file_path)