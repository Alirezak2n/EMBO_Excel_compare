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

# Function to interpolate colors from red to yellow
def interpolate_colors(num_colors):
    colors = []
    for i in range(num_colors):
        # Linear interpolation for the green channel
        green = int(255 * (i / (num_colors - 1)))  # Increment green channel linearly
        colors.append(f"FF{green:02X}00")  # Format as HEX, keeping red at 255 and blue at 0
    return colors

def extract_first_five_decimals(value):
    try:
        parts = str(value).split('.')
        if len(parts) > 1 and len(parts[1]) > 4:  # Check if there's a decimal component
            return int(parts[1][:5])  # Return only up to five digits after the decimal
    except:
        return None

# Process each file
# Process each file with a progress bar
for file_path in tqdm(excel_files, desc="Processing files"):
    tqdm.write(f"Reading file: {file_path}")
    # Load all sheets from the Excel file
    xls = pd.ExcelFile(file_path)
    workbook_modified = False
    decimal_duplication_found = False
    wb = Workbook()
    wb.remove(wb.active)  # Remove the default sheet

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)

        # Flatten the DataFrame to check duplicates across all values
        values = pd.Series(df.values.ravel())
        duplicated_values = values[values.duplicated(keep=False)].unique()  # Get unique duplicated values

        # Prepare to check for decimal similarities
        decimal_occurrences = {}
        for val in df.values.ravel():
            first_five = extract_first_five_decimals(val)
            if first_five is not None:  # Only consider values with the required decimal part
                if first_five in decimal_occurrences:
                    decimal_occurrences[first_five].append(val)
                else:
                    decimal_occurrences[first_five] = [val]

        # Continue only if there are duplicated values or decimal matches
                # Generate color gradient
        num_groups = len((decimal_occurrences ))
        # print(len(set(tuple(vals) for vals in decimal_occurrences.values() if len(vals) > 1)))
        if num_groups > 0:
            color_gradient = interpolate_colors(num_groups)
            color_mapping = {key: color for key, color in zip(decimal_occurrences, color_gradient)}
            if len(duplicated_values) > 1 or any(len(vals) > 1 for vals in decimal_occurrences.values()):
                workbook_modified = True
                if any(len(vals) > 1 for vals in decimal_occurrences.values()):
                    decimal_duplication_found = True  # Flag that there's a decimal duplication
                ws = wb.create_sheet(title=sheet_name)

                # Convert the DataFrame to rows in Excel, including headers and index
                for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
                    ws.append(row)
                    if r_idx > 1:  # Skip the header row for formatting
                        # Iterate over each cell in the current row
                        for c_idx, cell in enumerate(ws[r_idx]):
                            # Highlight cells with duplicated values
                            if cell.value in duplicated_values:
                                cell.fill = PatternFill(start_color="add8e6", end_color="add8e6", fill_type="solid")
                            # Highlight cells with matching first five decimal digits
                            first_five = extract_first_five_decimals(cell.value)
                            if first_five is not None and len(decimal_occurrences[first_five]) > 1:
                                fill_color = color_mapping[first_five]
                                cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")

    # Save the new formatted workbook if modifications were made
    if workbook_modified:
        base_name = os.path.basename(file_path)
        new_base_name = os.path.splitext(base_name)[0] + ('_duplicateDecimal' if decimal_duplication_found else '_duplicateCell') + os.path.splitext(base_name)[1]

        new_file_path = os.path.join(output_directory, new_base_name)
        wb.save(new_file_path)