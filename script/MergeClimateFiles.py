import pandas as pd
import os
from pathlib import Path

# https://towardsthecloud.com/get-relative-path-python#:~:text=A%20relative%20path%20starts%20with,path%20to%20the%20file%20want.
absolute_path = os.path.dirname(__file__)
relative_path = "../Dataset/archive/World Bank Climate/"
full_path = Path(os.path.join(absolute_path, relative_path))
print(full_path)

if not full_path.exists():
    print(f"The folder '{full_path}' does not exist.")
else:
    merged_data = pd.DataFrame()

    # Loop through each file in the folder
    for file_path in full_path.glob('*.xlsx'):
        # Read all sheets from the current Excel file into a dictionary of DataFrames
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names
        sheet_data = {sheet_name: xls.parse(sheet_name) for sheet_name in sheet_names}

        # Append the data from each sheet to the merged_data DataFrame
        for sheet_name, df in sheet_data.items():
            df['SheetName'] = sheet_name  # Add a column to identify the climate indicator
            merged_data = merged_data.append(df, ignore_index=True)

    merged_file_path = Path(os.path.join(absolute_path, '../Dataset/archive/Merged_climate_data.xlsx'))

    # Save the merged data to a new Excel file
    merged_data.to_excel(merged_file_path, index=False)