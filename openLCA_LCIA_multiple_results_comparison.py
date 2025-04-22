import pandas as pd

# Define file paths and sheet names
file_paths = [
    r"C:\Users\laura\OneDrive\Documents\UT_graduate\WEG_graduate_research\Hydrogen_fugitive_emissions_natural_h2\openLCA_ecoinvent\GeoH2_85molH2_H2_AverageFugitive_CH4Capture_data_export.xlsx",
    r"C:\Users\laura\OneDrive\Documents\UT_graduate\WEG_graduate_research\Hydrogen_fugitive_emissions_natural_h2\openLCA_ecoinvent\GeoH2_85molH2_H2_AverageFugitive_CH4Flare_data_export.xlsx",
    r"C:\Users\laura\OneDrive\Documents\UT_graduate\WEG_graduate_research\Hydrogen_fugitive_emissions_natural_h2\openLCA_ecoinvent\GeoH2_85molH2_H2_AverageFugitive_CH4Vent_data_export.xlsx"
]

sheet_names = ["Grouped_Process_Other", "Grouped_ISIC_Other", "Grouped_Flows_Outher"]

# Create dataframes for each sheet name in the first file
dataframes = {}
for sheet_name in sheet_names:
    try:
        df = pd.read_excel(file_paths[0], sheet_name=sheet_name)
        dataframes[sheet_name] = df
        print(f"Dataframe for {sheet_name} created successfully.")
    except Exception as e:
        print(f"Error reading {sheet_name} from {file_paths[0]}: {e}")

# Dictionary to store merged dataframes for each sheet name
merged_dataframes = {}

# Process each sheet name separately
for sheet_name in sheet_names:
    sheet_dataframes = []  # Temporary list to store dataframes for the current sheet name
    for file_path in file_paths:
        try:
            # Read the specific sheet from the current file
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            df['Source_File'] = file_path.split("\\")[-1]  # Add a column to track the source file
            df['Sheet_Name'] = sheet_name  # Add a column to track the sheet name
            sheet_dataframes.append(df)
        except Exception as e:
            print(f"Error reading {sheet_name} from {file_path}: {e}")
    
    # Merge all dataframes for the current sheet name
    if sheet_dataframes:
        merged_dataframes[sheet_name] = pd.concat(sheet_dataframes, ignore_index=True)

# Save each merged dataframe as a separate worksheet in a single Excel file
output_path = r"C:\Users\laura\OneDrive\Documents\UT_graduate\WEG_graduate_research\Hydrogen_fugitive_emissions_natural_h2\openLCA_ecoinvent\Merged_Data.xlsx"
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    for sheet_name, merged_df in merged_dataframes.items():
        merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"Merged data for {sheet_name} added as a worksheet.")
print(f"All merged data saved to {output_path}")

import matplotlib.pyplot as plt

# Check if the merged dataframe for "Grouped_Process_Other" exists
if "Grouped_Process_Other" in merged_dataframes:
    grouped_process_df = merged_dataframes["Grouped_Process_Other"]
    
    # Example plot: Count of rows per Source_File
    source_file_counts = grouped_process_df['Source_File'].value_counts()
    
    plt.figure(figsize=(10, 6))
    source_file_counts.plot(kind='bar', color='skyblue')
    plt.title("Row Counts per Source File for Grouped_Process_Other")
    plt.xlabel("Source File")
    plt.ylabel("Row Count")
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.show()
else:
    print("No merged data available for Grouped_Process_Other.")