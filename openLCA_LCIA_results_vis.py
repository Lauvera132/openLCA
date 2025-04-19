# Laura Rivera

import pandas as pd

# Load the Excel file and specify the worksheet name
file_path = r'C:\Users\laura\OneDrive\Documents\UT_graduate\WEG_graduate_research\Hydrogen_fugitive_emissions_natural_h2\openLCA_ecoinvent\Hydrogen_gas__at_processing__production_mixture__to_consumer__kg___US__85__mol_H2____H2_Average___CH4_Capture.xlsx'
sheet_name = 'Direct impact contributions'

# Read the specified worksheet into a DataFrame
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Drop the first column if it contains only empty values
if df.iloc[:, 0].isnull().all():
    df.drop(df.columns[0], axis=1, inplace=True)

# Transpose the DataFrame
df_transposed = df.transpose()

# Reset the index of the transposed DataFrame
df_transposed.reset_index(drop=True, inplace=True)

# Find the column index containing the specified string in row index 1
target_string = "climate change - global warming potential (GWP100) (Nitrous Oxide, Methane, Carbon Dioxide)"
column_index = df_transposed.columns[df_transposed.iloc[1] == target_string].tolist()

# If the column is found, retrieve its index; otherwise, return None
column_index = column_index[0] if column_index else None

# Create a copy of df_transposed with only the specified columns
columns_to_keep = [0, 1, 2, column_index]
filtered_df = df_transposed.iloc[:, columns_to_keep]

