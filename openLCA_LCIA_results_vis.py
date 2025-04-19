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
impact_category_name = "climate change - global warming potential (GWP100) (Nitrous Oxide, Methane, Carbon Dioxide)"
print(f"Impact category name: {impact_category_name}")
column_index = df_transposed.columns[df_transposed.iloc[1] == impact_category_name].tolist()

# If the column is found, retrieve its index; otherwise, return None
column_index = column_index[0] if column_index else None

# Create a copy of df_transposed with only the specified columns
columns_to_keep = [0, 1, 2, column_index]
filtered_df = df_transposed.iloc[:, columns_to_keep]

impact_category_uuid = filtered_df.iloc[0, 3]
impact_category_reference_unit = filtered_df.iloc[2, 3]
print(f"Impact category UUID: {impact_category_uuid}")
print(f"Impact category reference unit: {impact_category_reference_unit}")

# Delete rows 0-3 from the filtered DataFrame
filtered_df = filtered_df.iloc[4:].reset_index(drop=True)

# Rename the columns of the filtered DataFrame
filtered_df.columns = ['Process UUID', 'Process', 'Location', 'GWP100_Impact_kg_CO2eq']
print(filtered_df.head())

