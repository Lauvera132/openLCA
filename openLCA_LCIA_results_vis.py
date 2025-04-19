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

