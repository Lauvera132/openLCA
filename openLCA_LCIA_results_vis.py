# Laura Rivera

import pandas as pd
import matplotlib.pyplot as plt

product_system = 'GeoH2 (85% mol H2; H2 Average Fugitive; CH4 Capture)'

# Load the Excel file and specify the worksheet name
lca_results_file_path = r'C:\Users\laura\OneDrive\Documents\UT_graduate\WEG_graduate_research\Hydrogen_fugitive_emissions_natural_h2\openLCA_ecoinvent\Hydrogen_gas__at_processing__production_mixture__to_consumer__kg___US__85__mol_H2____H2_Average___CH4_Capture.xlsx'
direct_impact_sheet_name = 'Direct impact contributions'
impact_by_flow_sheet_name = 'Impact contributions by flow'

# Read the specified worksheet into a DataFrame
df_process_impact = pd.read_excel(lca_results_file_path, sheet_name=direct_impact_sheet_name)
df_flows_impact = pd.read_excel(lca_results_file_path, sheet_name=impact_by_flow_sheet_name)

# Drop the first column if it contains only empty values
if df_process_impact.iloc[:, 0].isnull().all():
    df_process_impact.drop(df_process_impact.columns[0], axis=1, inplace=True)

# Transpose the DataFrame
df_process_impact_transposed = df_process_impact.transpose()

# Reset the index of the transposed DataFrame
df_process_impact_transposed.reset_index(drop=True, inplace=True)

# Find the column index containing the specified string in row index 1
impact_category_name = "climate change - global warming potential (GWP100) (Nitrous Oxide, Methane, Carbon Dioxide)"
print(f"Impact category name: {impact_category_name}")
process_column_index = df_process_impact_transposed.columns[df_process_impact_transposed.iloc[1] == impact_category_name].tolist()


# If the column is found, retrieve its index; otherwise, return None
process_column_index = process_column_index[0] if process_column_index else None


# Create a copy of df_process_impact_transposed with only the specified columns
columns_to_keep = [0, 1, 2, process_column_index]
filtered_df = df_process_impact_transposed.iloc[:, columns_to_keep]

impact_category_uuid = filtered_df.iloc[0, 3]
impact_category_reference_unit = filtered_df.iloc[2, 3]
print(f"Impact category UUID: {impact_category_uuid}")
print(f"Impact category reference unit: {impact_category_reference_unit}")

# Delete rows 0-3 from the filtered DataFrame
filtered_df = filtered_df.iloc[4:].reset_index(drop=True)

# Rename the columns of the filtered DataFrame
filtered_df.columns = ['Process_UUID', 'Process_Name', 'Location', 'GWP100_Impact_kg_CO2eq']
print(filtered_df.head())

# Read the database_processes_summary_added file into a new DataFrame
database_process_summary_file_path = r'C:\Users\laura\OneDrive\Documents\UT_graduate\WEG_graduate_research\Hydrogen_fugitive_emissions_natural_h2\openLCA_ecoinvent\openLCA\database_process_summary_added.xlsx'
database_process_df = pd.read_excel(database_process_summary_file_path)

# Display the first few rows of the new DataFrame
print(database_process_df.head())

# Split column "ISIC_Category" into multiple columns using "/" as the delimiter
split_columns = database_process_df['ISIC_Category'].str.split('/', expand=True)

# Rename the new columns for clarity
split_columns.columns = ['ISIC_Category_Section', 'ISIC_Category_Division', 'ISIC_Category_Group', 'ISIC_Category_Class']

# Add the split columns as additional columns to database_process_df
database_process_df = pd.concat([database_process_df, split_columns], axis=1)

# Merge the filtered DataFrame with the database_process_df DataFrame on the "Process UUID" column
merged_df = pd.merge(filtered_df, database_process_df, on='Process_UUID', how='inner')

# Display the first few rows of the merged DataFrame
print(merged_df.head())

# Check if the "Process_Name" column exists in both DataFrames after the merge
if 'Process_Name_x' in merged_df.columns and 'Process_Name_y' in merged_df.columns:
    # Check if the "Process_Name" columns match in both DataFrames
    if (merged_df['Process_Name_x'] == merged_df['Process_Name_y']).all():
        # Drop one of the "Process_Name" columns and rename the other to "Process_Name"
        merged_df.drop(columns=['Process_Name_y'], inplace=True)
        merged_df.rename(columns={'Process_Name_x': 'Process_Name'}, inplace=True)
    else:
        print("Warning: Process_Name columns do not match completely.")
else:
    print("Error: Process_Name columns are missing in the merged DataFrame.")

# Display the updated DataFrame
print(merged_df.head())

# Convert the GWP100_Impact_kg_CO2eq column to numeric, coercing errors to NaN
merged_df['GWP100_Impact_kg_CO2eq'] = pd.to_numeric(merged_df['GWP100_Impact_kg_CO2eq'], errors='coerce')

# Drop rows with NaN values in the GWP100_Impact_kg_CO2eq column
merged_df.dropna(subset=['GWP100_Impact_kg_CO2eq'], inplace=True)

# Identify the top Process_Name with the largest GWP100_Impact_kg_CO2eq
top_5_processes = merged_df.nlargest(5, 'GWP100_Impact_kg_CO2eq')[['Process_Name', 'GWP100_Impact_kg_CO2eq']]
top_10_processes = merged_df.nlargest(10, 'GWP100_Impact_kg_CO2eq')[['Process_Name', 'GWP100_Impact_kg_CO2eq']]

# Display the top processes
print("Top 5 processes with the largest GWP100 impact:")
print(top_5_processes)
print("Top 10 processes with the largest GWP100 impact:")
print(top_10_processes)

# Group the data by Process_Name and sum the GWP100_Impact_kg_CO2eq for each process
grouped_data = merged_df.groupby('Process_Name')['GWP100_Impact_kg_CO2eq'].sum()

# Separate processes with less than impact threshold and group them into "Other"
threshold = 0.05
grouped_data = grouped_data.sort_values(ascending=False)
other_data = grouped_data[grouped_data < threshold].sum()
grouped_data = grouped_data[grouped_data >= threshold]
grouped_data['Other'] = other_data

##### Plotting #####
# Create a horizontal bar plot

plt.figure(figsize=(12, 5))
plt.style.use('dark_background')

colors = plt.cm.tab20.colors
cumulative_width = 0
bar_height = 0.3

for i, (process_name, value) in enumerate(grouped_data.items()):
    plt.barh(
        y=0,
        width=value,
        left=cumulative_width,
        height=bar_height,
        color=colors[i % len(colors)],
        edgecolor='white',
        linewidth=2,
        label=process_name
    )
    # Only show text if the segment is wide enough
    if value > 0.05:
        plt.text(
            x=cumulative_width + value / 2,
            y=0,
            s=f'{value:.2f}',
            va='center',
            ha='center',
            color='white',
            fontsize=12,
            fontweight='bold'
        )
    cumulative_width += value

# Add total at the end, with padding
plt.text(
    x=cumulative_width + 0.03 * cumulative_width,
    y=0,
    s=f'Total: {cumulative_width:.2f}',
    va='center',
    ha='left',
    color='white',
    fontsize=13,
    fontweight='bold'
)

plt.xlabel('GWP100 Impact (kg CO2eq per kg H2)', fontsize=13, color='white')
plt.title(product_system, fontsize=14, color='white', pad=15, wrap=True)
plt.xticks(color='white', fontsize=11)
plt.yticks([])

plt.xlim(0, cumulative_width * 1.12)

plt.legend(
    loc='upper center',
    bbox_to_anchor=(0.5, -0.2),
    ncol=1,
    fontsize=11,
    frameon=False
)

plt.tight_layout(pad=2)
plt.show()