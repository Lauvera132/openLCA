# Laura Rivera

import pandas as pd

product_system = 'Hydrogen_gas__at_processing__production_mixture__to_consumer__kg___US__85__mol_H2____H2_Average___CH4_Capture'

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
filtered_df.columns = ['Process_UUID', 'Process_Name', 'Location', 'GWP100_Impact_kg_CO2eq']
print(filtered_df.head())

# Read the database_processes_summary_added file into a new DataFrame
excel_file_path = r'C:\Users\laura\OneDrive\Documents\UT_graduate\WEG_graduate_research\Hydrogen_fugitive_emissions_natural_h2\openLCA_ecoinvent\openLCA\database_process_summary_added.xlsx'
process_df = pd.read_excel(excel_file_path)

# Display the first few rows of the new DataFrame
print(process_df.head())

# Split column "ISIC_Category" into multiple columns using "/" as the delimiter
split_columns = process_df['ISIC_Category'].str.split('/', expand=True)

# Rename the new columns for clarity
split_columns.columns = ['ISIC_Category_Section', 'ISIC_Category_Division', 'ISIC_Category_Group', 'ISIC_Category_Class']

# Add the split columns as additional columns to process_df
process_df = pd.concat([process_df, split_columns], axis=1)

# Merge the filtered DataFrame with the process_df DataFrame on the "Process UUID" column
merged_df = pd.merge(filtered_df, process_df, on='Process_UUID', how='inner')

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

import matplotlib.pyplot as plt
# Group the data by Process_Name and sum the GWP100_Impact_kg_CO2eq for each process
grouped_data = merged_df.groupby('Process_Name')['GWP100_Impact_kg_CO2eq'].sum()

# Separate processes with less than 0.01 impact and group them into "Other"
threshold = 0.05
grouped_data = grouped_data.sort_values(ascending=False)
other_data = grouped_data[grouped_data < threshold].sum()
grouped_data = grouped_data[grouped_data >= threshold]
grouped_data['Other'] = other_data

# Plot the results as a single stacked bar chart with a black background and white lines and font
plt.figure(figsize=(10, 6))
plt.style.use('dark_background')  # Set the background to black

# Create a single stacked bar chart
colors = plt.cm.tab20.colors  # Use a colormap for distinct colors
plt.bar(
    x=[product_system], 
    height=[grouped_data.sum()], 
    color=colors[:len(grouped_data)], 
    edgecolor='white'
)

# Add individual contributions as stacked segments
bottom = 0
for i, (process_name, value) in enumerate(grouped_data.items()):
    plt.bar(
        x=[product_system], 
        height=[value], 
        bottom=[bottom], 
        color=colors[i % len(colors)], 
        edgecolor='white', 
        label=process_name[:25]  # Limit legend to 25 characters
    )
    # Add impact numbers to the center of each segment
    plt.text(
        x=0, 
        y=bottom + value / 2, 
        s=f'{value:.2f}', 
        ha='center',  # Center horizontally
        va='center',  # Center vertically
        color='white', 
        fontsize=10
    )
    bottom += value

# Add total value on top of the bar
total_value = grouped_data.sum()
plt.text(0, total_value + 0.01, f'{total_value:.2f}', ha='center', va='bottom', color='white', fontsize=12)

# Add labels, title, and legend
plt.xlabel('', fontsize=12, color='white')
plt.ylabel('GWP100 Impact (kg CO2eq)', fontsize=12, color='white')
plt.title('GWP100 Impact by Process Name', fontsize=14, color='white')
plt.xticks(color='white')
plt.yticks(color='white')
plt.legend(
    title='Process Name', 
    fontsize=10, 
    title_fontsize=12, 
    loc='center left', 
    bbox_to_anchor=(1, 0.5),  # Move the legend box to the right of the chart
    facecolor='black', 
    edgecolor='white'
)
plt.tight_layout()

# Show the plot
plt.show()
