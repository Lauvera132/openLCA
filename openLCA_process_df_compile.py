# Laura Rivera

import pandas as pd
import os
# Define the folder path containing the Excel files
folder_path = r'C:\Users\laura\OneDrive\Documents\UT_graduate\WEG_graduate_research\Hydrogen_fugitive_emissions_natural_h2\openLCA_ecoinvent\ecoinvent_391_esther_v2_laura_v1_process_export'

# Initialize an empty list to store data from each file
data = []

# Initialize a counter for the number of files processed
file_counter = 0

# Iterate through all .xlsx files in the folder
for file_name in os.listdir(folder_path):
    if file_name.endswith('.xlsx'):
        file_path = os.path.join(folder_path, file_name)
        
        # Read the "General information" worksheet
        try:
            general_info_df = pd.read_excel(file_path, sheet_name='General information', header=None)
            
            # Extract UUID, name, and category
            uuid = general_info_df.iloc[1, 1]
            name = general_info_df.iloc[2, 1]
            category = general_info_df.iloc[3, 1]
            
            # Append the extracted data to the list
            data.append({'Process_UUID': uuid, 'Process_Name': name, 'ISIC_Category': category})
            
            # Increment the file counter
            file_counter += 1
            print("Number of files processed:", file_counter)
        except Exception as e:
            print(f"Error reading file {file_name}: {e}")

# Print the total number of files processed
print(f"Total files processed: {file_counter}")

# Create a DataFrame from the collected data
processes_df = pd.DataFrame(data)

# Display the first few rows of the DataFrame
print(processes_df.head())

# Save the DataFrame to a CSV file
output_file = r'C:\Users\laura\OneDrive\Documents\UT_graduate\WEG_graduate_research\Hydrogen_fugitive_emissions_natural_h2\openLCA_ecoinvent\openLCA\database_process_summary.csv'
processes_df.to_csv(output_file, index=False)
print(f"DataFrame saved to {output_file} CSV successfully.")

# Save the DataFrame to an Excel file and rename the file
excel_file = r'C:\Users\laura\OneDrive\Documents\UT_graduate\WEG_graduate_research\Hydrogen_fugitive_emissions_natural_h2\openLCA_ecoinvent\openLCA\database_process_summary_added.xlsx'
print(f"Saving DataFrame to Excel file: {excel_file}")
processes_df.to_excel(excel_file, index=False)
print(f"DataFrame saved to {excel_file} successfully.")
print("Please open the Excel file and review the 'ISIC_Category' column for any blank entries.")
print("If blank entries are found, refer to the 'International Standard Industrial Classification of All Economic Activities (ISIC), Rev. 4' to assign an appropriate ISIC Category for the corresponding unit process.")
print("The 'ISIC_Category' uses a four-level hierarchical structure - Section (letter), Division (2 digits), Group (3 digits), and Class (4 digits).")
print("The 'ISIC_Category' must be entered as a string with the format 'Section/Division/Group/Class'.")
print("Please save the Excel file after making any changes.")
