#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
from pathlib import Path

# Define the input file path
input_file_path = Path(r"C:\Users\User\Desktop\example.xlsx")
# Define the output folder path
output_folder_path = Path(r"C:\Users\User\Desktop\")

# Check if the input file exists
if not input_file_path.exists():
    print(f"Error: File not found at {input_file_path}")
else:
    # Read the Excel file
    xlsx = pd.ExcelFile(input_file_path)
    # Create an empty list to store the data
    all_data = []

    # Loop through each sheet in the Excel file
    for sheet_name in xlsx.sheet_names:
        # Read the sheet data into a DataFrame
        df = xlsx.parse(sheet_name)

        # Check if the required column names exist in the first row
        required_columns = ['Departure velocity', 'Time to peak 90', 'Time to baseline 90', 'Peak height', 'Return velocity']
        if all(col in df.columns for col in required_columns):
            # Extract the desired columns based on their column names
            departure_velocity_col = df['Departure velocity']
            time_to_peak_90_col = df['Time to peak 90']
            time_to_baseline_90_col = df['Time to baseline 90']
            peak_height_col = df['Peak height']
            return_velocity_col = df['Return velocity']

            # Add the sheet name as the first row
            sheet_data = [f"Sheet name: {sheet_name}"]
            sheet_data.append("Departure velocity\t" + "\t".join(map(str, departure_velocity_col.tolist())))
            sheet_data.append("Time to peak 90\t" + "\t".join(map(str, time_to_peak_90_col.tolist())))
            sheet_data.append("Time to baseline 90\t" + "\t".join(map(str, time_to_baseline_90_col.tolist())))
            sheet_data.append("Peak height\t" + "\t".join(map(str, peak_height_col.tolist())))
            sheet_data.append("Return velocity\t" + "\t".join(map(str, return_velocity_col.tolist())))

            # Add the sheet data to the list
            all_data.extend(sheet_data)
        else:
            print(f"Warning: Required column names not found in sheet '{sheet_name}'. Skipping this sheet.")

    # Write the result to a new text file
    output_file_path = output_folder_path / "combined_data.txt"
    with open(output_file_path, 'w') as file:
        for data in all_data:
            file.write(data + "\n")


# In[ ]:




