# Reads workbook path from config.py, which is written to via main script (build_roomv[x].py)

import os
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
from config import workbook_path  # Import the path from config file

# Use the imported workbook_path directly
file_path = workbook_path

# Load the spreadsheet
xl = pd.ExcelFile(file_path)

# Load sheets into DataFrames by name
df_42_items = xl.parse('4.2_Items')
df_br_items = xl.parse('BR_Items')

# Replace NaN values with 0 in 'NewCount' column for both dataframes
df_42_items['NewCount'].fillna(0, inplace=True)
df_br_items['NewCount'].fillna(0, inplace=True)

# Combine the two dataframes
combined_df = pd.concat([df_42_items, df_br_items])

# Group by 'Item' and sum the 'NewCount' values
grouped_df = combined_df.groupby('Item')['NewCount'].sum().reset_index()

# Create a horizontal bar chart for the summed inventory levels
plt.figure(figsize=(14 * 0.60, 10 * 0.60))
bars = plt.barh(grouped_df['Item'], grouped_df['NewCount'], color='#006aff')

# Define the spacing for the text
spacing = 1  # Adjust this value for more or less spacing

# Add the text with the summed count at the end of each bar
for bar in bars:
    width = bar.get_width()
    plt.text(width + spacing, bar.get_y() + bar.get_height()/2,
             f'{int(width)}', ha='left', va='center', color='black')

plt.ylabel('Item', fontsize=12)
plt.xlabel('Volume', fontsize=12)
plt.xlim(0, 120)
current_date = datetime.now().strftime('%d-%m-%Y')
plt.title(f'Total (combined) Inventory Levels (Perth) - {current_date}', fontsize=14)
plt.legend()
plt.tight_layout()

# Ensure 'Plots' folder exists
plots_folder = os.path.join(os.path.dirname(workbook_path), 'Plots')
if not os.path.exists(plots_folder):
    os.makedirs(plots_folder)

# Get current date and time for file name
current_datetime = datetime.now().strftime('%d.%m.%y-%H.%M.%S')

# Construct the full file path for saving the plot
file_name = os.path.join(plots_folder, f'combined_inventory_levels_{current_datetime}.png')

# Save and show the plot
plt.savefig(file_name)
plt.show()