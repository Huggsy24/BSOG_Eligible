import tkinter as tk
from tkinter import filedialog
import pandas as pd

# Create the Tkinter root window
root = tk.Tk()
root.withdraw()

# Prompt the user to select a text file
file_path = filedialog.askopenfilename()

# Read the text file with desired import options
data_frame = pd.read_csv(file_path, delimiter='\t', header=None, encoding='utf-8')

# Extract Distant 
data_frame[1] = data_frame[0].str[94:102]

# BSOG Eligbility F or N
data_frame[2] = data_frame[0].str[115]

# Extract Days of the Week
data_frame[3] = data_frame[0].str[85:94]

# Calculate the sums
total_sum = data_frame[1].astype(float).sum()
n_sum = data_frame.loc[data_frame[2] == 'N', 1].astype(float).sum()
f_sum = data_frame.loc[data_frame[2] == 'F', 1].astype(float).sum()

# Create a new DataFrame for the results
results_df = pd.DataFrame({
    "Measurement": ["Total Mileage", "Total Non-BSOG Mileage", "Total BSOG Mileage"],
    "Monday": [total_sum if 'M' in data_frame.loc[0, 3] else None, n_sum if 'M' in data_frame.loc[0, 3] else None, f_sum if 'M' in data_frame.loc[0, 3] else None],
    "Tuesday": [total_sum if 'T' in data_frame.loc[0, 3] else None, n_sum if 'T' in data_frame.loc[0, 3] else None, f_sum if 'T' in data_frame.loc[0, 3] else None],
    "Wednesday": [total_sum if 'W' in data_frame.loc[0, 3] else None, n_sum if 'W' in data_frame.loc[0, 3] else None, f_sum if 'W' in data_frame.loc[0, 3] else None],
    "Thursday": [total_sum if 'H' in data_frame.loc[0, 3] else None, n_sum if 'H' in data_frame.loc[0, 3] else None, f_sum if 'H' in data_frame.loc[0, 3] else None],
    "Friday": [total_sum if 'F' in data_frame.loc[0, 3] else None, n_sum if 'F' in data_frame.loc[0, 3] else None, f_sum if 'F' in data_frame.loc[0, 3] else None],
    "Saturday": [total_sum if 'S' in data_frame.loc[0, 3] else None, n_sum if 'S' in data_frame.loc[0, 3] else None, f_sum if 'S' in data_frame.loc[0, 3] else None],
    "Sunday": [total_sum if '$' in data_frame.loc[0, 3] else None, n_sum if '$' in data_frame.loc[0, 3] else None, f_sum if '$' in data_frame.loc[0, 3] else None]
})

# Add a total column to the results DataFrame
results_df["Total"] = results_df.iloc[:, 1:].sum(axis=1)

# Export the results DataFrame to an Excel file
export_path = filedialog.asksaveasfilename(filetypes=[("Excel files", "*.xlsx")])
results_df.to_excel(export_path, index=False)

# Print the first 10 rows of the data frame
print("# First 10 rows of the data frame:")
print(data_frame.head(10))

# Print the sums
print("Total Mileage:", total_sum)
print("Total Non-BSOG Mileage:", n_sum)
print("Total BSOG Mileage:", f_sum)
