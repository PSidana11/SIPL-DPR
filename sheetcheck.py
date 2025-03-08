import pandas as pd

# Load the Excel file
file_path = "C:/sitamarhi.xlsx"
xls = pd.ExcelFile(file_path)

# Print available sheet names
print("Available sheets:", xls.sheet_names)
