import pandas as pd

# Load the Excel file
file_path = "C:/DPR-04_April'25_.xlsx"
xls = pd.ExcelFile(file_path)

# Print available sheet names
print("Available sheets:", xls.sheet_names)

df=xls.parse("DPR Daily ", skiprows=10, nrows=19)
df_bilaspur = df.iloc[:, :12]
