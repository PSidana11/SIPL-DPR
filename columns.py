import pandas as pd 

file_path = "C:/pkn.xlsx"
df = pd.read_excel(file_path, header=[0,1])
print(df.head())   

df.columns = ['_'.join(str(x).strip() for x in col if 'Unnamed' not in str(x)) for col in df.columns.values]

print(df.columns)
# Drop initial rows with NaN values
df = df.dropna(how="all")

# Reset the index
df = df.reset_index(drop=True)

# Display cleaned data
print(df.head())
