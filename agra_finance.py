import pandas as pd
import mysql.connector

# Load Excel file
file_path = "C:/agra.xlsx"
xls = pd.ExcelFile(file_path)
df = xls.parse("sheet")

# 🔍 Debug: Print all available columns
print("\n🔍 Available Columns:", df.columns.tolist())

# Extract data from C to I (Excel columns 3 to 9 → Python index 2 to 8)
df_finance = df.iloc[35:43, 2:9]

# 🔍 Debug: Show extracted data
print("\n🔍 Extracted Data Before Cleaning:\n", df_finance)

# Flatten numeric columns into a single column
df_finance = df_finance.melt(ignore_index=False, var_name="source_column", value_name="metric_value")

# Drop empty rows
df_finance = df_finance.dropna(subset=["metric_value"])

# Add metric names from column C
df_finance["metric_name"] = df.iloc[df_finance.index, 0]  # Pick names from column C (Index 0)

# Reset index
df_finance = df_finance[["metric_name", "metric_value"]].reset_index(drop=True)

# Remove commas and percentage symbols, then convert to numeric
df_finance["metric_value"] = df_finance["metric_value"].astype(str).str.replace(",", "", regex=True).str.replace("%", "", regex=True)
df_finance["metric_value"] = pd.to_numeric(df_finance["metric_value"], errors="coerce")

# 🔍 Debug: Print cleaned data
print("\n✅ Final Cleaned Data (Before Insertion):\n", df_finance)

# Connect to MySQL
db_connection = mysql.connector.connect(
    host="localhost",
    user="root",    # Your MySQL username
    password="Hope@123",  # Your MySQL password
    database="dpr"
)
cursor = db_connection.cursor()

# Drop and create a table
cursor.execute("DROP TABLE IF EXISTS agra_finance")
cursor.execute("""
    CREATE TABLE agra_finance (
        id INT AUTO_INCREMENT PRIMARY KEY,
        metric_name VARCHAR(100),
        metric_value FLOAT
    )
""")

# Insert data into table
insert_query = "INSERT INTO agra_finance (metric_name, metric_value) VALUES (%s, %s)"
cursor.executemany(insert_query, df_finance.values.tolist())
db_connection.commit()

print("\n✅ Data successfully inserted into 'agra_finance' table!")

# Close connection
cursor.close()
db_connection.close()
