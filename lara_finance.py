import pandas as pd
import mysql.connector

# Configurations
file_path = "C:/lara.xlsx"
sheet_name = "sheet"  # Update if needed
db_config = {
    "host": "localhost",
    "user": "root",
    "password": "Hope@123",
    "database": "dpr"
}

try:
    # Load Excel file
    xls = pd.ExcelFile(file_path)
    df = xls.parse(sheet_name)

    # 🔍 Debug: Print all available columns
    print("\n🔍 Available Columns:", df.columns.tolist())

    # Extract data from **rows 32 to 39** (Excel rows are 1-based, Pandas is 0-based)
    df_finance = df.iloc[31:39, [0] + list(range(2, 9))]  # Column 0 for names, 2-8 for values

    # Rename first column to "metric_name"
    df_finance.rename(columns={df_finance.columns[0]: "metric_name"}, inplace=True)

    # 🔍 Debug: Show extracted data before cleaning
    print("\n🔍 Extracted Data Before Cleaning:\n", df_finance)

    # Convert wide format to long format (removing source_column)
    df_finance = df_finance.set_index("metric_name").stack().reset_index(level=1, drop=True).reset_index()
    df_finance.columns = ["metric_name", "metric_value"]  # Keeping only two columns

    # Remove empty rows
    df_finance = df_finance.dropna(subset=["metric_value"])

    # Remove commas, percentage symbols, and ensure numeric conversion
    df_finance["metric_value"] = (
        df_finance["metric_value"]
        .astype(str)
        .str.replace(",", "", regex=True)
        .str.replace("%", "", regex=True)
    )
    df_finance["metric_value"] = pd.to_numeric(df_finance["metric_value"], errors="coerce").fillna(0)

    # 🚀 **Fix: Remove duplicate metric names**
    df_finance = df_finance.drop_duplicates(subset=["metric_name"])

    # 🔍 Debug: Print cleaned data
    print("\n✅ Final Cleaned Data (Before Insertion):\n", df_finance)

    # Connect to MySQL
    db_connection = mysql.connector.connect(**db_config)
    cursor = db_connection.cursor()

    # Drop and create table
    cursor.execute("DROP TABLE IF EXISTS lara_finance")
    cursor.execute("""
        CREATE TABLE lara_finance (
            id INT AUTO_INCREMENT PRIMARY KEY,
            metric_name VARCHAR(255),
            metric_value DECIMAL(20, 2)
        )
    """)

    # Insert cleaned data into MySQL table
    insert_query = "INSERT INTO lara_finance (metric_name, metric_value) VALUES (%s, %s)"
    cursor.executemany(insert_query, df_finance[["metric_name", "metric_value"]].values.tolist())
    db_connection.commit()

    print("\n✅ Data successfully inserted into 'lara_finance' table!")

except mysql.connector.Error as err:
    print(f"\n❌ MySQL Error: {err}")

except Exception as e:
    print(f"\n❌ Error: {e}")

finally:
    # Close connection if established
    if 'cursor' in locals():
        cursor.close()
    if 'db_connection' in locals():
        db_connection.close()
