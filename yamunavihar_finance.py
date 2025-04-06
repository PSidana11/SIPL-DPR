import pandas as pd
import mysql.connector

# Configurations
from datetime import datetime
today = datetime.now().strftime("%Y-%m-%d")
file_path = fr"C:\Users\E01412\Desktop\REPORTS\DAILY\yamunavihar_{today}.xlsx"
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
    df_finance = df.iloc[34:41, :7]  # Column 0 for names, 2-8 for values

    # Rename first column to "metric_name"
    df_finance.rename(columns={df_finance.columns[0]: "metric_name"}, inplace=True)

    # 🔍 Debug: Show extracted data before cleaning
    print("\n🔍 Extracted Data Before Cleaning:\n", df_finance)


    df_finance = df_finance.melt(id_vars=["metric_name"], var_name="source_column", value_name="metric_value")

    # Keep only required columns
    df_finance = df_finance[["metric_name", "metric_value"]]

    # Convert 'metric_name' to string to avoid tuple issues
    df_finance["metric_name"] = df_finance["metric_name"].astype(str)





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
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS yamunavihar_finance (
            id INT AUTO_INCREMENT PRIMARY KEY,
            metric_name VARCHAR(255),
            metric_value DECIMAL(20, 2),
            entry_date DATE
)
    """)

    # Insert cleaned data into MySQL table
    insert_query = "INSERT INTO yamunavihar_finance (metric_name, metric_value, entry_date) VALUES (%s, %s, CURDATE())"
    cursor.executemany(insert_query, df_finance[["metric_name", "metric_value"]].values.tolist())
    db_connection.commit()

    print("\n✅ Data successfully inserted into 'yamuna' table!")

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
