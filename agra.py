import pandas as pd
import mysql.connector
from tabulate import tabulate

# **Step 1: Load Excel File (Read up to row 34)**
file_path = "C:/agra.xlsx"  # Update the path if needed
xls = pd.ExcelFile(file_path)

# Read only rows 9 to 34 (skip first 8 rows, then take next 26 rows)
df_sheet = xls.parse("sheet", skiprows=8, nrows=25)  # Reads from row 9 to 34

# Rename columns
df_agra = df_sheet.iloc[:, :16]  # Selecting first 16 columns
df_agra.columns = ["item_no", "item_description", "unit", "rate", "boq_qty", "this_month_target",
                   "achieved_upto_previous_month", "achieved_upto_previous_day_this_month",
                   "program_for_today", "today_achieved", "cumulative_achieved_this_month",
                   "percent_achieved_this_month", "cumulative_achieved_upto_date",
                   "percent_cumulative_achieved_upto_date", "tomorrow_program", "remark"]

# **Step 2: Data Cleaning**
df_agra["item_no"] = df_agra["item_no"].astype(str)  # Convert to string (handles numbers & text)
df_agra["item_no"] = df_agra["item_no"].replace("nan", None)  # Replace 'nan' with NULL values

# Handle missing descriptions
df_agra["item_description"] = df_agra["item_description"].fillna("Unknown")

# Convert numeric columns
numeric_cols = ["rate", "boq_qty", "this_month_target", "achieved_upto_previous_month",
                "achieved_upto_previous_day_this_month", "program_for_today", "today_achieved",
                "cumulative_achieved_this_month", "percent_achieved_this_month",
                "cumulative_achieved_upto_date", "percent_cumulative_achieved_upto_date",
                "tomorrow_program"]

df_agra[numeric_cols] = df_agra[numeric_cols].apply(pd.to_numeric, errors="coerce")

# **Step 3: Connect to MySQL**
db_connection = mysql.connector.connect(
    host="localhost",
    user="root",
    password="Hope@123",  # Update password
    database="dpr"
)
cursor = db_connection.cursor()

# **Step 4: Drop and Create Table**
cursor.execute("DROP TABLE IF EXISTS agra")
cursor.execute("""
CREATE TABLE agra (
    id INT AUTO_INCREMENT PRIMARY KEY,
    item_no VARCHAR(50),  -- Supports both numbers and text
    item_description TEXT NOT NULL,
    unit VARCHAR(50),
    rate FLOAT,
    boq_qty FLOAT,
    this_month_target FLOAT,
    achieved_upto_previous_month FLOAT,
    achieved_upto_previous_day_this_month FLOAT,
    program_for_today FLOAT,
    today_achieved FLOAT,
    cumulative_achieved_this_month FLOAT,
    percent_achieved_this_month FLOAT CHECK (percent_achieved_this_month BETWEEN 0 AND 100),
    cumulative_achieved_upto_date FLOAT,
    percent_cumulative_achieved_upto_date FLOAT CHECK (percent_cumulative_achieved_upto_date BETWEEN 0 AND 100),
    tomorrow_program FLOAT,
    remark TEXT
)
""")

# **Step 5: Insert Data into MySQL**
insert_query = """
INSERT INTO agra (item_no, item_description, unit, rate, boq_qty, this_month_target,
                  achieved_upto_previous_month, achieved_upto_previous_day_this_month,
                  program_for_today, today_achieved, cumulative_achieved_this_month,
                  percent_achieved_this_month, cumulative_achieved_upto_date,
                  percent_cumulative_achieved_upto_date, tomorrow_program, remark)
VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
"""

# Convert DataFrame to list of tuples
data = df_agra.values.tolist()

# Insert data
cursor.executemany(insert_query, data)
db_connection.commit()
print("\n✅ Data successfully inserted into 'agra' table!")

# **Step 6: Fetch and Display Data in Tabular Format**
def limit_words(text, word_limit=10):
    """Limit the number of words in a text field."""
    if isinstance(text, str):
        words = text.split()[:word_limit]
        return " ".join(words)
    return text

# Fetch column names dynamically
cursor.execute("SHOW COLUMNS FROM agra")
columns = [column[0] for column in cursor.fetchall()]

# Fetch all data
cursor.execute("SELECT * FROM agra LIMIT 20")  # Fetch first 20 rows
rows = cursor.fetchall()

# Convert to DataFrame
df_output = pd.DataFrame(rows, columns=columns)

# **Step 7: Format for Better Readability**
# Limit item_description to 10 words
df_output["item_description"] = df_output["item_description"].apply(lambda x: limit_words(x, 10))

# Set column width limit for better display
pd.set_option("display.max_colwidth", 30)  # Prevents column wrapping
pd.set_option("display.width", 1000)  # Ensures proper table width

# Print formatted table
print("\n🔹 Data in Tabular Format:")
print(tabulate(df_output, headers="keys", tablefmt="fancy_grid", showindex=False))

# **Step 8: Close Database Connection**
cursor.close()
db_connection.close()
