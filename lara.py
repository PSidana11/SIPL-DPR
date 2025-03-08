import pandas as pd
import mysql.connector

# **Step 1: Load Excel File**
file_path = "C:/lara.xlsx"
xls = pd.ExcelFile(file_path)

# Read Excel sheet
df_sheet = xls.parse("sheet", skiprows=9, nrows=20)

# Rename columns
df_lara = df_sheet.iloc[:, :17]
df_lara.columns = [
    "item_no", "item_description", "unit", "total_boq_quantity", "this_month_target", "rate",
    "amount_target_this_month", "achieved_upto_previous_month", "achieved_upto_previous_day_this_month",
    "program_for_today", "today_achieved", "cumulative_achieved_this_month", "percent_achieved_this_month",
    "cumulative_achieved_upto_date", "percent_cumulative_achieved_upto_date", "tomorrow_program", "remark"
]

# **Step 2: Data Cleaning**
df_lara["item_no"] = df_lara["item_no"].astype(str)
df_lara["item_no"] = df_lara["item_no"].replace(["nan", "None"], None)  # Replace invalid values

# Handle missing descriptions
df_lara["item_description"] = df_lara["item_description"].fillna("Unknown")

# Convert numeric columns
numeric_cols = [
    "total_boq_quantity", "this_month_target", "rate",
    "amount_target_this_month", "achieved_upto_previous_month", "achieved_upto_previous_day_this_month",
    "program_for_today", "today_achieved", "cumulative_achieved_this_month", "percent_achieved_this_month",
    "cumulative_achieved_upto_date", "percent_cumulative_achieved_upto_date", "tomorrow_program"
]
df_lara[numeric_cols] = df_lara[numeric_cols].apply(pd.to_numeric, errors="coerce")

# **Step 3: Connect to MySQL**
db_connection = mysql.connector.connect(
    host="localhost",
    user="root",
    password="Hope@123",
    database="dpr"
)
cursor = db_connection.cursor()

# **Step 4: Drop and Create Table**
cursor.execute("DROP TABLE IF EXISTS lara")
cursor.execute("""
CREATE TABLE lara (
    id INT AUTO_INCREMENT PRIMARY KEY,
    item_no VARCHAR(50),
    item_description VARCHAR(50),
    unit VARCHAR(50),
    total_boq_quantity FLOAT,
    this_month_target FLOAT, 
    rate FLOAT,
    amount_target_this_month FLOAT,
    achieved_upto_previous_month FLOAT,
    achieved_upto_previous_day_this_month FLOAT,
    program_for_today FLOAT,
    today_achieved FLOAT,   
    cumulative_achieved_this_month FLOAT, 
    percent_achieved_this_month FLOAT,
    cumulative_achieved_upto_date FLOAT,    
    percent_cumulative_achieved_upto_date FLOAT,
    tomorrow_program FLOAT,
    remark VARCHAR(50)
)
""")

# **Step 5: Insert Data into MySQL**
insert_query = """
INSERT INTO lara (item_no, item_description, unit, total_boq_quantity, this_month_target, rate, amount_target_this_month,
                  achieved_upto_previous_month, achieved_upto_previous_day_this_month, program_for_today, today_achieved,
                  cumulative_achieved_this_month, percent_achieved_this_month, cumulative_achieved_upto_date,
                  percent_cumulative_achieved_upto_date, tomorrow_program, remark)
VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
"""

# Convert dataframe to list of tuples
data = df_lara.replace({pd.NA: None}).values.tolist()
cursor.executemany(insert_query, data)
db_connection.commit()
print("\n✅ Data successfully inserted into 'lara' table!")

# **Step 6: Assign Unique `item_no` for Missing Values**
cursor.execute("""
UPDATE lara 
SET item_no = CONCAT('Item-', id)
WHERE item_no IS NULL OR item_no = '';
""")
db_connection.commit()
print("\n✅ Missing item_no values assigned!")

# **Step 7: Update `total_boq_quantity` Values Correctly**
update_queries = [
    ("A-1", 60),
    ("A", 327670),
    ("B", 162468),
    ("C", 106487),
    ("102", 549533),
    ("109", 173990),
    ("113", 1000),
    ("201", 16670),
    ("Item-13", 16670),
    ("203", 74465),
    ("203", 8786),  # Conflict: Two different values for "203". Choose one.
    ("301", 134966),
    ("403", 13466),
    ("404", 4893),
    ("405", 79600)
]

for item, value in update_queries:
    cursor.execute("UPDATE lara SET total_boq_quantity = %s WHERE TRIM(item_no) = %s;", (value, item))
    db_connection.commit()

print("\n✅ total_boq_quantity values updated!")

# **Step 8: Close Database Connection**
cursor.close()
db_connection.close()
print("\n🔄 Refresh Power BI to see the changes!")
