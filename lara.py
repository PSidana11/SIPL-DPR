import pandas as pd
import mysql.connector

# **Step 1: Load Excel File**
from datetime import datetime
today = datetime.now().strftime("%Y-%m-%d")
file_path = fr"C:\Users\E01412\Desktop\REPORTS\DAILY\lara_{today}.xlsx"

xls = pd.ExcelFile(file_path)

# Read Excel sheet
df_sheet = xls.parse("sheet", skiprows=9, nrows=21)

# Rename columns
df_lara = df_sheet.iloc[:, :17]
df_lara.columns = [
    "item_no", "item_description", "unit", "total_boq_quantity", "this_month_target", "rate",
    "amount_target_this_month", "achieved_upto_previous_month", "achieved_upto_previous_day_this_month",
    "program_for_today", "today_achieved", "cumulative_achieved_this_month", "percent_achieved_this_month",
    "cumulative_achieved_upto_date", "percent_cumulative_achieved_upto_date", "tomorrow_program", "remark"
]

# **Step 2: Data Cleaning**
df_lara["item_no"] = df_lara["item_no"].astype(str).replace(["nan", "None"], None)  # Handle missing item_no
df_lara["item_description"] = df_lara["item_description"].fillna("Unknown")  # Fill missing descriptions

# Convert numeric columns
numeric_cols = [
    "total_boq_quantity", "this_month_target", "rate",
    "amount_target_this_month", "achieved_upto_previous_month", "achieved_upto_previous_day_this_month",
    "program_for_today", "today_achieved", "cumulative_achieved_this_month", "percent_achieved_this_month",
    "cumulative_achieved_upto_date", "percent_cumulative_achieved_upto_date", "tomorrow_program"
]
df_lara[numeric_cols] = df_lara[numeric_cols].apply(pd.to_numeric, errors="coerce")

# **Fix: Replace first "203" with "205"**
changed_first_203 = False
for idx, row in df_lara.iterrows():
    if row["item_no"] == "203" and not changed_first_203:
        df_lara.at[idx, "item_no"] = "205"
        changed_first_203 = True

# **Fix: Assign unique item_no before insertion**
df_lara["item_no"] = df_lara["item_no"].fillna(df_lara.index.to_series().apply(lambda x: f"Item-{x+1}"))

# **Step 3: Connect to MySQL**
db_connection = mysql.connector.connect(
    host="localhost",
    user="root",
    password="Hope@123",
    database="dpr"
)
cursor = db_connection.cursor()

# **Step 4: Drop and Create Table**

cursor.execute("""
CREATE TABLE IF NOT EXISTS  lara (
    id INT AUTO_INCREMENT PRIMARY KEY,
    item_no VARCHAR(50),
    item_description VARCHAR(255),
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
    remark VARCHAR(50), 
    entry_date DATE 
)
""")

# **Step 5: Insert Data into MySQL**
insert_query = """
INSERT INTO lara (item_no, item_description, unit, total_boq_quantity, this_month_target, rate, amount_target_this_month,
                  achieved_upto_previous_month, achieved_upto_previous_day_this_month, program_for_today, today_achieved,
                  cumulative_achieved_this_month, percent_achieved_this_month, cumulative_achieved_upto_date,
                  percent_cumulative_achieved_upto_date, tomorrow_program, remark, entry_date)
VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, CURDATE())
"""

# Convert dataframe to list of tuples
data = df_lara.replace({pd.NA: None}).values.tolist()
cursor.executemany(insert_query, data)
db_connection.commit()
print("\n✅ Data successfully inserted into 'lara' table!")

# **Step 6: Optimized Bulk Update for `total_boq_quantity`**
update_mappings = {
    "A-1": 60, "A": 327670, "B": 162468, "C": 106487, "102": 549533, "109": 173990, "113": 1000,
    "201": 16670, "Item-13": 16670, "205": 74465, "203": 8786, "301": 134966, "403": 13466,
    "404": 4893, "405": 79600
}

# **Generate SQL Query for Bulk Update**
update_query = """
    UPDATE lara
    SET total_boq_quantity = CASE item_no
"""
for item, value in update_mappings.items():
    update_query += f" WHEN '{item}' THEN {value}"
update_query += " ELSE total_boq_quantity END;"

cursor.execute(update_query)
db_connection.commit()
print("\n✅ total_boq_quantity values updated!")

# **Step 7: Close Database Connection**
cursor.close()
db_connection.close()
print("\n🔄 Refresh Power BI to see the changes!")
