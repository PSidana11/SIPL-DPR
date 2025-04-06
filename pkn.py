import pandas as pd
import mysql.connector

# Load Excel file
from datetime import datetime
today = datetime.now().strftime("%Y-%m-%d")
file_path = fr"C:\Users\E01412\Desktop\REPORTS\DAILY\pkn_{today}.xlsx"

xls = pd.ExcelFile(file_path)

# Read the specific sheet
df = xls.parse("sheet", skiprows=10 ,nrows=35)  # Adjust skiprows if needed

# Drop completely empty columns
df.dropna(how='all', axis=1, inplace=True)

# Debugging Step: Print column names before renaming
print("\n🔍 Number of columns in Excel:", len(df.columns))
print("\n📌 Column Names Before Renaming:", df.columns.tolist())

#rena ing columns
df_pkn = df.iloc[:, :14]
df_pkn.columns = [
    "item_no", "item_description", "total_amount", "this_month_target_qty", "this_month_target_amount",
    "achieved_upto_previous_day_this_month", "program_today", "today_achieved",
    "cumm_achieved_this_month", "percentage_achieved_this_month", "cumm_acieved_upto_date",
    "percentage_cumm_achieved_upto_date",  "remarks"
]
# convert numeric columns 
numeric_cols= [
    "total_amount", "this_month_target_qty", "this_month_target_amount",
    "achieved_upto_previous_day_this_month", "program_today", "today_achieved",
    "cumm_achieved_this_month", "percentage_achieved_this_month", "cumm_acieved_upto_date",
    "percentage_cumm_achieved_upto_date"
]
df_pkn[numeric_cols] = df_pkn[numeric_cols].apply(pd.to_numeric, errors="coerce")


db_connection = mysql.connector.connect(
    host="localhost",
    user="root",
    password="Hope@123",  # Update this
    database="dpr"
)
cursor = db_connection.cursor()

# Drop & Create Table

cursor.execute("""
CREATE TABLE IF NOT EXISTS pkn ( 
    id INT AUTO_INCREMENT PRIMARY KEY,
    item_no VARCHAR(100),
    item_description  VARCHAR(100),
    total_amount FLOAT,
    this_month_target_qty FLOAT,
    this_month_target_amount FLOAT,
    achieved_upto_previous_day_this_month FLOAT,
    program_today FLOAT,
    today_achieved FLOAT,
    cumm_achieved_this_month FLOAT,
    percentage_achieved_this_month FLOAT,
    cumm_achieved_upto_date FLOAT,
    percentage_cumm_achieved_upto_date FLOAT,
    remarks VARCHAR(100),
    entry_date DATE 
)
""")

# Insert Data
insert_query = """
INSERT INTO pkn(item_no, item_description, total_amount,
                this_month_target_qty ,
                 this_month_target_amount,
                 achieved_upto_previous_day_this_month ,
                  program_today,
                  today_achieved,
                   cumm_achieved_this_month,
                   percentage_achieved_this_month,
                  cumm_achieved_upto_date,
                  percentage_cumm_achieved_upto_date,
                 remarks, 
                    entry_date
)
VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, CURDATE())
"""
data = df_pkn.values.tolist()
cursor.executemany(insert_query, data)
db_connection.commit()

print("\n✅ Data successfully inserted into 'pkn' table!")

cursor.close()
db_connection.close()
print("\n✅ Database connection closed!")
