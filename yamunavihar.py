import pandas as pd
import mysql.connector

from datetime import datetime
today = datetime.now().strftime("%Y-%m-%d")
file_path = fr"C:\Users\E01412\Desktop\REPORTS\DAILY\yamunavihar_{today}.xlsx"

df = pd.read_excel(file_path)
xls = pd.ExcelFile(file_path)

df = xls.parse("sheet", skiprows=10, nrows=22)

df_yamunavihar = df.iloc[:, :14]
df_yamunavihar.columns = [
    "item_no", "item_description","total_amount", "this_month_target_qty",
    "this_month_target_amount", "achieved_upto_previous_day_this_month",
    "program_for_today", "today_achieved", "cumm_achieved_this_month",
    "percentage_achieved_this_month", "cumm_achieved_upto_date",
    "percentage_cumm_achieved_upto_date", "tomorrow_program", "remarks"
]

numeric_cols= [
    "total_amount", "this_month_target_qty", "this_month_target_amount",
    "achieved_upto_previous_day_this_month", "program_for_today",   
    "today_achieved", "cumm_achieved_this_month", "percentage_achieved_this_month",
    "cumm_achieved_upto_date", "percentage_cumm_achieved_upto_date",
    "tomorrow_program"
]
df_yamunavihar[numeric_cols] = df_yamunavihar[numeric_cols].apply(pd.to_numeric, errors='coerce')

df_yamunavihar = df_yamunavihar[df_yamunavihar["item_description"].notna() & (df_yamunavihar["item_description"] != "")]

db_connection= mysql.connector.connect(
    host = "localhost",
    user = "root",
    password = "Hope@123",
    database = "dpr"
)
cursor = db_connection.cursor()


cursor.execute("""
CREATE TABLE IF NOT EXISTS yamunavihar(
    id INT AUTO_INCREMENT PRIMARY KEY,
    item_no TEXT,
    item_description TEXT,
    total_amount FLOAT,
    this_month_target_qty FLOAT,
    this_month_target_amount FLOAT,
    achieved_upto_previous_day_this_month FLOAT,
    program_for_today FLOAT,
    today_achieved FLOAT,
    cumm_achieved_this_month FLOAT,
    percentage_achieved_this_month FLOAT,
    cumm_achieved_upto_date FLOAT,
    percentage_cumm_achieved_upto_date FLOAT,
    tomorrow_program FLOAT,
    remarks VARCHAR(255),
    entry_date DATE 
)
""")
insert_query = """
INSERT INTO yamunavihar(
    item_no, item_description, total_amount, this_month_target_qty,
    this_month_target_amount, achieved_upto_previous_day_this_month,
    program_for_today, today_achieved, cumm_achieved_this_month,
    percentage_achieved_this_month, cumm_achieved_upto_date,
    percentage_cumm_achieved_upto_date, tomorrow_program, remarks, entry_date
) 
VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s, CURDATE())
"""
data = df_yamunavihar.values.tolist()
cursor.executemany(insert_query, data)
db_connection.commit()

cursor.close()
db_connection.close()
