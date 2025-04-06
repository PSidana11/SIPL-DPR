import pandas as pd 
import mysql.connector

#Lodaing the excel file 
from datetime import datetime
today = datetime.now().strftime("%Y-%m-%d")
file_path = fr"C:\Users\E01412\Desktop\REPORTS\DAILY\nilothi_{today}.xlsx"
xls = pd.ExcelFile(file_path)

#making it read the excel sheet and exact columns 
df = xls.parse("sheet", skiprows=10, nrows=47)

#droping the empty columns 
df.dropna(how='all', axis=1, inplace=True)

df_nilothi = df.iloc[:, :15]
df_nilothi.columns = [
    "item_no", "item_descriotion", "total_amount", "this_month_target_qty",
    "this_month_target_amount", "achieved_upto_last_month", 
    "achieved_upto_previous_day_this_month", "program_today", "today_achieved",
    "cumm_achieved_this_month", "percentage_achieved_this_month", "cumm_achieved_upto_date",
    "percentage_cumm_achieved_upto_date", "tomorrow_program", "remarks"
]

numeric_cols = [
    "total_amount", "this_month_target_qty", "this_month_target_amount",
    "achieved_upto_last_month", "achieved_upto_previous_day_this_month", "program_today",
    "today_achieved", "cumm_achieved_this_month", "percentage_achieved_this_month",
    "cumm_achieved_upto_date", "percentage_cumm_achieved_upto_date", "tomorrow_program"
]
df_nilothi[numeric_cols] = df_nilothi[numeric_cols].apply(pd.to_numeric, errors="coerce")
db_connection = mysql.connector.connect(
    host="localhost",
    user="root",
    password="Hope@123",
    database="dpr"
)
cursor = db_connection.cursor()

#drop and create table 

cursor.execute("""
CREATE TABLE  IF NOT EXISTS nilothi (
    id INT AUTO_INCREMENT PRIMARY KEY,
    item_no VARCHAR(100),
    item_description VARCHAR(100),
    total_amount FLOAT,
    this_month_target_qty FLOAT,
    this_month_target_amount FLOAT,
    achieved_upto_last_month FLOAT,
    achieved_upto_previous_day_this_month FLOAT,
    program_today FLOAT,
    today_achieved FLOAT,
    cumm_achieved_this_month FLOAT,
    percentage_achieved_this_month FLOAT,
    cumm_achieved_upto_date FLOAT,
    percentage_cumm_achieved_upto_date FLOAT,
    tomorrow_program FLOAT,
    remarks TEXT,
    entry_date DATE
)
""")
insert_query = """
INSERT INTO nilothi (item_no, item_description, total_amount, this_month_target_qty,
                    this_month_target_amount, achieved_upto_last_month,
                    achieved_upto_previous_day_this_month, program_today, today_achieved,
                    cumm_achieved_this_month, percentage_achieved_this_month,
                    cumm_achieved_upto_date, percentage_cumm_achieved_upto_date,
                    tomorrow_program, remarks, entry_date)
VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s, CURDATE())
"""
data = df_nilothi.values.tolist()
cursor.executemany(insert_query, data)
db_connection.commit()

cursor.close()
db_connection.close()   
print("✅ Data inserted successfully!")