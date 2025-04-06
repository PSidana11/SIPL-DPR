import pandas as pd 
import mysql.connector

from datetime import datetime
today = datetime.now().strftime("%Y-%m-%d")
file_path = fr"C:\Users\E01412\Desktop\REPORTS\DAILY\dvc_{today}.xlsx"

xls = pd.ExcelFile(file_path)

df = xls.parse("sheet", skiprows=7, nrows=18)

df_dvc = df.iloc[:, :12]
df_dvc.columns = [
    "sr_no", "item_description", "unit", "boq_quantity", "target_for_month",
    "program_for_today", "achieved_today", "achieved_this_month",
    "percentage_achieved_this_month", "cumm_percentage_achieved",
    "tomorrow_program", "remarks"
]
numeric_cols = [
    "boq_quantity", "target_for_month", "program_for_today", "achieved_today",
    "achieved_this_month", "percentage_achieved_this_month", "cumm_percentage_achieved",
    "tomorrow_program"
]
df_dvc[numeric_cols] = df_dvc[numeric_cols].apply(pd.to_numeric, errors="coerce")   

db_connection = mysql.connector.connect(
    host="localhost",
    user="root",
    password="Hope@123",
    database="dpr"   
)
cursor = db_connection.cursor()



cursor.execute("""
CREATE TABLE IF NOT EXISTS dvc (
    batch_id INT AUTO_INCREMENT PRIMARY KEY,
    sr_no VARCHAR(255),
    item_description TEXT,
    unit VARCHAR(255),
    boq_quantity FLOAT,
    target_for_month FLOAT,
    program_for_today FLOAT,
    achieved_today FLOAT,
    achieved_this_month FLOAT,
    percentage_achieved_this_month FLOAT,
    cumm_percentage_achieved FLOAT,
    tomorrow_program FLOAT,
    remarks VARCHAR(50),
    entry_date DATE 
)
""")

insert_query = """ 
INSERT INTO dvc (sr_no, item_description, unit, boq_quantity, target_for_month,
                program_for_today, achieved_today, achieved_this_month,
                percentage_achieved_this_month, cumm_percentage_achieved,
                tomorrow_program, remarks, entry_date)
VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s, CURDATE())

"""
data = df_dvc.values.tolist()
cursor.executemany(insert_query, data)
db_connection.commit()

cursor.close()
db_connection.close()
