import pandas as pd  
import mysql.connector

from datetime import datetime
today = datetime().strftime("%Y-%m-%d")
file_path = fr"C:\Users\E01412\Desktop\REPORTS\DAILY\gpr_{today}.xlsx"

xls = pd.ExcelFile(file_path)
df = xls.parse("summary", skiprows=4 ,nrows=8)

  #     Adjust skiprows if needed
df_gpr = df.iloc[:, :11]
print(df_gpr.columns.tolist())
df_gpr.columns = df_gpr.columns.str.strip()

df_gpr.columns = ["sr_no", "description", "target_amount_march_25", 
                  "cumm_planned_amount",
                    "work_planned", "upto_date", "upto_feb_month", 
                    "upto_date_achieved_amount_this_month",
                    "update_acheived_amount_til_prevous_day", 
                    "today_achieved_amount", "remarks"]

numeric_cols = ["target_amount_march_25", "cumm_planned_amount",
                "work_planned", "upto_date", "upto_feb_month", 
                "upto_date_achieved_amount_this_month",
                "update_acheived_amount_til_prevous_day", 
                "today_achieved_amount"]
df_gpr[numeric_cols] = df_gpr[numeric_cols].apply(pd.to_numeric, errors="coerce")

db_connection = mysql.connector.connect(
    host = "localhost",
    user = "root",
    password = "Hope@123",  # Update this
    database = "dpr"

)
cursor = db_connection.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS gpr (
    batch_id INT AUTO_INCREMENT PRIMARY KEY,
    sr_no VARCHAR(100),
    description VARCHAR(100),
    target_amount_march_25 FLOAT,
    cumm_planned_amount FLOAT,
    work_planned FLOAT,
    upto_date FLOAT,
    upto_feb_month FLOAT,
    upto_date_achieved_amount_this_month FLOAT,
    update_acheived_amount_til_prevous_day FLOAT,
    today_achieved_amount FLOAT,
    remarks VARCHAR(255),
    entry_date DATE
)
""")

insert_query = """
INSERT INTO gpr (sr_no, description, target_amount_march_25, cumm_planned_amount,
                work_planned, upto_date, upto_feb_month, upto_date_achieved_amount_this_month,
                update_acheived_amount_til_prevous_day, today_achieved_amount, remarks, entry_date)
VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, CURDATE())
"""
data = df_gpr.values.tolist()

cursor.executemany(insert_query, data)
db_connection.commit()

cursor.close()
db_connection.close()