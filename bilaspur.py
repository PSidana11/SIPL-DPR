
import mysql.connector
import pandas as pd 

from datetime import datetime
today = datetime.now().strftime("%Y-%m-%d")
file_path = fr"C:\Users\E01412\Desktop\REPORTS\DAILY\bilaspur_{today}.xlsx"

xls = pd.ExcelFile(file_path)

df=xls.parse("sheet", skiprows=10, nrows=19)
df_bilaspur = df.iloc[:, :12]


df_bilaspur.columns = ["item_no", "item_description", "unit", "program_for_today", "achieved_today",
                       "achieved_this_month", "percentage_achieved_this_month", "cumm_achieved_upto_date",
                        "percentage_cumm_achieved_upto_date", "tomorrow_program", "program_for_this_month", "remarks"]




#converting to numeric columns 
numeric_cols = ["program_for_today", "achieved_today", "achieved_this_month", "percentage_achieved_this_month",
                "cumm_achieved_upto_date", "percentage_cumm_achieved_upto_date", "tomorrow_program",
                "program_for_this_month"]



db_connection = mysql.connector.connect(
    host="localhost",
    user="root",
    password="Hope@123",
    database="dpr"
)
cursor = db_connection.cursor()

cursor.execute(""" 
CREATE TABLE IF NOT EXISTS bilaspur(
            id INT AUTO_INCREMENT PRIMARY KEY,
            item_no VARCHAR(255),
            item_description TEXT,
            unit text,
            program_for_today FLOAT,
            achieved_today FLOAT,
            achieved_this_month FLOAT,
            percentage_achieved_this_month FLOAT,
            cumm_achieved_upto_date FLOAT,
            percentage_cumm_achieved_upto_date FLOAT,
            tomorrow_program FLOAT,
            program_for_this_month FLOAT,
            remarks VARCHAR(50),
            entry_date DATE
               

)
""")

insert_query = """
INSERT INTO bilaspur (item_no, item_description, unit, program_for_today, achieved_today,
                    achieved_this_month, percentage_achieved_this_month,
                    cumm_achieved_upto_date, percentage_cumm_achieved_upto_date, tomorrow_program,
                      program_for_this_month,
                    remarks, entry_date)
VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, CURDATE())
"""
data = df_bilaspur.values.tolist()

cursor.executemany(insert_query, data)
db_connection.commit()

# **Step 7: Close Database Connection**
cursor.close()
db_connection.close()




