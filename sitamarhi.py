import pandas as pd 
import mysql.connector 

# loading the excel file 
from datetime import datetime
today = datetime.now().strftime("%Y-%m-%d")
file_path = fr"C:\Users\E01412\Desktop\REPORTS\DAILY\sitamarhi_{today}.xlsx"

xls = pd.ExcelFile(file_path)

#reading the excel sheet 
df_sheet = xls.parse("sheet1", skiprows=7, nrows=18)

#renaming the columns
df_sitamarhi = df_sheet.iloc[:,:18]
df_sitamarhi.columns = [
    "si_no", "description", "boq_item_no", "unit", "total_boq_quantity", "boq_rate",
    "boq_amount", "this_month_target",  "amount_target_this_month", "today_target",
    "cum_achieved_upto_previous_day", "today_achieved", "cum_achieved_upto_date",
    "balance_to_be_executed", "today_program_amount", "cum_amount_upto_previous_day",
    "today_achieved_amount", "cum_amount_upto_date"]



#convert numeric columns 
numeric_cols = [
    "si_no", "total_boq_quantity", "boq_rate", "boq_amount", "this_month_target",
    "amount_target_this_month", "today_target", "cum_achieved_upto_previous_day",
    "today_achieved", "cum_achieved_upto_date", "balance_to_be_executed",
    "today_program_amount", "cum_amount_upto_previous_day", "today_achieved_amount", "cum_amount_upto_date"
]
df_sitamarhi[numeric_cols] = df_sitamarhi[numeric_cols].apply(pd.to_numeric, errors="coerce")



#connect to mysql 
db_connection = mysql.connector.connect(
    host="localhost",
    user="root",
    password="Hope@123",
    database="dpr"
)
cursor = db_connection.cursor()

#drop and create table 

cursor.execute("""
CREATE TABLE IF NOT EXISTS sitamarhi ( 
            id INT AUTO_INCREMENT PRIMARY KEY,
            si_no INT,
            description VARCHAR(255),
            boq_item_no VARCHAR(50),
            unit VARCHAR(50),
            total_boq_quantity FLOAT,   
            boq_rate FLOAT,
            boq_amount FLOAT,
            this_month_target FLOAT,
            amount_target_this_month FLOAT,
            today_target FLOAT,
            cum_achieved_upto_previous_day FLOAT,
            today_achieved FLOAT,
            cum_achieved_upto_date FLOAT,
            balance_to_be_executed FLOAT,
            today_program_amount FLOAT,
            cum_amount_upto_previous_day FLOAT,
            today_achieved_amount FLOAT,
            cum_amount_upto_date FLOAT,
            entry_date DATE 
)
""")    
#inserting the data into mysql
insert_query = """
INSERT INTO sitamarhi (si_no, description, boq_item_no, unit, total_boq_quantity, boq_rate, boq_amount, this_month_target, 
                        amount_target_this_month,
                        today_target, cum_achieved_upto_previous_day,today_achieved, cum_achieved_upto_date, balance_to_be_executed,
                         today_program_amount, cum_amount_upto_previous_day, today_achieved_amount, cum_amount_upto_date, entry_date)
VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s, CURDATE())
"""
#convert dataframe to list of tuples 
data = df_sitamarhi.values.tolist()
cursor.executemany(insert_query, data)
db_connection.commit()  
print("\n✅ Data successfully inserted into 'sitamarhi' table!")

cursor.close()
db_connection.close()
print("\n✅ Database connection closed!")

print("\n📊 Checking first few rows before inserting into MySQL:")
print(df_sitamarhi.head())  # Shows first 5 rows