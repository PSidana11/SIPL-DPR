import pandas as pd
import mysql.connector  
from tabulate import tabulate

# **Step 1: Load Excel File**
file_path = "C:/sitamarhi.xlsx"
xls = pd.ExcelFile(file_path)

# **Step 2: Read Required Rows from Sheet**
df_sheet = xls.parse("sheet1", skiprows=7, nrows=22)

# **Step 3: Rename Columns**
df_sitamarhi = df_sheet.iloc[:, :17]
df_sitamarhi.columns = ["si_no", "description", "boq_item_no", "unit", "total_s_boq_quantity", "s_boq_rate", 
                         "s_BOQ_amount", "this_month_target", "today_program_quatnity",
                         "cum_achieved_upto_previous_day", "today_achieved", "cum_achieved_upto_date",
                         "balance_to_be_executed", "today_program_amount", "cum_achieved_upto_previous_day_amount",
                         "today_achieved_amount", "cum_achieved_amount_upto_date"]

# **Step 4: Data Cleaning**
df_sitamarhi["description"] = df_sitamarhi["description"].fillna("Unknown")  # Fill missing descriptions

# Convert numeric columns
numeric_cols = ["si_no", "total_s_boq_quantity", "s_boq_rate", "s_BOQ_amount",
                "this_month_target", "today_program_quatnity", "cum_achieved_upto_previous_day",
                "today_achieved", "cum_achieved_upto_date", "balance_to_be_executed", "today_program_amount",
                "cum_achieved_upto_previous_day_amount", "today_achieved_amount", "cum_achieved_amount_upto_date"]

df_sitamarhi[numeric_cols] = df_sitamarhi[numeric_cols].apply(pd.to_numeric, errors="coerce").fillna(0)

# **Step 5: Connect to MySQL**
db_connection = mysql.connector.connect(
    host="localhost",
    user="root",
    password="Hope@123",
    database="dpr"
)
cursor = db_connection.cursor()

# **Step 6: Drop and Recreate Table**
cursor.execute("DROP TABLE IF EXISTS sitamarhi")
cursor.execute("""
CREATE TABLE sitamarhi (
    id INT AUTO_INCREMENT PRIMARY KEY,
    si_no VARCHAR(50),
    description VARCHAR(100),
    boq_item_no VARCHAR(255),
    unit VARCHAR(50),
    total_s_boq_quantity FLOAT,
    s_boq_rate FLOAT,
    s_BOQ_amount FLOAT,
    this_month_target FLOAT,
    today_program_quatnity FLOAT,
    cum_achieved_upto_previous_day FLOAT,   
    today_achieved FLOAT,
    cum_achieved_upto_date FLOAT,
    balance_to_be_executed FLOAT,
    today_program_amount FLOAT,
    cum_achieved_upto_previous_day_amount FLOAT,           
    today_achieved_amount FLOAT,
    cum_achieved_amount_upto_date FLOAT
)
""")

# **Step 7: Insert Data into MySQL**
insert_query = """
INSERT INTO sitamarhi (si_no, description, boq_item_no, unit, total_s_boq_quantity, s_boq_rate,
                        s_BOQ_amount, this_month_target, today_program_quatnity, cum_achieved_upto_previous_day,
                        today_achieved, cum_achieved_upto_date, balance_to_be_executed, today_program_amount,
                        cum_achieved_upto_previous_day_amount, today_achieved_amount, cum_achieved_amount_upto_date)
VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
"""

# Convert dataframe to list of tuples
data = df_sitamarhi.replace({None: 0}).values.tolist()  # Replace None with 0 for insertion

# Insert Data
cursor.executemany(insert_query, data)
db_connection.commit()
print("\n✅ Data successfully inserted into 'sitamarhi' table!")

# **Step 8: Fetch and Display Data**
def limit_words(text, word_limit=10):
    if isinstance(text, str):
        words = text.split()[:word_limit]
        return " ".join(words) 
    return text

# Fetch all data 
cursor.execute("SELECT * FROM sitamarhi LIMIT 20")
rows = cursor.fetchall()

# Convert to DataFrame
columns = [
    "id", "si_no", "description", "boq_item_no", "unit", "total_s_boq_quantity",
    "s_boq_rate", "s_BOQ_amount", "this_month_target", "today_program_quatnity",
    "cum_achieved_upto_previous_day", "today_achieved", "cum_achieved_upto_date",
    "balance_to_be_executed", "today_program_amount", "cum_achieved_upto_previous_day_amount",
    "today_achieved_amount", "cum_achieved_amount_upto_date"            
]
df_output = pd.DataFrame(rows, columns=columns)

# **Step 9: Format for Better Readability**
df_output["description"] = df_output["description"].apply(lambda x: limit_words(x, 10))
pd.set_option("display.max_colwidth", 30)  # Prevents column wrapping
pd.set_option("display.width", 1000)  # Ensures proper table width

# Print formatted table
print("\n🔹 Data in Tabular Format:")
print(tabulate(df_output, headers="keys", tablefmt="fancy_grid", showindex=False))

# **Step 10: Close Database Connection**
cursor.close()
db_connection.close()
