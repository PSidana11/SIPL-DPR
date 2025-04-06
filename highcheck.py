import pdfplumber
import pandas as pd
import mysql.connector
pdf_path = "C:/bilaspur.pdf"

# Extract table data
extracted_data = []
with pdfplumber.open(pdf_path) as pdf:
    for page in pdf.pages:
        table = page.extract_table()  # Extract table from page
        if table:
            extracted_data.extend(table)  # Add table data to the list

# Convert to DataFrame
df_bilaspur = pd.DataFrame(extracted_data)

# Print all extracted rows (to manually check where column names exist)
print("Extracted Data (First 10 Rows):")
print(df_bilaspur.head(10))

# Identify the first row that looks like column names
for i, row in df_bilaspur.iterrows():
    if "Item No." in row.values:  # Adjust keyword based on your table
        header_index = i
        break


# Reassign proper column names and remove extra rows
df_bilaspur.columns = df_bilaspur.iloc[header_index]  # Use the correct row as column headers
df_bilaspur = df_bilaspur.iloc[header_index + 1:]  # Remove metadata rows above the table

# Reset index
df_bilaspur.reset_index(drop=True, inplace=True)

# Print final column names
print("\n✅ Corrected Column Names:")
print(df_bilaspur.columns)

df_bilaspur.columns = ["item_no", "item_description", "unit", "program_for_today", "achieved_today",
                       "achieved_this_month", "percentage_achieved_this_month", "cumm_achieved_upto_date",
                        "percentage_cumm_achieved_upto_date", "tomorrow_program", "program_for_this_month", "remarks"]

#data cleaning to handle both sting and numberss
df_bilaspur["item_no"]= df_bilaspur["item_no"].astype(str)
df_bilaspur["remarks"]= df_bilaspur["remarks"].astype(str) 


#converting to numeric columns 
numeric_cols = ["program_for_today", "achieved_today", "achieved_this_month", "percentage_achieved_this_month",
                "cumm_achieved_upto_date", "percentage_cumm_achieved_upto_date", "tomorrow_program",
                "program_for_this_month"]

for col in numeric_cols:
    df_bilaspur[col] = pd.to_numeric(df_bilaspur[col], errors="coerce")  # Convert numbers

db_connection = mysql.connector.connect(
    host="localhost",
    user="root",
    password="Hope@123",
    database="dpr"
)
cursor = db_connection.cursor()
cursor.execute("DROP TABLE IF EXISTS bilaspur")
cursor.execute(""" 
CREATE TABLE bilaspur(
            id INT AUTO_INCREMENT PRIMARY KEY,
            item_no VARCHAR(255),
            item_description VARCHAR(255),
            unit text,
            program_for_today FLOAT,
            achieved_today FLOAT,
            achieved_this_month FLOAT,
            percentage_achieved_this_month FLOAT,
            cumm_achieved_upto_date FLOAT,
            percentage_cumm_achieved_upto_date FLOAT,
            tomorrow_program FLOAT,
            program_for_this_month FLOAT,
            remarks VARCHAR(50)
)
""")

insert_query = """
INSERT INTO bilaspur (item_no, item_description, unit, program_for_today, achieved_today,
                    achieved_this_month, percentage_achieved_this_month,
                    cumm_achieved_upto_date, percentage_cumm_achieved_upto_date, tomorrow_program,
                      program_for_this_month,
                    remarks)
VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
"""
data = df_bilaspur.values.tolist()

cursor.executemany(insert_query, data)
db_connection.commit()

# **Step 7: Close Database Connection**
cursor.close()
db_connection.close()
