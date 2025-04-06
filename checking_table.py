import mysql.connector  # Import MySQL Connector
import pandas as pd  # Import Pandas




try:
    db_connection = mysql.connector.connect(
        host="localhost",
        user="root",
        password="Hope@123",
        database="dpr"
    )
    cursor = db_connection.cursor()

    # Drop table if it exists
    cursor.execute("DROP TABLE IF EXISTS bilaspur")

    # Create table
    cursor.execute(""" 
    CREATE TABLE bilaspur(
        id INT AUTO_INCREMENT PRIMARY KEY,
        item_no VARCHAR(255),
        item_description TEXT,
        unit VARCHAR(50),
        program_for_today FLOAT,
        achieved_today FLOAT,
        achieved_this_month FLOAT,
        percentage_achieved_this_month FLOAT,
        cumm_achieved_upto_date FLOAT,
        percentage_cumm_achieved_upto_date FLOAT,
        tomorrow_program FLOAT,
        program_for_this_month FLOAT,
        remarks TEXT
    )
    """)
    print("✅ Table 'bilaspur' created successfully!")

    # Insert Data
    insert_query = """
    INSERT INTO bilaspur (item_no, item_description, unit, program_for_today, achieved_today,
                          achieved_this_month, percentage_achieved_this_month,
                          cumm_achieved_upto_date, percentage_cumm_achieved_upto_date, tomorrow_program,
                          program_for_this_month, remarks)
    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
    """
    
    # Convert DataFrame to list
    data = df_bilaspur.values.tolist()

    cursor.executemany(insert_query, data)
    db_connection.commit()
    print("✅ Data inserted successfully!")

except mysql.connector.Error as err:
    print(f"❌ MySQL Error: {err}")

finally:
    cursor.close()
    db_connection.close()
    print("🔌 Connection closed.")
