import mysql.connector 
print("its working ")

#for connecting to my sql server 
conn = mysql.connector.connect(
    host=" localhost",
    user="root",    #username   
    password="Hope@123",  #password
)

#creating a cursor object using the cursor() method
cursor = conn.cursor()
#creating a database
cursor.execute("CREATE DATABASE IF NOT EXISTS dpr ")
#closing the connection
cursor.close()
conn.close()
print("Database created successfully")