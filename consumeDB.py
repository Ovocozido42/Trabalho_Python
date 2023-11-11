import pymysql 

connection = pymysql.connect(
    host='localhost',
    port=3306,  
    user="root",
    password="", 
    db='bot_estatistica'
)

cursor = connection.cursor(); 

cursor.execute('select * from Lucro'); 

rows = cursor.fetchall(); 

for row in rows:
    print(row[1]);
    
connection.close(); 