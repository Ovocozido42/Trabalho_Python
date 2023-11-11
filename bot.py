import pymysql
import random
import xlsxwriter

# Function to fetch data from database
def fetch_data():
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

    data = [(row[0], row[1]) for row in rows]

    connection.close(); 

    return data

# Fetch data from database
data = fetch_data()

# Generate random data for movimento
movimento = [random.randint(5, 30) for _ in range(24)]

# --------------------------------------------------------------------------------

# Create excel file
excel_file = 'dados_de_vendas_e_movimento.xlsx'
workbook = xlsxwriter.Workbook(excel_file)
worksheet_vendas = workbook.add_worksheet("Vendas Mensais")
worksheet_movimento = workbook.add_worksheet("Movimento por Hora")

# Write data to worksheets
for r_idx, (mes, venda) in enumerate(data, 1):
    worksheet_vendas.write(r_idx, 0, mes)
    worksheet_vendas.write(r_idx, 1, venda)

for r_idx, (hora, valor) in enumerate(zip(range(9, 24), movimento), 1):
    worksheet_movimento.write(r_idx, 0, hora)
    worksheet_movimento.write(r_idx, 1, valor)
    
# ---------------------------------------------------------------------------------

# Create charts
chart_vendas = workbook.add_chart({'type': 'column'})
chart_vendas.add_series({
    'categories': ['Vendas Mensais', 1, 0, 7, 0],
    'values': ['Vendas Mensais', 1, 1, 7, 1],
})
chart_vendas.set_title({'name': 'Vendas Mensais'})
worksheet_vendas.insert_chart('E2', chart_vendas)

chart_movimento = workbook.add_chart({'type': 'line'})
chart_movimento.add_series({
    'categories': ['Movimento por Hora', 1, 0, 7, 0],
    'values': ['Movimento por Hora', 1, 1, 7, 1],
})
chart_movimento.set_title({'name': 'Movimento por Hora'})
worksheet_movimento.insert_chart('E2', chart_movimento)

workbook.close()

# sender_email = 'policenorosa17@gmail.com'
# sender_password = 'xjcr csbw ubgg bric'
# recipient_email = 'blatarosa@gmail.com'
# email_subject = 'Relatório Mensal de Vendas e Movimento'
# email_body = 'Relatório de vendas e movimento anexado.'
