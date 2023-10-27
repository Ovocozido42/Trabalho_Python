meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho']
vendas = []
for mes in meses:
    venda = int(input(f'Informe as vendas para {mes}: '))
    vendas.append(venda)

horas = [9, 10, 11, 12, 13, 14, 15]
movimento = []
for hora in horas:
    valor = int(input(f'Informe o movimento para a hora {hora}: '))
    movimento.append(valor)

# --------------------------------------------------------------------------------

excel_file = 'dados_de_vendas_e_movimento.xlsx'
workbook = xlsxwriter.Workbook(excel_file)
worksheet_vendas = workbook.add_worksheet("Vendas Mensais")
worksheet_movimento = workbook.add_worksheet("Movimento por Hora")

for r_idx, (mes, venda) in enumerate(zip(meses, vendas), 1):
    worksheet_vendas.write(r_idx, 0, mes)
    worksheet_vendas.write(r_idx, 1, venda)

for r_idx, (hora, valor) in enumerate(zip(horas, movimento), 1):
    worksheet_movimento.write(r_idx, 0, hora)
    worksheet_movimento.write(r_idx, 1, valor)

#----------------------------------------------------------------------------

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

sender_email = 'policenorosa17@gmail.com'
sender_password = 'xjcr csbw ubgg bric'
recipient_email = 'blatarosa@gmail.com'
email_subject = 'Relatório Mensal de Vendas e Movimento'
email_body = 'Relatório de vendas e movimento anexado.'
