from openpyxl.workbook import Workbook

planilhas_contas = Workbook()
pagina1 = planilhas_contas.active

with open('anotações.txt','r',encoding='utf-8') as arquivo:
    for linha in arquivo:
        pagina1.append(linha.split(','))
planilhas_contas.save('contas_a_pagar.xlsx')