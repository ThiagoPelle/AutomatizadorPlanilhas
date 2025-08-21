# Automatizador_De_Planilhas

## Utilizando **openpyxl**

O *openpyxl* é uma biblioteca Python utilizada para **ler, criar e modificar arquivos do Excel no formato .xlsx**
Ela é amplamente usada para automação de planilhas, análise de dados e geração de relatórios sem precisar abrir o Excel manualmente.

A instalção é feita no *Windows*, no terminal da sua maquina.

`pip install openpyxl`

 ### 📌Exemplo Rápido

```python
from openpyxl import Workbook

# Criando um novo arquivo Excel
wb = Workbook()
sheet = wb.active
sheet.title = "Vendas"

# Inserindo dados
sheet["A1"] = "Produto"
sheet["B1"] = "Preço"
sheet.append(["Notebook", 4600])
sheet.append(["Mouse", 76.80])

# Salvando o arquivo
wb.save("planilha_vendas.xlsx")
```
```
