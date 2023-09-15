import os
import pandas as pd

# Define o diretório onde a planilha será salva
diretorio = "temp"

# Verifica se o diretório existe e o cria se não existir
if not os.path.exists(diretorio):
    os.mkdir(diretorio)

# Crie um DataFrame com os dados
data = {
    'Referência': [],  # Substitua com os valores reais
    'Fornecedor': [],  # Substitua com os valores reais
    'Email': [],  # Substitua com os valores reais
    'Empresa': [],  # Substitua com os valores reais
    'Imposto': [],  # Substitua com os valores reais
    'Valor': [],  # Substitua com os valores reais
}

df = pd.DataFrame(data)

# Salve o DataFrame em um arquivo Excel no diretório "temp"
caminho_planilha = os.path.join(diretorio, 'matriz.xlsx')
df.to_excel(caminho_planilha, index=False)
