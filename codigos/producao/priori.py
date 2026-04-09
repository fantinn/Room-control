import pyodbc
import pandas as pd
from datetime import date, timedelta

# Dados de conexão (ajuste para seu ambiente)
conn = pyodbc.connect(
    'DRIVER={SQL Server};SERVER=SEU_SERVIDOR;DATABASE=SEU_BANCO;UID=SEU_USUARIO;PWD=SUA_SENHA'
)

# Calcular o dia anterior
ontem = date.today() - timedelta(days=1)
data_str = ontem.strftime('%Y-%m-%d')

# Query para buscar o dia anterior
query = f"SELECT * FROM producao_diaria WHERE data = '{data_str}'"

# Ler os dados
df = pd.read_sql(query, conn)

# Salvar em Excel
df.to_excel('producao_diaria.xlsx', index=False)