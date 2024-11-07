#PETICOS - RPA BI - EXCEL 🐶😺❤️
#--------------------------------------------------------------------------
#IMPORTS:
import psycopg2 as pg
import os
from dotenv import load_dotenv
import pandas as pd
#--------------------------------------------------------------------------
#Pegando a senha do banco no .env:
load_dotenv()
password = os.getenv('PASSWORD12')

#Conectando com o banco:
conn = pg.connect(
    dbname = "dbPeticos_2ano",
    user = "avnadmin",
    password = password,
    host = "db-peticos-cardosogih.k.aivencloud.com",
    port = 16207
)
cursor = conn.cursor()
#--------------------------------------------------------------------------

#Lista com os nomes das tabelas do banco que serão utilizadas no BI de usuário que estará na área restrita:
nome_tabelas = ['address', 'user_', 'pet_register', 'specie', 'race', 'hair_color', 'size', 'weight', 'vaccine', 'user_phone']

#Pegando os valores de cada tabela:
for nome_tabela in nome_tabelas:
    #Definindo o caminho do excel:
    excel = '../../BI/Dados/' + nome_tabela + '.xlsx'

    #Lendo as tabelas e adicionando os dados para utilizar no BI:
    cursor.execute(f'SELECT * FROM {nome_tabela}')
    dados_tabela = cursor.fetchall()

    #Pegando o nome das colunas da tabela usando o description que a para cada valor, a posição [0] é equivalente ao nome da coluna:
    nome_colunas = [coluna[0] for coluna in cursor.description]

    # Leia os dados do banco de dados em um DataFrame
    df = pd.DataFrame.from_records(dados_tabela, columns=nome_colunas)

    #Salvando os dados no excel:
    df.to_excel(excel, index=False)

#Fechando a conexão com o banco:
cursor.close()
conn.close()