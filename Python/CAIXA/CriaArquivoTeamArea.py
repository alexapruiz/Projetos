import pandas as pd

#Definindo uma lista para nomear as colunas
colunas = ['area_projeto','team_area','matricula','funcao']

#Abrindo um arquivo csv
arquivo = pd.read_csv('c:\\projetos\_Arquivos\\CAIXA\\teste2.txt', sep=';', names=colunas)
print(arquivo.info())
print(arquivo.head)
print(arquivo.describe)
print(arquivo['area_projeto'])
#print(arquivo['area_projeto'].value_counts())

print(arquivo.groupby(['area_projeto']))
