import pandas as pd

arquivo = pd.read_csv('c:\projetos\_arquivos\Cardio\cardio.csv', sep=';')

# Selecionando apenas as pessoas com idade > 20.000
velhos = arquivo.loc[arquivo['age'] > 20000]
#print(velhos)

# Apagar uma coluna
arquivo.drop('gluc',1,inplace=True)
#print(arquivo)

#Apagar os registros dos velhos
arquivo.drop(arquivo[arquivo.age > 20000].index, inplace=True)
#print(arquivo)
#print('Média de idade sem os velhos: ' + str(arquivo['age'].mean()))

#Exibindo apenas as colunas 1,2,3 - A coluna 4 não está incluída
previsores = arquivo.iloc[:,1:4]
#print(previsores)

#Definindo a coluna 'cardio' como o target
classe = arquivo.iloc[:,-1]
#print(classe)

from sklearn.preprocessing import Imputer
Imputer = Imputer(missing_values = 'NaN', strategy='mean', axis=0)