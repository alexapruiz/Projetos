import pandas as pd
from sklearn.svm import LinearSVC
from sklearn.model_selection import train_test_split
from sklearn.metrics import accuracy_score

#Definindo uma lista para nomear as colunas
colunas = ['idade','genero','poliuria','Polidipsia','perda_peso','fraqueza','polifagia','tordo_genital','embacamento_visual','coceira','irritabilidade','cura_retardada','paresia_parcial','rigidez_muscular','alopecia','obesidade','classe']

#Explicando os termos médicos
#poliuria - produção de urina acima de 2,5 litros por dia
#Polidipsia - sede excessiva
#polifagia - fome excessiva
#tordo_genital - candidíase (mulheres)
#paresia_parcial - restrição/diminuição do movimento
#alopecia - queda de cabelo

#Abrindo um arquivo csv
arquivo = 'c:\\projetos\_Arquivos\\diabetes_data_upload_num.csv'
dados = pd.read_csv(arquivo, header=1, sep=';', names=colunas, low_memory=False)

#Explorando os dados
#Verificando se tem registros de tordo_genital em homens
total_registros = dados.count()[0]
total_homens = dados[(dados.genero == 1)].count()[0]
total_mulheres = dados[(dados.genero == 2)].count()[0]
print("Percentual de Homens com diabetes: " + str(dados[(dados.genero == 1) & (dados.classe == 1) ].count()[0] / total_homens * 100))
print("Percentual de Mulheres com diabetes: " + str(dados[(dados.genero == 2) & (dados.classe == 1) ].count()[0] / total_mulheres * 100))
print('')
print("Percentual de Obesos com diabetes: " + str(dados[(dados.obesidade == 1) & (dados.classe == 1) ].count()[0] / total_registros * 100))
print('')
print("Idade média dos diabéticos: " + str(dados['idade'].mean()))
print('')
print("Percentual de pessoas diabéticas e paresia parcial: " + str(dados[(dados.paresia_parcial == 1) & (dados.classe == 1) ].count()[0] / total_registros * 100))
print('')
print("Percentual de Diabéticos com Polidipsia: " + str(dados[(dados.Polidipsia == 1) & (dados.classe == 1) ].count()[0] / total_registros * 100))
print('')
print('Idade Mínima: ' + str(dados['idade'].min()))
print('Idade Máxima: ' + str(dados['idade'].max()))
print('Idade Média: ' + str(dados['idade'].mean()))
print('')
#Definindo os dados de treino e teste
x = dados[['idade','genero','poliuria','Polidipsia','perda_peso','fraqueza','polifagia','tordo_genital','embacamento_visual','coceira','irritabilidade','cura_retardada','paresia_parcial','rigidez_muscular','alopecia','obesidade']]
y = dados['classe']
treino_x, teste_x, treino_y, teste_y = train_test_split(x,y,test_size=0.2,random_state=42)

#Criando o modelo
modelo = LinearSVC()

#Treinando o modelo
modelo.fit(treino_x,treino_y)

#Criando as previsões
previsoes = modelo.predict(teste_x)

#Exibindo as previsões
print(previsoes)

#Calculando a eficácia das previsões
print(accuracy_score(teste_y,previsoes) * 100)