#Importação das bibliotecas
import pandas as pd
import seaborn as sns
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import accuracy_score
import pickle

#Importação dos Dados
clientes = pd.read_csv('clientes.csv')

#verifica se há valores duplicados
clientes.duplicated().sum()

#Substituindo caracteres por números
clientes['sexo']= clientes['sexo'].map({'Male':0, 'Female':1})
clientes['estado_civil']= clientes['estado_civil'].map({'No':0, 'Yes':1})
clientes['aprovacao_emprestimo']= clientes['aprovacao_emprestimo'].map({'N':0, 'Y':1})

# Eliminando valores Nulos
clientes = clientes.dropna()
clientes.isnull().sum()

# Separando variáveis Explicativas e Variável TARGET
X = clientes[['sexo', 'estado_civil', 'renda', 'emprestimo', 'historico_credito']]
y = clientes.aprovacao_emprestimo

# Realizando Amostragem dosa Dados
x_train, x_teste, y_train, y_teste = train_test_split(X,y, test_size = 0.2, random_state = 7)

# Criando a Máquina Preditiva com o Random Forest
maquina = RandomForestClassifier()
maquina.fit(x_train, y_train)

# Criando a Máquina Preditiva com o Random Forest
maquina = RandomForestClassifier(max_depth=4, random_state = 7)
maquina.fit(x_train, y_train)
pred_maquina_treino = maquina.predict(x_train)
pred_maquina_teste = maquina.predict(x_teste)

#Comando de Salvamento da Máquina Preditiva
pickle_out = open("maquina_preditiva.pkl", mode = "wb")
pickle.dump(maquina, pickle_out)
pickle_out.close()
print('Máquina preditiva criada com sucesso!!!')