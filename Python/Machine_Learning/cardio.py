import pandas as pd
from sklearn.svm import LinearSVC
from sklearn.model_selection import train_test_split
from sklearn.metrics import accuracy_score

colunas = ['id','idade','genero','altura','peso','pressao_max','pressao_min','colesterol','diabetes','fuma','bebe','ativo','coracao']
cardio = pd.read_csv('c:\\Projetos\\_Arquivos\\Cardio\\cardio2.csv', sep=';',names=colunas,header=1)
x = cardio[['idade','peso','pressao_max','colesterol','diabetes','fuma']]
y = cardio["coracao"]

treino_x, teste_x, treino_y, teste_y = train_test_split(x,y,test_size=0.1,random_state=42)

modelo = LinearSVC()
modelo.fit(treino_x,treino_y)
previsoes = modelo.predict(teste_x)
print(previsoes)
print(accuracy_score(teste_y,previsoes) * 100)