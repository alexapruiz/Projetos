from sklearn.svm import LinearSVC
from sklearn.metrics import accuracy_score

# Informações - pessoa = [# Fuma , # Bebe , # Obeso , #cardiaco]

pessoa1 = [1,1,1]
pessoa2 = [1,1,1]
pessoa3 = [1,1,0]
pessoa4 = [1,0,0]
pessoa5 = [0,0,0]
pessoa6 = [0,0,1]
pessoa7 = [0,1,1]
pessoa8 = [1,1,1]
pessoa9 = [0,0,0]

treinox = [pessoa1,pessoa2,pessoa3,pessoa4,pessoa5,pessoa6,pessoa7,pessoa8,pessoa9]
treinoy = [1,1,1,0,0,0,1,1,0] #1 - tem doenca coração, 0 - não tem doença coração

modelo = LinearSVC()
modelo.fit(treinox,treinoy)

previsoes = modelo.predict(treinox)
print(previsoes)
print(accuracy_score(treinoy, previsoes) * 100)
