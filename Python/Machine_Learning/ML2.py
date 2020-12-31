from sklearn.svm import LinearSVC
from sklearn.metrics import accuracy_score
# Caracteristicas
# Tem pelo longo
# Tem perna curta
# Faz au au
porco1 = [0,1,0]
porco2 = [0,0,0]
porco3 = [0,1,0]

cachorro1 = [0,1,1]
cachorro2 = [1,0,1]
cachorro3 = [1,1,1]

treinox = [porco1,porco2,porco3,cachorro1,cachorro2,cachorro3]
treinoy = [1,1,1,0,0,0] # 1 - porco, 0 - cachorro

modelo = LinearSVC()
modelo.fit(treinox,treinoy)

misterio1 = [1,1,1]
misterio2 = [1,1,0]
misterio3 = [0,1,1]

teste_x = [misterio1,misterio2,misterio3]
teste_y = [0,0,0]

previsoes = modelo.predict(teste_x)
print(previsoes)

print(accuracy_score(teste_y, previsoes) * 100)