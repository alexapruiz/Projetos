#Método tradicional
quadrados = []
for k in range(1,51):
    quadrados.append(k ** 2)

print(quadrados)

#Método simplificado
quadrados = [k ** 2 for k in range(1,51)]
print(quadrados)