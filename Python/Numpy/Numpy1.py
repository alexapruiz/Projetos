import numpy as np

#Lista com 1 dimensão
minha_lista = [1,2,3]
print(minha_lista)
print(np.array(minha_lista))

#Lista com 2 dimensões
minha_matriz = [[1,2,3],[4,5,6],[7,8,9]]
print(minha_matriz)
print(np.array(minha_matriz))

#Funções do Numpy
print(np.arange(0,10))
print(np.arange(0,10,2))
print(np.zeros(4))
print(np.ones(4))
print(np.ones((3,3)))
print()
print(np.eye(3))
print(np.eye(4))
print()
print(np.linspace(0,10,2))
print(np.linspace(0,10,3))
print()
print(np.random.rand(5))
print(np.random.randn(4))
print(np.random.randint(0,100,10))
print(np.random.randint(0,100,10))
print()
#arr = np.random.rand(25)
#print(arr)
arr = np.random.rand(5,5)
#arr = arr.reshape(5,5)
print(arr)
print(arr.shape)
print(arr.max())
print(arr.argmax())
print(arr[:][:])