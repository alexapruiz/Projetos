import numpy as np

class VetorOrdenado:
    def __init__(self, capacidade):
        self.capacidade = capacidade
        self.ultima_posicao = -1
        self.valores = np.empty(self.capacidade , dtype = int)

    def Imprime(self):
        if self.ultima_posicao == -1:
            print('O vetor está vazio')
        else:
            for i in range(self.ultima_posicao + 1):
                print(i,' - ', self.valores[i])

    def Insere(self,valor):
        if self.ultima_posicao == self.capacidade -1:
            print('Capacidade máxima atingida')
            return

        posicao = 0
        for i in range(self.ultima_posicao + 1):
            posicao = i
            if self.valores[i] > valor:
                break

        x = self.ultima_posicao
        while x >= posicao:
            self.valores[x+1] = self.valores[x]
            x -= 1

        self.valores[posicao] = valor
        self.ultima_posicao += 1


vetor = VetorOrdenado(10)

vetor.Insere(7)
vetor.Insere(3)
vetor.Insere(6)
vetor.Insere(5)
vetor.Insere(1)
vetor.Insere(8)
vetor.Imprime()