class Pilha:

    def __init__(self):
        self.__pl = []

    def empilhar(self,numero):
        self.__pl.append(numero)

    def desempilhar(self):
        valor = self.__pl[-1]
        del self.__pl[-1]
        return valor

    def mostrar_pilha(self):
        print(self.__pl)

obj_pilha = Pilha()

obj_pilha.empilhar(10)
obj_pilha.empilhar(15)
obj_pilha.empilhar(19)
obj_pilha.mostrar_pilha()