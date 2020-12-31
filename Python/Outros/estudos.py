#Definindo a lista de cores
cores = {'limpa':'\033[m',
         'azul':'\033[34m',
         'amarelo':'\033[33m',
         'pretoebranco':'\033[7;30m',}

a = input("Informe um número")
b = input("Informe um número")
c = input("Informe um número")

#Verificando o menor
menor = a
if (b < a) and (b < c):
    menor = b
if (c < a) and (c < b):
    menor = c

#Verificando o maior
maior = a
if (b > a) and (b > c):
    maior = b
if (c > a) and (c > b):
    maior = c

print('O menor valor informado foi: {} '.format(cores['azul']),menor)

print('O MAIOR valor informado foi: {} '.format(cores['amarelo']),maior)