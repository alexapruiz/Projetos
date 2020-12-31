#Jogo da forca
import random

#define as palavras possíveis
palavras = ("ATHENA", "APOLLO", "ROCKY", "LILI", "PICOLA", "FLOKY", "SMURF", "NINA", "MILI", "MARIA")

#Escolhe uma palavra
palavrasecreta = random.choice(palavras)

#Define a máscara
letrascertas = []
x=1
while x <= len(palavrasecreta):
    letrascertas.append("_")
    x = x + 1

#Exibe a máscara para o usuário
print(letrascertas)
acertos = 0
while (acertos < len(palavrasecreta)):
    palpite = input("Informe uma letra... ")
    posicao = 0
    for letra in palavrasecreta:
        if (palpite.upper() == letra.upper()):
            #Acertou a letra
            acertos = acertos + 1
            letrascertas[posicao] = str(palpite.upper())

        posicao = posicao + 1

    print(letrascertas)
print("************************ ")
print("Parabéns, você acertou. A palavra secreta era: " + palavrasecreta)