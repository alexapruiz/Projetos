from random import randint

print('*****************************')
print('*    Jogo da adivinhação    *')
print('*****************************')

qtde_chances=3
numero_secreto = randint(1,9)

acertou = False
while (acertou == False) and (qtde_chances > 0):
    escolha = int(input("Adivinhe um número entre 1 e 9: "))

    if (escolha == numero_secreto):
        print("Você acertou !!!")
        acertou = True
    else:
        print("Você errou. Tente novamente... ")
        qtde_chances = qtde_chances - 1


if acertou == False:
    print("")
    print("O número correto era : " + str(numero_secreto))