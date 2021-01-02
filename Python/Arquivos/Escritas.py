#Abrindo um arquivo no modo Write
#arquivo = open('palavras.txt', 'w')


#Abrindo um arquivo no modo Append
arquivo = open("palavras.txt", "w")

#Gravando no arquivo
arquivo.write("banana \n")
arquivo.write("maça \n")
arquivo.write("mamao \n")
arquivo.write("uva \n")
arquivo.write("laranja \n")
arquivo.close

# Arquivos de imagem
#imagem = open("Fotos.jpg", "rb")
arquivo = open("palavras.txt", "r")
#print("Imprimindo o arquivo palavras.txt...")
#print (arquivo.read())
x = 1
for linha in arquivo:
    print("lendo a linha : " + str(x))
    x = x + 1
    print(linha)

arquivo.close()

print("Usando o WITH")

with open('palavras.txt') as arquivo:
    for linha in arquivo:
        print(linha)

print("    _______________         ")
print("   /               \        ")
print("  /                 \       ")
print("//                   \/\    ")
print("\|   XXXX     XXXX   | /    ")
print(" |   XXX       XXX   |      ")
print(" |                   |      ")
print(" \__      XXX      __/      ")
print("   |\     XXX     /|        ")
print("   | |           | |        ")
print("   | I I I I I I I |        ")
print("   |  I I I I I I  |        ")
print("   \_             _/        ")
print("     \_         _/          ")
print("       \_______/            ")