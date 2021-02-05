#Tuplas
tp1 = (1,2,3,4)
tp2 = (10,20,30,40)

print(tp1 + tp2)
print(tp1 * 2)

#Verifica se o item está na tupla
print(1 in tp1)
print(5 in tp1)

print()
print('3')
print()

#Dicionarios
dic1 = {'SP': 'São Paulo', 'RJ':'Rio de Janeiro', 'MG':'Minas Gerais'}

for chave in dic1.keys():
    print(chave + ' - ' + dic1[chave])

print()
print('4')
print()

for palavra in dic1.values():
    print(palavra + ' --')

print(dic1['SP'])
print(dic1.items())