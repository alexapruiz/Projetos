import pandas as pd

#Definindo uma lista para nomear as colunas
colunas = ['id','idade','genero','altura','peso','pressao_min','pressao_max','colesterol','glicemia','fumante','alcool','atividade_fisica','doenca_coracao']

#Abrindo um arquivo csv
arquivo = 'c:\\projetos\_Arquivos\\Cardio\\cardio_original.csv'
#dados = pd.read_csv(arquivo, header=1, sep=';', names=colunas, low_memory=False, dtype={"id": int, "idade": int, "genero":int, "altura":int, "peso":float, "pressao_min":int, "pressao_max":int, "colesterol":int, "glicemia":int, "fumante":int, "alcool":int, "atividade_fisica":int, "doenca_coracao":int})
dados = pd.read_csv(arquivo, header=1, sep=';', names=colunas, low_memory=False)
#df = pd.DataFrame(dados)
print(dados.head(5))

#Loop para contagens e calculos
#qtde_fumantes = 0
#qtde_nao_fumantes = 0
#qtde_peso = 0
#for item in dados.peso:
#    if (item > 10000):
#        print(item)
#        qtde_peso += 1

#print('Qtde de fumantes: ' + str(qtde_fumantes))
#print('Qtde de NÃO fumantes: ' + str(qtde_nao_fumantes))

#Definindo um filtro para pessoas com mais de 60 anos e fumante
#velhos_fumantes = dados[(dados.idade > 21900) & (dados.fumante == 1)]
#print(velhos_fumantes.sort_values('idade', ascending=False).head(10))

#Tudo numa linha só
print("Total: " + str(dados.count()[0]))
print("Qtde de Pessoas Totalmente Saudáveis: " + str(dados[(dados.colesterol == 1) & (dados.glicemia == 1) & (dados.fumante == 0) & (dados.alcool == 0) & (dados.atividade_fisica == 1) & (dados.doenca_coracao == 0) & (dados.pressao_max <= 120)].count()[0]))
print("Qtde de Fumantes: " + str(dados[(dados.fumante == 1)].count()[0]))
print("Qtde de Cardíacos: " + str(dados[(dados.doenca_coracao == 1)].count()[0]))
print("Qtde de Fumantes E Cardíacos: " + str(dados[(dados.fumante == 1) & (dados.doenca_coracao == 1)].count()[0]))
print("Qtde de Sedentários: " + str(dados[(dados.atividade_fisica == 0)].count()[0]))
print("Qtde de Fumantes E Sedentários: " + str(dados[(dados.fumante == 1) & (dados.atividade_fisica == 0)].count()[0]))
print("Qtde de Sedentários E Cardíacos: " + str(dados[(dados.atividade_fisica == 0) & (dados.doenca_coracao == 1)].count()[0]))
print("Qtde de Sedentários E Cardíacos E Fumantes: " + str(dados[(dados.atividade_fisica == 0) & (dados.doenca_coracao == 1) & (dados.fumante == 1)].count()[0]))
print("Qtde de Pressão Alta (acima de 15): " + str(dados[(dados.pressao_max > 150)].count()[0]))
print("Qtde de Pressão Alta e Coração OK: " + str(dados[(dados.pressao_max > 150) & (dados.doenca_coracao == 0)].count()[0]))
print("Qtde de Cardíacos com Pressão Alta (acima de 15): " + str(dados[(dados.doenca_coracao == 1) & (dados.pressao_max > 150)].count()[0]))
print("Qtde de Fumantes Saudáveis (Sem doenças / sintomas): " + str(dados[(dados.fumante == 1) & (dados.doenca_coracao == 0) & (dados.pressao_max <= 120) & (dados.colesterol == 1) & (dados.glicemia == 1) ].count()[0]))

print('Idade Média: ' + str(dados['idade'].mean()))