import six
import sys
sys.modules['sklearn.externals.six'] = six
import mlrose

pessoas = [('Lisboa', 'LIS'),
           ('Madrid', 'MAD'),
           ('Paris', 'CDG'),
           ('Dublin', 'DUB'),
           ('Bruxelas', 'BRU'),
           ('Londres', 'LHR')]

destino = 'FCO'

voos = {}
for linha in open('c:/Projetos/Python/Curso_ML_IA/voos.txt'):
    origem, destino, saida, chegada, preco = linha.split(',')
    voos.setdefault((origem, destino), [])
    voos[(origem, destino)].append((saida, chegada, int(preco)))

def imprimir_voos(agenda):
    id_voo = -1
    total_preco = 0
    for i in range(len(agenda) // 2):
        nome = pessoas[i][0]
        origem = pessoas[i][1]
        id_voo += 1
        ida = voos[(origem, destino)][agenda[id_voo]]
        total_preco += ida[2]
        id_voo += 1
        volta = voos[(destino, origem)][agenda[id_voo]]
        total_preco += volta[2]
        print('%10s%10s %5s-%5s %3s %5s-%5s %3s' % (nome, origem, ida[0], ida[1], ida[2],volta[0], volta[1], volta[2]))
    print('Preço total: ', total_preco)


def fitness_function(agenda):
    id_voo = -1
    total_preco = 0
    for i in range(len(agenda) // 2):
        origem = pessoas[i][1]
        id_voo += 1
        ida = voos[(origem, destino)][agenda[id_voo]]
        total_preco += ida[2]
        id_voo += 1
        volta = voos[(destino, origem)][agenda[id_voo]]
        total_preco += volta[2]

    return total_preco

#Usando algoritmos de otimização
#Busca randomicamente pela melhor solução
#fitness = mlrose.CustomFitness(fitness_function)
#problema = mlrose.DiscreteOpt(length=12, fitness_fn=fitness,maximize = False, max_val = 10)

#Hill Climb
#melhor_solucao, melhor_custo = mlrose.hill_climb(problema, random_state = 1)
#imprimir_voos(melhor_solucao)

#Simulated Annealing
#melhor_solucao, melhor_custo = mlrose.simulated_annealing(problema)
#imprimir_voos(melhor_solucao)

#Algorimo Genetico
#melhor_solucao, melhor_custo = mlrose.genetic_alg(problema, pop_size=500, mutation_prob=0.2)
#imprimir_voos(melhor_solucao)


#Exercicio - Caminhão com eletrônicos
produtos = [('Refrigerador A', 0.751, 999.90),
            ('Celular', 0.0000899, 2911.12),
            ('TV 55', 0.400, 4346.99),
            ('TV 50', 0.290, 3999.90),
            ('TV 42', 0.200, 2999.00),
            ('Notebook A', 0.00350, 2499.90),
            ('Ventilador', 0.496, 199.90),
            ('Microondas A', 0.0424, 308.66),
            ('Microondas B', 0.0544, 429.90),
            ('Microondas C', 0.0319, 299.29),
            ('Refrigerador B', 0.635, 849.00),
            ('Refrigerador C', 0.870, 1199.89),
            ('Notebook B', 0.498, 1999.90),
            ('Notebook C', 0.527, 3999.00)]

espaco_disponivel = 3

def imprimir_solucao(solucao):
    for i in range(len(solucao)):
        if solucao[i] == 1:
            print('%s - %s' % (produtos[i][0], produtos[i][2]))

def fitness_function(solucao):
    custo = 0
    soma_espaco = 0
    for i in range(len(solucao)):
        if solucao[i] == 1:
            custo += produtos[i][2]
            soma_espaco += produtos[i][1]
    if soma_espaco > espaco_disponivel:
        custo = 1
    return custo

fitness = mlrose.CustomFitness(fitness_function)
problema = mlrose.DiscreteOpt(length = 14, fitness_fn = fitness,maximize = True, max_val = 2)

melhor_solucao, melhor_custo = mlrose.hill_climb(problema)
melhor_solucao, melhor_custo = mlrose.simulated_annealing(problema)

melhor_solucao, melhor_custo = mlrose.genetic_alg(problema, pop_size=500, mutation_prob=0.2)
print(melhor_solucao, melhor_custo)