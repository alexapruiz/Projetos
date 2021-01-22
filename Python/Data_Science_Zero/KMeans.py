import numpy as np

class KMeans:
    """executa agrupamentos k-means"""
    def __init__(self, k):
        self.k = k # número de agrupamentos
        self.means = None # ponto médio de agrupamentos


    def classify(self, input):
        """retorna o índice do agrupamento mais próximo da entrada"""
        return min(range(self.k),
        key=lambda i: squared_distance(input, self.means[i]))


    def train(self, inputs):
        # escolha pontos k aleatórios como média inicial
        self.means = random.sample(inputs, self.k)
        assignments = None
        while True:
            # encontre novas associações
            new_assignments = map(self.classify, inputs)
            # se nenhuma associação mudou, terminamos.
            if assignments == new_assignments:
                return

            # senão, mantenha as novas associações,
            assignments = new_assignments
            # e compute novas médias, baseado nas novas associações
            for i in range(self.k):
                # encontre todos os pontos associados ao agrupamento i
                i_points = [p for p, a in zip(inputs, assignments) if a == i]
                # certifique-se que i_points não está vazio,
                # para não dividir por 0
                if i_points:
                    self.means[i] = vector_mean(i_points)

np.random.seed(0) # para que você consiga os mesmos
clusterer = KMeans(3) # resultados que eu
clusterer.train(inputs)
print(clusterer.means)