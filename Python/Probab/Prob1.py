from scipy.stats import binom

#Probabilidade de jogar uma moeda 5 vezes e dar cara 3 vezes
print(binom.pmf(3,5,0.5))

#Passar por 4 semáforos de 4 tempos, qual a probabilidade de pegar sinal verde
# eventos , experimentos, probabilidade de cada evento
#0-Nenhuma, 1-Uma vez, 2-Duas vezes, 3-Três vezes, 4-Quatro vezes
print(float(binom.pmf(0,4,0.25)))
print(float(binom.pmf(1,4,0.25)))
print(float(binom.pmf(2,4,0.25)))
print(float(binom.pmf(3,4,0.25)))

