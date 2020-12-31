from pylab import *
ent = arange(0., 20.1, .1)
# Calcula os cossenos da entrada
sai = cos(ent)
# Plota a curva
plot(ent, sai)
# Texto para o eixo X
xlabel('entrada')
# Texto para o eixo Y
ylabel('cosseno')
# Texto no topo da figura
title('Cossenos')
# Ativa a grade
grid(True)

# Apresenta a figura resultante na tela
show()