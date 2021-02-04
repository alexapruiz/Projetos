from matplotlib import pyplot as plt

# Eixo_x, Eixo_y
vendas = [3000, 2300, 1000, 500]
labels = ['E-commerce', 'Loja Física', 'e-mail', 'Marketplace']

plt.pie(vendas, labels=labels)
plt.show()