from collections import Counter
import matplotlib.pyplot as plt

#x = [10,20,30,40]
#y = [1,4,2,3]
plt.bar([10,20,30,40],[1,4,2,3])
plt.axis([0, 50, 0, 10])
plt.title("Histograma da Contagem de Amigos")
plt.xlabel("# de amigos")
plt.ylabel("# de pessoas")
plt.show()