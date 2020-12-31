import matplotlib.pyplot as plt

a=[]
b=[]
c=[]
d=[]

#Preenchendo as listas com loop
for x in range(0,12,1):
    a.append(x)
    b.append(x * 1.2)
    c.append(x * 1.3)
    d.append(x * 1.5)

eixo_X = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez']
eixo_Y = ['1','2','3','4','5','6','7','8','9','10','11','12']

plt.scatter(a,b, s=20,c='blue', marker='^')
plt.scatter(a,c, s=20,c='red', marker='+')
plt.scatter(b,c, s=20,c='green', marker='o')
plt.scatter(c,d, s=20,c='black', marker='x')
plt.xticks(a,eixo_X)
plt.yticks(b,eixo_Y)
plt.legend('leg 1', loc='upper left', frameon=True)
plt.legend('leg 2', loc='upper left', frameon=True)
plt.legend('leg 3', loc='upper left', frameon=True)
plt.legend('leg 4', loc='upper left', frameon=True)
plt.show()