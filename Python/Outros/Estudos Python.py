def Fatorial(x):
  #Calcular n!
  i=1
  fatorial=1
  while i <= x:
    fatorial = fatorial * i
    i = i + 1

  return fatorial

resultado = Fatorial(3)
resultado

def Prob(x,p,n,tipo_saida):

  n_x = 0
  n_x = Fatorial(n) / (Fatorial(x) * (Fatorial(n-x)))

  if tipo_saida == 1:
    return (n_x * (p ** x) * (1 - p) ** (n - x))
  else:
    return (str((n_x * (p ** x) * (1 - p) ** (n - x)) * 100) + "%")
    

#print(Prob(3,0.5,5,1))

print(Prob(0,0.25,12,1))