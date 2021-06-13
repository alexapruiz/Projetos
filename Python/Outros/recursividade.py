import numpy as np

def fib(n):
    if (n == 0) or (n == 1):
        return 1

    resultado = fib(n-1) + fib(n-2)
    return resultado

fibonacci = np.arange(7)
for element in fibonacci:
    print(fib(element))