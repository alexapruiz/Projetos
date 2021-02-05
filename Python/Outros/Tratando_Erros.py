import sys
x = 40
y = 0

try:
    print(x / y)
except ZeroDivisionError:
    print('Deu erro de divisão por zero')
except:
    print('Deu erro: ' + str(sys.exc_info()))
finally:
    print('Mensagem final que sempre aparece')