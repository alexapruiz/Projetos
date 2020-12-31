import time
import datetime

data_futura = datetime.datetime(2020,12,31,23,59,59)
data_atual = datetime.datetime.today()
# Quanto tempo falta para 31/12/2020
falta = data_futura - data_atual
print("Faltam '" + str(falta) + "' para o final do ano")