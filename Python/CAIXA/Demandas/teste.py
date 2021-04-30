def DefinePeriodo(PRAZO_FINAL):
    if int(PRAZO_FINAL[8:10]) > 20:
        if (int(PRAZO_FINAL[5:7]) < 12):
            return PRAZO_FINAL[:4] + '-' + str('0' + str(int(PRAZO_FINAL[5:7]) + 1))[-2:]
        else:
            return str(int(PRAZO_FINAL[:4]) + 1) + '-' + str('01')
    return PRAZO_FINAL[:-3]

print(DefinePeriodo('2016-11-21'))