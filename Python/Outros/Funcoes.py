def velocidade_media(distancia, tempo, 	*args):
    return (distancia / tempo) * int(args[0]) * int(args[2])

print(velocidade_media(100,1,4,33,0))