from sklearn import tree

def ExecutaArvore(clf,peso,tipo_casca):
    fruta = clf.predict([[peso, tipo_casca]])

    if fruta == 5:
        print('Peso: ' + str(peso) + ' - Casca: ' + RetornaTipoCasca(tipo_casca) + ' -> Maçã')
    elif fruta == 10:
        print('Peso: ' + str(peso) + ' - Casca: ' + RetornaTipoCasca(tipo_casca) + ' -> Laranja')


def RetornaTipoCasca(tipo_casca):
    if tipo_casca == 1:
        return 'Lisa'
    else:
        return 'Irregular'

maca = 5
laranja = 10
lisa = 1
irregular = 0

X = [[90, lisa], [100, lisa], [250, irregular], [370, irregular]]
Y = [maca, maca, laranja, laranja]
clf = tree.DecisionTreeClassifier()
clf = clf.fit(X, Y)

ExecutaArvore(clf,100,irregular)
ExecutaArvore(clf,100,lisa)
ExecutaArvore(clf,300,irregular)
ExecutaArvore(clf,370,lisa)