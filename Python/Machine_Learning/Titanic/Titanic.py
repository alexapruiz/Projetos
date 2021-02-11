import pandas as pd
import numpy as np
from sklearn.ensemble import RandomForestClassifier

train = pd.read_csv('..\\..\\_Arquivos\\Titanic\\train.csv')
test = pd.read_csv('..\\..\\_Arquivos\\Titanic\\test.csv')

print(train.head())

modelo = RandomForestClassifier(n_estimators=100, n_jobs=-1,random_state=0)

