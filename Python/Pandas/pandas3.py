import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

df = pd.read_csv("c:\\Projetos\_Arquivos\\911.csv",sep=",")
print(df['title'].iloc[0].split(':')[0])
df['Reason'] = df['title'].apply(lambda title: title.split(':')[0])

print(df['Reason'])
sns.countplot(x='Reason',data=df)