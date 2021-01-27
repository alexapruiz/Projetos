import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
from plotly.offline import download_plotlyjs, init_notebook_mode, plot, iplot
import cufflinks as cf

from plotly import __version__
print(__version__)

df = pd.DataFrame(np.random.randn(100,4), columns='A B C D'.split())
df2 = pd.DataFrame({'Categoria':['A', 'B', 'C'], 'Valores':[32,43,50]})

init_notebook_mode(connected=True)
cf.go_offline()
df.iplot(kind='scatter', x='A', y='B')