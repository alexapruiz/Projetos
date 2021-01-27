import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from plotly.offline import download_plotlyjs, init_notebook_mode, plot, iplot
import cufflinks as cf
import chart_studio.plotly
import plotly.graph_objs as go
from plotly.offline import download_plotlyjs, init_notebook_mode, plot, iplot

init_notebook_mode(connected=True)
cf.go_offline()
df = pd.DataFrame(np.random.randn(100,4), columns='A B C D'.split())


data = dict(type='choropleth',
            locations=['Brazil','Canada','NY'],
            locationmode='country names',
            colorscale='Portland',
            text=['Texto1','Texto2','Texto3'],
            z=[1.0,2.0,3.0],
            colorbar={'title':'Título da barra de cores'})

layout = dict(geo = {'scope':'world'})
choromap = go.Figure(data = [data], layout = layout)
iplot(choromap)