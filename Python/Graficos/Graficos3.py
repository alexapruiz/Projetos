# -*- coding: latin1 -*-
import os
import matplotlib as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_agg import FigureCanvasAgg

def pie(filename, labels, values):
    """
    Gera um diagrama de Pizza e salva em arquivo.
    """
    # Use a biblioteca Anti-Grain Geometry
    plt.use('Agg')
    # Cores personalizadas
    colors = ['seagreen', 'lightslategray', 'lavender','khaki', 'burlywood', 'cornflowerblue']
    # Altera as opções padrão
    plt.rc('patch', edgecolor='#406785', linewidth=1, antialiased=True)
    # Altera as dimensões da imagem
    plt.rc('figure', figsize=(8., 7.))
    # Inicializa a figura
    fig = Figure()
    fig.clear()
    axes = fig.add_subplot(111)
    if values:
        #Diagrama
        chart = axes.pie(values, colors=colors, autopct='%2.0f%%')
        # Legenda
        pie_legend = axes.legend(labels)
        pie_legend.pad = 0.3

        # Altera o tamanho da fonte
        for i in range(len(chart[0])):
            chart[-1][i].set_fontsize(12)
            pie_legend.texts[0].set_fontsize(10)
    else:
        #Mensagem de erro
        #Desliga o diagrama
        axes.set_axis_off()
        #Mostra a mensagem
        axes.text(0.5, 0.5, 'Sem dados', horizontalalignment='center', verticalalignment='center',fontsize=32, color='#6f7c8c')

    # Salva a figura
    canvas = FigureCanvasAgg(fig)
    canvas.print_figure(os.path.dirname(__file__) + "\\" + filename, dpi=600)

if __name__ == '__main__':
    #Testes
    pie('fig1.png', [], [])
    pie('fig2.png', ['A', 'B', 'C', 'D', 'E'],[6.7, 5.6, 4.5, 3.4, 2.3])
