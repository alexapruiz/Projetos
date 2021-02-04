import tkinter as tk
from matplotlib import pyplot as plt

class App(tk.Frame):
    def __init__(self, master=None):
        tk.Frame.__init__(self, master)
        self.pack()
        self.criarbotoes("Exibe Gráfico")
        self.crialabel("Nome do campo")
        self.CriaTextBox()
        self.BotaoSair()

    def criarbotoes(self,caption):
        self.btCriar = tk.Button(self, text=caption, fg="black", command=self.ExibeGrafico())
        self.btCriar["text"] = caption
        self.btCriar.pack(side="top")

    def crialabel(self,texto):
        self.label = tk.Label(self)
        self.label["text"] = texto
        self.label.pack(side="top")

    def CriaTextBox(self):
        self.edit = tk.Entry(self)
        self.edit.pack(side="top")

    def BotaoSair(self):
        self.btSair = tk.Button(self, text="sair", fg="red", command=root.destroy)
        self.btSair.pack(side="bottom")

    def ExibeGrafico(self):
        vendas = [3000, 2300, 1000, 500]
        labels = ['E-commerce', 'Loja Física', 'e-mail', 'Marketplace']
        plt.pie(vendas, labels=labels)
        plt.show()

root = tk.Tk()
# criando a aplicação
minhaAplicacao = App(master=root)
minhaAplicacao.master.title("Exemplo de tela")
minhaAplicacao.master.maxsize(800, 600)
minhaAplicacao.master.geometry("800x600")
minhaAplicacao.master.positionfrom=100

# inicia a aplicacao
minhaAplicacao.mainloop()