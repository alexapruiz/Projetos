from tkinter import *
from tkinter import ttk
import tkinter

def Exibe():
    Label1['text'] = 'Alex Ruiz'

#Criando a janela
janela = Tk()

#Definindo propriedades da janela
janela.title('Exemplo de Formulário')
janela.geometry('800x450')
#janela.resizable(False,False)

#Criando objetos na janela
Label1 = Label(janela,text='Rótulo')
Label1.grid(row=1, column=1)

Text1 = Entry(janela, textvariable='Alex')
Text1.grid(row=1, column=2)

Botao1 = Button(janela,command=Exibe, text='Calcular')
Botao1.grid(row=1, column=3)

#Criando um frame
#frame1 = Frame(janela, width=200, height=200, bg='white').grid(row=0, column=0)
#frame2 = Frame(janela, width=200, height=200, bg='white').grid(row=0, column=1)
#frame3 = Frame(janela, width=200, height=200, bg='white').grid(row=0, column=2)
#Label(frame1, text='Labels no frame 1').grid(row=0, column=0)
#Label(frame2, text='Label no frame 2').grid(row=0, column=1)
#Label(frame3, text='Label no frame 3').grid(row=0, column=2)


#Criando um Treeview
tree = ttk.Treeview(janela,selectmode='browse',column=('Coluna1','Coluna2','Coluna3', 'Coluna4'), show='headings')

tree.column('Coluna1',width=70, minwidth=40, stretch=NO)
tree.heading('#1',text='Código')

tree.column('Coluna2',width=250, minwidth=200, stretch=NO)
tree.heading('#2',text='Nome')

tree.column('Coluna3',width=70, minwidth=40, stretch=NO)
tree.heading('#3',text='Idade')

tree.column('Coluna4',width=300, minwidth=250, stretch=NO)
tree.heading('#4',text='Endereço')
tree.grid(row=6,column=4)

#Inserindo dados na Treeview
elementos = ['1','Alex Ruiz','42','Rua Avedis Kamalakian, 764']
for x in range(1,5):
    tree.insert('',END,values=elementos,tag='1')

#Exibindo a janela
janela.mainloop()