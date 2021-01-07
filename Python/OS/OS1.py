import PySimpleGUI
import os

class Tela :

    def __init__ (self) :
        layout = [
            [PySimpleGUI.Text ('ip'),PySimpleGUI.Input(key='ip')],
            [PySimpleGUI.Button('enviar')]
        ]
        janela = PySimpleGUI.Window('Dados').layout(layout)
        self.button, self.values = janela.Read()

    def iniciar (self):
        print (self.values)
        ip = self.values ['ip']
        os.system ('ping -n 4 {} '.format(ip))
        return ''


tela1 = Tela()
tela1.iniciar ()