VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Elevador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private sAndar As Integer
Public Property Get Andar() As Variant

    Andar = sAndar
End Property
Public Property Let Andar(ByVal vNewValue As Variant)

    sAndar = vNewValue
End Property
Public Sub Subir(ByRef AndarAnterior As PictureBox, ByRef ProximoAndar As PictureBox)

    'Move o elevador 1 andar para cima
    Me.Andar = Me.Andar + 1

    'Pinta o andar atual
    AndarAnterior.BackColor = 12632256
    ProximoAndar.BackColor = 65535
End Sub
Public Function Descer(ByRef AndarAnterior As PictureBox, ByRef ProximoAndar As PictureBox)

    'Move o elevador 1 andar para baixo
    Me.Andar = Me.Andar - 1

    'Pinta o andar atual
    AndarAnterior.BackColor = 12632256
    ProximoAndar.BackColor = 65535
End Function
