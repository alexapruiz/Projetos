VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Serpente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private sPosicao As Integer
Private sDirecao As Integer
Public Property Get Posicao() As Variant

    Posicao = sPosicao
End Property
Public Property Let Posicao(ByVal vNewValue As Variant)

    sPosicao = vNewValue
End Property
Public Sub Mover(ByVal Posicao As Integer, ByRef ObjSerpente As SSPanel)

    'Calculando o LEFT
    Select Case (Posicao Mod 10)
    Case 0
        ObjSerpente.Left = 3600
    Case 1
        ObjSerpente.Left = 0
    Case 2
        ObjSerpente.Left = 400
    Case 3
        ObjSerpente.Left = 800
    Case 4
        ObjSerpente.Left = 1200
    Case 5
        ObjSerpente.Left = 1600
    Case 6
        ObjSerpente.Left = 2000
    Case 7
        ObjSerpente.Left = 2400
    Case 8
        ObjSerpente.Left = 2800
    Case 9
        ObjSerpente.Left = 3200
    End Select

    Select Case Posicao
    Case 1 To 10
        ObjSerpente.Top = 0
    Case 11 To 20
        ObjSerpente.Top = 400
    Case 21 To 30
        ObjSerpente.Top = 800
    Case 31 To 40
        ObjSerpente.Top = 1200
    Case 41 To 50
        ObjSerpente.Top = 1600
    Case 51 To 60
        ObjSerpente.Top = 2000
    Case 61 To 70
        ObjSerpente.Top = 2400
    Case 71 To 80
        ObjSerpente.Top = 2800
    Case 81 To 90
        ObjSerpente.Top = 3200
    Case 91 To 100
        ObjSerpente.Top = 3600
    End Select

    Me.Posicao = Posicao
End Sub
Public Property Get DIRECAO() As Variant

    DIRECAO = sDirecao
End Property
Public Property Let DIRECAO(ByVal vNewValue As Variant)

    sDirecao = vNewValue
End Property
