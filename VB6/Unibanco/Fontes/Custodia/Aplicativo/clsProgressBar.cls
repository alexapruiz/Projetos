VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Dim lngValorMaximo        As Long
    Dim lngValorMinimo        As Long
    Dim lngAtualValue         As Long
    Dim strDescricaoProcesso  As String
    
Private Sub Class_Initialize()
    
    lngValorMinimo = 0
    lngValorMaximo = 0
    lngAtualValue = 0
    strDescricaoProcesso = ""
    
End Sub

Public Sub InicializaProgressBar()
     
     'Inicializa Progress Bar ( Form Principal )
     Principal.ProgressBar1.Min = 0 'lngValorMinimo
     Principal.ProgressBar1.Max = lngValorMaximo
     Principal.ProgressBar1.Visible = True
     Principal.StatusBarPrincipal.Panels(StatusBar.Col_ContadorProgressBar).Text = ""
     Principal.StatusBarPrincipal.Panels(StatusBar.Col_Descrição).Text = strDescricaoProcesso

End Sub
Public Property Let ValorMaximo(lValorMaximo As Long)

     lngValorMaximo = lValorMaximo
     
End Property

Public Property Let ValorMinimo(lValorMinimo As Long)

     lngValorMinimo = lValorMinimo
     lngAtualValue = lValorMinimo
     
End Property

Public Property Let DescricaoProcesso(sDescricaoProcesso As String)

     strDescricaoProcesso = sDescricaoProcesso
     
End Property
Public Property Let AtualValue(lAtualValue As Long)

     lngAtualValue = lAtualValue
     
End Property
Public Sub AtualizaBarra()

     'Atualiza Progress Bar
     Principal.ProgressBar1.Value = lngAtualValue
     Principal.StatusBarPrincipal.Panels(StatusBar.Col_ContadorProgressBar).Text = lngAtualValue & "/" & lngValorMaximo

End Sub

Private Sub Class_Terminate()
    
     'Encerra progress bar
     Principal.StatusBarPrincipal.Panels(StatusBar.Col_Descrição).Text = ""
     Principal.ProgressBar1.Visible = False
     Principal.StatusBarPrincipal.Panels(StatusBar.Col_ContadorProgressBar).Text = ""

End Sub
