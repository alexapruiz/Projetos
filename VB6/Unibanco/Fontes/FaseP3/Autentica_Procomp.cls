VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Autentica_Procomp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Imprimir(ByVal buffer As String, ByVal autenticar As Boolean) As Long
    Dim strOut As String
    strOut = "  " & buffer & Chr(10)
    buffer = Chr(15) & buffer & Chr(10)
    
    If autenticar Then
        Imprimir = WinImprimePrtCx(1, 1, buffer, Len(buffer))
    Else
        Imprimir = WinImprimePrtCx(0, 0, strOut, Len(strOut))
    End If
End Function

Public Function Status(ByRef Buff_st As String) As Integer
    Status = WinStatusPrtCx(Buff_st)
End Function

Public Function Finaliza() As Integer
    Finaliza = 1
End Function

Public Function Inicia() As Integer
    Inicia = WinIniciaPrtCx()
End Function
