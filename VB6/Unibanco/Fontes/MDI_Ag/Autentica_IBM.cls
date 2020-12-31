VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Autentica_IBM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function Imprimir(ByVal buffer As String, ByVal autenticar As Boolean) As Long
    If autenticar Then
        Imprimir = autentica_prn(1, Len(buffer), buffer)
    Else
        Imprimir = imprime_prn(1, Len(buffer), buffer)
        buffer = Chr(13) & Chr(10)
        Imprimir = imprime_prn(1, Len(buffer), buffer)
    End If
End Function

Public Function Status(ByRef Buff_st As String) As Integer
    Buff_st = Chr(0) & Chr(2) & Chr(0)
    Status = status_impressora_prn(0, 0, 0, 0, 0, 0, 0)
End Function

Public Function Finaliza() As Integer
    Finaliza = fim_prn()
End Function

Public Function Inicia() As Integer
    Inicia = inicia_prn("IMPRESSORA", "AUTENTICADORA")
End Function
