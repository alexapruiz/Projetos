VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Autenticadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''''''''''''''''''''''''''''''''''''''''
' Versão 3.3 (67)                        '
' Escolher o tipo de autenticadora usada '

Private Declare Function inicia_prn Lib "PERIWOSA.DLL" (ByVal impressora As String, ByVal Autenticadora As String) As Long
Private Declare Function fim_prn Lib "PERIWOSA.DLL" () As Long
Private Declare Function imprime_prn Lib "PERIWOSA.DLL" (ByVal timeout As Long, ByVal size As Long, ByVal buffer As String) As Long
Private Declare Function autentica_prn Lib "PERIWOSA.DLL" (ByVal timeout As Long, ByVal size As Long, ByVal buffer As String) As Long
Private Declare Function status_impressora_prn Lib "PERIWOSA.DLL" (ByVal timeout As Long, device As Long, media As Long, paper As Long, toner As Long, retractbin As Long, retractcount As Long) As Long
Private Declare Function status_autenticadora_prn Lib "PERIWOSA.DLL" (ByVal timeout As Long, device As Long, media As Long, paper As Long, toner As Long, retractbin As Long, retractcount As Long) As Long

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
