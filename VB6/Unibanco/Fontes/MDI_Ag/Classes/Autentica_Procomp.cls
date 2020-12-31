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
Attribute VB_Exposed = True
Option Explicit

'interface de impressao autenticadora procomp
Private Declare Function WinIniciaPrtCx Lib "P32PRTCX.DLL" () As Integer
Private Declare Function WinStatusPrtCx Lib "P32PRTCX.DLL" (ByVal Buf_x As String) As Integer
Private Declare Function WinImprimePrtCx Lib "P32PRTCX.DLL" (ByVal Buf_A As Integer, ByVal buf_b As Integer, ByVal buf_c As String, ByVal buf_d As Integer) As Integer
Private Declare Function WinLineFeedPrtCx Lib "P32PRTCX.DLL" (ByVal n_linhas As Integer) As Integer

Public Function Imprimir(ByVal buffer As String, ByVal autenticar As Boolean) As Long
    Dim strOut As String
    strOut = "  " & buffer
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
