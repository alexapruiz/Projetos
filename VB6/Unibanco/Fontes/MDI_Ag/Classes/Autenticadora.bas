Attribute VB_Name = "Autenticadora"
Option Explicit

'interface de impressao autenticadora procomp
Declare Function WinIniciaPrtCx Lib "P32PRTCX.DLL" () As Integer
Declare Function WinStatusPrtCx Lib "P32PRTCX.DLL" (ByVal Buf_x As String) As Integer
Declare Function WinImprimePrtCx Lib "P32PRTCX.DLL" (ByVal Buf_A As Integer, ByVal buf_b As Integer, ByVal buf_c As String, ByVal buf_d As Integer) As Integer
Declare Function WinLineFeedPrtCx Lib "P32PRTCX.DLL" (ByVal n_linhas As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''
' Versão 3.3 (67)                        '
' Escolher o tipo de autenticadora usada '

Declare Function inicia_prn Lib "PERIWOSA.DLL" (ByVal impressora As String, ByVal autenticadora As String) As Long
Declare Function fim_prn Lib "PERIWOSA.DLL" () As Long
Declare Function imprime_prn Lib "PERIWOSA.DLL" (ByVal timeout As Long, ByVal size As Long, ByVal buffer As String) As Long
Declare Function autentica_prn Lib "PERIWOSA.DLL" (ByVal timeout As Long, ByVal size As Long, ByVal buffer As String) As Long
Declare Function status_impressora_prn Lib "PERIWOSA.DLL" (ByVal timeout As Long, device As Long, media As Long, paper As Long, toner As Long, retractbin As Long, retractcount As Long) As Long
Declare Function status_autenticadora_prn Lib "PERIWOSA.DLL" (ByVal timeout As Long, device As Long, media As Long, paper As Long, toner As Long, retractbin As Long, retractcount As Long) As Long
''''''''''''''''''''''''''''''''''''''''''

