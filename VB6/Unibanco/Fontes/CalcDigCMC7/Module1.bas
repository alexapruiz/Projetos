Attribute VB_Name = "CalcDig"
Option Explicit

Declare Function SetWindowPos Lib "User32" (ByVal h&, ByVal hb&, ByVal x&, ByVal y&, ByVal cx&, ByVal cy&, ByVal f&) As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1

Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Sub SelecionarTexto(ByVal pObjeto As Object)


    On Error Resume Next
    
    pObjeto.SelStart = 0
    pObjeto.SelLength = Len(pObjeto)
'    pObjeto.SetFocus
'
'    If Err <> 0 Then Err = 0
    

End Sub

Public Function SetTopWindow(hWnd As Long, bState As Boolean) As Boolean
  If bState = True Then 'Put the window on top
    SetTopWindow = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
  ElseIf bState = False Then ' Turn off the TopMost flag
    SetTopWindow = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
  Else
    Debug.Print "bState Unknown."
    SetTopWindow = False
  End If
End Function
Public Function Modulo11(ByVal pviNumero As Double) As Boolean
   '--------------------------------------------------
   '--------- MODULO 11 (2 BASE 9) -------------------
   ' Esta rotina serve para calcular linha 1 do cmc7 -
   '--------------------------------------------------
   Dim soma As Integer
   Dim resto As Integer
   Dim digito_11 As Byte
   Dim p As Integer
   Dim peso As Integer
   Dim nova_str As String
   
   nova_str = CStr(pviNumero)
   soma = 0
   resto = 0
   digito_11 = 0        'calculado pelo módulo 11

   '*************************************************************
   'número do envelope: (8+1)             0 0 9 9 9 9 9 9 9 9 - D
   '                                      x x x x x x x x x x
   'multiplica da direita para esquerda:  3 2 9 8 7 6 5 4 3 2
   '*************************************************************
   
   peso = 2    'começa multiplicar da direita para esquerda
   p = Len(nova_str) - 1
   
   Do
      '*** Peso de 2 a 9 (multiplicação dos caracteres de 2 a 9) ***
      soma = soma + Mid(nova_str, p, 1) * peso
      p = p - 1            'ponteiro
      peso = peso + 1      'peso
      If (peso = 10) Then
         peso = 2
      End If
      If (p = 0) Then
         Exit Do
      End If
   Loop
   
   resto = soma Mod 11        'resto da divisão
   If resto = 0 Or resto = 1 Then
        digito_11 = 0
   Else
        digito_11 = 11 - resto     'digito verificador
   End If
   
   '*** se o calculo for igual a 10 ou 11, muda-se. ***
   If (digito_11 = 11) Then
      digito_11 = 1
   End If
   
   If (digito_11 = 10) Then
      digito_11 = 0
   End If
    
   If Right(nova_str, 1) = digito_11 Then
        Modulo11 = True
   Else
        Modulo11 = False
   End If
   
End Function

