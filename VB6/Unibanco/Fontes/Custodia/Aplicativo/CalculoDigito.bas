Attribute VB_Name = "CalculoDigito"
Option Explicit
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
Public Function Modulo11Simplificado(ByVal pviNumero As Double) As Boolean
   
    Dim resto As Integer
    Dim quociente As Double
    Dim Numero As Double
    Dim digito As Byte
    
    Numero = Left(CStr(pviNumero), Len(CStr(pviNumero)) - 1)
    digito = Right(CStr(pviNumero), 1)
    
    quociente = Fix(Numero / 11)
    resto = (Numero - (quociente * 11))

    If resto = 10 Then
        resto = 0
    End If
    
    If resto = digito Then
        Modulo11Simplificado = True
    End If

End Function
Public Function RetornaDigitoModulo11Simplificado(ByVal pviNumero As Double) As Byte
   
    Dim resto As Integer
    Dim quociente As Double
        
    quociente = Fix(pviNumero / 11)
    resto = (pviNumero - (quociente * 11))

    If resto = 10 Then
        resto = 0
    End If
    
    RetornaDigitoModulo11Simplificado = resto

End Function


