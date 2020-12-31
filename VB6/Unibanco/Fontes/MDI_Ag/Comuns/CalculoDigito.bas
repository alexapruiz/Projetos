Attribute VB_Name = "CalculoDigito"
Option Explicit


Public Function Modulo11U(ByVal pviNumero As Long) As Byte
   '---------------------------------------------
   'Versão 3.3 - C/E (1)
   '--------- MODULO 11 (2 BASE 8) --------------
   ' Esta rotina serve para calcular:
   ' 1) número do ENVELOPE  - 3a. regra modulo11
   ' 2) 8+1
   '---------------------------------------------
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
   'número do envelope: (7+1)             0 0 9 9 9 9 9 9 9 9 - D
   '                                      x x x x x x x x x x
   'multiplica da direita para esquerda:  4 3 2 8 7 6 5 4 3 2
   '*************************************************************
   
   peso = 2    'começa multiplicar da direita para esquerda
   p = Len(nova_str)
   
   Do
      '*** Peso de 2 a 8 (multiplicação dos caracteres de 2 a 8) ***
      soma = soma + Mid(nova_str, p, 1) * peso
      p = p - 1            'ponteiro
      peso = peso + 1      'peso
      If (peso = 9) Then
         peso = 2
      End If
      If (p = 0) Then
         Exit Do
      End If
   Loop
   
   resto = soma Mod 11        'resto da divisão
   digito_11 = 11 - resto     'digito verificador
   
   Modulo11U = digito_11
   
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

Public Function Modulo10Arrecadacao(base_calculo As String, Tamanho As Integer) As Boolean
   
   '---------------------------------
   '--------- MODULO 10 -------------
   ' Esta rotina serve para calcular:
   ' 1) codigo de barras
   '---------------------------------
   
   Dim soma1 As Integer, digito1_10 As Integer, unico As Integer
   Dim troca As Integer, p As Integer, verif_carac As Integer
   Dim resto As Integer
   Dim digito_op As String
   Dim str_cb As String

   soma1 = 0
   digito1_10 = 0       'calculado pelo módulo 10
   
   troca = 0
   unico = 0
   
   digito_op = ""       'caracter digitado pelo operador (digito verificador)
   
   
   '*************************************************************
   'número codigo barras (43)             9 9 9 9 9 9 9 9 9 9
   '                                      x x x x x x x x x x
   'multiplica da direita para esquerda:  1 2 1 2 1 2 1 2 1 2
   '*************************************************************
   
   p = Tamanho        'tamanho da string
   
   Do
      '*** Base 2 (multiplicação dos caracteres por 1 e 2) ***
      If (troca = 0) Then
         unico = Mid(base_calculo, p, 1) * 2    'multiplica por 1
         troca = 1
      Else
         unico = Mid(base_calculo, p, 1) * 1    'multiplica por 2
         troca = 0
      End If

      '*** Verifica se um destes caracteres que foram multiplicados é >= 10.
      '*** Se for maior ou igual, soma o primeiro com o segundo caracter.
      '*** Ex.: 14 = (1 + 4),  18 = (1 + 8), etc...
      If (unico < 10) Then
         soma1 = soma1 + unico
      Else
         soma1 = soma1 + Mid(unico, 1, 1) + Mid(unico, 2, 1)
      End If
      
      '*** aponta para próximo caracter a ser multiplicado ***
      p = p - 1
      If (p = 1) Then      'o 1º número é o digito, então não poderá ser calculado
         Exit Do
      End If
   
   Loop
   
   resto = soma1 Mod 10       'resto da divisão
   digito1_10 = 10 - resto    'digito verificador
   
   '*** se o calculo for igual a 10 ou 11, muda-se para 0 ***
   If (digito1_10 = 11) Or (digito1_10 = 10) Then
      digito1_10 = 0
   End If

   digito_op = Mid$(base_calculo, 1, 1)
   
   If CStr(digito1_10) <> (digito_op) Then
      Modulo10Arrecadacao = False   'digito não confere
   Else
      Modulo10Arrecadacao = True    'digito confere
   End If
   
End Function

Public Function Modulo10(ByVal base_calculo As String, ByVal Tamanho As Integer) As Boolean
   
   '----------------------------------
   '--------- MODULO 10 --------------
   ' Esta rotina serve para calcular:
   ' 1) número da agencia + conta
   ' 2) 4, 6+1
   '---------------------------------
   
   Dim soma1 As Integer, digito1_10 As Integer, unico As Integer
   Dim troca As Integer
   Dim p As Integer
   Dim dec As Integer
   Dim verif_carac As Integer
   Dim digito_op As String

   soma1 = 0
   digito1_10 = 0       'calculado pelo módulo 10
   
   troca = 0
   unico = 0
   dec = 0

   digito_op = ""       'caracter digitado pelo operador (digito verificador)
   
   
   '*************************************************************
   'número da agencia+conta: (4+6+1)      9 9 9 9  9 9 9 9 9 9 - D
   '                                      x x x x  x x x x x x
   'multiplica da direita para esquerda:  1 2 1 2  1 2 1 2 1 2
   '*************************************************************
   
   p = Tamanho - 1      'string sem o digito
   
   Do
      '*** Base 2 (multiplicação dos caracteres por 1 e 2) ***
      If (troca = 0) Then
         unico = Mid(base_calculo, p, 1) * 2    'multiplica por 1
         troca = 1
      Else
         unico = Mid(base_calculo, p, 1) * 1    'multiplica por 2
         troca = 0
      End If

      '*** Verifica se um destes caracteres que foram multiplicados é >= 10.
      '*** Se for maior ou igual, soma o primeiro com o segundo caracter.
      '*** Ex.: 14 = (1 + 4),  18 = (1 + 8), etc...
      If (unico < 10) Then
         soma1 = soma1 + unico
      Else
         soma1 = soma1 + Mid(unico, 1, 1) + Mid(unico, 2, 1)
      End If
      
      '*** aponta para próximo caracter a ser multiplicado ***
      p = p - 1
      If (p = 0) Then
         Exit Do
      End If
   Loop
   
   '*** Calculo do número decimal para subtração do valor encontrado ***
   'deverá ser um decimal acima da soma encontrada
   If (soma1 > 9) Then
      If Mid(soma1, 2, 1) = 0 Then
         dec = soma1
      Else
         dec = Mid(soma1, 1, 1) + 1 & "0"
      End If
   Else
      If soma1 = 0 Then
         dec = 0
      Else
         dec = 10
      End If
   End If
   

   '*** Digito verificador calculado pelo módulo 10 ***
   digito1_10 = dec - soma1
   
   '*** Digito verificador digitado pelo operador ***
   digito_op = Mid$(base_calculo, Tamanho, 1)
   
   '*** Verifica se usuário digitou corretamente ***
   If Val(digito1_10) <> Val(digito_op) Then
      Modulo10 = False   'digito do campo 1 não confere
   Else
      Modulo10 = True   'digito do campo 1 confere
   End If
End Function
Public Function Modulo11INSS(ByVal pviNumero As Long) As Byte
   '---------------------------------------------
   '--------- MODULO 11 (2 BASE 9) --------------
   ' Esta rotina serve para calcular:
   ' 1) número do INSS
   ' 2) 8+1
   '---------------------------------------------
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
   p = Len(nova_str)
   
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
   digito_11 = 11 - resto     'digito verificador
   
   '*** se o calculo for igual a 10 ou 11, muda-se. ***
   If (digito_11 = 10) Or (digito_11 = 11) Then
      digito_11 = 0
   End If

   Modulo11INSS = digito_11

End Function

Public Function Modulo11Simplificado(ByVal pviNumero As Double) As Byte
   
    Dim resto As Integer
    Dim quociente As Double
    
    quociente = Fix(pviNumero / 11)
    resto = (pviNumero - (quociente * 11))

    If resto = 10 Then
        resto = 0
    End If
    
    Modulo11Simplificado = resto

End Function

Public Function Modulo11UBB(ByVal pviNumero As Double) As Byte
   '---------------------------------------------
   '--------- MODULO 11 (2 BASE 9) --------------
   ' Esta rotina serve para calcular:
   ' 1) número do ENVELOPE pessoa fisica - SOMENTE
   ' 2) 8+1
   '---------------------------------------------
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
   p = Len(nova_str)
   
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
   digito_11 = 11 - resto     'digito verificador
   
   '*** se o calculo for igual a 10 ou 11, muda-se. ***
   If (digito_11 = 11) Then
      digito_11 = 1
   End If
   
   If (digito_11 = 10) Then
      digito_11 = 0
   End If

   Modulo11UBB = digito_11
   
End Function

Public Function Modulo10CMC7(ByVal pvsCMC7 As String) As String
   Dim soma1 As Integer, digito1_10 As Integer
   Dim unico As Integer, troca As Integer
   Dim p As Integer
   Dim dec As Integer
   Dim verif_carac As Integer
   Dim soma2 As Integer, digito2_10 As Integer
   Dim soma3 As Integer, digito3_10 As Integer
   Dim digito_op As String
   Dim Campo1 As String, Campo2 As String, Campo3 As String
   Dim Ok1 As Byte
   Dim Ok2 As Byte
   Dim Ok3 As Byte

   soma1 = 0
   digito1_10 = 0       'calculado pelo módulo 10
   soma2 = 0
   digito2_10 = 0       'calculado pelo módulo 10
   soma3 = 0
   digito3_10 = 0       'calculado pelo módulo 10
   
   troca = 0
   unico = 0
   dec = 0

   digito_op = ""       'caracter digitado pelo operador (digito verificador)
   
   Campo1 = ""          'armazena parte 1 da banda magnética
   Campo2 = ""          'armazena parte 2 da banda magnética
   Campo3 = ""          'armazena parte 3 da banda magnética

   '*************************************************************************
   Campo1 = Mid$(pvsCMC7, 1, 7)
   Campo2 = Mid$(pvsCMC7, 9, 10)
   Campo3 = Mid$(pvsCMC7, 20, 10)

   
   '*************************************************************************
   
   '***************************
   '*** Cálculo do campo 01 ***
   '***************************
   p = 1             'início do campo 01
   Do
      '*** Verifica se o caracter é numérico ***
      verif_carac = Asc(Mid$(Campo1, p, 1))

      If (verif_carac < 48) Or (verif_carac > 57) Then
         Ok1 = 0
         'alteraçaõ versão 3.3
         Modulo10CMC7 = CStr(Ok1) & "00"
         Exit Function
      End If
      
      '*** Base 2 (multiplicação dos caracteres por 1 e 2) ***
      If (troca = 0) Then
         unico = Mid(Campo1, p, 1) * 2    'multiplica por 2
         troca = 1
      Else
         unico = Mid(Campo1, p, 1) * 1    'multiplica por 1
         troca = 0
      End If

      '*** Verifica se um destes caracteres que foram multiplicados é >= 10.
      '*** Se for maior ou igual, soma o primeiro com o segundo caracter.
      '*** Ex.: 14 = (1 + 4),  18 = (1 + 8), etc...
      If (unico < 10) Then
         soma1 = soma1 + unico
      Else
         soma1 = soma1 + Mid(unico, 1, 1) + Mid(unico, 2, 1)
      End If
      
      '*** aponta para próximo caracter a ser multiplicado ***
      p = p + 1
      If (p = 8) Then
         Exit Do
      End If
   Loop
   
   '*** Calculo do número decimal para subtração do valor encontrado ***
   If (soma1 > 9) Then
      If Mid(soma1, 2, 1) = 0 Then
         dec = soma1
      Else
         dec = Mid(soma1, 1, 1) + 1 & "0"
      End If
   Else
      If soma1 = 0 Then
         dec = 0
      Else
         dec = 10
      End If
   End If
   

   '*** Digito verificador calculado pelo módulo 10 ***
   digito1_10 = dec - soma1
   
   '*** Digito verificador digitado pelo operador ***
   digito_op = Mid(pvsCMC7, 19, 1)
   
   '*** Verifica se usuário digitou corretamente ***
   If CStr(digito1_10) <> (digito_op) Then
      Ok1 = 0      'digito do campo 1 não confere
   Else
      Ok1 = 1      'digito do campo 1 confere
   End If

   
   '***********************************************************************************
   
   '***************************
   '*** Cálculo do campo 02 ***
   '***************************
   p = 1          'início do campo 02
   troca = 0
   unico = 0
   dec = 0
   Do
      '*** Verifica se o caracter é numérico ***
      verif_carac = Asc(Mid$(Campo2, p, 1))

      If (verif_carac < 48) Or (verif_carac > 57) Then
         Ok2 = 0
         'alteraçaõ versão 3.3
         Modulo10CMC7 = CStr(Ok1) & CStr(Ok2) & "0"
         Exit Function
      End If
      
      '*** Base 2 (multiplicação dos caracteres por 1 e 2) ***
      If (troca = 0) Then
         unico = Mid(Campo2, p, 1) * 1    'multiplica por 1
         troca = 1
      Else
         unico = Mid(Campo2, p, 1) * 2    'multiplica por 2
         troca = 0
      End If

      '*** Verifica se um destes caracteres que foram multiplicados é >= 10.
      '*** Se for maior ou igual, soma o primeiro com o segundo caracter.
      '*** Ex.: 14 = (1 + 4),  18 = (1 + 8), etc...
      If (unico < 10) Then
         soma2 = soma2 + unico
      Else
         soma2 = soma2 + Mid(unico, 1, 1) + Mid(unico, 2, 1)
      End If
      
      '*** aponta para próximo caracter a ser multiplicado ***
      p = p + 1
      If (p = 11) Then
         Exit Do
      End If
   Loop
   
   '*** Calculo do número decimal para subtração do valor encontrado ***
   If (soma2 > 9) Then
      If Mid(soma2, 2, 1) = 0 Then
         dec = soma2
      Else
         dec = Mid(soma2, 1, 1) + 1 & "0"
      End If
   Else
      If soma2 = 0 Then
         dec = 0
      Else
         dec = 10
      End If
   End If

   '*** Digito verificador calculado pelo módulo 10 ***
   digito2_10 = dec - soma2
   
   '*** Digito verificador digitado pelo operador ***
   digito_op = Mid(pvsCMC7, 8, 1)
   
   '*** Verifica se usuário digitou corretamente ***
   If CStr(digito2_10) <> (digito_op) Then
      Ok2 = 0      'digito do campo 2 não confere
   Else
      Ok2 = 1      'digito do campo 2 confere
   End If
   
   
   '*******************************************************************************
   
   '***************************
   '*** Cálculo do campo 03 ***
   '***************************
   p = 1          'início do campo 03
   troca = 0
   unico = 0
   dec = 0
   Do
      '*** Verifica se o caracter é numérico ***
      verif_carac = Asc(Mid$(Campo3, p, 1))

      If (verif_carac < 48) Or (verif_carac > 57) Then
         Ok3 = 0
         'alteraçaõ versão 3.3
         Modulo10CMC7 = CStr(Ok1) & CStr(Ok2) & CStr(Ok3)
         Exit Function
      End If

      '*** Base 2 (multiplicação dos caracteres por 1 e 2) ***
      If (troca = 0) Then
         unico = Mid(Campo3, p, 1) * 1    'multiplica por 1
         troca = 1
      Else
         unico = Mid(Campo3, p, 1) * 2    'multiplica por 2
         troca = 0
      End If

      '*** Verifica se um destes caracteres que foram multiplicados é >= 10.
      '*** Se for maior ou igual, soma o primeiro com o segundo caracter.
      '*** Ex.: 14 = (1 + 4),  18 = (1 + 8), etc...
      If (unico < 10) Then
         soma3 = soma3 + unico
      Else
         soma3 = soma3 + Mid(unico, 1, 1) + Mid(unico, 2, 1)
      End If
      
      '*** aponta para próximo caracter a ser multiplicado ***
      p = p + 1
      If (p = 11) Then
         Exit Do
      End If
   Loop
   
   '*** Calculo do número decimal para subtração do valor encontrado ***
   If soma3 > 9 Then
      If Mid(soma3, 2, 1) = 0 Then
         dec = soma3
      Else
         dec = Mid(soma3, 1, 1) + 1 & "0"
      End If
   Else
      If soma3 = 0 Then
         dec = 0
      Else
         dec = 10
      End If
   End If
   
   '*** Digito verificador calculado pelo módulo 10 ***
   digito3_10 = dec - soma3
   
   '*** Digito verificador digitado pelo operador ***
   digito_op = Mid(pvsCMC7, 30, 1)
   
   '*** Verifica se usuário digitou corretamente ***
   If CStr(digito3_10) <> (digito_op) Then
      Ok3 = 0      'digito do campo 3 não confere
   Else
      Ok3 = 1      'digito do campo 3 confere
   End If
   
   Modulo10CMC7 = CStr(Ok1) & CStr(Ok2) & CStr(Ok3)

End Function

Public Function DV10(ByVal str As String) As String
    Dim ContDv, soma, Dig, i As Integer
    
    ContDv = 2
    soma = 0
    For i = Len(str) To 1 Step -1
        Dig = Val(Mid(str, i, 1)) * ContDv
        soma = soma + (Dig Mod 10) + (Int(Dig / 10))
        ContDv = 3 - ContDv
    Next
    DV10 = CStr((10 - soma Mod 10) Mod 10)
End Function

Public Function DV11(ByVal str As String) As String
    Dim ContDv, soma, Dig, i As Integer
    
    ContDv = 2
    soma = 0
    For i = Len(str) To 1 Step -1
        Dig = Val(Mid(str, i, 1)) * ContDv
        soma = soma + Dig
        ContDv = ContDv + 1
        If ContDv = 10 Then
            ContDv = 2
        End If
    Next
    DV11 = CStr(((11 - soma Mod 11) Mod 11) Mod 10)
End Function

