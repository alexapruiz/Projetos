VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CalculoCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''
'Exemplo de chamada'
''''''''''''''''''''

'    Dim x As New clsCheque
'
'    x.Comp = "006"
'    x.Banco = "415"
'    x.Agencia = "0588"
'    x.Conta = "0588003806"
'    x.NumeroCheque = "986557"
'    x.Tipificacao = "5"
'
'    Call x.Calcula
'
'    MsgBox x.CMC7


Private m_Campo1        As String
Private m_Campo2        As String
Private m_Campo3        As String
Private m_Linha1        As String
Private m_Comp          As String
Private m_Banco         As String
Private m_Agencia       As String
Private m_Conta         As String
Private m_NumeroCheque  As String
Private m_C1            As String
Private m_C2            As String
Private m_C3            As String
Private m_Tipificacao   As String
Private m_CMC7          As String
Private m_DV1           As String
Private m_DV2           As String
Private m_DV3           As String
Private m_Status        As String
Private m_IdCheque      As Long
Private m_Valor         As Double
''''''''''''''''''''''''''''''''''
'Define os campos atraves do CMC7'
''''''''''''''''''''''''''''''''''
Private Sub DefineCampos()

    If Trim(m_CMC7) = "" Then Exit Sub

    m_Comp = Mid(m_CMC7, 9, 3)
    m_Banco = Left(m_CMC7, 3)
    m_Agencia = Mid(m_CMC7, 4, 4)
    m_Conta = Mid(m_CMC7, 20, 10)
    m_NumeroCheque = Mid(m_CMC7, 12, 6)
    
'    m_C1 = Mid(m_CMC7, 19, 1)
'    m_C2 = Mid(m_CMC7, 8, 1)
'    m_C3 = Right(m_CMC7, 1)
    m_DV1 = Mid(m_CMC7, 19, 1)
    m_DV2 = Mid(m_CMC7, 8, 1)
    m_DV3 = Mid(m_CMC7, 30, 1)
    
    m_Tipificacao = Mid(m_CMC7, 18, 1)
    

End Sub

Public Property Get DV1() As String
    DV1 = m_DV1
End Property

Public Property Get DV2() As String
    DV2 = m_DV2
End Property


Public Property Get DV3() As String
    DV3 = m_DV3
End Property



Public Property Let IdCheque(ByVal pIdCheque As Long)
    m_IdCheque = pIdCheque
End Property

Public Property Get IdCheque() As Long
    IdCheque = m_IdCheque
End Property

Public Property Let Status(ByVal pStatus As String)
    m_Status = pStatus
End Property

Public Property Get Status() As String
    Status = m_Status
End Property

Public Property Let Tipificacao(ByVal pTipificacao As String)
    m_Tipificacao = pTipificacao
End Property
Public Property Get Tipificacao() As String
    Tipificacao = m_Tipificacao
End Property
Public Property Let C3(ByVal pC3 As String)
    m_C3 = pC3
End Property
Public Property Get C3() As String
    C3 = m_C3
End Property
Public Property Let C2(ByVal pC2 As String)
    m_C2 = pC2
End Property
Public Property Get C2() As String
    C2 = m_C2
End Property
Public Property Let C1(ByVal pC1 As String)
    m_C1 = pC1
End Property
Public Property Get C1() As String
    C1 = m_C1
End Property
Public Property Let NumeroCheque(ByVal pNumeroCheque As String)
    m_NumeroCheque = pNumeroCheque
End Property
Public Property Get NumeroCheque() As String
    NumeroCheque = m_NumeroCheque
End Property
Public Property Let Conta(ByVal pConta As String)
    m_Conta = pConta
End Property
Public Property Get Conta() As String
    Conta = m_Conta
End Property
Public Property Let Agencia(ByVal pAgencia As String)
    m_Agencia = pAgencia
End Property
Public Property Get Agencia() As String
    Agencia = m_Agencia
End Property
Public Property Let Banco(ByVal pBanco As String)
    m_Banco = pBanco
End Property
Public Property Get Banco() As String
    Banco = m_Banco
End Property

Public Function Calcula() As Boolean
    Dim iTipo           As Integer
    Dim sAgencia        As String
    Dim sConta          As String
    Dim sBanco          As String
    Dim sComp           As String
    Dim sCampo1         As String
    Dim sCampo2         As String
    Dim sCampo3         As String
    Dim sNumeroCheque   As String
    Dim sDV1            As String
    Dim sDV2            As String
    Dim sDV3            As String
    Dim sLinha1         As String
    
    Calcula = False

    DefineCampos


'    ''''''''''''''''''''
'    'Consiste os campos'
'    ''''''''''''''''''''
'    If Trim(m_Comp) = "" Or _
'       Trim(m_Banco) = "" Or _
'       Trim(m_Agencia) = "" Or _
'       Trim(m_Conta) = "" Or _
'       Trim(m_NumeroCheque) = "" Or _
'       Trim(m_Tipificacao) = "" Then
'
'        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        'Define os campos Comp, Banco, Agencia, Conta e Numero do Cheque'
'        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        DefineCampos
'
'        Exit Function
'    End If
    
    iTipo = Val(m_Tipificacao)
    ''''''''''''''''''''''
    'Tipificação inválida'
    ''''''''''''''''''''''
    If iTipo <> 5 And _
       iTipo <> 6 And _
       iTipo <> 8 And _
       iTipo <> 9 Then
        Exit Function
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Se foram passados os campos C1,C2 e C3, o cálculo é diferente'
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If m_C1 <> "" And m_C2 <> "" And m_C3 <> "" Then
        '''''''''''''''''''''''''''''''''
        'Validar Primeiro Lote da Linha1'
        '''''''''''''''''''''''''''''''''
        If Not Modulo11(m_Comp & m_Banco & m_Agencia & m_C1) Then
            Exit Function
        End If
        ''''''''''''''''''''''''''''''''
        'Validar Segundo Lote da Linha1'
        ''''''''''''''''''''''''''''''''
        If Not Modulo11(m_Conta & m_C2) Then
            Exit Function
        End If
        '''''''''''''''''''''''''''''''''
        'Validar Terceiro Lote da Linha1'
        '''''''''''''''''''''''''''''''''
        If Not Modulo11(m_NumeroCheque & m_C3) Then
            Exit Function
        End If
        
        m_CMC7 = RetornaCMC7(m_Comp, _
                             m_Banco, _
                             m_Agencia, _
                             m_Conta, _
                             m_NumeroCheque, _
                             m_Tipificacao)
        Calcula = True
    Else
        '''''''''''''''''''''''''''''''''''''''''
        'Não foram passados os campos C1,C2 e C3'
        '''''''''''''''''''''''''''''''''''''''''
        
        m_CMC7 = RetornaCMC7(m_Comp, _
                             m_Banco, _
                             m_Agencia, _
                             m_Conta, _
                             m_NumeroCheque, _
                             m_Tipificacao)
        Calcula = True
    End If
    
    
    
    
    
    
'    Exit Function
'
'
'
'
'
'
'
'
'
'    If m_Campo1 = "" Or m_Campo2 = "" Or m_Campo3 = "" Then
'        ''''''''''''''''''''''''
'        'Verifica a Tipificacao'
'        ''''''''''''''''''''''''
'        iTipo = Right(m_Linha1, 1)
'        If iTipo <> 5 And _
'           iTipo <> 6 And _
'           iTipo <> 8 And _
'           iTipo <> 9 Then
'
'            ''''''''''''''''''''''
'            'Tipificação inválida'
'            ''''''''''''''''''''''
'            CMC7 = ""
'            Exit Property
'        End If
'
'        sLinha1 = Left(m_Linha1, Len(m_Linha1) - 1)
'
'        '''''''''''''''''''''
'        'Pega DVs da Linha 1'
'        '''''''''''''''''''''
'        sDV1 = Mid(sLinha1, 11, 1)
'        sDV2 = Mid(sLinha1, 19, 1)
'        'sDV3 = Mid(m_Linha1, 26, 1)
'        '''''''''''''''''''''''''''''
'        'Pega Compensacao da linha 1'
'        '''''''''''''''''''''''''''''
'        sComp = Left(sLinha1, 3)
'        sLinha1 = Mid(sLinha1, 4)
'        '''''''''''''''''''''''
'        'Pega Banco da Linha 1'
'        '''''''''''''''''''''''
'        sBanco = Left(sLinha1, 3)
'        sLinha1 = Mid(sLinha1, 4)
'
'        '''''''''''''''''''''''''
'        'Pega Agencia da Linha 1'
'        '''''''''''''''''''''''''
'        sAgencia = Left(sLinha1, 4)
'        sLinha1 = Mid(sLinha1, 5)
'        '''''''''''''''''''''''
'        'Pega numero do Cheque'
'        '''''''''''''''''''''''
'        sNumeroCheque = Right(sLinha1, 6)
'        sLinha1 = Mid(sLinha1, 1, Len(sLinha1) - Len(sNumeroCheque))
'        ''''''''''''''''''''''''''''''''''
'        'pega Conta Corrente da linha 1  '
'        ''''''''''''''''''''''''''''''''''
'        'sConta = Mid(m_Linha1, 12, 7)
'        sConta = sLinha1
'        '''''''''''''''''''''''''''''''''
'        'Validar Primeiro Lote da Linha1'
'        '''''''''''''''''''''''''''''''''
''        If Not Modulo11(sComp & sBanco & sAgencia & sDV1) Then
''            Exit Property
''        End If
'
'        ''''''''''''''''''''''''''''''''
'        'Validar Segundo Lote da Linha1'
'        ''''''''''''''''''''''''''''''''
''        If Not Modulo11(sConta & sDV2) Then
''            Exit Property
''        End If
'
'        '''''''''''''''''''''''''''''''''
'        'Validar Terceiro Lote da Linha1'
'        '''''''''''''''''''''''''''''''''
''        If Not Modulo11(sNumeroCheque & sDV3) Then
''            Exit Property
''        End If
'
'
'        CMC7 = RetornaCMC7(sComp, sBanco, sAgencia, sConta, sNumeroCheque, CStr(iTipo))
'    Else
'        CMC7 = m_Campo1 & m_Campo2 & m_Campo3
'    End If
'



End Function

Public Property Let Campo1(ByVal pCampo1 As String)
    m_Campo1 = pCampo1
End Property

Public Property Get Campo1() As String

    Campo1 = m_Campo1

End Property

Public Property Let Campo2(ByVal pCampo2 As String)
    m_Campo2 = pCampo2
End Property

Public Property Get Campo2() As String
    Campo2 = m_Campo2
End Property


Public Property Let Campo3(ByVal pCampo3 As String)
    m_Campo3 = pCampo3
End Property

Public Property Get Campo3() As String
    Campo3 = m_Campo3
End Property

Public Property Get CMC7() As String
    CMC7 = m_CMC7
End Property

Public Property Let Comp(ByVal pComp As String)
    m_Comp = pComp
End Property

Public Property Get Comp() As String
    Comp = m_Comp
End Property

Public Property Get Linha1() As String
    Linha1 = m_Linha1
End Property

Private Function RetornaCMC7(ByVal pComp As String, _
                             ByVal pBanco As String, _
                             ByVal pAgencia As String, _
                             ByVal pContaCorrente As String, _
                             ByVal pNumeroCheque As String, _
                             ByVal pTipificacao As String) As String

  On Error GoTo ERRO_CALCULARDVLINHA1

  Dim soma1 As Integer, digito1 As Integer
  Dim soma2 As Integer, digito2 As Integer
  Dim soma3 As Integer, digito3 As Integer
  Dim resto As Integer
  Dim p As Integer
  Dim Campo1, Campo2, Campo3 As String
  Dim unico As Integer, troca As Integer
  Dim dec As Integer

  troca = 0
  unico = 0
  soma1 = 0
  digito1 = 0       'calculado
  soma2 = 0
  digito2 = 0       'calculado
  soma3 = 0
  digito3 = 0       'calculado
  resto = 0

  Campo1 = pBanco & pAgencia
  Campo2 = pComp & pNumeroCheque & pTipificacao
  Campo3 = pContaCorrente

  '***************************
  '*** CÁLCULO DO CAMPO 01 ***
  '***************************
  p = Len(Campo1)         'fim do campo 01
  Do
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
      p = p - 1
      If (p = 0) Then
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
  digito1 = dec - soma1
  troca = 0
 
  '***************************
  '*** CÁLCULO DO CAMPO 02 ***
  '***************************
  p = Len(Campo2)         'fim do campo 02
  Do
      '*** Base 2 (multiplicação dos caracteres por 1 e 2) ***
      If (troca = 0) Then
         unico = Mid(Campo2, p, 1) * 2    'multiplica por 2
         troca = 1
      Else
         unico = Mid(Campo2, p, 1) * 1    'multiplica por 1
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
      p = p - 1
      If (p = 0) Then
          Exit Do
      End If
  Loop
  
  '*** Calculo do DV2 do CMC7 ***
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
  digito2 = dec - soma2
  
  
  '***************************
  '*** CÁLCULO DO CAMPO 03 ***
  '***************************
  p = Len(Campo3)         'fim do campo 03
  Do
      '*** Base 2 (multiplicação dos caracteres por 1 e 2) ***
      If (troca = 0) Then
          unico = Mid(Campo3, p, 1) * 2    'multiplica por 2
          troca = 1
      Else
          unico = Mid(Campo3, p, 1) * 1    'multiplica por 1
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
      p = p - 1
      If (p = 0) Then
          Exit Do
      End If
  Loop
  
  '*** Calculo do DV3 do CMC7 ***
  '*** Calculo do número decimal para subtração do valor encontrado ***
  If (soma3 > 9) Then
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
  digito3 = dec - soma3
  
  m_Campo1 = pBanco & pAgencia & digito2
  m_Campo2 = pComp & pNumeroCheque & pTipificacao
  m_Campo3 = digito1 & Format(pContaCorrente, "0000000000") & digito3
  
  RetornaCMC7 = m_Campo1 & m_Campo2 & m_Campo3

                

  Exit Function

ERRO_CALCULARDVLINHA1:
  Screen.MousePointer = vbDefault
End Function


'
'
Public Property Let CMC7(ByVal pCMC7 As String)

'    m_Campo1 = Mid(pCMC7, 1, 8)
'    m_Campo2 = Mid(pCMC7, 9, 10)
'    m_Campo3 = Mid(pCMC7, 19, 12)
'
    ''''''''''''''''''''''''''''''''''''''
    'Define variavel membro para trabalho'
    ''''''''''''''''''''''''''''''''''''''
    m_CMC7 = pCMC7

End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Esta propriedade ja define no momento de sua atribuição as outras'
'propriedades como Compensacao,Banco,Agencia,NumeroCheque e Conta.'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let Linha1(ByVal pLinha1 As String)

    Dim sLinha1     As String
    
    MsgBox "Ainda não tem uma definição."
    
    Exit Property

    m_Linha1 = pLinha1
    
    sLinha1 = m_Linha1
    
    '''''''''''''''''''''''''''''
    'Pega Compensacao da linha 1'
    '''''''''''''''''''''''''''''
    m_Comp = Left(sLinha1, 3)
    sLinha1 = Mid(sLinha1, 4)
    '''''''''''''''''''''''
    'Pega Banco da Linha 1'
    '''''''''''''''''''''''
    m_Banco = Left(sLinha1, 3)
    sLinha1 = Mid(sLinha1, 4)
    '''''''''''''''''''''''''
    'Pega Agencia da Linha 1'
    '''''''''''''''''''''''''
    m_Agencia = Left(sLinha1, 4)
    sLinha1 = Mid(sLinha1, 5)
    '''''''''''''''''''''''
    'Pega numero do Cheque'
    '''''''''''''''''''''''
    m_NumeroCheque = Right(sLinha1, 6)
    sLinha1 = Mid(sLinha1, 1, Len(sLinha1) - Len(m_NumeroCheque))
    ''''''''''''''''''''''''''''''''
    'pega Conta Corrente da linha 1'
    'QUE É O QUE SOBROU            '
    ''''''''''''''''''''''''''''''''
    m_Conta = sLinha1

End Property




Public Property Let Valor(ByVal pValor As Double)
    m_Valor = pValor
End Property

Public Property Get Valor() As Double
    Valor = m_Valor
End Property


