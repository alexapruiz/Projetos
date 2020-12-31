VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataDeposito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"DataDeposito"
Option Explicit

'''''''''''''''''''''''''''''''''''''
'Receber datas no formato mm/dd/yyyy'
'''''''''''''''''''''''''''''''''''''
Private m_DataProcessamento     As String
Private m_DataDeposito          As String
Private m_Quantidade            As Integer
Private m_ValorDeposito         As String
Private m_IdBordero             As Long
Private m_DataDivergente        As Boolean
Private m_SomatoriaQtdCheques   As Integer
Private m_SomatoriaVlrCheques   As String
Private m_ValorAD               As Double
'''''''''''''''''''''''''''''''
'Define variaveis de diferença'
'''''''''''''''''''''''''''''''
Private m_DiferencaQuantidade   As Integer
Private m_DiferencaValor        As String


Public Cheque                   As New clsCheque

Public Function Calcula(ByVal pConnection As ADODB.Connection) As Boolean

    Dim Proc_Selecionar     As New Custodia.Selecionar
    Dim rst                 As New ADODB.Recordset
    Dim rstCalculo          As New ADODB.Recordset


    Calcula = False
    ''''''''''''''''''''''''''''''''''''''''''''''''
    'Seleciona valores referente à Data de Deposito'
    ''''''''''''''''''''''''''''''''''''''''''''''''
    Set rst = pConnection.Execute(Proc_Selecionar.GetDataDepositoBordero( _
                                  m_DataProcessamento, _
                                  m_IdBordero, _
                                  m_DataDeposito))

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Seleciona os valores dos cheques referente à esta Data de Deposito'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set rstCalculo = pConnection.Execute(Proc_Selecionar.GetSomatoriaChequesBordero( _
                                         m_DataProcessamento, _
                                         m_IdBordero, _
                                         m_DataDeposito))

    If (Not rst.EOF()) And (Not rstCalculo.EOF()) Then
    
        m_SomatoriaVlrCheques = Format(rstCalculo!TotCheques, MASK_VALOR)
        m_SomatoriaQtdCheques = rstCalculo!QtdeCheques
    
        ''''''''''''''''''''
        'Os campos estão ok'
        ''''''''''''''''''''
        If rst!QuantidadeCheques = rstCalculo!QtdeCheques And _
           rst!ValorDeposito = rstCalculo!TotCheques Then

            '''''''''''''''''''''''''''''''''''''''''
            'Define as diferenças à serem retornadas'
            '''''''''''''''''''''''''''''''''''''''''
            m_DiferencaValor = 0
            m_DiferencaQuantidade = 0

            m_DataDivergente = False
            Calcula = True
        '''''''''''''''''''''''''''''''
        'Somente a Quantidade não bate'
        '''''''''''''''''''''''''''''''
        ElseIf rst!QuantidadeCheques <> rstCalculo!QtdeCheques And _
               rst!ValorDeposito = rstCalculo!TotCheques Then
            
            '''''''''''''''''''''''''''''''''''''''''
            'Define as diferenças à serem retornadas'
            '''''''''''''''''''''''''''''''''''''''''
            m_DiferencaValor = 0
            'm_DiferencaQuantidade = Abs(rst!QuantidadeCheques - rstCalculo!QtdeCheques)
            m_DiferencaQuantidade = rst!QuantidadeCheques - rstCalculo!QtdeCheques
            
            m_DataDivergente = True
            Calcula = False
        ''''''''''''''''''''''''''
        'Somente o Valor não bate'
        ''''''''''''''''''''''''''
        ElseIf rst!QuantidadeCheques = rstCalculo!QtdeCheques And _
               rst!ValorDeposito <> rstCalculo!TotCheques Then
            
            '''''''''''''''''''''''''''''''''''''''''
            'Define as diferenças à serem retornadas'
            '''''''''''''''''''''''''''''''''''''''''
            'm_DiferencaValor = Format(Abs(rst!ValorDeposito - rstCalculo!TotCheques), MASK_VALOR)
            m_DiferencaValor = Format((rst!ValorDeposito - rstCalculo!TotCheques), MASK_VALOR)
            m_DiferencaQuantidade = 0
            
            m_DataDivergente = True
            Calcula = False
        ''''''''''''''''''''''''''''''
        'Valor e Quantidade não batem'
        ''''''''''''''''''''''''''''''
        ElseIf rst!QuantidadeCheques <> rstCalculo!QtdeCheques And _
               rst!ValorDeposito <> rstCalculo!TotCheques Then
            
            '''''''''''''''''''''''''''''''''''''''''
            'Define as diferenças à serem retornadas'
            '''''''''''''''''''''''''''''''''''''''''
            'm_DiferencaValor = Format(Abs(rst!ValorDeposito - rstCalculo!TotCheques), MASK_VALOR)
            m_DiferencaValor = Format((rst!ValorDeposito - rstCalculo!TotCheques), MASK_VALOR)
            'm_DiferencaQuantidade = Abs(rst!QuantidadeCheques - rstCalculo!QtdeCheques)
            m_DiferencaQuantidade = (rst!QuantidadeCheques - rstCalculo!QtdeCheques)
            
            m_DataDivergente = True
            Calcula = False

        End If
    ElseIf Not rst.EOF() And rstCalculo.EOF() Then
    
        m_DiferencaValor = rst!ValorDeposito
        m_DiferencaQuantidade = rst!QuantidadeCheques
        
        m_DataDivergente = True
        Calcula = False
    Else
        ''''''''''''''''''''''''''''''''
        'Não é obrigatorio ter o cheque'
        ''''''''''''''''''''''''''''''''
        Calcula = True
    End If

    rst.Close
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''
'Retorna TRUE se sucesso, caso contrario FALSE'
'''''''''''''''''''''''''''''''''''''''''''''''
Public Function CalculaChequesAD(ByVal pConnection As ADODB.Connection) As Boolean

    Dim Proc_Selecionar         As New Custodia.Selecionar
    Dim rst                     As New ADODB.Recordset
    Dim dValorAD                As Double
    
    On Error GoTo Erro_CalculaChequesAD
    
    CalculaChequesAD = False

    '''''''''''''''''''''''''''''''''''''''''''''
    'Seleciona os cheques desta Data de Deposito'
    '''''''''''''''''''''''''''''''''''''''''''''
    Set rst = pConnection.Execute(Proc_Selecionar.GetCheques(m_DataProcessamento, _
                                                             m_IdBordero, _
                                                             CLng(m_DataDeposito)))
    
    ''''''''''''''''''''''''''
    'Preenche o objeto cheque'
    ''''''''''''''''''''''''''
    Do While Not rst.EOF()
        Me.Cheque.Add rst!CMC7
        Me.Cheque.Item(rst.AbsolutePosition).Valor = rst!Valor
        
        rst.MoveNext
    Loop

    dValorAD = m_ValorAD

    '''''''''''''''''''''''''''
    'Começa de baixo para cima'
    '''''''''''''''''''''''''''
    rst.MoveLast
    Do While Not rst.BOF()
    
    
        If Me.Cheque.Item(rst.AbsolutePosition).Valor <= dValorAD Then
            '''''''''''''''''''''''''''''''''''''''''''''
            'Flag para dizer que é de Aviso de Diferenca'
            '''''''''''''''''''''''''''''''''''''''''''''
            Me.Cheque.Item(rst.AbsolutePosition).Status = "AD"
            
            dValorAD = dValorAD - Me.Cheque.Item(rst.AbsolutePosition).Valor
            
        End If
        rst.MovePrevious
    Loop
    
    rst.Close
    
    CalculaChequesAD = True
    
    Exit Function
    
Erro_CalculaChequesAD:

    CalculaChequesAD = False


End Function


'''''''''''''''''''''''''''''''
'Retorna no formato mm/dd/yyyy'
'''''''''''''''''''''''''''''''
Public Property Get DataDeposito() As String
    DataDeposito = m_DataDeposito
End Property

Public Property Let DataDeposito(ByVal pDataDeposito As String)
    m_DataDeposito = pDataDeposito
End Property

Public Property Let DataDivergente(ByVal pDivergencia As Boolean)
    m_DataDivergente = pDivergencia
End Property

Public Property Get DataDivergente() As Boolean
    DataDivergente = m_DataDivergente
End Property

Public Property Get DataProcessamento() As String
    DataProcessamento = m_DataProcessamento
End Property

Public Property Let DataProcessamento(ByVal pDataProcessamento As String)
    m_DataProcessamento = pDataProcessamento
End Property

Public Property Get DiferencaQuantidade() As Integer
    DiferencaQuantidade = m_DiferencaQuantidade
End Property


Public Property Get DiferencaValor() As String
    DiferencaValor = m_DiferencaValor
End Property

Public Property Let IdBordero(ByVal pIdBordero As Long)
    m_IdBordero = pIdBordero
End Property

Public Property Get IdBordero() As Long
    IdBordero = m_IdBordero
End Property

Public Property Get Quantidade() As Integer
    Quantidade = m_Quantidade
End Property


Public Property Let Quantidade(ByVal pQuantidade As Integer)
    m_Quantidade = pQuantidade
End Property



Public Property Get SomatoriaQuantidadeCheques() As Integer
    SomatoriaQuantidadeCheques = m_SomatoriaQtdCheques
End Property

Public Property Get SomatoriaValoresCheques() As String
    SomatoriaValoresCheques = m_SomatoriaVlrCheques
End Property


Public Property Let ValorAD(ByVal pValorAD As Double)
    m_ValorAD = pValorAD
End Property

Public Property Get ValorAD() As Double
    ValorAD = m_ValorAD
End Property

Public Property Get ValorDeposito() As String
    ValorDeposito = m_ValorDeposito
End Property

Public Property Let ValorDeposito(ByVal pValor As String)
    m_ValorDeposito = pValor
End Property


