VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SomatoriaDatas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'''''''''''''''''''''''''''''''''
'Exemplo de chamada desta classe'
'''''''''''''''''''''''''''''''''
'Dim clSD            As New SomatoriaDatas
'
'clSD.SetConnection g_cMainConnection
'clSD.DataProcessamento = m_DataProcessamento
'clSD.IdBordero = m_IdBordero
'
'clSD.Calcula

''''''''''''''''''''''''''''''''
'Definição do objeto de Conexão'
''''''''''''''''''''''''''''''''
Private m_Connection            As New ADODB.Connection

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Definição das variaveis privadas de propriedades da classe'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private m_SomatoriaDatas        As String
Private m_SomatoriaQuantidades  As String
Private m_SomatoriaValores      As String
Private m_SomatoriaControle     As String

'''''''''''''''''''''''''''''''''''''''''
'Receber DataProcessamento em mm/dd/yyyy'
'''''''''''''''''''''''''''''''''''''''''
Private m_IdBordero             As Long
Private m_DataProcessamento     As Long
Public Sub Calcula()

    Dim Proc_Selecionar         As New Custodia.Selecionar
    Dim rst                     As New ADODB.Recordset
    Dim lDataDeposito           As Long
    Dim lQuantidade             As Integer
    Dim lValor                  As Double

    '''''''''''''''''''''''''''''''''''''''''''
    'Seleciona as Datas de Deposito do Bordero'
    '''''''''''''''''''''''''''''''''''''''''''
    
    Set rst = m_Connection.Execute(Proc_Selecionar.GetDataDeposito( _
                                   m_DataProcessamento, _
                                   m_IdBordero))


'    Set rst = m_Connection.Execute(Proc_Selecionar.GetSomatoriaChequesBordero( _
'                                   m_DataProcessamento, _
'                                   m_IdBordero))

    Do While Not rst.EOF()
        lDataDeposito = lDataDeposito + Format(Format(rst!DataDeposito, "0000/00/00"), "ddmmyy")
        lQuantidade = lQuantidade + rst!QuantidadeCheques
        lValor = lValor + rst!ValorDeposito
        
        rst.MoveNext
    Loop
    
    '''''''''''''''''''''''''''''''''
    'Atribui valores às propriedades'
    '''''''''''''''''''''''''''''''''
    m_SomatoriaDatas = lDataDeposito
    m_SomatoriaQuantidades = lQuantidade
    m_SomatoriaValores = RetiraPonto(Format(lValor, MASK_VALOR))
    '''''''''''''''''''''''''''''''''''''''''
    'Atribui valores à Somatoria de Controle'
    '''''''''''''''''''''''''''''''''''''''''
    m_SomatoriaControle = CDbl(m_SomatoriaDatas) + _
                          CDbl(m_SomatoriaQuantidades) + _
                          CDbl(m_SomatoriaValores)
    
    rst.Close

End Sub

Public Property Let DataProcessamento(ByVal pDataProcessamento As Long)
    m_DataProcessamento = pDataProcessamento
End Property

Public Property Get DataProcessamento() As Long
    DataProcessamento = m_DataProcessamento
End Property

Public Property Let IdBordero(ByVal pIdBordero As Long)
    m_IdBordero = pIdBordero
End Property

Public Property Get IdBordero() As Long
    IdBordero = m_IdBordero
End Property

Public Sub SetConnection(ByVal pConnection As ADODB.Connection)

    Set m_Connection = pConnection

End Sub


Public Property Get SomatoriaControle() As String
    SomatoriaControle = m_SomatoriaControle
End Property

Public Property Get SomatoriaDatas() As String
    SomatoriaDatas = m_SomatoriaDatas
End Property


Public Property Get SomatoriaQuantidades() As String
    SomatoriaQuantidades = m_SomatoriaQuantidades
End Property


Public Property Get SomatoriaValores() As String
    SomatoriaValores = m_SomatoriaValores
End Property


