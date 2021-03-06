VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Arquivo_AvisoDiferen�a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum enumTipoGeracaoAviso
    AvisoDiferenca
End Enum

'---------------------------------------------------------------------------------------------
'                                    Cria��o do Aviso de Diferen�a
'---------------------------------------------------------------------------------------------

Private Type tpDetalheArquivo
    Rotulo                      As String * 6       '"CHADIF"
    CGCEnderecamento            As String * 14      ' CGC de Endere�amento
    NumBordero                  As String * 18      ' N�mero do Border�
    Carteira                    As String * 2       ' Carteira
    DataOcorrencia              As String * 8       ' Data de Ocorrencia
    NumeroOcorrencia            As String * 9      ' Numero da Ocorrencia
    CodAgencia                  As String * 4       ' C�digo da Agencia Border�
    ContaCorrente               As String * 7       ' Conta Corrente Border�
    CodDevolucao                As String * 2       ' 01-Total  02-Parcial
    CodCompe                    As String * 3       ' C�digo da Compensa��o
    NumBco                      As String * 4       ' C�digo do Banco
    AgEmitente                  As String * 4       ' Agencia Emitente
    ContaEmitente               As String * 11      ' Conta corrente emitente
    NumChEmitente               As String * 10      ' N�mero do Cheque Emitente
    TipoCheque                  As String * 1       ' Tipo do Cheque
    TipoInscricao               As String * 2       ' Inscri��o
    InscEmitente                As String * 14      ' Inscri��o Emitente
    DataDeposito                As String * 8       ' Data Dep�sito
    ValorCheque                 As String * 13      ' Valor do Cheque
    CMC7                        As String * 34      ' CMC7 do Cheque
    MotivoDevolucao             As String * 2       ' Motivo devolu��o
    OcorrCHADIF                 As String * 1       ' Ocorrencia do CHADIF
    CrLf                        As String * 2       '
End Type


Private Type tpArquivoAviso
    Detalhe As tpDetalheArquivo
End Type

Private m_AvisoDiferen�a        As tpArquivoAviso
Private m_DataProcessamento     As Long
Private m_ProgressBar           As Object
Private m_DiretorioSaida        As String
Private m_ArquivoSaida          As String
Private m_Connection            As New ADODB.Connection

Public Property Let ArquivoSaida(ByVal pArquivoSaida As String)
    m_ArquivoSaida = pArquivoSaida
End Property

Public Property Let DataProcessamento(ByVal pDataProcessamento As Long)
    m_DataProcessamento = pDataProcessamento
End Property

Public Property Get DataProcessamento() As Long
    DataProcessamento = m_DataProcessamento
End Property

Public Property Get ArquivoSaida() As String
    ArquivoSaida = m_ArquivoSaida
End Property

Public Property Let DiretorioSaida(ByVal pDiretorioSaida As String)
    m_DiretorioSaida = pDiretorioSaida & IIf(Right(pDiretorioSaida, 1) <> "\", "\", "")
End Property

Public Property Get DiretorioSaida() As String
    DiretorioSaida = m_DiretorioSaida
End Property

Public Sub SetConnection(ByVal pConnection As ADODB.Connection)
    Set m_Connection = pConnection
End Sub

Public Sub SetProgressBar(ByRef pProgressBar As Object)
    Set m_ProgressBar = pProgressBar
End Sub


' Gera��o do Aviso de Diferen�a
Public Function Gera_AvisoDiferen�a() As Boolean

    Dim sstr                As String
    Dim rstDetAviso         As New ADODB.Recordset
    Dim rstMotivoAviso      As New ADODB.Recordset
    Dim Proc_Selecionar     As New Custodia.Selecionar
    Dim Proc_Atualizar      As New Custodia.Atualizar
    Dim strCMC7             As New CalculoCheque
    Dim ProgressBar         As New clsProgressBar
    Dim iFile               As Integer
    Dim lRetorno            As Long
    Dim bTransacaoAberta    As Boolean
    
    On Error GoTo Erro_Gerar
    
    bTransacaoAberta = False
    
    Gera_AvisoDiferen�a = False
    Screen.MousePointer = vbHourglass
    
    Set rstDetAviso = m_Connection.Execute(Proc_Selecionar.GetRecAvisoDiferenca(m_DataProcessamento, True))
    
    If rstDetAviso.RecordCount < 1 Then
        ' N�o existem registros para gerar o Arquivo de AD
        MsgBox "N�o Existe Aviso de Diferen�a para Transmiss�o", vbExclamation + vbOKOnly, App.Title
        Gera_AvisoDiferen�a = False
        Screen.MousePointer = vbDefault
        Exit Function
        
    End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                         CRIA��O E ABERTURA DO AVISO DE DIFEREN�A
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If FileExist(m_DiretorioSaida & m_ArquivoSaida) Then
       Kill m_DiretorioSaida & m_ArquivoSaida
    End If '

    iFile = FreeFile
    Open m_DiretorioSaida & m_ArquivoSaida For Binary As #iFile
    
    g_cMainConnection.BeginTrans
    bTransacaoAberta = True
    
    'Inicia progress bar
    ProgressBar.ValorMinimo = 0
    ProgressBar.ValorMaximo = rstDetAviso.RecordCount
    ProgressBar.DescricaoProcesso = "Gerando Aviso de Diferen�a para o CH ..."
    ProgressBar.InicializaProgressBar
    
    
    Do While Not rstDetAviso.EOF
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                               DETALHES DO ARQUIVO
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     
            
            'Atualiza Progress Bar
            ProgressBar.AtualValue = rstDetAviso.AbsolutePosition
            ProgressBar.AtualizaBarra
            
            
            Set rstMotivoAviso = m_Connection.Execute(Proc_Selecionar.GetCodMotivoAD(rstDetAviso!CodigoMotivo))
    
            
            With strCMC7
            
                .Banco = Format(rstDetAviso!BancoEmitente, String(3, "0"))
                .Agencia = Format(rstDetAviso!AgenciaEmitente, String(4, "0"))
                .Comp = Format(rstDetAviso!CodigoCompensacao, String(3, "0"))
                .NumeroCheque = Format(rstDetAviso!NrChequeEmitente, String(6, "0"))
                .Tipificacao = Format(rstDetAviso!TipoCheque, String(1, "0"))
                .Conta = Format(rstDetAviso!CcEmitente, String(10, "0"))
                .CMC7 = Space(1)
            
            End With
            
            strCMC7.Calcula
                        
            With m_AvisoDiferen�a.Detalhe
                .Rotulo = "CHADIF"
                .CGCEnderecamento = Format(g_Parametros.CNPJ_Terceira, String(14, "0"))
                .NumBordero = Format(rstDetAviso!Num_Bordero, String(Len(.NumBordero), "0"))
                .Carteira = Format(rstDetAviso!CodigoCarteira, String(Len(.Carteira), "0"))
                .DataOcorrencia = rstDetAviso!DataOcorrencia
                .NumeroOcorrencia = Format(rstDetAviso!CodigoOcorrencia, String(Len(.NumeroOcorrencia), "0"))
                .CodAgencia = Format(rstDetAviso!Agencia, String(Len(.CodAgencia), "0"))
                .ContaCorrente = Format(rstDetAviso!Conta, String(Len(.ContaCorrente), "0"))
                .CodDevolucao = Format(rstDetAviso!CodigoDevolucao, String(Len(.CodDevolucao), "0"))
                .CodCompe = Format(rstDetAviso!CodigoCompensacao, String(Len(.CodCompe), "0"))
                .NumBco = Format(rstDetAviso!BancoEmitente, String(Len(.NumBco), "0"))
                .AgEmitente = Format(rstDetAviso!AgenciaEmitente, String(4, "0"))
                .ContaEmitente = Format(rstDetAviso!CcEmitente, String(Len(.ContaEmitente), "0"))
                .NumChEmitente = Format(rstDetAviso!NrChequeEmitente, String(10, "0"))
                .TipoCheque = rstDetAviso!TipoCheque
                .TipoInscricao = Format(rstDetAviso!TipoInscricao, String(Len(.InscEmitente), "0"))
                .InscEmitente = Format(rstDetAviso!InscricaoEmitente, String(14, "0"))
                .DataDeposito = rstDetAviso!DataDeposito
                .ValorCheque = Format(rstDetAviso!Valor * 100, String(13, "0"))
                .CMC7 = Space(1) & strCMC7.Campo1 & Space(1) _
                & strCMC7.Campo2 & Space(1) & strCMC7.Campo3 & Space(1)
                .MotivoDevolucao = Format(rstMotivoAviso!CodigoMotivo, String(Len(.MotivoDevolucao), "0"))
                .OcorrCHADIF = 0
                .CrLf = vbCrLf
                
            End With
            
                If strCMC7.Campo1 = "" Then
                    m_AvisoDiferen�a.Detalhe.CMC7 = String(34, "0")
                End If
            
                Put #iFile, , m_AvisoDiferen�a.Detalhe
                                
                ' Gravar "T" Transmitido em Aviso de Diferen�a
                Call g_cMainConnection.Execute(Proc_Atualizar.AtualizaStatusAviso _
                (rstDetAviso!DataOcorrencia, rstDetAviso!CodigoOcorrencia, "T"), _
                lRetorno, _
                adCmdText)
            
                rstDetAviso.MoveNext
                
                If rstDetAviso.EOF Then Exit Do
              
            
        Loop

    'Fecha Arquivo txt
    Close
    'Fecha Transa��o
    If bTransacaoAberta Then g_cMainConnection.CommitTrans

    Gera_AvisoDiferen�a = True
    Screen.MousePointer = vbDefault
    
    Exit Function
    
Erro_Gerar:
    
    'Fecha arquivo txt
    Close
    'Cancela Transa��o
    If bTransacaoAberta Then g_cMainConnection.RollbackTrans
    
    Screen.MousePointer = vbDefault
    MsgBox Err.Description
    
    'Apaga arquivo gerado
    If FileExist(m_DiretorioSaida & m_ArquivoSaida) Then
        Kill m_DiretorioSaida & m_ArquivoSaida
    End If
    
End Function

