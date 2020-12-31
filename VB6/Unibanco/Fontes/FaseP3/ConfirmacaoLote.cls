VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfirmacaoLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim m_IdLote                        As Long
Dim m_Err                           As Long
Dim m_ArquivoOrigem                 As String
Dim m_ArquivoCapaCriacao            As String
Dim m_ArquivoDoctoCriacao           As String
Dim m_DiretorioCriacao              As String
Dim m_DiretorioOrigem               As String
Dim m_DiretorioDestino              As String
Dim m_DiretorioBULK                 As String
Dim m_FormatoData                   As String
Dim m_DataProcessamento             As String
Dim m_MsgError                      As String
Dim m_Login                         As String
Dim m_TotalCapas                    As Integer
Dim m_TotalDoctos                   As Integer
Dim m_ExcluirArquivoCriacao         As Boolean
Dim m_Connection                    As New RDO.rdoConnection

Private Const DELIMITADOR = ";"

'''''''''''''''''''''''''''''''''''''''''
'Precisa ser identico à tabela Documento'
'''''''''''''''''''''''''''''''''''''''''
'Sempre o campo + 1 para "|" (Delimitador)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ja existe um TpCapa e TpDocumento, coloquei cls para diferenciar'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Type clsTpCapa
    DataProcessamento               As String * 9
    IdCapa                          As String * 2
    IdLote                          As String * 10
    IdEnv_Mal                       As String * 2
    Capa                            As String * 15
    Num_Malote                      As String * 2
    AgOrig                          As String * 5
    Status                          As String * 2
    DataCriacao                     As String * 20
    Alcada                          As String * 1
    PendenciaValor                  As String * 1
    Supervisor                      As String * 1
    VinculoManual                   As String * 1
    IgnorarProva0                   As String * 1
    Ocorrencia                      As String * 1
    Conta                           As String * 1
    Dinheiro                        As String * 1
    Diferenca                       As String * 1
    Duplicidade                     As String * 2
    HoraAtual                       As String * 1
    RecepcionadoIK                  As String * 1
    CrLf                            As String * 2
End Type

Private Type clsTpDocumento
    DataProcessamento               As String * 9
    IdDocto                         As String * 2
    IdCapa                          As String * 6
    OrdemCaptura                    As String * 4
    TipoDocto                       As String * 2
    Ocorrencia                      As String * 1
    CodMotivo                       As String * 2
    Leitura                         As String * 49
    Frente                          As String * 20
    Verso                           As String * 20
    Status                          As String * 2
    Alcada                          As String * 2
    Autenticado                     As String * 2
    Cortado                         As String * 1
    OcorrenciaOK                    As String * 1
    AutenticacaoDigital             As String * 1
    Ordem                           As String * 2
    NSU                             As String * 1
    Terminal                        As String * 1
    Vinculo                         As String * 2
    AgenciaVinculo                  As String * 1
    ContaVinculo                    As String * 1
    CMC7Associado                   As String * 1
    Duplicidade                     As String * 2
    CodCenape                       As String * 2
    CodBarComplem                   As String * 2
    Valor                           As String * 15
    RetornoTransacao                As String * 1
    CrLf                            As String * 2
End Type

Private Type clsTpLog
    DataProcessamento               As String * 9
    IdCapa                          As String * 6
    IdDocto                         As String * 6
    Data                            As String * 20
    Login                           As String * 11
    Acao                            As String * 4
    CrLf                            As String * 2
End Type
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Recebe o nome do arquivo de capas que será gerado no cliente e servidor'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let ArquivoCapaCriacao(ByVal pArquivoCapaCriacao As String)
    m_ArquivoCapaCriacao = pArquivoCapaCriacao
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Retorna o nome do arquivo de capas que será gerado no cliente e servidor'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get ArquivoCapaCriacao() As String
    ArquivoCapaCriacao = m_ArquivoCapaCriacao
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Retorna o nome do arquivo de documentos que foi gerado no cliente e servidor'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let ArquivoDoctoCriacao(ByVal pArquivoDoctoCriacao As String)
    m_ArquivoDoctoCriacao = pArquivoDoctoCriacao
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Retorna o nome do arquivo de documentos que será gerado no cliente e servidor'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get ArquivoDoctoCriacao() As String
    ArquivoDoctoCriacao = m_ArquivoDoctoCriacao
End Property

''''''''''''''''''''''''''''''''''''''''''
'Recebe o nome do arquivo gerado pela DLL'
''''''''''''''''''''''''''''''''''''''''''
Public Property Let ArquivoOrigem(ByVal pArquivoOrigem As String)
    m_ArquivoOrigem = pArquivoOrigem
End Property

Public Property Let DataProcessamento(ByVal pDataProcessamento As String)
    m_DataProcessamento = pDataProcessamento
End Property

Public Property Get DataProcessamento() As String
    DataProcessamento = m_DataProcessamento
End Property

Public Property Let DiretorioBULK(ByVal pDiretorioBULK As String)
    m_DiretorioBULK = pDiretorioBULK
End Property

Public Property Get DiretorioBULK() As String
    DiretorioBULK = m_DiretorioBULK
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Recebe o diretorio de criacao onde serão gerados os arquivos'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let DiretorioCriacao(ByVal pDiretorioCriacao As String)

    m_DiretorioCriacao = IIf(Right(pDiretorioCriacao, 1) = "\", pDiretorioCriacao, pDiretorioCriacao & "\")
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Retorna o diretorio de criacao dos arquivos de importacao'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get DiretorioCriacao() As String
    DiretorioCriacao = m_DiretorioCriacao
End Property



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Recebe o endereço no servidor onde será gravado o arquivo de importação'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let DiretorioDestino(ByVal pDiretorioDestino As String)

    m_DiretorioDestino = IIf(Right(pDiretorioDestino, 1) = "\", pDiretorioDestino, pDiretorioDestino & "\")
    
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Retorna o nome do servidor onde fora atribuido anteriormente'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get DiretorioDestino() As String
    DiretorioDestino = m_DiretorioDestino
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Retorna o diretorio de origem onde foi gerado o arquivo pela DLL'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let DiretorioOrigem(ByVal pDiretorioOrigem As String)

    m_DiretorioOrigem = IIf(Right(pDiretorioOrigem, 1) = "\", pDiretorioOrigem, pDiretorioOrigem & "\")
    
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Diretorio onde se encontra o arquivo gerado pela DLL'
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get DiretorioOrigem() As String
    DiretorioOrigem = m_DiretorioOrigem
End Property

Public Property Get Err() As Long
    Err = m_Err
End Property

Public Property Get Error() As String
    Error = m_MsgError
End Property


Public Property Let ExcluirArquivoCriacao(ByVal pExcluir As Boolean)
    m_ExcluirArquivoCriacao = pExcluir
End Property

Public Property Get ExcluirArquivoCriacao() As Boolean
    ExcluirArquivoCriacao = m_ExcluirArquivoCriacao
End Property

Public Property Let FormatoData(ByVal pFormatoData As String)
    m_FormatoData = pFormatoData
End Property

Public Property Get FormatoData() As String
    FormatoData = m_FormatoData
End Property

Public Property Let IdLote(ByVal pIdLote As Long)
    m_IdLote = pIdLote
End Property

Public Property Get IdLote() As Long
    IdLote = m_IdLote
End Property

Public Property Let Login(ByVal pLogin As String)
    m_Login = pLogin
End Property

Public Property Get Login() As String
    Login = m_Login
End Property

Public Function Processar() As Boolean

    If Geral.VIPSDLL = eDllNovaUBB Then
        Processar = ProcessarLayoutNovo
    Else
        Processar = ProcessarLayoutAntigo
    End If

End Function

Public Sub SetConnection(ByVal pConnection As RDO.rdoConnection)

    Set m_Connection = pConnection
    
End Sub

Private Function TrataLeitura(ByVal Leitura As String) As String
    Dim bInvalido As Boolean
    Dim Count As Integer
    Dim Result As String
    Dim Char As String * 1
    
    Result = ""
    Leitura = Trim(Leitura)
    bInvalido = False
    For Count = 1 To Len(Leitura)
        Char = Mid(Leitura, Count, 1)
        If Char <> "<" And Char <> ">" And Char <> ":" And Char <> ";" Then
            If (Not bInvalido) And (Not IsNumeric(Char)) Then
                bInvalido = True
            End If
            If bInvalido Then
                Result = Result & "0"
            Else
                Result = Result & Char
            End If
        End If
    Next
    If Val(Result) > 0 Then
        TrataLeitura = Result
    Else
        TrataLeitura = ""
    End If
End Function

Private Function FormataString(ByVal pOque As Variant, _
                               ByVal pCompletarCom As Variant, _
                               ByVal pFieldLen As Integer, _
                               ByVal pAEsquerda As Boolean) As Variant
    
    
    If pFieldLen <= 0 Then FormataString = pOque: Exit Function
    If pCompletarCom = "" Then FormataString = pOque: Exit Function
    If pFieldLen < Len(pOque) Then FormataString = pOque: Exit Function

    If pAEsquerda Then
        FormataString = Right(String(pFieldLen - Len(pOque), pCompletarCom) & pOque, pFieldLen)
    Else
        FormataString = Left(pOque & String(pFieldLen - Len(pOque), pCompletarCom), pFieldLen)
    End If
    
End Function



Private Function VerificaSeCapa(ByVal Leitura As String, ByRef IdEnv_Mal As String) As Boolean
    Dim Campo1 As String
    Dim Campo2 As String
    Dim Campo3 As String
    Dim Valor As String
    
    VerificaSeCapa = False
    IdEnv_Mal = "E"
    Leitura = Trim(Leitura)
    
    If Leitura = "" Then
        Exit Function
    End If
    If Len(Leitura) = 8 Then ' Envelope
        If IsNumeric(Leitura) Then
            If Right(Leitura, 1) = Modulo11UBB(Val(Left(Leitura, Len(Leitura) - 1))) Or _
               Right(Leitura, 1) = Modulo11Simplificado(Val(Left(Leitura, Len(Leitura) - 1))) Or _
               Right(Leitura, 1) = Modulo11U(Val(Left(Leitura, Len(Leitura) - 1))) Then
                VerificaSeCapa = True
                IdEnv_Mal = "E"
            End If
        End If
    Else ' Malote
        If Len(Leitura) >= 30 Then
            If TratarCamposCMC7(Leitura, Campo1, Campo2, Campo3, Valor) Then
                'If Mid(Campo3, 1, 4) = "0600" And Mid(Campo1, 1, 3) = "409" And Mid(Campo2, 10, 1) = "4" Then
                If Mid(Campo2, 1, 3) = "600" And Mid(Campo1, 1, 3) = "409" And Mid(Campo2, 10, 1) = "4" Then
                    VerificaSeCapa = True
                    IdEnv_Mal = "M"
                End If
            End If
        End If
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''
'Retorna o nome do arquivo gerado pela DLL'
'''''''''''''''''''''''''''''''''''''''''''
Public Property Get ArquivoOrigem() As String
    ArquivoOrigem = m_ArquivoOrigem
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Propriedade Read Only que retorna o numero de capas do arquivo txt'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get TotalCapas() As Integer
    TotalCapas = m_TotalCapas
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Propriedade Read Only que retorna o numero de documentos do arquivo txt'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get TotalDocumentos() As Integer
    TotalDocumentos = m_TotalDoctos
End Property

Private Function ProcessarLayoutAntigo() As Boolean

    Dim vCapa()             As tpRetornoVips
    Dim vwCapa()            As clsTpCapa        'Tipo para gravar no arquivo de BULK INSERT
    Dim vwDocto()           As clsTpDocumento   'Tipo para gravar no arquivo de BULK INSERT
    Dim rstIdCapas          As RDO.rdoResultset
    Dim rstIdDoctos         As RDO.rdoResultset
    Dim vIndiceDocto()      As String
    Dim iRFile              As Integer
    Dim iWFile              As Integer
    Dim iContaLinhas_RF     As Integer          'Contador de linhas do arquivo de leitura
    Dim iCountVWCapa        As Integer
    Dim iCountVWDocto       As Integer
    Dim iOrdemCaptura       As Integer
    Dim sCampo1             As String
    Dim sCampo2             As String
    Dim sCampo3             As String
    Dim svalor              As String
    Dim sIdEnv_Mal          As String
    Dim IdCapa              As Long
    Dim sTipoDocto          As String
    Dim iIndexDocto         As Integer
    Dim bVirtual            As Boolean
    Dim i                   As Integer
    Dim j                   As Integer
    Dim sStr                As String
    Dim sStr2               As String
    Dim sAtualizaStatusCapa As String
    Dim sCapa               As String
    Dim sCapa2              As String
    Dim ExisteDuplicidadeCaptura As Boolean
    
    On Error GoTo Erro_ProcessarLayoutAntigo
    
    ProcessarLayoutAntigo = False
    
    Screen.MousePointer = vbHourglass
    
    iRFile = FreeFile
    IdCapa = 0
    iIndexDocto = 0
    sCapa = ""
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Abre o arquivo de dados gerado pela DLL para a leitura dos documentos'
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Open m_DiretorioOrigem & m_ArquivoOrigem For Binary As #iRFile
    
    Do While Not EOF(iRFile)
        sTipoDocto = 0
        
        '''''''''''''''''''''''''''''''
        'Redimensiona o vetor de capas'
        '''''''''''''''''''''''''''''''
        ReDim Preserve vCapa(iContaLinhas_RF) As tpRetornoVips

        Get #iRFile, , vCapa(iContaLinhas_RF)
        
        If vCapa(iContaLinhas_RF).Tipo <> "A" And vCapa(iContaLinhas_RF).Tipo <> "B" Then Exit Do
        
        '''''''''''''''''''''''''''''''''''''''''''
        'Tratamento da linha do arquivo de leitura'
        '''''''''''''''''''''''''''''''''''''''''''
        vCapa(iContaLinhas_RF).Leitura = TrataLeitura(vCapa(iContaLinhas_RF).Leitura)
        
        If VerificaSeCapa(vCapa(iContaLinhas_RF).Leitura, sIdEnv_Mal) Then
            If sIdEnv_Mal = "M" Then
                TratarCamposCMC7 vCapa(iContaLinhas_RF).Leitura, sCampo1, sCampo2, sCampo3, svalor
            End If
            sTipoDocto = 1 ' Capa
            bVirtual = False
            ' gravar capa

            ReDim Preserve vwCapa(iCountVWCapa) As clsTpCapa

            With vwCapa(iCountVWCapa)
                '.PosCapa = "   "
                .DataProcessamento = m_DataProcessamento & DELIMITADOR
                .IdEnv_Mal = sIdEnv_Mal & DELIMITADOR
                .Capa = FormataString(IIf(sIdEnv_Mal = "E", Val(Trim(vCapa(iContaLinhas_RF).Leitura)), "0" & Mid(sCampo2, 1, 9) & Mid(sCampo1, 4, 4)), " ", Len(.Capa) - 1, False) & DELIMITADOR
                .IdCapa = "0" & DELIMITADOR
                .IdLote = FormataString(m_IdLote, "0", Len(.IdLote) - 1, True) & DELIMITADOR
                .IdEnv_Mal = sIdEnv_Mal & DELIMITADOR
                .Num_Malote = "0" & DELIMITADOR
                .AgOrig = FormataString("0", "0", Len(.AgOrig) - 1, True) & DELIMITADOR
                .Status = "U" & DELIMITADOR
                .DataCriacao = Format(Now(), m_FormatoData) & DELIMITADOR '"dd-mm-yyyy hh:mm:00") & DELIMITADOR
                
                .Alcada = DELIMITADOR
                .PendenciaValor = DELIMITADOR
                .Supervisor = DELIMITADOR
                .VinculoManual = DELIMITADOR
                .IgnorarProva0 = DELIMITADOR
                .Ocorrencia = DELIMITADOR
                .Conta = DELIMITADOR
                .Dinheiro = DELIMITADOR
                .Diferenca = DELIMITADOR
                .Duplicidade = "0" & DELIMITADOR
                .HoraAtual = DELIMITADOR
                .RecepcionadoIK = "N" 'não vai delimitador
                .CrLf = vbCrLf
            End With
            sCapa = vwCapa(iCountVWCapa).Capa
            m_TotalCapas = m_TotalCapas + 1
            iOrdemCaptura = 1
            iCountVWCapa = iCountVWCapa + 1
            
        ElseIf sCapa = "" Then 'so pode entrar aqui quando primeira linha
        
            sTipoDocto = 1 ' Capa
            bVirtual = True
            sCapa = "9"

            ReDim Preserve vwCapa(iCountVWCapa) As clsTpCapa
            
            With vwCapa(iCountVWCapa)
                .DataProcessamento = m_DataProcessamento & DELIMITADOR
                .IdEnv_Mal = sIdEnv_Mal & DELIMITADOR
                .Capa = FormataString("9", "0", Len(.Capa) - 1, True) & DELIMITADOR
                .IdCapa = "0" & DELIMITADOR
                .IdLote = FormataString(m_IdLote, "0", Len(.IdLote) - 1, True) & DELIMITADOR
                .IdEnv_Mal = "E" & DELIMITADOR
                .Num_Malote = "0" & DELIMITADOR
                .AgOrig = FormataString("0", "0", Len(.AgOrig) - 1, True) & DELIMITADOR
                .Status = "U" & DELIMITADOR
                .DataCriacao = Format(Now(), m_FormatoData) & DELIMITADOR '"dd-mm-yyyy hh:mm:00") & DELIMITADOR
                .Alcada = DELIMITADOR
                .PendenciaValor = DELIMITADOR
                .Supervisor = DELIMITADOR
                .VinculoManual = DELIMITADOR
                .IgnorarProva0 = DELIMITADOR
                .Ocorrencia = DELIMITADOR
                .Conta = DELIMITADOR
                .Dinheiro = DELIMITADOR
                .Diferenca = DELIMITADOR
                .Duplicidade = "0" & DELIMITADOR
                .HoraAtual = DELIMITADOR
                .RecepcionadoIK = "N" 'não vai delimitador
                .CrLf = vbCrLf
            End With

            iCountVWCapa = iCountVWCapa + 1
            m_TotalCapas = m_TotalCapas + 1
            iOrdemCaptura = 1
            
        End If
        ' gravar documento
        svalor = "000"
        
        ReDim Preserve vIndiceDocto(iIndexDocto) As String
        ReDim Preserve vwDocto(iCountVWDocto) As clsTpDocumento

        
        vIndiceDocto(iIndexDocto) = vwCapa(iCountVWCapa - 1).Capa
        iIndexDocto = iIndexDocto + 1
        
        vwDocto(iCountVWDocto).DataProcessamento = m_DataProcessamento & DELIMITADOR
        vwDocto(iCountVWDocto).TipoDocto = sTipoDocto & DELIMITADOR
        
        
        If sIdEnv_Mal = "M" And sTipoDocto = 1 Then ' Capa de Malote
            If Not bVirtual Then
                vwDocto(iCountVWDocto).Leitura = FormataString("0" & Mid(sCampo2, 1, 9) & Mid(sCampo1, 4, 4), " ", Len(vwDocto(iCountVWDocto).Leitura) - 1, False) & DELIMITADOR
            Else
                vwDocto(iCountVWDocto).Leitura = FormataString("9", " ", Len(vwDocto(iCountVWDocto).Leitura) - 1, False) & DELIMITADOR 'Capa Virtual
            End If
        
        ElseIf sIdEnv_Mal = "E" And sTipoDocto = 1 Then ' Capa de Envelope
            If Not bVirtual Then
                vwDocto(iCountVWDocto).Leitura = FormataString(Trim(vCapa(iContaLinhas_RF).Leitura), " ", Len(vwDocto(iCountVWDocto).Leitura) - 1, False) & DELIMITADOR
            Else
                vwDocto(iCountVWDocto).Leitura = FormataString("9", " ", Len(vwDocto(iCountVWDocto).Leitura) - 1, False) & DELIMITADOR 'Capa Virtual
            End If
        
        ElseIf vCapa(iContaLinhas_RF).Tipo = "A" Then ' Docto com CMC7
            Valor = ""
            TratarCamposCMC7 vCapa(iContaLinhas_RF).Leitura, sCampo1, sCampo2, sCampo3, svalor
            vwDocto(iCountVWDocto).Leitura = FormataString(sCampo1 & sCampo2 & sCampo3, " ", Len(vwDocto(iCountVWDocto).Leitura) - 1, False) & DELIMITADOR
        
        ElseIf vCapa(iContaLinhas_RF).Tipo = "B" Then ' Docto com Cod Barras
            vwDocto(iCountVWDocto).Leitura = FormataString(RPad(Trim(vCapa(iContaLinhas_RF).Leitura), 44), " ", Len(vwDocto(iCountVWDocto).Leitura) - 1, False) & DELIMITADOR
        Else
            vwDocto(iCountVWDocto).Leitura = FormataString(Trim(vCapa(iContaLinhas_RF).Leitura), " ", Len(vwDocto(iCountVWDocto).Leitura) - 1, False) & DELIMITADOR
        End If

        ''''''''''''''''''''''''''''''''''''
        'Campos em branco porem necessarios'
        ''''''''''''''''''''''''''''''''''''
        With vwDocto(iCountVWDocto)
            .IdDocto = "0" & DELIMITADOR
            .Ocorrencia = DELIMITADOR
            .CodMotivo = "0" & DELIMITADOR
            .Status = "0" & DELIMITADOR
            .Alcada = "N" & DELIMITADOR
            .Autenticado = "0" & DELIMITADOR
            .Cortado = DELIMITADOR
            .OcorrenciaOK = DELIMITADOR
            .AutenticacaoDigital = DELIMITADOR
            .Ordem = "0" & DELIMITADOR
            .NSU = DELIMITADOR
            .Terminal = DELIMITADOR
            .Vinculo = "0" & DELIMITADOR
            .AgenciaVinculo = DELIMITADOR
            .ContaVinculo = DELIMITADOR
            .CMC7Associado = DELIMITADOR
            .Duplicidade = "0" & DELIMITADOR
            .CodCenape = "0" & DELIMITADOR
            .CodBarComplem = "N" & DELIMITADOR
            .RetornoTransacao = "" 'não vai delimitador
        End With
        ''''''''''''''''''''''''''''''''''
        'Campos necessarios para inserção'
        ''''''''''''''''''''''''''''''''''
        vwDocto(iCountVWDocto).Valor = FormataString(Val(Valor) / 100, "0", Len(vwDocto(iCountVWDocto).Valor) - 1, True) & DELIMITADOR
        vwDocto(iCountVWDocto).Frente = vCapa(iContaLinhas_RF).Frente & DELIMITADOR
        vwDocto(iCountVWDocto).Verso = IIf(Trim(vCapa(iContaLinhas_RF).Verso) = "", vCapa(iContaLinhas_RF).Frente, vCapa(iContaLinhas_RF).Verso) & DELIMITADOR
        vwDocto(iCountVWDocto).OrdemCaptura = FormataString(iOrdemCaptura, "0", Len(vwDocto(iCountVWDocto).OrdemCaptura) - 1, True) & DELIMITADOR
        vwDocto(iCountVWDocto).CrLf = vbCrLf
        m_TotalDoctos = m_TotalDoctos + 1
        iOrdemCaptura = iOrdemCaptura + 1
        
        iContaLinhas_RF = iContaLinhas_RF + 1
        iCountVWDocto = iCountVWDocto + 1
    Loop
    Close #iRFile
    
    sCapa = ""
    
    '
    'Caso exista duplicidade de captura, apagar todas referencias do mesmo CMC7 ou Cd.Barras
    '
    If iCountVWDocto > 0 Then
        For j = 0 To (iCountVWDocto - 2)
            ExisteDuplicidadeCaptura = False
            For i = (j + 1) To (iCountVWDocto - 1)
                If (Trim(vwDocto(j).Leitura) <> "") And (Trim(vwDocto(i).Leitura) <> "") Then
                    If vwDocto(j).Leitura = vwDocto(i).Leitura Then
                        If vwDocto(i).TipoDocto <> "1;" Then
                            ExisteDuplicidadeCaptura = True
                            vwDocto(i).Leitura = FormataString("", " ", Len(vwDocto(i).Leitura) - 1, False) & DELIMITADOR
                        End If
                    End If
                End If
                If vwDocto(i).Frente = vwDocto(j).Frente Then
                    ExisteDuplicidadeCaptura = True
                    vwDocto(i).Leitura = FormataString("", " ", Len(vwDocto(i).Leitura) - 1, False) & DELIMITADOR
                End If
            Next i
            If ExisteDuplicidadeCaptura Then
                'Nao excluir o primeiro docto referente a capa
                If j <> 0 Then
                    vwDocto(j).Leitura = FormataString("", " ", Len(vwDocto(j).Leitura) - 1, False) & DELIMITADOR
                End If
            End If
        Next j
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Gera o arquivo de saida para ser executado o BULK INSERT'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ''''''''''''''''''''''''''''''''''''''''''''
    'Se o diretorio de criação não existe, cria'
    'o diretorio com base na DataProcessamento '
    ''''''''''''''''''''''''''''''''''''''''''''
    If Not DirExists(m_DiretorioCriacao) Then
        If Not CriaDir(m_DiretorioCriacao) Then GoTo Erro_ProcessarLayoutAntigo
    End If
    
    '''''''''''''''''''''''''''''''
    'Se existe os arquivos, exclui'
    '''''''''''''''''''''''''''''''
    If FileExist(m_DiretorioCriacao & m_ArquivoCapaCriacao) Then
        Kill m_DiretorioCriacao & m_ArquivoCapaCriacao
    End If
    If FileExist(m_DiretorioCriacao & m_ArquivoDoctoCriacao) Then
        Kill m_DiretorioCriacao & m_ArquivoDoctoCriacao
    End If
    

    '''''''''''''''''''''''''
    'Grava o arquivo de capa'
    '''''''''''''''''''''''''
    iWFile = FreeFile
    Open m_DiretorioCriacao & m_ArquivoCapaCriacao For Binary As #iWFile
        Put #iWFile, , vwCapa
    Close #iWFile
    FileCopy m_DiretorioCriacao & m_ArquivoCapaCriacao, m_DiretorioDestino & m_ArquivoCapaCriacao
    
    '''''''''''''''''''''''''''''''
    'Monta SQL de seleção de capas'
    '''''''''''''''''''''''''''''''
           sStr = "SELECT IdCapa, Capa"
    sStr = sStr & "  FROM CAPA"
    sStr = sStr & " WHERE DataProcessamento = " & m_DataProcessamento
'    sStr = sStr & "   AND IdLote = " & m_IdLote
'    sStr = sStr & "   AND Status = 'U'"
    sStr = sStr & "   AND "
    
    sStr2 = ""
    For i = 0 To UBound(vwCapa)
        If Len(sStr2) > 0 Then sStr2 = sStr2 & " OR "
        sStr2 = sStr2 & "(IdLote = " & m_IdLote & " AND Status = 'U' AND "
        sStr2 = sStr2 & "Capa = " & Val(Left(vwCapa(i).Capa, Len(vwCapa(i).Capa) - 1))
        sStr2 = sStr2 & " AND DataCriacao = '" & Left(vwCapa(i).DataCriacao, Len(vwCapa(i).DataCriacao) - 1) & "')"
    Next i
    sStr = sStr & sStr2
    
    '''''''''''''''''''''''''''
    'Insere o arquivo de capas'
    '''''''''''''''''''''''''''
    m_Connection.Execute "BULK INSERT Capa From '" & m_DiretorioBULK & m_ArquivoCapaCriacao & "' WITH ( " & _
                         "DATAFILETYPE = 'char'," & _
                         "FIELDTERMINATOR = ';', " & _
                         "FIRSTROW = 1," & _
                         "MAXERRORS = 0, " & _
                         "ROWTERMINATOR = '\n'" & _
                         ")"

    Set rstIdCapas = m_Connection.OpenResultset(sStr, rdOpenStatic)
    
    If Not rstIdCapas.EOF() Then
        rstIdCapas.MoveNext
        rstIdCapas.MoveFirst
        
        For j = 0 To UBound(vIndiceDocto)
        
            sCapa = vIndiceDocto(j)
            
            If sCapa <> sCapa2 Then
                sCapa2 = sCapa
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'é a maneira que eu encontrei para certificar de que estou no IdCapa correto'
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                rstIdCapas.MoveFirst
                Do While Val(rstIdCapas!Capa) <> Val(Left(vIndiceDocto(j), Len(vIndiceDocto(j)) - 1))
                    rstIdCapas.MoveNext
                Loop
            End If
            
            vwDocto(j).IdCapa = FormataString(rstIdCapas!IdCapa, "0", Len(vwDocto(j).IdCapa) - 1, True) & DELIMITADOR
        Next j
    End If
    
    '''''''''''''''''''''''''''''''
    'Grava o arquivo de documentos'
    '''''''''''''''''''''''''''''''
    iWFile = FreeFile
    Open m_DiretorioCriacao & m_ArquivoDoctoCriacao For Binary As #iWFile
        Put #iWFile, , vwDocto
    Close #iWFile
    FileCopy m_DiretorioCriacao & m_ArquivoDoctoCriacao, m_DiretorioDestino & m_ArquivoDoctoCriacao
    
    
    ''''''''''''''''''''''''''''''''
    'Insere o arquivo de documentos'
    ''''''''''''''''''''''''''''''''
    m_Connection.Execute "BULK INSERT Documento From '" & m_DiretorioBULK & m_ArquivoDoctoCriacao & "' WITH ( " & _
                         "DATAFILETYPE = 'char'," & _
                         "FIELDTERMINATOR = ';', " & _
                         "FIRSTROW = 1," & _
                         "MAXERRORS = 0, " & _
                         "ROWTERMINATOR = '\n'" & _
                         ")"
    
    '''''''''''''''''''''''''''''''''''''''''
    'Por último, atualiza a tabela documento'
    '''''''''''''''''''''''''''''''''''''''''

    rstIdCapas.MoveFirst
    
    sStr2 = ""
    
           sStr = "UPDATE CAPA SET "
    sStr = sStr & "STATUS = '1'"
    sStr = sStr & " WHERE DataProcessamento = " & m_DataProcessamento
    sStr = sStr & "   AND IdCapa IN ("
    For i = 0 To rstIdCapas.RowCount - 1
        If Len(sStr2) > 0 Then sStr2 = sStr2 & ","
        sStr2 = sStr2 & rstIdCapas!IdCapa
        rstIdCapas.MoveNext
    Next i
    
    rstIdCapas.MoveFirst
    Do While Not rstIdCapas.EOF()
        GravaLog rstIdCapas!IdCapa, 0, 40
        rstIdCapas.MoveNext
    Loop
    
    sAtualizaStatusCapa = sStr & sStr2 & ") AND Status = 'U'"

    m_Connection.Execute sAtualizaStatusCapa
    
    rstIdCapas.Close
   
    ProcessarLayoutAntigo = True
    
    '''''''''''''''''''''''''''''''''''''''''''''''
    'Excluir os arquivos criados, LOCAL e SERVIDOR'
    '''''''''''''''''''''''''''''''''''''''''''''''
    If m_ExcluirArquivoCriacao Then
    
        On Error Resume Next
        'LOCAL
        '''''''''''''''''''''''''''
        'Exclui o arquivo de capas'
        '''''''''''''''''''''''''''
        Kill m_DiretorioCriacao & m_ArquivoCapaCriacao
        ''''''''''''''''''''''''''''''''
        'Exclui o arquivo de documentos'
        ''''''''''''''''''''''''''''''''
        Kill m_DiretorioCriacao & m_ArquivoDoctoCriacao
        
        'SERVIDOR
        '''''''''''''''''''''''''''
        'Exclui o arquivo de capas'
        '''''''''''''''''''''''''''
        Kill m_DiretorioDestino & m_ArquivoCapaCriacao
        ''''''''''''''''''''''''''''''''
        'Exclui o arquivo de documentos'
        ''''''''''''''''''''''''''''''''
        Kill m_DiretorioDestino & m_ArquivoDoctoCriacao
        
        On Error GoTo 0
    
    End If
    
    Screen.MousePointer = vbDefault

    Exit Function
Erro_ProcessarLayoutAntigo:

    Screen.MousePointer = vbDefault

    m_Err = Err
    m_MsgError = Error

End Function


Private Function ProcessarLayoutNovo() As Boolean

    Dim vCapa()             As tpRetornoVipsNovaDLL
    Dim vwCapa()            As clsTpCapa        'Tipo para gravar no arquivo de BULK INSERT
    Dim vwDocto()           As clsTpDocumento   'Tipo para gravar no arquivo de BULK INSERT
    Dim rstIdCapas          As RDO.rdoResultset
    Dim rstIdDoctos         As RDO.rdoResultset
    Dim vIndiceDocto()      As String
    Dim iRFile              As Integer
    Dim iWFile              As Integer
    Dim iContaLinhas_RF     As Integer          'Contador de linhas do arquivo de leitura
    Dim iCountVWCapa        As Integer
    Dim iCountVWDocto       As Integer
    Dim iOrdemCaptura       As Integer
    Dim sCampo1             As String
    Dim sCampo2             As String
    Dim sCampo3             As String
    Dim svalor              As String
    Dim sIdEnv_Mal          As String
    Dim IdCapa              As Long
    Dim sTipoDocto          As String
    Dim iIndexDocto         As Integer
    Dim bVirtual            As Boolean
    Dim i                   As Integer
    Dim j                   As Integer
    Dim sStr                As String
    Dim sStr2               As String
    Dim sAtualizaStatusCapa As String
    Dim sCapa               As String
    Dim sCapa2              As String
    Dim ExisteDuplicidadeCaptura As Boolean
    
    On Error GoTo Erro_ProcessarLayoutNovo
    
    ProcessarLayoutNovo = False
    
    Screen.MousePointer = vbHourglass
    
    iRFile = FreeFile
    IdCapa = 0
    iIndexDocto = 0
    sCapa = ""
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Abre o arquivo de dados gerado pela DLL para a leitura dos documentos'
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Open m_DiretorioOrigem & m_ArquivoOrigem For Binary As #iRFile
    
    Do While Not EOF(iRFile)
        sTipoDocto = 0
        
        '''''''''''''''''''''''''''''''
        'Redimensiona o vetor de capas'
        '''''''''''''''''''''''''''''''
        ReDim Preserve vCapa(iContaLinhas_RF) As tpRetornoVipsNovaDLL

        Get #iRFile, , vCapa(iContaLinhas_RF)
        
        If vCapa(iContaLinhas_RF).Tipo <> "A" And vCapa(iContaLinhas_RF).Tipo <> "B" _
            And vCapa(iContaLinhas_RF).Tipo <> "E" Then Exit Do
        
        '''''''''''''''''''''''''''''''''''''''''''
        'Tratamento da linha do arquivo de leitura'
        '''''''''''''''''''''''''''''''''''''''''''
        vCapa(iContaLinhas_RF).Leitura = TrataLeitura(vCapa(iContaLinhas_RF).Leitura)
        
        If VerificaSeCapa(vCapa(iContaLinhas_RF).Leitura, sIdEnv_Mal) Then
            If sIdEnv_Mal = "M" Then
                TratarCamposCMC7 vCapa(iContaLinhas_RF).Leitura, sCampo1, sCampo2, sCampo3, svalor
            End If
            sTipoDocto = 1 ' Capa
            bVirtual = False
            ' gravar capa

            ReDim Preserve vwCapa(iCountVWCapa) As clsTpCapa

            With vwCapa(iCountVWCapa)
                '.PosCapa = "   "
                .DataProcessamento = m_DataProcessamento & DELIMITADOR
                .IdEnv_Mal = sIdEnv_Mal & DELIMITADOR
                .Capa = FormataString(IIf(sIdEnv_Mal = "E", Val(Trim(vCapa(iContaLinhas_RF).Leitura)), "0" & Mid(sCampo2, 1, 9) & Mid(sCampo1, 4, 4)), " ", Len(.Capa) - 1, False) & DELIMITADOR
                .IdCapa = "0" & DELIMITADOR
                .IdLote = FormataString(m_IdLote, "0", Len(.IdLote) - 1, True) & DELIMITADOR
                .IdEnv_Mal = sIdEnv_Mal & DELIMITADOR
                .Num_Malote = "0" & DELIMITADOR
                .AgOrig = FormataString("0", "0", Len(.AgOrig) - 1, True) & DELIMITADOR
                .Status = "U" & DELIMITADOR
                .DataCriacao = Format(Now(), m_FormatoData) & DELIMITADOR '"dd-mm-yyyy hh:mm:00") & DELIMITADOR
                
                .Alcada = DELIMITADOR
                .PendenciaValor = DELIMITADOR
                .Supervisor = DELIMITADOR
                .VinculoManual = DELIMITADOR
                .IgnorarProva0 = DELIMITADOR
                .Ocorrencia = DELIMITADOR
                .Conta = DELIMITADOR
                .Dinheiro = DELIMITADOR
                .Diferenca = DELIMITADOR
                .Duplicidade = "0" & DELIMITADOR
                .HoraAtual = DELIMITADOR
                .RecepcionadoIK = "N" 'não vai delimitador
                .CrLf = vbCrLf
            End With
            sCapa = vwCapa(iCountVWCapa).Capa
            m_TotalCapas = m_TotalCapas + 1
            iOrdemCaptura = 1
            iCountVWCapa = iCountVWCapa + 1
            
        ElseIf sCapa = "" Then 'so pode entrar aqui quando primeira linha
        
            sTipoDocto = 1 ' Capa
            bVirtual = True
            sCapa = "9"

            ReDim Preserve vwCapa(iCountVWCapa) As clsTpCapa
            
            With vwCapa(iCountVWCapa)
                .DataProcessamento = m_DataProcessamento & DELIMITADOR
                .IdEnv_Mal = sIdEnv_Mal & DELIMITADOR
                .Capa = FormataString("9", "0", Len(.Capa) - 1, True) & DELIMITADOR
                .IdCapa = "0" & DELIMITADOR
                .IdLote = FormataString(m_IdLote, "0", Len(.IdLote) - 1, True) & DELIMITADOR
                .IdEnv_Mal = "E" & DELIMITADOR
                .Num_Malote = "0" & DELIMITADOR
                .AgOrig = FormataString("0", "0", Len(.AgOrig) - 1, True) & DELIMITADOR
                .Status = "U" & DELIMITADOR
                .DataCriacao = Format(Now(), m_FormatoData) & DELIMITADOR '"dd-mm-yyyy hh:mm:00") & DELIMITADOR
                .Alcada = DELIMITADOR
                .PendenciaValor = DELIMITADOR
                .Supervisor = DELIMITADOR
                .VinculoManual = DELIMITADOR
                .IgnorarProva0 = DELIMITADOR
                .Ocorrencia = DELIMITADOR
                .Conta = DELIMITADOR
                .Dinheiro = DELIMITADOR
                .Diferenca = DELIMITADOR
                .Duplicidade = "0" & DELIMITADOR
                .HoraAtual = DELIMITADOR
                .RecepcionadoIK = "N" 'não vai delimitador
                .CrLf = vbCrLf
            End With

            iCountVWCapa = iCountVWCapa + 1
            m_TotalCapas = m_TotalCapas + 1
            iOrdemCaptura = 1
            
        End If
        ' gravar documento
        svalor = "000"
        
        ReDim Preserve vIndiceDocto(iIndexDocto) As String
        ReDim Preserve vwDocto(iCountVWDocto) As clsTpDocumento

        
        vIndiceDocto(iIndexDocto) = vwCapa(iCountVWCapa - 1).Capa
        iIndexDocto = iIndexDocto + 1
        
        vwDocto(iCountVWDocto).DataProcessamento = m_DataProcessamento & DELIMITADOR
        vwDocto(iCountVWDocto).TipoDocto = sTipoDocto & DELIMITADOR
        
        
        If sIdEnv_Mal = "M" And sTipoDocto = 1 Then ' Capa de Malote
            If Not bVirtual Then
                vwDocto(iCountVWDocto).Leitura = FormataString("0" & Mid(sCampo2, 1, 9) & Mid(sCampo1, 4, 4), " ", Len(vwDocto(iCountVWDocto).Leitura) - 1, False) & DELIMITADOR
            Else
                vwDocto(iCountVWDocto).Leitura = FormataString("9", " ", Len(vwDocto(iCountVWDocto).Leitura) - 1, False) & DELIMITADOR 'Capa Virtual
            End If
        
        ElseIf sIdEnv_Mal = "E" And sTipoDocto = 1 Then ' Capa de Envelope
            If Not bVirtual Then
                vwDocto(iCountVWDocto).Leitura = FormataString(Trim(vCapa(iContaLinhas_RF).Leitura), " ", Len(vwDocto(iCountVWDocto).Leitura) - 1, False) & DELIMITADOR
            Else
                vwDocto(iCountVWDocto).Leitura = FormataString("9", " ", Len(vwDocto(iCountVWDocto).Leitura) - 1, False) & DELIMITADOR 'Capa Virtual
            End If
        
        ElseIf vCapa(iContaLinhas_RF).Tipo = "A" Then ' Docto com CMC7
            Valor = ""
            TratarCamposCMC7 vCapa(iContaLinhas_RF).Leitura, sCampo1, sCampo2, sCampo3, svalor
            vwDocto(iCountVWDocto).Leitura = FormataString(sCampo1 & sCampo2 & sCampo3, " ", Len(vwDocto(iCountVWDocto).Leitura) - 1, False) & DELIMITADOR
        
        ElseIf vCapa(iContaLinhas_RF).Tipo = "B" Then ' Docto com Cod Barras
            vwDocto(iCountVWDocto).Leitura = FormataString(RPad(Trim(vCapa(iContaLinhas_RF).Leitura), 44), " ", Len(vwDocto(iCountVWDocto).Leitura) - 1, False) & DELIMITADOR
        Else
            vwDocto(iCountVWDocto).Leitura = FormataString(Trim(vCapa(iContaLinhas_RF).Leitura), " ", Len(vwDocto(iCountVWDocto).Leitura) - 1, False) & DELIMITADOR
        End If

        ''''''''''''''''''''''''''''''''''''
        'Campos em branco porem necessarios'
        ''''''''''''''''''''''''''''''''''''
        With vwDocto(iCountVWDocto)
            .IdDocto = "0" & DELIMITADOR
            .Ocorrencia = DELIMITADOR
            .CodMotivo = "0" & DELIMITADOR
            .Status = "0" & DELIMITADOR
            .Alcada = "N" & DELIMITADOR
            .Autenticado = "0" & DELIMITADOR
            .Cortado = DELIMITADOR
            .OcorrenciaOK = DELIMITADOR
            .AutenticacaoDigital = DELIMITADOR
            .Ordem = "0" & DELIMITADOR
            .NSU = DELIMITADOR
            .Terminal = DELIMITADOR
            .Vinculo = "0" & DELIMITADOR
            .AgenciaVinculo = DELIMITADOR
            .ContaVinculo = DELIMITADOR
            .CMC7Associado = DELIMITADOR
            .Duplicidade = "0" & DELIMITADOR
            .CodCenape = "0" & DELIMITADOR
            .CodBarComplem = "N" & DELIMITADOR
            .RetornoTransacao = "" 'não vai delimitador
        End With
        ''''''''''''''''''''''''''''''''''
        'Campos necessarios para inserção'
        ''''''''''''''''''''''''''''''''''
        vwDocto(iCountVWDocto).Valor = FormataString(Val(Valor) / 100, "0", Len(vwDocto(iCountVWDocto).Valor) - 1, True) & DELIMITADOR
        vwDocto(iCountVWDocto).Frente = vCapa(iContaLinhas_RF).Frente & DELIMITADOR
        vwDocto(iCountVWDocto).Verso = IIf(Trim(vCapa(iContaLinhas_RF).Verso) = "", vCapa(iContaLinhas_RF).Frente, vCapa(iContaLinhas_RF).Verso) & DELIMITADOR
        vwDocto(iCountVWDocto).OrdemCaptura = FormataString(iOrdemCaptura, "0", Len(vwDocto(iCountVWDocto).OrdemCaptura) - 1, True) & DELIMITADOR
        vwDocto(iCountVWDocto).CrLf = vbCrLf
        m_TotalDoctos = m_TotalDoctos + 1
        iOrdemCaptura = iOrdemCaptura + 1
        
        iContaLinhas_RF = iContaLinhas_RF + 1
        iCountVWDocto = iCountVWDocto + 1
    Loop
    Close #iRFile
    
    sCapa = ""
    
    '
    'Caso exista duplicidade de captura, apagar todas referencias do mesmo CMC7 ou Cd.Barras
    '
    If iCountVWDocto > 0 Then
        For j = 0 To (iCountVWDocto - 2)
            ExisteDuplicidadeCaptura = False
            For i = (j + 1) To (iCountVWDocto - 1)
                If (Trim(vwDocto(j).Leitura) <> "") And (Trim(vwDocto(i).Leitura) <> "") Then
                    If vwDocto(j).Leitura = vwDocto(i).Leitura Then
                        If vwDocto(i).TipoDocto <> "1;" Then
                            ExisteDuplicidadeCaptura = True
                            vwDocto(i).Leitura = FormataString("", " ", Len(vwDocto(i).Leitura) - 1, False) & DELIMITADOR
                        End If
                    End If
                End If
                If vwDocto(i).Frente = vwDocto(j).Frente Then
                    ExisteDuplicidadeCaptura = True
                    vwDocto(i).Leitura = FormataString("", " ", Len(vwDocto(i).Leitura) - 1, False) & DELIMITADOR
                End If
            Next i
            If ExisteDuplicidadeCaptura Then
                'Nao excluir o primeiro docto referente a capa
                If j <> 0 Then
                    vwDocto(j).Leitura = FormataString("", " ", Len(vwDocto(j).Leitura) - 1, False) & DELIMITADOR
                End If
            End If
        Next j
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Gera o arquivo de saida para ser executado o BULK INSERT'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ''''''''''''''''''''''''''''''''''''''''''''
    'Se o diretorio de criação não existe, cria'
    'o diretorio com base na DataProcessamento '
    ''''''''''''''''''''''''''''''''''''''''''''
    If Not DirExists(m_DiretorioCriacao) Then
        If Not CriaDir(m_DiretorioCriacao) Then GoTo Erro_ProcessarLayoutNovo
    End If
    
    '''''''''''''''''''''''''''''''
    'Se existe os arquivos, exclui'
    '''''''''''''''''''''''''''''''
    If FileExist(m_DiretorioCriacao & m_ArquivoCapaCriacao) Then
        Kill m_DiretorioCriacao & m_ArquivoCapaCriacao
    End If
    If FileExist(m_DiretorioCriacao & m_ArquivoDoctoCriacao) Then
        Kill m_DiretorioCriacao & m_ArquivoDoctoCriacao
    End If
    

    '''''''''''''''''''''''''
    'Grava o arquivo de capa'
    '''''''''''''''''''''''''
    iWFile = FreeFile
    Open m_DiretorioCriacao & m_ArquivoCapaCriacao For Binary As #iWFile
        Put #iWFile, , vwCapa
    Close #iWFile
    FileCopy m_DiretorioCriacao & m_ArquivoCapaCriacao, m_DiretorioDestino & m_ArquivoCapaCriacao
    
    '''''''''''''''''''''''''''''''
    'Monta SQL de seleção de capas'
    '''''''''''''''''''''''''''''''
           sStr = "SELECT IdCapa, Capa"
    sStr = sStr & "  FROM CAPA"
    sStr = sStr & " WHERE DataProcessamento = " & m_DataProcessamento
'    sStr = sStr & "   AND IdLote = " & m_IdLote
'    sStr = sStr & "   AND Status = 'U'"
    sStr = sStr & "   AND "
    
    sStr2 = ""
    For i = 0 To UBound(vwCapa)
        If Len(sStr2) > 0 Then sStr2 = sStr2 & " OR "
        sStr2 = sStr2 & "(IdLote = " & m_IdLote & " AND Status = 'U' AND "
        sStr2 = sStr2 & "Capa = " & Val(Left(vwCapa(i).Capa, Len(vwCapa(i).Capa) - 1))
        sStr2 = sStr2 & " AND DataCriacao = '" & Left(vwCapa(i).DataCriacao, Len(vwCapa(i).DataCriacao) - 1) & "')"
    Next i
    sStr = sStr & sStr2
    
    '''''''''''''''''''''''''''
    'Insere o arquivo de capas'
    '''''''''''''''''''''''''''
    m_Connection.Execute "BULK INSERT Capa From '" & m_DiretorioBULK & m_ArquivoCapaCriacao & "' WITH ( " & _
                         "DATAFILETYPE = 'char'," & _
                         "FIELDTERMINATOR = ';', " & _
                         "FIRSTROW = 1," & _
                         "MAXERRORS = 0, " & _
                         "ROWTERMINATOR = '\n'" & _
                         ")"

    Set rstIdCapas = m_Connection.OpenResultset(sStr, rdOpenStatic)
    
    If Not rstIdCapas.EOF() Then
        rstIdCapas.MoveNext
        rstIdCapas.MoveFirst
        
        For j = 0 To UBound(vIndiceDocto)
        
            sCapa = vIndiceDocto(j)
            
            If sCapa <> sCapa2 Then
                sCapa2 = sCapa
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'é a maneira que eu encontrei para certificar de que estou no IdCapa correto'
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                rstIdCapas.MoveFirst
                Do While Val(rstIdCapas!Capa) <> Val(Left(vIndiceDocto(j), Len(vIndiceDocto(j)) - 1))
                    rstIdCapas.MoveNext
                Loop
            End If
            
            vwDocto(j).IdCapa = FormataString(rstIdCapas!IdCapa, "0", Len(vwDocto(j).IdCapa) - 1, True) & DELIMITADOR
        Next j
    End If
    
    '''''''''''''''''''''''''''''''
    'Grava o arquivo de documentos'
    '''''''''''''''''''''''''''''''
    iWFile = FreeFile
    Open m_DiretorioCriacao & m_ArquivoDoctoCriacao For Binary As #iWFile
        Put #iWFile, , vwDocto
    Close #iWFile
    FileCopy m_DiretorioCriacao & m_ArquivoDoctoCriacao, m_DiretorioDestino & m_ArquivoDoctoCriacao
    
    
    ''''''''''''''''''''''''''''''''
    'Insere o arquivo de documentos'
    ''''''''''''''''''''''''''''''''
    m_Connection.Execute "BULK INSERT Documento From '" & m_DiretorioBULK & m_ArquivoDoctoCriacao & "' WITH ( " & _
                         "DATAFILETYPE = 'char'," & _
                         "FIELDTERMINATOR = ';', " & _
                         "FIRSTROW = 1," & _
                         "MAXERRORS = 0, " & _
                         "ROWTERMINATOR = '\n'" & _
                         ")"
    
    '''''''''''''''''''''''''''''''''''''''''
    'Por último, atualiza a tabela documento'
    '''''''''''''''''''''''''''''''''''''''''

    rstIdCapas.MoveFirst
    
    sStr2 = ""
    
           sStr = "UPDATE CAPA SET "
    sStr = sStr & "STATUS = '1'"
    sStr = sStr & " WHERE DataProcessamento = " & m_DataProcessamento
    sStr = sStr & "   AND IdCapa IN ("
    For i = 0 To rstIdCapas.RowCount - 1
        If Len(sStr2) > 0 Then sStr2 = sStr2 & ","
        sStr2 = sStr2 & rstIdCapas!IdCapa
        rstIdCapas.MoveNext
    Next i
    
    rstIdCapas.MoveFirst
    Do While Not rstIdCapas.EOF()
        GravaLog rstIdCapas!IdCapa, 0, 40
        rstIdCapas.MoveNext
    Loop
    
    sAtualizaStatusCapa = sStr & sStr2 & ") AND Status = 'U'"

    m_Connection.Execute sAtualizaStatusCapa
    
    rstIdCapas.Close
   
    ProcessarLayoutNovo = True
    
    '''''''''''''''''''''''''''''''''''''''''''''''
    'Excluir os arquivos criados, LOCAL e SERVIDOR'
    '''''''''''''''''''''''''''''''''''''''''''''''
    If m_ExcluirArquivoCriacao Then
    
        On Error Resume Next
        'LOCAL
        '''''''''''''''''''''''''''
        'Exclui o arquivo de capas'
        '''''''''''''''''''''''''''
        Kill m_DiretorioCriacao & m_ArquivoCapaCriacao
        ''''''''''''''''''''''''''''''''
        'Exclui o arquivo de documentos'
        ''''''''''''''''''''''''''''''''
        Kill m_DiretorioCriacao & m_ArquivoDoctoCriacao
        
        'SERVIDOR
        '''''''''''''''''''''''''''
        'Exclui o arquivo de capas'
        '''''''''''''''''''''''''''
        Kill m_DiretorioDestino & m_ArquivoCapaCriacao
        ''''''''''''''''''''''''''''''''
        'Exclui o arquivo de documentos'
        ''''''''''''''''''''''''''''''''
        Kill m_DiretorioDestino & m_ArquivoDoctoCriacao
        
        On Error GoTo 0
    
    End If
    
    Screen.MousePointer = vbDefault

    Exit Function
Erro_ProcessarLayoutNovo:

    Screen.MousePointer = vbDefault

    m_Err = Err
    m_MsgError = Error

End Function


