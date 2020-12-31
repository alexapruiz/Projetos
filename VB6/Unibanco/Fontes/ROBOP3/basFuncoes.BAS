Attribute VB_Name = "basFuncoes"
Option Explicit

Type tpModuloBas
    AgenciaVinculo          As String
    AgenCob                 As Long
    BHAceitaCobranca        As Byte
    Capa                    As String
    CapaBack                As Integer
    CaixaAberto             As Boolean
    CaixaAnterior           As Integer
    CodCVTCob               As Long
    CodOcorrencia           As Integer
    CodTransacao            As String
    Criptografia            As Boolean
    Transacao               As String
    ContaCob                As Double
    ContaVinculo            As String
    DataMovAnt              As Long
    DataProcessamento       As String
    Deposito                As String
    ehVinculoManual         As Boolean
    Evento                  As Integer
    GereiLog                As Integer
    FecharCaixa             As Boolean
    Hora                    As Integer
    hsSQLa                  As String
    hsSQLb                  As String
    IdentDep                As Long
    idEnvMal                As String * 1
    IndTransac              As String
    JAM                     As String
    Multa                   As String
    NossoNumCob             As String
    NroDeposito             As String
    PreparouLog             As Integer
    PrimeiraVez             As Boolean
    RecebendoCapa           As Boolean
    RetDigMod11             As Integer
    RetTransacao            As Integer
    
    rst                     As Recordset
    rstCapa                 As Recordset
    rstDoctos               As Recordset
    rstDocto                As Recordset
    
    TpCtaBHQQ               As Integer
    SeqBHQC                 As Long
    SeqPagto                As Long
    SeqRecto                As Long
    TipoOperacaoDeposito    As String
    TipoConta               As String

    TipoTransacao           As Integer
    TpRep                   As Integer
    ValorAbat               As String
    ValorCob                As String
    ValorDesc               As String
    ValorMora               As String
    ValorTrans              As String
    VencCob                 As Long
    VincProcAnt             As Double
    Vinculo                 As Double
 End Type
 Global Geral As tpModuloBas
 
 Type tpOpcoes
    ClearCapaCx     As Boolean
'   ClearDoctoCX    As Boolean
    FechaCx         As Boolean
    InterValo       As Integer
    InterFixo       As Boolean
 End Type
 Global AntIniOpcoes As tpOpcoes
  
 Enum EnumOrigem
    EnumCheque = 0
    EnumADCC = 1
    EnumOutros = 2
    EnumDeposito = 3
 End Enum
 
 Enum EnumCores
    nulo = 0
    azul = vbBlue
    Verde = vbGreen
    Vermelho = vbRed
    PRETO = vbBlack
    Amarelo = vbYellow
    Rosa = vbMagenta
    Cian = vbCyan
 End Enum
 
 Public Enum enumRetornoMensagemErro
    eSair
    eRepetir
    eContinuar
    eFinalizar
 End Enum

 Enum EnumErroAcao
    eCapa = 0
    eDoctoProcesso = 1
    eDoctoSubidaLog = 2
    eDefault = 3
 End Enum
 
'Funcoes
 Public Const MSG_CapaProblemaBatimentoCSP = 0
 Public Const MSG_CapaProblemaBatimento = 1
 Public Const MSG_DiferencaFechamentoCx = 2
 Public Const MSG_AjusteAgenContaInexis = 3
 Public Const MSG_TipoDoctoNaoPodeEstor = 4
 Public Const MSG_EstornoNaoAutorizadoBH = 5
 Public Const MSG_AgenCapanaoCadastrada = 6
 Public Const MSG_FalhaEstornoDocumento = 7
 Public Const MSG_NaoPermitidaAgCentral = 8
 Public Const MSG_AjusteAgenContaInexisCSP = 9
 
'Declaração de função para a leitura do .INI
 Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
 Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'Declaração da DLL calculo do NSU e Calculo do SDV
 Public Declare Function QXGetSDV Lib "qxnsusdv32.dll" (ByVal PCC As String, ByVal Caixa As String, ByVal DataAutentic As String, ByVal Tipo As String, ByVal Valor As String, ByVal CIF As String, ByVal Ret As String) As Integer
 Public Declare Function QXCalNsu Lib "qxnsusdv32.dll" (ByVal PCC As String, ByVal Caixa As String, ByVal DataAutentic As String, ByVal Tipo As String, ByVal Valor As String, ByVal Ret As String) As Integer

'Declaracao da DLL QXHOST - comunicação com Host do Unibanco
 Public Declare Function qxhostnt Lib "qxhostnt.DLL" (ByVal A1 As String, ByVal A2 As String, ByVal A3 As String) As Integer
 
 Public MDIQuery        As New clsQueryMDI
 Public UBBQuery        As New ClsQueryUBB
 Public Caixa           As New ClsCaixa
 Public Parametros      As New clsParametro
 
 Public Const TotalTabelas = 10
 Public Const LenLabelProgress = 4005
 
'constantes para estorno
 Public Const CodigoDoctoNoEstorno = "F"
 Public Const CodigoDoctoEstornado = "S"
 Public Const CodigoDoctoEmEstorno = "P"
 
'Constantes Status
 Public Const IdModuloRobo = 33
 Public Const ST_CapaTransmitida = "T"
 Public Const ST_CapaParaIlegivel = 5
 Public Const ST_CapaParaCSP = "N"
 Public Const ST_CapaOcorrencia = "X"
 Public Const ST_CapaCorrecaoAgConta = "Y"
 Public Const ST_CapaVinculoManual = "7"
 Public Const ST_DoctoDeletadaRobo = "F"
 Public Const ST_DoctoCorrecaoAgConta = "Y"
 Sub Main()
   
    ''''''''''''''''''''''''''''''''''''''
    ' Chave do algorito de criptografia  '
    ''''''''''''''''''''''''''''''''''''''
    Key(0) = 97
    Key(1) = 150
    Key(2) = 127
    Key(3) = 254
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Verificar se o programa foi aberto mais de 1 vez '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    If App.PrevInstance Then
        MsgBox "Programa já esta sendo executado, não é possível executar outra cópia.", vbExclamation + vbOKOnly, App.Title
        End
    End If
    
    Call ValidaDLL(True)
                
    frmPrincipal.Show vbModeless
    
    DoEvents
    FrmPassword.Show vbModal
       
    If FrmPassword.SenhaOk Then
       Unload FrmPassword
       frmComunica.Show vbModal
    End If
    
    Call CopiaLogs
        
    End
End Sub
Public Sub GravarOpcaoINI(ByVal pvsSecao As String, pvsItem As String, ByVal pvsValor As String)
    Dim iRet As Long
    
    iRet = WritePrivateProfileString(pvsSecao, pvsItem, pvsValor, App.Path & "\Robo.ini")
End Sub
Public Function PegarOpcaoINI(ByVal pvsSecao As String, pvsItem As String, ByVal pvsDefault As String) As String
    
    Dim iRet        As Integer
    Dim sDado       As String
    Dim sDadoAux    As String
    Dim i           As Integer
    Dim sArquivoINI As String
    
    sDado = String(255, " ")
    
   'Abrir arquivo INI conforme opção
    If pvsSecao = "Conexao" And (pvsItem = "Senha" Or pvsItem = "Usuario") Then
        sArquivoINI = App.Path & "\MDI_Conexao.INI"
    Else
        sArquivoINI = App.Path & "\Robo.INI"
    End If

    sDado = String(255, " ")
    iRet = GetPrivateProfileString(pvsSecao, pvsItem, pvsDefault, sDado, 255, sArquivoINI)
    
    sDado = Trim(sDado)
    sDadoAux = ""
    
    For i = 1 To Len(sDado)
        If pvsSecao = "Conexao" And pvsItem = "Senha" Then
            If Asc(Mid(sDado, i, 1)) <> 0 Then
                sDadoAux = sDadoAux & Mid(sDado, i, 1)
            End If
        Else
            If Asc(Mid(sDado, i, 1)) >= 32 And Asc(Mid(sDado, i, 1)) <= 122 Then
                sDadoAux = sDadoAux & Mid(sDado, i, 1)
            End If
        End If
    Next
    
    If pvsSecao = "Conexao" And pvsItem = "Senha" Then
        PegarOpcaoINI = Decript(Trim(sDadoAux))
    Else
        PegarOpcaoINI = Trim(sDadoAux)
    End If
    
End Function
Function formata(ByVal pValor As String, Optional ByVal DivideCem As Boolean) As String
   
    Dim sValor As String
    
    If DivideCem Then
        sValor = Format(pValor / 100, "000000000000.00")
    Else
        sValor = Format(pValor, "000000000000.00")
    End If
           
    formata = sValor
   
End Function
Function formataValor(ByVal pValor As String, Optional ByVal DivideCem As Boolean) As String
   
   Dim sValor As String
    
    If DivideCem Then
        sValor = Format(pValor / 100, "000000000000.00")
    Else
        sValor = Format(pValor, "000000000000.00")
    End If
    
    Mid(sValor, InStr(sValor, ","), 1) = "."
       
    formataValor = sValor
       
End Function
Sub GetCaixa()
   
On Error GoTo TrataErro

    Dim spRetorno   As Integer
    Dim pCaixa      As Integer
    Dim pNSU        As Integer
    

   '==============================='
   ' Obtem nro de caixa disponivel '
   '==============================='
   
    spRetorno = MDIQuery.getCaixaNSU(Geral.DataProcessamento, Caixa.Estacao, Caixa.CaixaIni, pCaixa, pNSU)
    
    If spRetorno <> 0 Then
        MsgBox "Ocorreu algum erro com leitura do número do caixa disponível. Saia e tente novamente.", vbOKOnly + vbCritical, "Atenção"
        Exit Sub
    End If
    
    Caixa.Caixa = pCaixa
    Caixa.BaseNSU = pNSU
    
    If Caixa.Caixa > Caixa.CaixaFim Then
        MsgBox "Atenção. Não é possível disponibilizar um novo caixa para esta estação pois o número máximo de caixas para este módulo já foi atingido.", vbOKOnly + vbCritical, "Atenção"
        Close #20
        End
    End If

    Exit Sub
    
TrataErro:

    Screen.MousePointer = 0
    Select Case TratamentoErro("Falha na leitura do número de caixa disponível no MDI-Ubb.", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
    End Select
End Sub
Public Sub LocalLog(pLog As String)
   Print #20, Format(Now(), "HH:MM:SS") & " = " & pLog & "    {" & Timer & "}"
End Sub
Function MaloteNovo() As Boolean

    If (Left(CStr(Geral.rstCapa!Num_malote), 2) = "09" And _
              Len(Geral.rstCapa!Num_malote) = 12) Or _
       (Left(CStr(Geral.rstCapa!Num_malote), 1) = "9" And _
              Len(Geral.rstCapa!Num_malote) = 11) Then
       MaloteNovo = True
    Else
       MaloteNovo = False
    End If

End Function
Public Sub Espera(sSegundos As Double)
    Dim Start As Double
    
    Start# = IIf(Timer > 86398.999 And Timer < 86399.999, 0, Timer) 'Segurança para 24:00 hs
        
    While Start# + sSegundos > Timer
        DoEvents
    Wend

End Sub
Sub Modulo11(ByVal nova_str As String)
   
   '---------------------------------------------
   '--------- MODULO 11 (2 BASE 9) --------------
   ' Esta rotina serve para calcular:
   ' 1) número da Capa pessoa fisica
   ' 2) 8+1
   '---------------------------------------------
   
   Dim soma, resto, digito_11, p, peso, ok_1 As Integer
   Dim digito_rv As String
   
   soma = 0
   resto = 0
   digito_11 = 0        'calculado pelo módulo 11
   digito_rv = ""       'caracter digitado pelo operador

   '*************************************************************
   'número da Capa: (8+1)             0 0 9 9 9 9 9 9 9 9 - D
   '                                      x x x x x x x x x x
   'multiplica da direita para esquerda:  3 2 9 8 7 6 5 4 3 2
   '*************************************************************
   
   peso = 2    'começa multiplicar da direita para esquerda
   p = 10
   
   Do
      '*** Peso de 2 a 9 (multiplicação dos caracteres de 2 a 9) ***
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
   
   '*** se o calculo for igual a 10 ou 11, muda-se para 0 ***
   If (digito_11 = 11) Or (digito_11 = 10) Then
      digito_11 = 0
   End If

   Geral.RetDigMod11 = digito_11
   
   digito_rv = Mid(nova_str, 11, 1)    'digito verificador
   
   If CStr(digito_11) <> (digito_rv) Then
      ok_1 = 0                         'digito não confere
   Else
      ok_1 = 1                         'digito confere
   End If
   
End Sub
Public Function Modulo10(ByVal base_calculo As String, ByVal Tamanho As Integer) As Integer
   
   Dim soma1 As Integer, digito1_10 As Integer, unico As Integer
   Dim troca As Integer
   Dim p As Integer
   Dim dec As Integer
   Dim verif_carac As Integer

   soma1 = 0
   
   troca = 0
   unico = 0
   dec = 0

   
   '***********************************************************
   'número da agencia+conta: (4+6+1)      9 9 9 9  9 9 9 9 9 9
   '                                      x x x x  x x x x x x
   'multiplica da direita para esquerda:  1 2 1 2  1 2 1 2 1 2
   '***********************************************************
   
   p = Tamanho
   
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
   
   ' Digito verificador calculado pelo módulo 10
   digito1_10 = dec - soma1
   
   Modulo10 = digito1_10
   
End Function
Public Function TratamentoErro(ByVal pDescricaoModulo As String, _
                               ByRef pError As ErrObject, _
                               Optional ByRef pErroAcao As EnumErroAcao = eDefault, _
                               Optional ByVal ExibeCMDContinua As Boolean = False) As enumRetornoMensagemErro
    
    Dim spRetorno       As Integer
    
   'se for Trigger Error gravalog, ocorrencia e encerra aplicação
    If TratamentoTerror(pDescricaoModulo, pError) Then
        TratamentoErro = eSair
        Exit Function
    End If
    
    If pErroAcao = eCapa Then
       'Falha na Capa - Enviar Capa para CSP e Continuar
        Geral.PreparouLog = 4
        MDIQuery.insLogErro Geral.DataProcessamento, Caixa.Estacao, Err.Number, "Erro: " & pDescricaoModulo & " - Módulo: " & Err.Description
        TratamentoErro = eFinalizar
    ElseIf pErroAcao = eDoctoProcesso Then
       'Falha na Obtenção do Docto - Enviar capa para CSP sem ocorrencia no Docto e continuar
       'Call MDIQuery.updDoctoPendente(Geral.DataProcessamento, Geral.rstCapa!idcapa)
        Geral.PreparouLog = 4
    
        MDIQuery.insLogErro Geral.DataProcessamento, Caixa.Estacao, Err.Number, pDescricaoModulo & " - " & Err.Description
        TratamentoErro = eFinalizar
    ElseIf pErroAcao = eDoctoSubidaLog Then
        MDIQuery.insLogErro Geral.DataProcessamento, Caixa.Estacao, Err.Number, pDescricaoModulo & " - " & Err.Description
        
        Geral.CodOcorrencia = 999
        Geral.RetTransacao = 89
        
        Geral.PreparouLog = 4
        
        Call DbRejeitaDocto(Geral.rstCapa!idcapa, Geral.rstDoctos!iddocto)
        TratamentoErro = eFinalizar
    Else
        'Se ErroAcao = eDefault continua na Tela de Tratamento
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '* Tratamento de Erros Ocorridos durante Processamento *
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        TratamentoErro = (frmMensagemErro.ShowModal(pDescricaoModulo, pError, ExibeCMDContinua))
        
        If TratamentoErro = eSair Then
            Unload frmShow
            Call CopiaLogs
        End If
        
    End If
    

End Function
Public Sub FormataText(pObjeto As Object, pSelIni As Long, pSellen As Long, Optional pNomeFonte As String, _
                Optional pAlinhamento As AlignmentConstants, Optional pFontSize As Long, Optional pFontColor As EnumCores, _
                Optional pNegrito As Boolean, Optional PGrifado As Boolean)
                

    pObjeto.SelStart = pSelIni
    pObjeto.SelLength = pSellen
    pObjeto.SetFocus
    
    pObjeto.SelFontName = IIf(pNomeFonte = "", "Courrier New", pNomeFonte)
    pObjeto.SelAlignment = IIf(pAlinhamento = 0, vbLeftJustify, pAlinhamento)
    pObjeto.SelFontSize = IIf(pFontSize = 0, 10, pFontSize)
    pObjeto.SelColor = IIf(pFontColor = nulo, vbBlack, pFontColor)
    pObjeto.SelBold = CBool(pNegrito)
    pObjeto.SelUnderline = CBool(PGrifado)
    
    pObjeto.SelStart = 0
    pObjeto.SelLength = 0
        
End Sub
Sub ValidaDLL(InOut As Boolean)
    Dim lPath As String
    
    lPath = PegarOpcaoINI("Path", "Biblio", "")
    
    If Dir(lPath & "\QxHostNT.DLL") = "" Then
        MsgBox "Library [QxHostNT.DLL] não Localizado, verifique !", vbCritical + vbOKOnly, App.Title
        End
    End If
    
    If Dir(lPath & "\ADOUBB.DLL") = "" Then
        MsgBox "Library [ADOUBB.DLL] não Localizado, verifique !", vbCritical + vbOKOnly, App.Title
        End
    End If
    
    If Dir(lPath & "\QxConnect_dll.dll") = "" Then
        MsgBox "Library [QxConnect_dll.dll] não Localizado, verifique !", vbCritical + vbOKOnly, App.Title
        End
    End If
    
    If Dir(lPath & "\Encripta.dll") = "" Then
        MsgBox "Library [ENCRIPTA.dll] não Localizado, verifique !", vbCritical + vbOKOnly, App.Title
        End
    End If
    
    If Dir(lPath & "\qxnsusdv32.dll") = "" Then
        MsgBox "Library [QXNSUSDV32.dll] não Localizado, verifique !", vbCritical + vbOKOnly, App.Title
        End
    End If

'    If InOut Then
'        HwndLibQxHost = LoadLibrary(lPath & "\QxHostNT.dll")
'        HwndLibAdoUBB = LoadLibrary(lPath & "\ADOUBB.DLL")
'        HwndLibConn = LoadLibrary(lPath & "\QxConnect_dll.dll")
'    Else
'        FreeLibrary (HwndLibQxHost)
'        FreeLibrary (HwndLibAdoUBB)
'        FreeLibrary (HwndLibConn)
'    End If
    
   'Se ambiente não for desenvolvimento validar ActiveX
    If UCase(Command) <> "DEBUG" Then
        lPath = PegarOpcaoINI("Path", "ActiveX", "")
        
        If Dir(lPath & "\dbMDI.DLL") = "" Then
            MsgBox "Library [ " & lPath & "\dbMDI.DLL ] não Localizado, verifique !", vbCritical + vbOKOnly, App.Title
            End
        End If
        
        If Dir(lPath & "\dbUBB.DLL") = "" Then
            MsgBox "Library [ " & lPath & "\dbMUBB.DLL ] não Localizado, verifique !", vbCritical + vbOKOnly, App.Title
            End
        End If
    
'        If InOut Then
'            HwndLibDBMDI = LoadLibrary(lPath & "\dbMDI.dll")
'            HwndLibDBUBB = LoadLibrary(lPath & "\dbUBB.dll")
'        Else
'            FreeLibrary (HwndLibDBMDI)
'            FreeLibrary (HwndLibDBUBB)
'        End If
        
    End If
End Sub
Sub CopiaLogs()

    Dim DirSvr     As String
    Dim Arquivo    As String

    On Error Resume Next
        
    DoEvents
    Close #20
    
    If CBool(PegarOpcaoINI("Diversos", "CopiaLog", "0")) Then
    
        frmPrincipal.LblMSG.Caption = "Aguarde.. Copiando Log's para o Servidor."
        frmPrincipal.LblMSG.Visible = True
        frmPrincipal.Refresh
        
        DirSvr = PegarOpcaoINI("Path", "LogSvr", "0")
        Arquivo = "EST" & Trim(Caixa.Estacao) & Parametros.AgenciaCentral & "_" & Format(Now, "ddmm") & ".TXT"
    
        If DirSvr <> "0" And Dir(DirSvr, vbDirectory) <> "" Then
        
            Kill DirSvr & "\" & Arquivo
            FileCopy App.Path & "\Logs\" & Arquivo, DirSvr & "\" & Arquivo
            
        End If
        
    End If
    
End Sub
Function ValidaCMC7Deposito(ByVal pCMC7 As String) As Boolean

    If Trim(pCMC7) = "" Then
        Exit Function
    End If

    If Mid(pCMC7, 1, 3) <> "409" Then
        Exit Function
    End If
    
   'Verificar se as tres primeiras posições do segundo campo do CMC7 devem ser iguais à : 999
    If Mid(pCMC7, 9, 3) <> "999" Then
        Exit Function
    End If
    
    ValidaCMC7Deposito = True
End Function
Function ValidaCMC7Cheque(ByVal pCMC7 As String) As Boolean

    If Trim(pCMC7) = "" Then
        Exit Function
    End If
    
    If Len(Trim(pCMC7)) <> 30 Then
        Exit Function
    End If
        
    ValidaCMC7Cheque = True
End Function
Function Abrelinha(Optional pTransacao As String) As Integer

On Error GoTo TrataErro

    Dim FuncaoAux       As String * 14
    Dim MsgIdaAux       As String
    Dim MsgRetornoAux   As String

   'Aloca Logical Unit no Host
    LocalLog "***** Aloca Linha - Hora Abertura: " & Format(Now, "hh:mm:ss")
    FuncaoAux = "000000000****"
    MsgIdaAux = ""
    MsgRetornoAux = ""

    Abrelinha = qxhostnt(FuncaoAux, MsgIdaAux, MsgRetornoAux)
    LocalLog "Hora Final abertura linha: " & Format(Now, "hh:mm:ss") & " com Retorno: " & Format(Abrelinha, "00") & " - Enviando Transacao: " & pTransacao
    
    If Abrelinha <> 0 Then
        Err.Raise 965, App.Title, "Falha na Abertura da Linha de comunicacao com BH [LU]"
    End If
    
    Exit Function
    
TrataErro:
    Err.Raise 965, App.Title, "Falha na Abertura da Linha de comunicacao com BH [LU]"

End Function
Function FechaLinha(Optional ByVal pTransacao As String) As Integer

On Error GoTo TrataErro

    Dim FuncaoAux       As String * 14
    Dim MsgIdaAux       As String
    Dim MsgRetornoAux   As String

   'DESALOCA LOGICAL UNIT
    FuncaoAux = "200000000****"
    MsgIdaAux = ""
    MsgRetornoAux = ""
     
    LocalLog "Finalizada Transacao: " & pTransacao & " - Tempo: " & Format(Now, "hh:mm:ss")
    FechaLinha = qxhostnt(FuncaoAux, MsgIdaAux, MsgRetornoAux)
    LocalLog "***** Desaloca Linha " & Format(Now, "hh:mm:ss") & " com Retorno: " & Format(FechaLinha, "00")
    
    If FechaLinha <> 0 Then
        Err.Raise 964, App.Title, "Falha no Fechamento da Linha de comunicacao com BH [LU]"
    End If
    
    Exit Function
    
TrataErro:
    Err.Raise 964, App.Title, "Falha no Fechamento da Linha de comunicacao com BH [LU]"

End Function
Function GetDocumentoTransmitido(pOrigem As EnumOrigem) As Boolean

On Error GoTo TrataErro

    Dim RstUBB          As Recordset
    Dim spRetorno       As Integer
    
    Dim MsgIda          As String
    Dim MsgRetorno      As String
    Dim Funcao          As String * 14
    Dim RetQX           As Integer
    Dim HeaderTx        As String
    Dim TamIda          As String
        
   '******************************************************************
   ' Verifica se Houve Falha no Processamento Apos docto Transmitido '
   '******************************************************************
   
    If Not IsNull(Geral.rstDoctos!NSU) Then
       'verifica se nsu existe na tabela de caixa
        Set RstUBB = UBBQuery.getNsuTabelacaixa(Caixa.Caixa)
    
       'se encontrou nsu na tabela de caixa
        If RstUBB("tcxfsnuseqt") = Val(Geral.rstDoctos!NSU) Then
           'se docto for ADCC verificar conta
            If pOrigem = EnumADCC And _
               Not IsNull(Geral.rstDoctos!RetornoTransacao And _
               Geral.rstDoctos!RetornoTransacao = 75) Then
               
               'Conta Unibanco reinformada Corretamente
                MDIQuery.updCancelarRetornoTransacao Geral.DataProcessamento, _
                                                     Geral.rstCapa!idcapa, _
                                                     Geral.rstDoctos!iddocto, _
                                                     Caixa.Caixa
                                                
            ElseIf pOrigem = EnumDeposito Then
            
                Geral.CodOcorrencia = 999
                Geral.RetTransacao = 96

                Call DevolveDocumentos
                
                GetDocumentoTransmitido = True
                Exit Function
                
            End If

           'seta documento como transmtido
            spRetorno = MDIQuery.updDoctoTransmitido(Geral.DataProcessamento, _
                                                     Geral.rstCapa!idcapa, _
                                                     Geral.rstDoctos!iddocto, _
                                                     Geral.rstDoctos!NSU, _
                                                     Caixa.Caixa, "N")
    
            If spRetorno <> 0 Then
                MsgBox "3987. ATENÇÃO! Documento já enviado Log, não foi atualizado no SQL. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
            End If
    
            Geral.GereiLog = 1
            GetDocumentoTransmitido = True
        End If
    
    End If

Exit Function

TrataErro:

    Select Case TratamentoErro("Falha na Verificação de Docto ja Transmitido.", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
    End Select
    
End Function
Public Function TratamentoTerror(ByVal pDescricaoModulo As String, ByRef pError As ErrObject) As Boolean
    
    Dim ehTriggerError  As Integer
    Dim spRetorno       As Integer
    
    ehTriggerError = InStr(1, pError.Description, "TERROR=", vbTextCompare)
    
    If ehTriggerError <> 0 Then

        MDIQuery.insLogErro Geral.DataProcessamento, Caixa.Estacao, pError.Number, "Erro: " & pError.Description & "Módulo: " & pDescricaoModulo
        
        If InStr(1, "234_261_369_370_371", Mid(pError.Description, ehTriggerError + 7, 3), vbTextCompare) <> 0 Then
            LocalLog "CodigoTriggerError Retornado: " & Mid(pError.Description, ehTriggerError + 7, 3) & " - Retorno: " & pError.Description
        Else
            LocalLog "CodigoTriggerError Retornado(Desconhecido): " & Mid(pError.Description, ehTriggerError + 7, 3) & " - Retorno: " & pError.Description
        End If
             
        Geral.CodOcorrencia = "999"
        Geral.RetTransacao = "90"
        
        MsgBox "Documento Rejeitado, Status já Processado. O Robo deverá ser Reinicializado", vbCritical + vbOKOnly, "Mensagem do Sistema de Caixa"
        
        Call DevolveDocumentos
                        
        If Geral.PreparouLog = 4 Then
            LocalLog "Envia Capa para CSP " & Geral.rstCapa!Capa

            MDIQuery.insLog Geral.DataProcessamento, Geral.rstCapa!idcapa, "0", Caixa.UsuarioAtual, "126"

            spRetorno = MDIQuery.UpdCapaStatusCaixaControle(Geral.DataProcessamento, Geral.rstCapa!idcapa, ST_CapaParaCSP, Caixa.Caixa)
            If spRetorno <> 0 Then MsgBox "Falha Procedure [UpdCapaStatusCaixaControle]", vbCritical + vbOKOnly

            Call GaugeTitulo(4)
            Espera (0.3)

            '''''''''''''''''''''''''''''''''''
            'Grava a procedure de fim de ciclo'
            '''''''''''''''''''''''''''''''''''
            LogFimCiclo

        ElseIf Geral.PreparouLog <> 3 Then

            LocalLog "Atualiza Capa Transmitida " & Geral.rstCapa!Capa
            spRetorno = MDIQuery.UpdCapaStatusCaixaControle(Geral.DataProcessamento, Geral.rstCapa!idcapa, ST_CapaTransmitida, Caixa.Caixa)
    
            If spRetorno <> 0 Then
                MsgBox "1. Problema ao atualizar registro da capa que terminou de gerar LOG!", 16, "Atenção"
            End If
             
            '''''''''''''''''''''''''''''''''''
            'Grava a procedure de fim de ciclo'
            '''''''''''''''''''''''''''''''''''
             
            LogFimCiclo
             
        End If
        
        If Geral.CaixaAberto Then
             LogFechamentoCaixa ("A")
             Geral.CaixaAberto = False
        End If
          
        If Geral.FecharCaixa Then
            Call GetCaixa
            Geral.FecharCaixa = False
        End If
        
        Unload frmShow
        Call CopiaLogs
        
        TratamentoTerror = True
        
    End If

End Function
Sub CalculaNSU(Optional ByVal outProcess As Boolean)
    Dim Ret         As Integer
    Dim NSU         As String * 6
    Dim SDV         As String * 1
    Dim Data        As String
    Dim nroCaixa    As String
    Dim Agencia     As String
    Dim Valor       As String
    Dim Operador    As String
        
    Data = Right(Geral.DataProcessamento, 2) & Mid(Geral.DataProcessamento, 5, 2) & Mid(Geral.DataProcessamento, 3, 2)
    Caixa.AddNSU
    NSU = Caixa.BaseNSU
    nroCaixa = Caixa.Caixa
    Agencia = IIf(Parametros.AgenciaSatelite = Empty, "0", Parametros.AgenciaSatelite)
    Valor = IIf(Geral.ValorTrans = Empty Or outProcess, 0, Geral.ValorTrans)
    Operador = Caixa.CIF
        
    If CDbl(NSU) >= 9000 Then
        Geral.FecharCaixa = True
    End If
    
    NSU = Format(NSU, "0000") & "00"
    Ret = QXCalNsu(Agencia, nroCaixa, Data, "9", Valor, NSU)
    LocalLog "Agencia: " & Agencia & " / nroCaixa: " & nroCaixa & " / Data: " & Data & "/ Tipo: 9 / Valor: " & Valor & " / NSU: " & NSU
    Caixa.NSU = NSU
    
    Ret = QXGetSDV(Agencia, nroCaixa, Data, "9", Valor, Operador, SDV)
    Caixa.SDV = SDV
    
End Sub


