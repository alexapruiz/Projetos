Attribute VB_Name = "Global"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Constantes para Mascaras e etc
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const MASK_VALOR = "###,###,###,##0.00"
Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const BIF_DONTGOBELOWDOMAIN = &H2
Public Const BIF_STATUSTEXT = &H4
Public Const BIF_RETURNFSANCESTORS = &H8
Public Const BIF_BROWSEFORCOMPUTER = &H1000
Public Const BIF_BROWSEFORPRINTER = &H2000

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Constantes para Conexão
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const PROVIDER = "SQLOLEDB"
Public Const SERVER = "MDI_NT2"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Declaração de Types
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public g_Parametros             As tpParametro
Public Type tpParametro
    QuantidadeCheques           As Integer
    QuantidadeDatas             As Integer
    QuantidadeMinimaDias        As String '* 3
    DiretorioTransmissao        As String '* 50
    DiretorioRecepcao           As String '* 50
    Sequencia_Bordero           As Integer         'Conterá o próximo número de Borderô
    Gerar_Arquivo_CEL           As Boolean
    Numero_Lote_CEL             As Long
    Comp_Origem_CEL             As Integer
    Numero_Versao_Inicial_CEL   As Integer
    Numero_Versao_Final_CEL     As Integer
    HeaderAV                    As Boolean
    chkSoma                     As Boolean
    Codigo_USB                  As Integer
    CPD_Origem                  As String '* 3
    CPD_Destino                 As String '* 3
    Codigo_Terceira             As String '* 4
    CNPJ_Terceira               As String '* 14
    Cidade_Terceira             As String
    Nome_Terceira               As String
    UF_Terceira                 As String '* 2
    Seq_Ocorrencia              As Long
    CodigoAgAcolhed             As Integer
    Num_Remessa_TER             As String '* 3
    ValorChequeLimite           As Currency
    CodigoAplicacao             As String '* 3
    TMP_Pendente                As Long
    Scanner                     As Integer
    PortaCom                    As Integer
    EmUso                       As Integer
End Type

Public Type tpGrupoUsuario
     Supervisor               As Integer       'Identificador de Grupo Supervisor
     Digitadores              As Integer       'Identificador de Grupo Digitadores
End Type

Public Geral                  As tpGeral
Public Type tpGeral
     UsuarioLogin             As String * 10
     UsuarioNome              As String * 50
     UsuarioId                As Integer
     SenhaDesenv              As String
     DataProcessamento        As Long
     GrupoUsuario             As tpGrupoUsuario
End Type

Type SHITEMID
     cb                         As Long
     abID                       As Byte
End Type

Public Type ITEMIDLIST
     mkid                       As SHITEMID
End Type

Public Type BROWSEINFO
     hOwner                     As Long
     pidlRoot                   As Long
     pszDisplayName             As String
     lpszTitle                  As String
     ulFlags                    As Long
     lpfn                       As Long
     lParam                     As Long
     iImage                     As Long
End Type

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'         Define type para paineis do Status Bar       '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public StatusBar As PainelStatusBar
Public Type PainelStatusBar
     Col_Descrição            As Integer
     Col_ProgressBar          As Integer
     Col_ContadorProgressBar  As Integer
     Col_Usuario              As Integer
     Col_Data                 As Integer
End Type

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Declaração de Enumeração
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Enum enumRetornoUsuario
    eSUPERVISOR
    eNAO_EXISTENTE
    eSENHA_INCORRETA
    eOK
End Enum

Public Enum enumRetornoModal
    eRetornoCancelar
    eRetornoOK
End Enum

Public Enum enumFormatoData
    DD_MM_AAAA
    MM_DD_AAAA
    AAAA_MM_DD
    DD_MM
    DDMM
End Enum
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Enumeracao publica do tipo de exportação de arquivo CEL'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Enum enumTipoArquivoCEL
    eCheque_Limite
    eCheque_Superior
    eCheque_Unibanco
    eArquivo_TER
End Enum

' Scroll Bar Commands
Public Const SB_LINEUP = 0
Public Const SB_LINELEFT = 0
Public Const SB_LINEDOWN = 1
Public Const SB_LINERIGHT = 1
Public Const SB_PAGEUP = 2
Public Const SB_PAGELEFT = 2
Public Const SB_PAGEDOWN = 3
Public Const SB_PAGERIGHT = 3
Public Const SB_THUMBPOSITION = 4
Public Const SB_THUMBTRACK = 5
Public Const SB_TOP = 6
Public Const SB_LEFT = 6
Public Const SB_BOTTOM = 7
Public Const SB_RIGHT = 7
Public Const SB_ENDSCROLL = 8

Public Const WM_HSCROLL = &H114
Public Const WM_VSCROLL = &H115

' Scroll Bar Constants
Public Const SB_HORZ = 0
Public Const SB_VERT = 1
Public Const SB_CTL = 2
Public Const SB_BOTH = 3


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Declaração de Variáveis
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public g_cMainConnection        As New ADODB.Connection

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Declaração de API'S
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long 'ITEMIDLIST
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Global Erro As String
Public Function Compactar_BancoDados(ByVal BD As String) As Boolean

On Error GoTo Compactar_BancoDados_Err:

Dim Diretorio, DirData  As String
Dim Erro                As Integer
Dim Repete              As Boolean

Repete = True
DirData = Trim(Str(Year(Now))) + Trim(Format(Month(Now), "00")) + Trim(Format(Day(Now), "00"))

Diretorio = PegarOpcaoINI("Backup", "Database", "") + "\" + DirData

If Dir(Diretorio + "\Custodia.mdb") <> "" Then
   Kill Diretorio + "\Custodia.mdb"
End If
   
If Dir(Diretorio, vbDirectory) <> "" Then
   RmDir Diretorio
End If

MkDir Diretorio

Repetir:

DBEngine.RepairDatabase BD
DBEngine.CompactDatabase BD, Diretorio + "\Custodia.mdb", , dbEncrypt + dbVersion20

Compactar_BancoDados = True

Exit Function

Compactar_BancoDados_Err:
    Select Case Err.Number
        Case 3356
           'Call TratamentoErro("Erro Durante a Compactação do Banco de Dados. Podem Haver Usuários Conectados", Err, Repete)
            Call TratamentoErro("Há Outros Usuários Conectados. Favor Providenciar Para Que Todos Saiam.", Err, Repete)
            If Repete Then
                Resume Repetir:
            Else
                Compactar_BancoDados = False
            End If
        Case Else
            Call TratamentoErro("Falha no Modulo de Compactação.", Err, Repete)
            If Repete Then
                Resume Repetir:
            Else
                Compactar_BancoDados = False
            End If
        
        End Select
           
End Function
Public Function LimpaTabelas(ByVal vDataProcessamento As Long, vDataQuinzena As Long) As Boolean

Dim LimpaTabela   As New Custodia.Excluir
Dim rsLimpaTabela As ADODB.Recordset
Dim Banco         As String
Dim Erro          As Integer

Dim Repete        As Boolean
Dim Transacao     As Boolean

On Error GoTo LimpaTabelas_Err:

Repetir:

Repete = True

    g_cMainConnection.Close
    
    Transacao = False
    
    Banco = PegarOpcaoINI("CONEXAO", "DATABASE", "")
    
    If Compactar_BancoDados(Banco) Then
    
        If Not SetConnection(g_cMainConnection, True) Then
           Erro = MsgBox("Não Foi Possível Re-estabelecer a Conexão com o Banco de Dados. O Sistema Será Finalizado", vbCritical, "Sistema Custódia")
           End
        End If
            
        g_cMainConnection.BeginTrans
        
        Transacao = True
        
        Set rsLimpaTabela = g_cMainConnection.Execute(LimpaTabela.LimpaDataDeposito(vDataProcessamento))
        Set rsLimpaTabela = g_cMainConnection.Execute(LimpaTabela.LimpaBordero(vDataProcessamento))
        Set rsLimpaTabela = g_cMainConnection.Execute(LimpaTabela.LimpaCheque(vDataProcessamento))
        Set rsLimpaTabela = g_cMainConnection.Execute(LimpaTabela.LimpaRejeicaoRemessa(vDataProcessamento))
        Set rsLimpaTabela = g_cMainConnection.Execute(LimpaTabela.LimpaChequeDataBoa(vDataProcessamento))
        Set rsLimpaTabela = g_cMainConnection.Execute(LimpaTabela.LimpaParametro(vDataProcessamento))
        If vDataQuinzena <> 0 Then
             Set rsLimpaTabela = g_cMainConnection.Execute(LimpaTabela.LimpaAlteracaoData(vDataQuinzena))
             Set rsLimpaTabela = g_cMainConnection.Execute(LimpaTabela.LimpaChequesBaixados(vDataQuinzena))
        End If
        
        g_cMainConnection.CommitTrans
        
        LimpaTabelas = True
        
    Else
    
       If Not SetConnection(g_cMainConnection, True) Then
           Erro = MsgBox("Não Foi Possível Re-estabelecer a Conexão com o Banco de Dados. O Sistema Será Finalizado", vbCritical, "Sistema Custódia")
           End
       End If
        
       LimpaTabelas = False
        
    End If
    
    Exit Function

LimpaTabelas_Err:
    
    If Transacao Then
       g_cMainConnection.RollbackTrans
    End If
    
    Call TratamentoErro("Houve Falha ao Efetuar: BACKUP/LIMPEZA Da Base", Err, Repete)
    If Repete Then Resume Repetir
    
    LimpaTabelas = False

End Function
Public Function VerificaLimpezaTabelas() As Boolean

    Dim rsPegaDataProcessamento As ADODB.Recordset
    Dim rsPegaCapaLimpeza       As ADODB.Recordset
    Dim vNroLinhas              As Integer
    Dim vResp                   As Integer
    Dim vDataProcessamento      As Long
    Dim vDataQuinzena           As Long
    Dim TesteLimpeza            As New Custodia.Selecionar
    
    On Error GoTo ErroVerificaLimpezaTabelas:
    
   'Verifica Dias Para Limpeza
    vNroLinhas = NroLinhas()
    If vNroLinhas = 0 Then
       MsgBox "Não é possível Inicializar o Sistema sem Parâmetro", vbOKOnly + vbCritical
       VerificaLimpezaTabelas = False
       End
    End If
        
    'Obtem data para limpeza de tabelas
    Set rsPegaDataProcessamento = g_cMainConnection.Execute(TesteLimpeza.GetDataProcessamento(vNroLinhas))
        
    If Not rsPegaDataProcessamento.EOF And rsPegaDataProcessamento.RecordCount >= vNroLinhas Then
        
       rsPegaDataProcessamento.MoveLast
       vDataProcessamento = rsPegaDataProcessamento!DataProcessamento
           
      'Verifica se há Capas com status D e E
       Set rsPegaCapaLimpeza = g_cMainConnection.Execute(TesteLimpeza.GetCapaLimpeza(vDataProcessamento))
       
       If rsPegaCapaLimpeza.RecordCount = 0 Then
         Set rsPegaDataProcessamento = Nothing
         Set rsPegaCapaLimpeza = Nothing
         Exit Function
       End If
       
       vResp = MsgBox("Movimentos Inferiores a " & FormataData(vDataProcessamento, DD_MM_AAAA) & " Possui Borderôs Não-Confirmados Pelo CH. Confirma Limpeza?", vbYesNo + vbInformation, "Sistema Custódia")
          
       If vResp = vbYes Then
          'Obtem data para limpeza da tabela "AlteracaoData" fixo em 15 dias inferiores do ultimo movimento
          If Not (rsPegaDataProcessamento Is Nothing) Then Set rsPegaDataProcessamento = Nothing
          Set rsPegaDataProcessamento = g_cMainConnection.Execute(TesteLimpeza.GetDataProcessamento(15))
          If rsPegaDataProcessamento.EOF Then
               vDataQuinzena = 0
          Else
               rsPegaDataProcessamento.MoveLast
               vDataQuinzena = rsPegaDataProcessamento!DataProcessamento
          End If
         
         'Faz a Limpeza das Tabelas
          If LimpaTabelas(vDataProcessamento, vDataQuinzena) Then
             VerificaLimpezaTabelas = True
          Else
             VerificaLimpezaTabelas = False
          End If
       Else
          VerificaLimpezaTabelas = False
       End If
       
    Else
       VerificaLimpezaTabelas = False
    End If
    
    If Not (rsPegaDataProcessamento Is Nothing) Then Set rsPegaDataProcessamento = Nothing
    
    Exit Function
    
ErroVerificaLimpezaTabelas:
    If Not (rsPegaDataProcessamento Is Nothing) Then Set rsPegaDataProcessamento = Nothing
    VerificaLimpezaTabelas = False
    Call TratamentoErro("Selecionar/Indicar Data Limite para Limpeza da Base", Err)

End Function
Public Function NroLinhas() As Integer

    Dim rsNroLinhas As ADODB.Recordset
    Dim Procedures  As New Custodia.Selecionar
    
    Set rsNroLinhas = g_cMainConnection.Execute(Procedures.GetDiasLimpeza())
    
    If Not rsNroLinhas.EOF Then
       NroLinhas = rsNroLinhas!DiasLimpeza
    End If
 
End Function
Public Function FileExist(ByVal pFileName As String) As Boolean

    Dim lclFileNum As Integer

    On Error Resume Next
    
    lclFileNum = FreeFile
    
    Open pFileName For Input As lclFileNum

    FileExist = IIf(Err = 0, True, False)

    Close lclFileNum

    Err = 0

End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Retorna True se o bordero estiver "Em ..."
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function EstaPara(ByVal pStatusAtual As String, _
                         ByRef pMsg As String, _
                         ParamArray pStatusPara()) As Boolean


    Dim lRst                As New ADODB.Recordset
    Dim Proc_Selecionar     As New Custodia.Selecionar
    Dim i                   As Integer
    
    EstaPara = False
    
    '''''''''''''''''''''''''''''''''''''''
    'Já está com o mesmo status, então sai'
    '''''''''''''''''''''''''''''''''''''''
    For i = 0 To UBound(pStatusPara())
        If pStatusAtual = pStatusPara(i) Then Exit Function
    Next i
    
    '''''''''''''''''''''''''''''''''''''
    'Busca todas as descrições de Status'
    '''''''''''''''''''''''''''''''''''''
    Set lRst = g_cMainConnection.Execute(Proc_Selecionar.GetStatusBordero())
    
    If Not lRst.EOF() Then
        lRst.Find "Status = '" & pStatusAtual & "'"
        If Not lRst.EOF() Then
            EstaPara = True
            pMsg = lRst!Descricao
        End If
    End If

    lRst.Close
   
End Function
Public Function FormataData(ByVal pData As Long, ByVal pFormato As enumFormatoData) As String

    Dim sData       As String
    Dim sFormato    As String
    
    ''''''''''''''''''''''''''''''
    'Formata a data para aaaammdd'
    ''''''''''''''''''''''''''''''
    sData = Format(pData, "0000/00/00")
    
    
    ''''''''''''''''''''''''''''''''''''''
    'Formata a data para o que foi pedido'
    ''''''''''''''''''''''''''''''''''''''
    If pFormato = AAAA_MM_DD Then
        sFormato = "YYYY/MM/DD"
    ElseIf pFormato = DD_MM_AAAA Then
        sFormato = "DD/MM/YYYY"
    ElseIf pFormato = MM_DD_AAAA Then
        sFormato = "MM/DD/YYYY"
    ElseIf pFormato = DD_MM Then
        sFormato = "DD/MM"
    ElseIf pFormato = DDMM Then
        sFormato = "DDMM"
    End If
    
    FormataData = Format(sData, sFormato)

End Function

Function ResultTimer(ByVal phWnd As Long, _
                     ByVal uMsg As Long, _
                     ByVal wParam As Long, _
                     ByVal lParam As Long) As Long
    
'    ResultTimer = CallWindowProc(lpWndTimer, phWnd, uMsg, wParam, lParam)
End Function
Public Function TratamentoErro(ByVal p1vsTexto As String, _
                               ByVal pvoErro As ErrObject, _
                               Optional ByRef ExibeCMDRepete As Boolean, _
                               Optional ByRef ExibeCMDHelpScan As Boolean) As VbMsgBoxResult

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               * Tratamento de Erros Ocorridos durante Processamento *                     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Screen.MousePointer = Default
    Call MensagemErro.ShowModal(p1vsTexto, pvoErro.Number, pvoErro.Description, ExibeCMDRepete, ExibeCMDHelpScan)
End Function
Public Function RetiraPonto(ByVal Valor As String) As String
    Dim Result As String
    Dim Count As Integer
    
    Valor = Trim(Valor)
    If InStr(1, Valor, ".", 1) = 0 And InStr(1, Valor, ",", 1) = 0 Then
        Result = Valor & "00"
    Else
        For Count = 1 To Len(Valor)
            If Mid(Valor, Count, 1) <> "." And Mid(Valor, Count, 1) <> "," Then
                Result = Result + Mid(Valor, Count, 1)
            End If
        Next
    End If
    RetiraPonto = Result
End Function

Public Function PegarOpcaoINI(ByVal pvsSecao As String, pvsItem As String, ByVal pvsDefault As String) As String
    Dim iRet As Long
    Dim sDado As String
    Dim sDadoAux As String
    Dim i As Integer
    
    sDado = String(255, " ")
    iRet = GetPrivateProfileString(pvsSecao, pvsItem, pvsDefault, sDado, 255, App.path & "\Custodia.ini")
    
    sDado = Trim(sDado)
    sDadoAux = ""
    
    For i = 1 To Len(sDado)
        If Asc(Mid(sDado, i, 1)) >= 32 And Asc(Mid(sDado, i, 1)) <= 122 Then
            sDadoAux = sDadoAux & Mid(sDado, i, 1)
        End If
    Next
    
    PegarOpcaoINI = Trim(sDadoAux)
End Function
Public Sub GravarOpcaoINI(ByVal pvsSecao As String, pvsItem As String, ByVal pvsValor As String)
    Dim iRet As Long
    
    iRet = WritePrivateProfileString(pvsSecao, pvsItem, pvsValor, App.path & "\Custodia.INI")
End Sub

Public Function FormataString( _
    ByVal pOque As Variant, _
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

Public Function InserePonto(ByVal Valor As String) As String
    Valor = Format(Valor, "000")
    If Val(Valor) = 0 Then
        InserePonto = "0.00"
    Else
        InserePonto = Left(Valor, Len(Valor) - 2) & "." & Right(Valor, 2)
    End If
End Function
Public Sub SetStatusBar(ByRef pForm As Form, ByVal pMensagem As String)

    pForm.StatusBar.Panels(0).Text = pMensagem

End Sub

Public Sub SoNumero(ByRef KeyAscii As Integer)

    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then KeyAscii = 0

End Sub
Public Sub SelecionarTexto(ByVal pObjeto As Object)

    On Error Resume Next
    
    pObjeto.SelStart = 0
    pObjeto.SelLength = Len(pObjeto)
    
End Sub
Public Sub Main()

    Dim RstAux      As New ADODB.Recordset
    Dim Proc        As New Custodia.Selecionar
    Dim sstr        As String
    Dim Diretorio   As String
    
    On Error GoTo Erro_Main:
         
    ''''''''''''''''''''''''''''''''''''''''
    'Abre nova conexão com o Banco de Dados'
    ''''''''''''''''''''''''''''''''''''''''
    If Not SetConnection(g_cMainConnection, True) Then GoTo Erro_Main
    
    Load Principal
    
    '''''''''''''''''''''''''''''''''''''''''''''''
    '     Definição da senha do usuário DESENV    '
    '''''''''''''''''''''''''''''''''''''''''''''''
    Geral.SenhaDesenv = "VENUS"                   'Senha do Usuário DESENV (Sempre UPPER CASE)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    '     Define código de ID do Grupo de Usuários     '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
     Geral.GrupoUsuario.Supervisor = 1            'Identificador de Grupo Supervisor
     Geral.GrupoUsuario.Digitadores = 2           'Identificador de Grupo Digitadores

    '''''''''''''''''''''''''''''''''''''''''''''''
    '     Chave do algorito de criptografia       '
    '''''''''''''''''''''''''''''''''''''''''''''''
    Key(0) = 97
    Key(1) = 150
    Key(2) = 127
    Key(3) = 254
    
    Geral.DataProcessamento = Val(Format(Now, "yyyymmdd"))
    
    While Not Login.SenhaOk
        Login.Show vbModal, Principal
        If Login.Cancelou Then
            Unload Principal
            End
        End If
    Wend
    
    Unload Login
     
     'Define numeração de paineis da Status Bar
     With StatusBar
          .Col_Descrição = 1
          .Col_ProgressBar = 2
          .Col_ContadorProgressBar = 3
          .Col_Usuario = 4
          .Col_Data = 5
     End With
    
    
    '''''''''''''''''''''''''''''
    'Obtem parametros do sistema'
    '''''''''''''''''''''''''''''
    Set RstAux = g_cMainConnection.Execute(Proc.GetParametros(Geral.DataProcessamento))
    If RstAux.EOF() Then
        RstAux.Close
        MsgBox "Não é possível iniciar o sistema sem os parâmetros.", vbExclamation
        Call SetConnection(g_cMainConnection, False)
        End
    End If
        
    With g_Parametros
        .QuantidadeCheques = RstAux!QuantidadeCheques & ""
        .QuantidadeDatas = RstAux!QuantidadeDatas & ""
        .DiretorioTransmissao = Trim(RstAux!DiretorioTransmissao) & ""
        .DiretorioRecepcao = RstAux!DiretorioRecepcao & ""
        .Sequencia_Bordero = Val(RstAux!Seq_Bordero & "")
        .Gerar_Arquivo_CEL = CBool(RstAux!GerarArquivo_CEL)
        .Comp_Origem_CEL = Val(RstAux!Comp_Origem_CEL & "")
        .Numero_Lote_CEL = Val(RstAux!Num_Lote_CEL & "")
        .Numero_Versao_Inicial_CEL = Val(RstAux!Num_Versao_Inicial_CEL & "")
        .Numero_Versao_Final_CEL = Val(RstAux!Num_Versao_Final_CEL & "")
        .HeaderAV = CBool(RstAux!HeaderAV)
        .chkSoma = CBool(RstAux!CriticaSoma)
        .UF_Terceira = RstAux!UF_Terceira & ""
        .CodigoAplicacao = RstAux!CodigoAplicacao & ""
        .Codigo_USB = Val(RstAux!Codigo_USB & "")
        .CPD_Origem = RstAux!CPD_Origem & ""
        .CPD_Destino = RstAux!CPD_Destino & ""
        .Codigo_Terceira = RstAux!Codigo_Terceira & ""
        .CNPJ_Terceira = RstAux!CNPJ_Terceira & ""
        .Seq_Ocorrencia = IIf(IsNull(RstAux!Seq_Ocorrencia), 0, RstAux!Seq_Ocorrencia) & ""
        .CodigoAgAcolhed = IIf(IsNull(RstAux!CodigoAgAcolhed), 0, RstAux!CodigoAgAcolhed) & ""
        .Num_Remessa_TER = RstAux!Num_Remessa_TER & ""
        .ValorChequeLimite = IIf(IsNull(RstAux!ValorChequeLimite), 0, RstAux!ValorChequeLimite)
        .TMP_Pendente = IIf(IsNull(RstAux!TMP_Pendente), 0, RstAux!TMP_Pendente) & ""
        .QuantidadeMinimaDias = IIf(IsNull(RstAux!QuantidadeMinimaDias), 0, RstAux!QuantidadeMinimaDias)
        .Cidade_Terceira = RstAux!Cidade_Terceira & ""
        .Nome_Terceira = RstAux!Nome_Terceira & ""
        
        
       'Verifica .ini se estação usará scanner
        sstr = Trim(PegarOpcaoINI("Scanner", "Tipo", ""))
        
        If sstr = "1" Or sstr = "2" Then
            .Scanner = CInt(sstr)
            
           'Verifica porta(com) especificada
            sstr = Trim(PegarOpcaoINI("Scanner", "PortaCOM", ""))
            If sstr = "1" Or sstr = "2" Then
               .PortaCom = CInt(sstr)
               
               'Procura DLL de scanner espeficado no arquivo .ini
                If Not AchaScannerDLL(.Scanner) Then
                    Err.Raise 989, App.Title, "Não Localizado Arquivo (DLL) de Utilização do Scanner." & vbCrLf _
                    & "Contate o Administrador do Sistema."
                    .Scanner = 0
                    .PortaCom = 0
                Else
                
                   'Verifica se instalado driver LA93
                    If .Scanner = 2 And VerRegSerialLA93 <> Trim(PegarOpcaoINI("Scanner", "Serial", "")) Then
                        Err.Raise 979, App.Title, "Driver( VipsDrv) do Scanner LA93 não Instalado." & vbCrLf _
                        & "Contate o Administrador do Sistema."
                        .Scanner = 0
                        .PortaCom = 0
                    End If
                End If
            Else
               MsgBox "ATENÇÃO !!! - Parâmetro ( PortaCom ) Inválida, Verifique...", vbInformation + vbOKOnly, App.Title & " - Custodia.ini"
               .Scanner = 0
               .PortaCom = 0
            End If
        End If
        
    End With
    RstAux.Close

    Principal.Show

    Exit Sub

Erro_Main:
    Screen.MousePointer = vbDefault
    If Err.Number = 989 Or Err.Number = 979 Then
        Call TratamentoErro("Falha na Recuperação de parâmetros do Scanner." & vbCrLf & vbTab & "Verifique o problema em [Descrição], para continuar tecle [Sair]", Err, , True)
        Resume Next
    Else
        Call TratamentoErro("Erro na inicialização do sistema.", Err)
    End If
    
End Sub
Public Function SetConnection(ByRef pConnection As ADODB.Connection, ByVal pConectar As Boolean) As Boolean

    Dim sstr            As String

    On Error GoTo Erro_SetConnection:

    SetConnection = False

    If pConectar Then

        '''''''''''''''''''''''''''''''
        'Cria nova conexão com o Banco'
        '''''''''''''''''''''''''''''''
        pConnection.ConnectionTimeout = 10
        pConnection.CursorLocation = adUseClient
        'pConnection.PROVIDER = "Microsoft.JET.OLEDB.3.51"
        pConnection.PROVIDER = "Microsoft.JET.OLEDB.4.0"
        'pConnection.PROVIDER = "SQLOLEDB"

        sstr = PegarOpcaoINI("CONEXAO", "DATABASE", "")
        
        If Trim(sstr) = "" Then Exit Function
        
        'pConnection.Open "Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & sStr
        pConnection.Open sstr, "admin", ""
        
    Else
        '''''''''''''''''''''''''''''
        'Fecha a conexão com o Banco'
        '''''''''''''''''''''''''''''
        If Not pConnection Is Nothing Then pConnection.Close
        
    End If
    
    SetConnection = True
    
Erro_SetConnection:

End Function
Public Function VerificaCGC(ByVal CGC As String) As Boolean
   
   '------------------------------------------------
   '--------- MODULO 11 (2 BASE 9) -----------------
   ' Esta rotina serve para conferir o CGC: tam = 15
   '------------------------------------------------
   
   Dim soma As Integer, resto As Integer
   Dim digito_11 As Integer, p As Integer, peso As Integer
   Dim digito_rv As String
   Dim bOk As Boolean
   
   bOk = True           'default - OK
   
   soma = 0
   resto = 0
   digito_11 = 0        'calculado pelo módulo 11
   digito_rv = ""       'caracter digitado pelo operador

   '*************************************************************
   'número do CGC: (13+2)                0 9 9.9 9 9.9 9 9/9 9 9 9 - D D
   '                                     x x x x x x x x x x x x x
   'multiplica da direita para esquerda: 6 5 4 3 2 9 8 7 6 5 4 3 2
   '*************************************************************
   
   peso = 2    'começa multiplicar da direita para esquerda
   p = 13      'tamanho do campo se o digito
   
   Do
      '*** Peso de 2 a 9 (multiplicação dos caracteres de 2 a 9) ***
      soma = soma + Mid(CGC, p, 1) * peso
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
   
   '*** se o calculo for igual a 10 ou 11, muda-se para 0 ***
   If (digito_11 = 11) Or (digito_11 = 10) Then
      digito_11 = 0
   End If

   digito_rv = Mid(CGC, 14, 1)  '1º digito verificador
   
   If CStr(digito_11) <> (digito_rv) Then
      bOk = False               'digito não confere
      VerificaCGC = bOk
      Exit Function
   End If

   soma = 0
   resto = 0
   digito_11 = 0        'calculado pelo módulo 11
   digito_rv = ""       'caracter digitado pelo operador
   peso = 2             'começa multiplicar da direita para esquerda
   p = 14               'tamanho do campo (13) + 1º digito = 14
   
   Do
      '*** Peso de 2 a 9 (multiplicação dos caracteres de 2 a 9) ***
      soma = soma + Mid(CGC, p, 1) * peso
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
   
   '*** se o calculo for igual a 10 ou 11, muda-se para 0 ***
   If (digito_11 = 11) Or (digito_11 = 10) Then
      digito_11 = 0
   End If

   digito_rv = Mid(CGC, 15, 1)  '2º digito verificador
   
   If CStr(digito_11) <> (digito_rv) Then
      bOk = False               'digito não confere
   End If

   VerificaCGC = bOk

End Function
Public Function VerificaCPF(ByVal CPF As String) As Boolean
   
   '---------------------------------------------
   '--------- MODULO 11 (2 BASE 9) --------------
   ' Esta rotina serve para consistir numero do CPF:
   '---------------------------------------------
   
   Dim soma As Integer, resto As Integer
   Dim digito_11 As Integer, p As Integer, peso As Integer
   Dim digito_rv As String
   Dim bOk As Boolean
   
   bOk = True           'default - OK
   
   soma = 0
   resto = 0
   digito_11 = 0        'calculado pelo módulo 11
   digito_rv = ""       'caracter digitado pelo operador
   
   '*************************************************************
   'número do CPF: (9+2)                  0 0 9 9 9 9 9 9 9 - D D
   '                                      x x x x x x x x x   x x
   'multiplica da direita para esquerda: 10 9 8 7 6 5 4 3 2
   '*************************************************************
   
   peso = 2    'começa multiplicar da direita para esquerda
   p = 9
   
   Do
      '*** Peso de 2 a 9 (multiplicação dos caracteres de 2 a 9) ***
      soma = soma + Mid(CPF, p, 1) * peso
      p = p - 1            'ponteiro
      peso = peso + 1      'peso
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

   digito_rv = Mid(CPF, 10, 1)  'digito verificador
   
   If CStr(digito_11) <> (digito_rv) Then
      bOk = False               'digito não confere
      VerificaCPF = bOk
      Exit Function
   End If

   
   soma = 0
   resto = 0
   digito_11 = 0        'calculado pelo módulo 11
   digito_rv = ""       'caracter digitado pelo operador

   '*************************************************************
   'número do CPF: (9+2)                  0 0 9 9 9 9 9 9 9 - D D
   '                                      x x x x x x x x x   x x
   'multiplica da direita para esquerda: 10 9 8 7 6 5 4 3 2
   '*************************************************************
   
   peso = 2    'começa multiplicar da direita para esquerda
   p = 10
   
   Do
      '*** Peso de 2 a 9 (multiplicação dos caracteres de 2 a 9) ***
      soma = soma + Mid(CPF, p, 1) * peso
      p = p - 1            'ponteiro
      peso = peso + 1      'peso
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

   digito_rv = Mid(CPF, 11, 1)  'digito verificador
   
   If CStr(digito_11) <> (digito_rv) Then
      bOk = False               'digito não confere
   End If
   
   VerificaCPF = bOk

End Function
Function LimpaTela(Form As Variant)
'* Limpa todos os objetos do Form *'
    Dim Controle As Control

    For Each Controle In Form.Controls
        If TypeName(Controle) = "UbbEdit" Then
            Controle.Text = ""
        End If
        If TypeName(Controle) = "TextBox" Then
            Controle.Text = ""
        End If
        If TypeName(Controle) = "DateEdit" Then
            Controle.Text = ""
        End If
        If TypeName(Controle) = "CurrencyEdit" Then
            Controle.Text = ""
        End If
    Next Controle

End Function
'''''''''''''''''''''''''''''''''''''''''''''''
' Retorna True se a data é válida             '
' Data deve ser informada no formato DDMMAAAA '
'''''''''''''''''''''''''''''''''''''''''''''''
Public Function DataOk(ByVal pviData As Long) As Boolean
    Dim iDia As Byte
    Dim iMes As Byte
    Dim iAno As Integer
    Dim sData As String
    Dim iUltimoDia As Byte
    Dim bOk As Boolean
    
    bOk = True
    
    sData = Format(pviData, "00000000")
    
    iDia = Left(sData, 2)
    iMes = Mid(sData, 3, 2)
    iAno = Right(sData, 4)

    If iAno < 1998 Or iAno > 2050 Then
        bOk = False
    Else
        Select Case iMes
            Case 1, 3, 5, 7, 8, 10, 12 ' 31 dias
                iUltimoDia = 31
            Case 2 ' 28/29 dias
                If iAno Mod 4 = 0 Then ' ano é bissexto
                    iUltimoDia = 29
                Else
                    iUltimoDia = 28
                End If
            Case 4, 6, 9, 11 ' 30 dias
                iUltimoDia = 30
            Case Else
                bOk = False
        End Select
        
        If bOk Then
            If iDia < 1 Or iDia > iUltimoDia Then
                bOk = False
            End If
        End If
    End If
    
    DataOk = bOk
End Function
Public Function DataDDMMAAAA(ByVal pviData As Long) As Long
    Dim sData As String
    
    sData = CStr(pviData)
    
    DataDDMMAAAA = Val(Right(sData, 2) & Mid(sData, 5, 2) & Left(sData, 4))
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''
' Converter data do formato DDMMAAAA para AAAAMMDD '
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DataAAAAMMDD(ByVal pviData As Long) As Long
    Dim sData As String
    
    sData = Format(pviData, "00000000")
    
    DataAAAAMMDD = Val(Right(sData, 4) & Mid(sData, 3, 2) & Left(sData, 2))
End Function
Public Function GeraNovaDataProc(pIdUsuario As Integer) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'           * Gera nova data de Processamento para a data de Processamento atual *           '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    GeraNovaDataProc = False
    
    ''''''''''''''''''''''''''
    '* Variáveis de Conexão *'
    ''''''''''''''''''''''''''
    Dim RsParametro         As New ADODB.Recordset
    Dim RsGrupoUsuario      As New ADODB.Recordset
    Dim RsSelecionar        As New Custodia.Selecionar
    Dim RsInsere            As New Custodia.Inserir
        
    ''''''''''''''''''''''''''''
    '* Variáveis de Parâmetro *'
    ''''''''''''''''''''''''''''
    Dim QtdCheques          As Integer
    Dim QtdDatas            As Integer
    Dim QtdDias             As Integer
    Dim DirTrans            As String * 50
    Dim DirRecep            As String * 50
    Dim SeqBordero          As Integer
    Dim GerarArquivoCEL     As Boolean
    Dim NumLoteCEL          As Integer
    Dim CompOrigem          As Integer
    Dim VersaoInicialCEL    As Integer
    Dim VersaoFinalCEL      As Integer
    Dim HeaderAV            As Boolean
    Dim chkSoma             As Boolean
    Dim CodigoUSB           As Integer
    Dim CodigoAplicacao     As String * 3
    Dim CPDOrigem           As String * 3
    Dim CPDDestino          As String * 3
    Dim CodigoTerceira      As String * 4
    Dim NomeTerceira        As String * 40
    Dim CNPJTerceira        As String * 14
    Dim UFTerceira          As String
    Dim Seq_Ocorrencia      As Long
    Dim AgAcolhed           As Integer
    Dim NumRemessa          As Integer
    Dim ValorChequeLimite   As Currency
    Dim TmpPendente         As Integer
    Dim Resp                As Integer
    Dim Fim                 As Integer
    Dim DiasLimpeza         As Integer
    Dim CidadeTerceira      As String

    ''''''''''''''''''''''''''''
    '*  Variáveis Auxiliares  *'
    ''''''''''''''''''''''''''''
    Dim nContador           As Integer
    Dim bSupervisor         As Boolean

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '* Verifica se Data  de Processamento atual existe na Tabela de Parametrôs *'
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set RsParametro = g_cMainConnection.Execute(RsSelecionar.GetParametros _
                                               (Geral.DataProcessamento))

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '* Se Data de Proc. atual não existir na Tabela de Parametrôs faz INSERT *'
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If RsParametro.EOF Then

        '''''''''''''''''''''''''''''''''''''''''''''''
        '* Verifica se Grupo do Usuario é Supervisor *'
        '''''''''''''''''''''''''''''''''''''''''''''''
        Set RsGrupoUsuario = g_cMainConnection.Execute(RsSelecionar.GetGrupoUsuario(pIdUsuario))
    
        If Not RsGrupoUsuario.EOF Then
            For nContador = 1 To RsGrupoUsuario.RecordCount
                If UCase(Trim(RsGrupoUsuario!Descricao)) = "SUPERVISOR" Then
                    bSupervisor = True
                    Exit For
                End If
                RsGrupoUsuario.MoveNext
            Next
        End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '* Se não for Supervisor não pode inicializar Sistema *'
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        If bSupervisor = False Then
            MsgBox "Sistema não pode ser inicializado por um Usuário que não é Supervisor.", vbInformation + vbOKOnly, App.Title
            Exit Function
        Else
            
            If (MsgBox("Deseja inicializar o Sistema.", vbExclamation + vbYesNo, App.Title)) = vbYes Then
            
            
            
            ' Se existe movimento não transmitido para o CH não inicializa o Sistema
            If Not VerTransmissao() Then
              Exit Function
            End If
            
                Erro = ""
            
                If VerificaLimpezaTabelas() Then
            
                   If Len(Trim(Erro)) <> 0 Then
                      MsgBox Erro
                   Else
                      Resp = MsgBox("Limpeza das Tabelas Concluída", vbOKOnly, App.Title)
                   End If
                   
               End If
            
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '* Seleciona o maior  Data de Processamento da Tabela de Parâmetros *'
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Set RsParametro = g_cMainConnection.Execute _
                                 (RsSelecionar.GetParametros(Geral.DataProcessamento, 1))
        
                '''''''''''''''''''''''''''''''''''''
                '* Preenche Variáveis de Parâmetro *'
                '''''''''''''''''''''''''''''''''''''
                If Not RsParametro.EOF Then
                    QtdCheques = RsParametro!QuantidadeCheques
                    QtdDatas = RsParametro!QuantidadeDatas
                    QtdDias = RsParametro!QuantidadeMinimaDias
                    DirTrans = RsParametro!DiretorioTransmissao
                    DirRecep = RsParametro!DiretorioRecepcao
                    SeqBordero = 0
                    GerarArquivoCEL = RsParametro!GerarArquivo_CEL
                    NumLoteCEL = RsParametro!Num_Lote_CEL
                    CompOrigem = RsParametro!Comp_Origem_CEL
                    VersaoInicialCEL = RsParametro!Num_Versao_Inicial_CEL
                    VersaoFinalCEL = RsParametro!Num_Versao_Final_CEL
                    HeaderAV = RsParametro!HeaderAV
                    chkSoma = RsParametro!CriticaSoma
                    CodigoUSB = RsParametro!Codigo_USB
                    CPDOrigem = RsParametro!CPD_Origem
                    CPDDestino = RsParametro!CPD_Destino
                    CodigoTerceira = RsParametro!Codigo_Terceira
                    NomeTerceira = RsParametro!Nome_Terceira & ""
                    CidadeTerceira = RsParametro!Cidade_Terceira & ""
                    CNPJTerceira = RsParametro!CNPJ_Terceira
                    Seq_Ocorrencia = RsParametro!Seq_Ocorrencia
                    AgAcolhed = RsParametro!CodigoAgAcolhed
                    NumRemessa = 0
                    ValorChequeLimite = RsParametro!ValorChequeLimite
                    TmpPendente = RsParametro!TMP_Pendente
                    UFTerceira = RsParametro!UF_Terceira
                    CodigoAplicacao = Format(RsParametro!CodigoAplicacao, "000")
                    DiasLimpeza = RsParametro!DiasLimpeza
                            
                    ''''''''''''''''''''''''''''''''''''''
                    '* Atualizar a Tabela de Parâmetros *'
                    ''''''''''''''''''''''''''''''''''''''
                    Set RsParametro = g_cMainConnection.Execute _
                                        (RsInsere.InsereParametro(Geral.DataProcessamento _
                                        , QtdCheques _
                                        , QtdDatas _
                                        , QtdDias _
                                        , DirTrans _
                                        , DirRecep _
                                        , SeqBordero _
                                        , GerarArquivoCEL _
                                        , NumLoteCEL _
                                        , CompOrigem _
                                        , VersaoInicialCEL _
                                        , VersaoFinalCEL _
                                        , HeaderAV, chkSoma _
                                        , CodigoUSB _
                                        , CPDOrigem _
                                        , CPDDestino _
                                        , CodigoTerceira _
                                        , CNPJTerceira _
                                        , Seq_Ocorrencia _
                                        , AgAcolhed _
                                        , NumRemessa _
                                        , ValorChequeLimite _
                                        , TmpPendente _
                                        , UFTerceira, CodigoAplicacao, DiasLimpeza, CidadeTerceira, NomeTerceira))
                End If
                
             Else
                Fim = MsgBox("Inicialização Obrigatória. O Sistema Será Finalizado", vbInformation, "Sistema Custódia")
                End
             End If
            
        End If
        
    End If
    
    GeraNovaDataProc = True
    
End Function
Public Function FormataCpfCnpj(ByVal CNPJCPF As String) As String

Dim sCodigo As String

sCodigo = Trim(CNPJCPF)

'Verifica se é CPF ou CNPJ
If Len(sCodigo) > 11 Then
     sCodigo = Right(String(14, "0") & Trim(CNPJCPF), 14)
     FormataCpfCnpj = Format(sCodigo, "@@.@@@.@@@/@@@@-@@")
Else
     sCodigo = Right(String(11, "0") & Trim(CNPJCPF), 11)
     FormataCpfCnpj = Format(sCodigo, "@@@.@@@.@@@-@@")
End If

End Function

Public Function VerTransmissao() As Boolean

    Dim rsPegaDataProcessamento As ADODB.Recordset
    Dim rsNaoTransmitidos       As ADODB.Recordset
    Dim vNroLinhas              As Integer
    Dim vDataProcessamento      As Long
    Dim Transmitidos            As New Custodia.Selecionar
    
    On Error GoTo ErroVerTransmissao:
    
   'Dias Para Verificação do movimento anterior
    vNroLinhas = NroLinhas()
    If vNroLinhas = 0 Then
       MsgBox "Não é possível Inicializar o Sistema sem Parâmetro", vbOKOnly + vbCritical
       VerTransmissao = False
       End
    End If
        
    'Obtem data para limpeza de tabelas
    Set rsPegaDataProcessamento = g_cMainConnection.Execute(Transmitidos.GetDataProcessamento(vNroLinhas))
        
    If Not rsPegaDataProcessamento.EOF Then
        
       rsPegaDataProcessamento.MoveLast
       vDataProcessamento = rsPegaDataProcessamento!DataProcessamento
           
      'Verifica se há Borderôs não transmitidos
       Set rsNaoTransmitidos = g_cMainConnection.Execute(Transmitidos.GetNaoTransmitidos(vDataProcessamento))
       
       If rsNaoTransmitidos.RecordCount = 0 Then
         VerTransmissao = True
       Else
         MsgBox "Há Borderô(s) não Transmitido(s) em Movimento(s) Anterior(es). Favor Verificar!" & vbCrLf & vbCrLf & Space(33) & "O Sistema será finalizado.", vbInformation, App.Title
         VerTransmissao = False
       End If
       
       Set rsPegaDataProcessamento = Nothing
       Set rsNaoTransmitidos = Nothing
       
       Exit Function
       
     End If
ErroVerTransmissao:
    If Not (rsPegaDataProcessamento Is Nothing) Then Set rsPegaDataProcessamento = Nothing
    VerTransmissao = False
    Call TratamentoErro("Verificar Movimentos Não Transmitidos", Err)

End Function


Public Function NumeroAD(ByVal d_DtaProcessamento As Long)

On Error GoTo TrataErro
Dim Selecao      As New Custodia.Selecionar
Dim Alteracao    As New Custodia.Atualizar
Dim rsNumAviso   As New ADODB.Recordset
Dim iretorno     As Integer
Dim NumAviso     As Integer
    
    Set rsNumAviso = g_cMainConnection.Execute(Selecao.GetNumRemessaMov(d_DtaProcessamento))
                                                
    If Not rsNumAviso.EOF Then
        NumAviso = IIf(IsNull(rsNumAviso!Num_Remessa_MOV), 0, rsNumAviso!Num_Remessa_MOV)
    End If

    Call g_cMainConnection.Execute(Alteracao.AtualizaNumRemessaMOV_Parametro(d_DtaProcessamento, NumAviso), iretorno, adCmdText)
    
    If iretorno = 0 Then
        GoTo TrataErro
    Else
        NumeroAD = NumAviso + 1
    End If
        
Exit Function
TrataErro:
    Call RemoveArquivo
    Call TratamentoErro("Erro ao atualizar Número do Aviso de Diferença no Parametro.", Err)
    
End Function


