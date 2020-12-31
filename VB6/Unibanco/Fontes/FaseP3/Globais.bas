Attribute VB_Name = "Globais"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''
' Definição do tipo de variáveis globais '
''''''''''''''''''''''''''''''''''''''''''
Public bInicializou As Boolean
Public strImagePath As String

'-------------------------------------
'           TIPO DE CAPA
'-------------------------------------
Type tpCapa
    IdCapa                      As Long
    IdLote                      As Long
    IdEnv_Mal                   As String
    Capa                        As Double
    Num_Malote                  As Double
    AgOrig                      As Integer
    Duplicidade                 As Integer
    Status                      As String
    'Informação da AGENF
    agefsdtmvan                 As Long
    agefsdtmvat                 As Long
    agefsstmovi                 As Integer
    agefsestado                 As String
End Type

'''''''''''''''''''''''''''''
'Tipo para Browse for Folder'
'''''''''''''''''''''''''''''
Type SHITEMID
     cb                         As Long
     abID                       As Byte
End Type

Type ITEMIDLIST
     mkid                       As SHITEMID
End Type

Type BROWSEINFO
     hOwner                     As Long
     pidlRoot                   As Long
     pszDisplayName             As String
     lpszTitle                  As String
     ulFlags                    As Long
     lpfn                       As Long
     lParam                     As Long
     iImage                     As Long
End Type

Public Enum cam_BrowseForFolder
    camDefualtBrowse = 0
    camTheDesktop = 0
    camProgramsFolders = 2
    camControlPanel = 3
    camPrinters = 4
    camDocumentsFolder = 5
    camFavoritesFolder = 6
    camStartupFolder = 7
    camRecentFolder = 8
    camSendToFolder = 9
    camRecycleBin = 10
    camStartMenuFolder = 11
    camDesktopFolder = 16
    camMyComputer = 17
    camNetworkNeighborhood = 18
    camNetHoodFolder = 19
    camFontsFolder = 20
    camShellNewFolder = 21
End Enum

'Constantes Globais para manipulação de diretórios
Global Const gstrSEP_URLDIR$ = "/"                      ' Separator for dividing directories in URL addresses.
Global Const gstrSEP_DIR$ = "\"                         ' Directory separator character

Global Const BIF_RETURNONLYFSDIRS = &H1
Global Const BIF_DONTGOBELOWDOMAIN = &H2
Global Const BIF_STATUSTEXT = &H4
Global Const BIF_RETURNFSANCESTORS = &H8
Global Const BIF_BROWSEFORCOMPUTER = &H1000
Global Const BIF_BROWSEFORPRINTER = &H2000

'-------------------------------------
'           TIPO DE DOCUMENTO
'-------------------------------------
Type TpDocumento
'    DataProcessamento As Long
    IdDocto                     As Long
    IdCapa                      As Long
    TipoDocto                   As Integer
    Leitura                     As String
    Frente                      As String
    Verso                       As String
    Status                      As String
    Alcada                      As String
    Autenticado                 As String
    Ocorrencia                  As Long
    OcorrenciaOK                As String
    Ordem                       As String
    ValorTotal                  As Currency
    NSU                         As String
    Terminal                    As Integer
    Vinculo                     As Long
    CMC7Associado               As String
    Duplicidade                 As Integer
    Atualizacao                 As Long
    Transacao                   As Integer
    Efetivado                   As Boolean
    PagtoTerceiro               As String
    TotalVinculado              As Currency
    Excluido                    As Boolean
    AjusteInterno               As Boolean
    Agencia                     As Integer
    Conta                       As Long
    AgenciaVinculo              As Integer
    ContaVinculo                As Long
    Estornado                   As String
End Type

''''''''''''''''''''''''''''''''''''''''''
' Definição do tipo de variáveis globais '
''''''''''''''''''''''''''''''''''''''''''
Type tpGlobais
    Capa                        As tpCapa
    Documento                   As TpDocumento
    qryLeituraParametro         As rdoQuery
    qryCriarParametro           As rdoQuery
    AgenciaCentral              As String
    Banco                       As rdoConnection
    Backup                      As Boolean
    DataProcessamento           As Long
    DiretorioDados              As String
    DiretorioImagens            As String
    DiretorioTrabalho           As String
    Estacao                     As Integer
    Scanner                     As enumScanner
    autenticadora               As enumAutentica    'alteração versão 3.3 (67)
    VIPSDLL                     As enumVipsDll      ' Unibanco ou Proservi
    'Dados do usuario
    idUsuario                   As Long             'Identificação do Usuário
    Usuario                     As String
    NomeUsuario                 As String           'Nome do usuário
    GrupoUsuario                As String
    DataUltimaTrocaSenhaUsuario As Long             'Data formato (AAAAMMDD)
    
    RetornoFinal                As String
    Intervalo                   As Integer          'usado no timer para atualizar DataAtual da capa
    Atualizacao                 As Integer          'usado no timer de atualizacao dos forms
    ValorChqInferior            As Currency
    ValorMaxADCC                As Currency
    StringConexao               As String
    DataFinalRegraAntiga_Mal    As Long             'Data Limite para aceitar Malotes na regra antiga    LimiteMaxDifLancto_Mal      As Currency         'Valor máximo permitido para diferencas a Debito ou Credito em Lancamentos Internos
    ValorAlcadaCoord_Mal        As Currency
    QtdeDiasTrocaSenha          As Integer          'Quant. de dias úteis para forçar troca senha
    
End Type
'''''''''''''''''''''''''''''''''''''''''
' Estrutura do arquivo de Retorno Final '
'''''''''''''''''''''''''''''''''''''''''
Type tpRetornoFinal
    Status                      As String * 2
    Tipo                        As String * 1
    Leitura                     As String * 63
    Frente                      As String * 12
    Verso                       As String * 12
    Origem                      As String * 1
    Estacao                     As String * 2
    CrLf                        As String * 2
End Type

Type tpRetornoVips
    Tipo                        As String * 1
    Leitura                     As String * 63
    Frente                      As String * 19
    Verso                       As String * 19
    Origem                      As String * 1
    CrLf                        As String * 2
End Type

Type tpRetornoVipsNovaDLL
    Status                      As String * 2
    Tipo                        As String * 1
    Leitura                     As String * 63
    Frente                      As String * 19
    Verso                       As String * 19
    IdScanner                   As String * 2
    Estacao                     As String * 2
    Cr                          As String * 1
    Lf                          As String * 1
End Type

'''''''''''''''''''''''''''''
' Principal variável global '
'''''''''''''''''''''''''''''
Global Geral                    As tpGlobais

'Constante contendo a cor de fundo do objeto desabilitado
Public Const G_ColorGray = &H8000000F
Public Const G_ColorBackGround = &H80000005
Public Const G_ColorBlue = &H800000    '&HC00000


'''''''''''''''''''''''
' Tratar Arquivos INI '
'''''''''''''''''''''''
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long 'ITEMIDLIST
Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

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

Type RECT
    Left                        As Long
    Top                         As Long
    Right                       As Long
    Bottom                      As Long
End Type

Public Type TpUsuario
    Usuario                     As String
    Senha                       As String
End Type

'Declarações da API
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Declare Function GetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long) As Long
Declare Function SetScrollRange Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nMinPos As Long, ByVal nMaxPos As Long, ByVal bRedraw As Long) As Long
Declare Function GetScrollRange Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long
Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
Declare Function EnableScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wSBflags As Long, ByVal wArrows As Long) As Long
'Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Any, ByVal bErase As Long) As Long
Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
'Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdata As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function GetActiveWindow Lib "user32" () As Long
Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Declare Function GetComputerName& Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long)

'Declarações de APIs da VIPSDLL
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public iVIPSDLL As Long
Public iVIPSGRIM As Long
Public iVIPSDRV As Long
Public iVIPSSERIE As Long
Public iVIPSPROD As Long
Public iVIPSCODE As Long
Public iVIPSXPGMK As Long


'Declaração da Função de Criptografia
Private Declare Function Encripta Lib "Encripta.dll" (ByVal lngIn As Long, ByVal strOut As String) As Long

' Variáveis Globais para manipulação de Janelas (LeadTools)
Global hCtl                     As OLE_HANDLE
Global IsMove                   As Boolean
Global Xold                     As Single
Global Yold                     As Single
Global Xatual                   As Single
Global Yatual                   As Single
Global Atualiza                 As Integer
Global Autentica                As Object

' Variáveis e funções para manipulação de Data e Hora
Type SYSTEMTIME
        wYear                   As Integer
        wMonth                  As Integer
        wDayOfWeek              As Integer
        wDay                    As Integer
        wHour                   As Integer
        wMinute                 As Integer
        wSecond                 As Integer
        wMilliseconds           As Integer
End Type

Declare Function SetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long
Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Declare Function SetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long
Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Function ResolveMapa(ByVal pIdUsuario As Integer, _
                            ByVal pRetorno As enumResolveMapa, _
                            Optional ByVal pMapa As Variant = Empty) As String

    Dim sStr                As String * 10
    Dim rsGrupos            As RDO.rdoResultset
    Dim rsGrupoUsuario      As RDO.rdoResultset
    Dim qryGetGrupoUsuario  As RDO.rdoQuery
    Dim qryGetAllGrupos     As RDO.rdoQuery
    Dim IdGrupo             As Integer
    Dim sStrDescricao       As String

    '''''''''''''''''''''''''''''''
    'Retorna em Mapa Ex. "   X X "'
    '''''''''''''''''''''''''''''''
    If pRetorno = eRetornaMapa Then
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Seleciona os grupos do usuario para resolver o mapa'
        '''''''''''''''''''''''''''''''''''''''''''''''''''''
        Set qryGetGrupoUsuario = Geral.Banco.CreateQuery("", "{call GetGrupoUsuario (?)}")
        Set qryGetAllGrupos = Geral.Banco.CreateQuery("", "{call GetAllGrupos }")
        qryGetGrupoUsuario.rdoParameters(0) = pIdUsuario
        Set rsGrupoUsuario = qryGetGrupoUsuario.OpenResultset(rdOpenStatic, rdConcurReadOnly)
        
        Set rsGrupos = qryGetAllGrupos.OpenResultset(rdOpenStatic, rdConcurReadOnly)
        
    '   1 - AUX Auxiliar de Supervisor
    '   2 - COO Coordenador
    '   3 - Dig Digitação
    '   4 - LID Lider
    '   5 - PES Pesquisa
    '   6 - REC Recepção
    '   7 - SPT Suporte
    '   8 - SUP Supervisor
    '   9 - TER Terceiros
    
        Do While Not rsGrupoUsuario.EOF()
            rsGrupos.MoveFirst
            Do While Not rsGrupos.EOF()
                If UCase(rsGrupoUsuario!IdGrupo) = UCase(rsGrupos!IdGrupo) Then
                    Mid(sStr, rsGrupos.AbsolutePosition, 1) = "X"
                Else
                    If Mid(sStr, rsGrupos.AbsolutePosition, 1) <> "X" Then
                        Mid(sStr, rsGrupos.AbsolutePosition, 1) = " "
                    End If
                End If
                rsGrupos.MoveNext
            Loop
            rsGrupoUsuario.MoveNext
        Loop
        
        rsGrupoUsuario.Close
        
        ''''''''''''''''
        'Retorna o Mapa'
        ''''''''''''''''
        ResolveMapa = sStr
    Else
        '''''''''''''''''''''''''''''''
        'Retorna em String Ex. SUP-REC'
        '''''''''''''''''''''''''''''''
        If IsEmpty(pMapa) Then
            sStr = ResolveMapa(pIdUsuario, eRetornaMapa)
        Else
            sStr = pMapa
        End If
        
        '   1 - AUX Auxiliar de Supervisor
        '   2 - COO Coordenador
        '   3 - Dig Digitação
        '   4 - LID Lider
        '   5 - PES Pesquisa
        '   6 - REC Recepção
        '   7 - SPT Suporte
        '   8 - SUP Supervisor
        '   9 - TER Terceiros
        
        'XXX      X'
        
        '''''''''''''''''''
        'Pega o primeiro X'
        '''''''''''''''''''
        IdGrupo = InStr(sStr, "X")
        Do While IdGrupo
        
            If Len(sStrDescricao) > 0 Then sStrDescricao = sStrDescricao & "-"
        
            sStrDescricao = sStrDescricao & Switch(IdGrupo = 1, "AUX", _
                                                   IdGrupo = 2, "COO", _
                                                   IdGrupo = 3, "DIG", _
                                                   IdGrupo = 4, "LID", _
                                                   IdGrupo = 5, "PES", _
                                                   IdGrupo = 6, "REC", _
                                                   IdGrupo = 7, "SPT", _
                                                   IdGrupo = 8, "SUP", _
                                                   IdGrupo = 9, "TER")
            
            Mid(sStr, IdGrupo, 1) = " "
            IdGrupo = InStr(sStr, "X")
        Loop
        
        ResolveMapa = sStrDescricao
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

Function CriaDir(ByVal strDirName As String) As Boolean
    Dim strPath As String
    Dim intOffset As Integer
    Dim intAnchor As Integer
    Dim strOldPath As String

    On Error Resume Next

    '
    'Add trailing backslash
    '
    If Right$(strDirName, 1) <> gstrSEP_DIR Then
        strDirName = strDirName & gstrSEP_DIR
    End If

    strOldPath = CurDir$
    CriaDir = False
    intAnchor = 0

    '
    'Loop and make each subdir of the path separately.
    '
    intOffset = InStr(intAnchor + 1, strDirName, gstrSEP_DIR)
    intAnchor = intOffset 'Start with at least one backslash, i.e. "C:\FirstDir"
    Do
        intOffset = InStr(intAnchor + 1, strDirName, gstrSEP_DIR)
        intAnchor = intOffset

        If intAnchor > 0 Then
            strPath = Left$(strDirName, intOffset - 1)
            ' Determine if this directory already exists
            Err = 0
            ChDir strPath
            If Err Then
                ' We must create this directory
                Err = 0
                MkDir strPath
            End If
        End If
    Loop Until intAnchor = 0
    
    CriaDir = DirExists(strDirName)

Done:
    ChDir strOldPath

    Err = 0
End Function


Public Function DirExists(ByVal strDirName As String) As Integer
    Const strWILDCARD$ = "*.*"

    Dim strDummy As String

    On Error Resume Next

    AddDirSep strDirName
    strDummy = Dir$(strDirName & strWILDCARD, vbDirectory)
    DirExists = Not (strDummy = vbNullString)

    Err = 0
End Function


Public Sub AddDirSep(strPathName As String)
    If Right(Trim(strPathName), Len(gstrSEP_URLDIR)) <> gstrSEP_URLDIR And _
       Right(Trim(strPathName), Len(gstrSEP_DIR)) <> gstrSEP_DIR Then
        strPathName = RTrim$(strPathName) & gstrSEP_DIR
    End If
End Sub


Public Function ComputerName() As String
    Dim l_sStr  As String
    Dim l_lBuff As Long
    
    l_lBuff = 199
    
    l_sStr = Space(255)
    
    ComputerName = "Desconhecido"
    
    If GetComputerName(l_sStr, l_lBuff) = 0 Then Exit Function
    
    ComputerName = Left(l_sStr, l_lBuff)

End Function


Public Function getPerformance(mStart As Boolean) As String
'* Calcula quanto tempo em média demora uma transação Visual Basic *'

Static start    As Double
Dim finish      As Double
    
    If mStart Then
        start# = Timer
        getPerformance = ""
    Else
        finish# = Timer
        getPerformance = Format$(finish# - start#, "##.######") & " secs."
    End If

End Function

Public Function InsereControleCapa(ByVal pDataProcessamento As Long, _
                                   ByVal pIdCapa As Long, _
                                   ByVal pComentario As String, _
                                   ByVal pIdModulo As Long) As Boolean

    Dim qryInsereControleCapa       As RDO.rdoQuery

    Set qryInsereControleCapa = Geral.Banco.CreateQuery("", "{? = call InsereControleCapa (?,?,?,?)}")

    qryInsereControleCapa.rdoParameters(0).Direction = rdParamReturnValue
    qryInsereControleCapa.rdoParameters(1) = pDataProcessamento
    qryInsereControleCapa.rdoParameters(2) = pIdCapa
    qryInsereControleCapa.rdoParameters(3) = pComentario
    qryInsereControleCapa.rdoParameters(4) = pIdModulo
    
    qryInsereControleCapa.Execute
    
    If qryInsereControleCapa.rdoParameters(0) <> 0 Then
        InsereControleCapa = False
    Else
        InsereControleCapa = True
    End If

End Function

Public Function GetControleCapa(ByVal pDataProcessamento As Long, _
                                ByVal pIdCapa As Long) As RDO.rdoResultset

    Dim qryGetControleCapa      As RDO.rdoQuery
    
    Set qryGetControleCapa = Geral.Banco.CreateQuery("", "{call GetControleCapa (?,?)}")

    qryGetControleCapa.rdoParameters(0) = pDataProcessamento
    qryGetControleCapa.rdoParameters(1) = pIdCapa
    
    Set GetControleCapa = qryGetControleCapa.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    
    

End Function


Public Function PreparaMenu(Form As Variant) As Boolean

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       *  Prepara/Habilita Menu de Acordo com Grupo que Usuário pertence *        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Control             As Control
Dim rsHabilitaGrupo     As rdoResultset
Dim qryGetHabilitaGrupo As rdoQuery

Dim bBaseAtual          As Boolean
Dim aOpcaoMenu(53)      As String
Dim iOpcaoMenu          As Integer
Dim bHabilita           As Boolean
Dim sGrupoModulo        As String
    
    PreparaMenu = False

    'Se Base de Dados em uso seja BACKUP, habilitar somente os ítens abaixo no array
    '*****   Obs.:  Adicionar o nome do ítem do menu em minúsculo  *****
    aOpcaoMenu(0) = "mnuexpedicao"                      'Menu expedicao
    aOpcaoMenu(1) = "mnusupervisao"                     'Menu supervisão
    aOpcaoMenu(2) = "mnuconsultasrelatorios"            'Menu consulta
    aOpcaoMenu(3) = "mnualtdatamovimento"               'Menu troca de Data
    aOpcaoMenu(4) = "mnusupbloquearaplicacao"           'Menu bloqueio de estação
    aOpcaoMenu(5) = "mnusobre"                          'Menu Info. do Sistema
    aOpcaoMenu(6) = "mnusupauditoria"                   'Sub Menu Auditoria"
    aOpcaoMenu(7) = "mnusupgrpmod"                      'Sub Menu Grupo de usuário
    aOpcaoMenu(8) = "mnuconconsulta"                    'Sub Menu Consulta
    aOpcaoMenu(9) = "mnuconrelatorio"                   'Sub Menu Relatorios
    aOpcaoMenu(10) = "mnusobre"                         'Sub Menu Sobre
    aOpcaoMenu(11) = "mnuconrelatorio"                  'Sub Menu Relatorios
    aOpcaoMenu(12) = "mnusair"                          'Menu Sair
    aOpcaoMenu(13) = "mnuconrelrecepcao"                'Sub Menu Relatorio
    aOpcaoMenu(14) = "mnuconreldifrecep"                'Sub Menu Relatorio
    aOpcaoMenu(15) = "reltotais"                        'Sub Menu Relatorio
    aOpcaoMenu(16) = "mnucontroleexpedidos"             'Sub Menu Relatorio
    aOpcaoMenu(17) = "mnucapasnaofinalizadas"           'Sub Menu Relatorio
    aOpcaoMenu(18) = "mnuconrelprocanalitico"           'Sub Menu Relatorio
    aOpcaoMenu(19) = "mnuconrelprocconsolidado"         'Sub Menu Relatorio
    aOpcaoMenu(20) = "mnuconsrelestorno"                'Sub Menu Relatorio
    aOpcaoMenu(21) = "mnurelsegdoctos"                  'Sub Menu Relatorio
    aOpcaoMenu(22) = "mnurelmovcomp"                    'Sub Menu Relatorio
    aOpcaoMenu(23) = "mnurelmovtochqubbcomp"            'Sub Menu Relatorio
    aOpcaoMenu(24) = "mnucaixaexpresso"                 'Sub Menu Relatorio
    aOpcaoMenu(25) = "mnucaixaexpressochequesuperior"   'Sub Menu Relatorio
    aOpcaoMenu(26) = "mnuconrelvaletransporte"          'Sub Menu Relatorio
    aOpcaoMenu(27) = "mnurellancamentointerno"          'Sub Menu Relatorio
    aOpcaoMenu(28) = "mnureldiferencascaixa"            'Sub Menu Relatorio
    aOpcaoMenu(29) = "mnuenvmalex"                      'Sub Menu Relatorio
    aOpcaoMenu(30) = "mnuenvelopesfininvestproc"        'Sub Menu Rel. Env. Finivest Recepcionado
    aOpcaoMenu(31) = "mnurelconcessionarias"            'Sub Menu Relatorio
    aOpcaoMenu(32) = "mnurelacompproducao"              'Sub Menu Relatorio
    aOpcaoMenu(33) = "mnurelenvfininvestdevolvido"      'Sub Menu Rel. Env. Finivest com ocorrência
    aOpcaoMenu(34) = "mnurelchequesunibanco"            'Menu Relatorio
    aOpcaoMenu(35) = "mnurelchqubbcompvalor"            'Sub Menu Relatorio
    aOpcaoMenu(36) = "mnurelenvfininvestporagencia"     'Sub Menu Rel Env. Fininvest por agência
    aOpcaoMenu(37) = "mnuenvelopefininvest"             'Sub Menu Rel Env. Fininvest
    aOpcaoMenu(38) = "mnuexpedicao"                     'Menu Expedição
    aOpcaoMenu(39) = "mnusupacompanhamento"             'Sub Menu Acomp. Produção
    aOpcaoMenu(40) = "mnusupacompatividade"             'Sub Menu Acomp. Atividades
    aOpcaoMenu(41) = "mnusupacomprecepcao"              'Sub Menu Acomp. Recepção
    aOpcaoMenu(42) = "mnusupacompusers"                 'Sub Menu Acomp. usuários
    aOpcaoMenu(43) = "mnusupexclusao"                   'Sub Menu Exclusão
    aOpcaoMenu(44) = "mnusupparametros"                 'Sub Menu Cad. Parâmetros
    aOpcaoMenu(45) = "mnusupcadastrousuario"            'Sub Menu Cad. Usuários
    aOpcaoMenu(46) = "mnurelrelacao_tc_ar"              'Sub Menu Relatório de TC e AR
    aOpcaoMenu(47) = "mnurelopcoestotais"               'Sub Menu Relatório de Totais
    aOpcaoMenu(48) = "mnureltotalconsolidado"           'Menu Relat. Total Consolidado
    aOpcaoMenu(49) = "mnurelpercdoctomodulo"            'Menu Relat. Percentual de Doctos por Módulo
    aOpcaoMenu(50) = "mnureltotalportipodocto"          'Menu Relat. Total Aberto por Tipo Docto
    aOpcaoMenu(51) = "mnureltotdocumentoporcliente"     'Menu Relat. Totais de Doctos por Cliente
    aOpcaoMenu(52) = "mnurelestdocagencia"              'Menu Relat. Estatística de Documentos por agência
    aOpcaoMenu(53) = "mnumotivodoctosilegiveis"         'Menu Relat. Docto Enviados para Ilegíveis
    '*****   Obs.:  Adicionar o nome do ítem do menu em minúsculo  *****
    
    
    'Verifica qual banco de dados está conectado
    If InStr(LCase(Geral.Banco.Connect), "mdi_ubbbackup") <> 0 Then
        bBaseAtual = False
    Else
        bBaseAtual = True
    End If

    '''''''''''''''''''''''''''''''''''''''''''
    ' *  Query de Pesquisa de Grupo-Modulos * '
    '''''''''''''''''''''''''''''''''''''''''''
    Set qryGetHabilitaGrupo = Geral.Banco.CreateQuery("", "{Call GetHabilitaGrupo(?)}")

    With qryGetHabilitaGrupo
        .rdoParameters(0).Value = Geral.idUsuario
        Set rsHabilitaGrupo = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With

    If Not rsHabilitaGrupo.EOF Then
        
        sGrupoModulo = "000*"
        
        Do While Not rsHabilitaGrupo.EOF
            sGrupoModulo = sGrupoModulo & Format(rsHabilitaGrupo!IdModulo, "000") & "*"
            rsHabilitaGrupo.MoveNext
        Loop

        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' * Define quais itens de Menu serão habilitados * '
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        For Each Control In Form.Controls
            'Habilita opção de Alteração de Senha
            If LCase(Geral.Usuario) <> "desenv" And bBaseAtual Then
                If LCase(Control.Name) = "mnusupalterasenha" Then Control.Enabled = True
            Else
                If LCase(Control.Name) = "mnusupalterasenha" Then Control.Enabled = False
            End If
        
            If Not bBaseAtual Then
                'Desabilita opções do menu para base backup
                bHabilita = False
                If TypeName(Control) = "Menu" Then
                    For iOpcaoMenu = 0 To UBound(aOpcaoMenu)
                        If LCase(Trim(Control.Name)) = aOpcaoMenu(iOpcaoMenu) And _
                            InStr(sGrupoModulo, Format(Control.Index, "000")) <> 0 Then
                            bHabilita = True
                            Exit For
                        End If
                    Next
                    If bHabilita Then
                        Control.Enabled = True
                    Else
                        Control.Enabled = False
                    End If
                End If
            Else
                If TypeName(Control) = "Menu" Then
                    If InStr(sGrupoModulo, Format(Control.Index, "000")) <> 0 Then
                        Control.Enabled = True
                    End If
                End If
            End If
        Next
    
    End If
    
    If Geral.Scanner <> escnSemScanner Then
        'Verificar se o usuário tem acesso ao menu Captura
        If VerificaAcessoUsuario(Geral.idUsuario, 5) Then
            'Habilitar menu no form principal
            Principal.mnuCapCaptura(5).Enabled = True
        Else
            'Desabilitar menu no form principal
            Principal.mnuCapCaptura(5).Enabled = False
        End If
    Else
        'Desabilitar menu no form principal
        Principal.mnuCapCaptura(5).Enabled = False
    End If

    PreparaMenu = True

End Function
Public Function ObtemRetornoTransacao(ByVal pRetornoTransacao As Integer, ByRef pDescricao As String) As Boolean

    Dim qryGetRetornoTransacao      As RDO.rdoQuery

    Set qryGetRetornoTransacao = Geral.Banco.CreateQuery("", "{Call GetRetornoTransacao(?,?)}")

    ObtemRetornoTransacao = False

    With qryGetRetornoTransacao
        .rdoParameters(0).Direction = rdParamInput
        .rdoParameters(1).Direction = rdParamOutput

        .rdoParameters(0) = pRetornoTransacao
        .Execute

        pDescricao = ""
        If IsNull(.rdoParameters(1).Value) Then
            pDescricao = "Retorno de Mensagem nao Tratado"
            Set qryGetRetornoTransacao = Nothing
            Exit Function
        End If

        '''''''''''''''''''''
        'Retorna a descricao'
        '''''''''''''''''''''
        pDescricao = Trim(.rdoParameters(1).Value)
    End With

    ObtemRetornoTransacao = True
    Set qryGetRetornoTransacao = Nothing
    
End Function


Public Function AtualizaAtividade(IdModulo As Integer)
'* Objetivo: Controlar todas as Atividades do Usuário *'
   
On Error GoTo Exit_AtualizaAtividade

    Dim qryAtualizaAtividade As rdoQuery
   
   rdoErrors.Clear
   
   If Geral.idUsuario = 0 Then Exit Function
   
   Set qryAtualizaAtividade = Geral.Banco.CreateQuery("", "{? = Call AtualizaAtividadeUsuario(?,?,?)}")
   
   With qryAtualizaAtividade
      .rdoParameters(0).Direction = rdParamReturnValue
      .rdoParameters(1) = Geral.DataProcessamento
      .rdoParameters(2) = Geral.idUsuario
      .rdoParameters(3) = IdModulo
      .Execute
      'Verifica se houve erro na atualização
      If .rdoParameters(0).Value <> 0 Then GoTo Exit_AtualizaAtividade
   End With
   
   qryAtualizaAtividade.Close
   
Exit Function
'Sai da Função
Exit_AtualizaAtividade:
    Select Case TratamentoErro("Erro ao atualizar atividades do usuário.", Err, rdoErrors)
        Case vbCancel, vbRetry
    End Select
        
End Function
Public Function RemoveAtividade() As Boolean
'* Objetivo: Remover Atividades do usuário ao encerrar aplicação *'

Dim qryRemoveAtividade As rdoQuery
   
   rdoErrors.Clear
   
   On Error GoTo Exit_RemoveAtividade
   
   Set qryRemoveAtividade = Geral.Banco.CreateQuery("", "{? = Call RemoveAtividade(?)}")
   
   With qryRemoveAtividade
      .rdoParameters(0).Direction = rdParamReturnValue
      .rdoParameters(1) = Geral.idUsuario
      .Execute
      'Verifica se houve erro na atualização
      If .rdoParameters(0).Value <> 0 Then GoTo Exit_RemoveAtividade
   End With
   
   qryRemoveAtividade.Close
   
Exit Function
'Sai da Função
Exit_RemoveAtividade:
   TratamentoErro "Erro ao atualizar atividades do usuário.", Err, rdoErrors
   
End Function

Public Function AtualizaStatusCapa(lIdCapa As Long, sStatus As String) As Boolean

    Dim qryAtualizaStatusCapa   As RDO.rdoQuery

    On Error GoTo Exit_AtualizaStatusCapa
    
    AtualizaStatusCapa = False
    
    'Atualiza Status da Capa
    Set qryAtualizaStatusCapa = Geral.Banco.CreateQuery("", "{? = call AtualizaStatusCapa (?,?,?)}")
        'Parâmetros (1)-Data (2)-IdCapa (3)-Status
        
    
    With qryAtualizaStatusCapa
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = lIdCapa
        .rdoParameters(3) = sStatus
        .Execute
        'Verifica se houve erro na atualização
        If .rdoParameters(0).Value = 0 Then AtualizaStatusCapa = True
        
    End With
    
    qryAtualizaStatusCapa.Close
    

    
    'Sai da função
Exit_AtualizaStatusCapa:

End Function


Public Function AtualizaStatusDocumento(ByVal IdDocto As Long, ByVal Status As String) As Boolean

    Dim qryAtualizaStatusDocumento      As rdoQuery
    
    On Error GoTo AtualizaStatusDocumento_Err


    AtualizaStatusDocumento = False

    Set qryAtualizaStatusDocumento = Geral.Banco.CreateQuery("", "{? = call AtualizaStatusDocumento (?,?,?)}")
    With qryAtualizaStatusDocumento
        .rdoParameters(0).Direction = rdParamReturnValue     'Parametro de Retorno
        .rdoParameters(1) = Geral.DataProcessamento          'Data Processamento
        .rdoParameters(2) = IdDocto                          'IdDocto
        .rdoParameters(3) = Status                           'Status do Documento
        .Execute
    End With

    If qryAtualizaStatusDocumento(0).Value <> 0 Then Exit Function

    AtualizaStatusDocumento = True

    Exit Function

AtualizaStatusDocumento_Err:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro ao Atualizar o Status do documento.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
End Function

Public Function BrowseForFolder(ByVal prmForm As Object, ByVal prmFolder As cam_BrowseForFolder) As String

    Dim bi                                  As BROWSEINFO
    Dim idl                                 As ITEMIDLIST
    Dim rtn                                 As Long
    Dim pidl                                As Long
    Dim path                                As String
    Dim Pos                                 As Integer
    Dim lresult                             As Long
    Dim X                                   As String
  
    bi.hOwner = prmForm.hwnd
    rtn& = SHGetSpecialFolderLocation(ByVal prmForm.hwnd, ByVal prmFolder, idl)
    
    bi.pidlRoot = idl.mkid.cb
    bi.lpszTitle = "Selecione a pasta"
    bi.ulFlags = BIF_RETURNONLYFSDIRS And BIF_DONTGOBELOWDOMAIN And BIF_STATUSTEXT _
        And BIF_RETURNFSANCESTORS And BIF_BROWSEFORCOMPUTER And BIF_BROWSEFORPRINTER
    
    pidl& = SHBrowseForFolder(bi) 'show the dialog box
    
    path$ = Space$(512) 'set the maximum returned path
    lresult = SHGetPathFromIDList(ByVal pidl&, ByVal path$)  'get the folder selected
    
    BrowseForFolder = ""
    If lresult Then 'if a folder was selected the
       Pos% = InStr(path$, Chr$(0)) 'extract the path
       BrowseForFolder = Left(path$, Pos - 1)
       'MsgBox "The folder you selected was:" + Chr$(10) + Chr$(10) + Left(path$, Pos - 1), vbInformation 'display the returned path
    End If
    
    

End Function

'
'Rotina que envia o documento para confirmação de agencia e conta quando usuario for terceiro
'
'
Public Function ConfirmaAgConta(ByVal pIdDocto As Long) As Boolean

    Dim qryAtualizaStatusDocumento      As RDO.rdoQuery

    On Error GoTo Erro_ConfirmaAgConta:

    ConfirmaAgConta = False

    Set qryAtualizaStatusDocumento = Geral.Banco.CreateQuery("", "{? = Call AtualizaStatusDocumento(?,?,?)}")
    With qryAtualizaStatusDocumento
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = pIdDocto
        .rdoParameters(3) = "L"
        .Execute
        If .rdoParameters(0).Value <> 0 Then
            GoTo Erro_ConfirmaAgConta
        End If
    End With
    

    ConfirmaAgConta = True
    
Erro_ConfirmaAgConta:

End Function

Public Function GrupoUsuario(ByVal pUsuario As String, ByVal pGrupo As enumGrupoUsuario) As Boolean

    Dim qryGetUsuario   As rdoQuery
    Dim rsUsuario       As rdoResultset
    Dim sIdGrupo        As String * 3
    
    On Error GoTo ErroUsuario
    
    If UCase(pUsuario) = "DESENV" Then
        GrupoUsuario = True
        Exit Function
    End If
    
    rdoErrors.Clear
    
    sIdGrupo = Switch(pGrupo = eG_AUX_SUPERVISOR, "AUX", _
                      pGrupo = eG_DIGITACAO, "DIG", _
                      pGrupo = eG_LIDER, "LID", _
                      pGrupo = eG_PESQUISA, "PES", _
                      pGrupo = eG_RECEPCAO, "REC", _
                      pGrupo = eG_SUPORTE, "SPT", _
                      pGrupo = eG_SUPERVISOR, "SUP", _
                      pGrupo = eG_TERCEIRO, "TER", _
                      pGrupo = eG_COORDENADOR, "COO")
    
    GrupoUsuario = False
'    Screen.MousePointer = vbHourglass
    
    Set qryGetUsuario = Geral.Banco.CreateQuery("", "{call GetUsuario (?)}")
    qryGetUsuario.rdoParameters(0) = pUsuario
    Set rsUsuario = qryGetUsuario.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    Do While Not rsUsuario.EOF
        If UCase(rsUsuario!IdGrupo) = sIdGrupo Then
            GrupoUsuario = True
            Exit Do
        End If
        rsUsuario.MoveNext
    Loop
    
    rsUsuario.Close
    qryGetUsuario.Close
'    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Function
    
ErroUsuario:
'    Screen.MousePointer = vbDefault
    Call TratamentoErro("Erro na obtenção do grupo do usuário.", Err, rdoErrors)

End Function

Public Function VerificaCEI(ByVal CEI As String) As Boolean

    On Error GoTo ERRO_VERIFICACEI
    
    'Esta rotina serve para conferir o CEI: tam = 12
    
    Dim soma As Integer, soma2 As Integer
    Dim dezena As String, i As Integer
    Dim peso(11) As Integer
    Dim digito As Integer
    Dim digito_cei As String
    Dim digito_rv As String
    Dim bOk As Boolean

    bOk = True           'default - OK
    
    soma = 0
    digito = 0           'calculado pelo módulo
    digito_rv = ""       'caracter digitado pelo operador
    
    'número do CEI: (11+1)                 N N N N N N N N N N A -D
    '                                      x x x x x x x x x x x  x
    'multiplica da direita para esquerda:  7 4 1 8 5 2 1 6 3 7 4 -
    
    If Val(Mid(CEI, 1, 2)) <= 0 Then
        bOk = False               'dois primeiros numeros devem ser maior que zero
        VerificaCEI = bOk
        Exit Function
    End If

    If Mid(CEI, 11, 1) <> 0 And Mid(CEI, 11, 1) <> 6 And Mid(CEI, 11, 1) <> 7 And Mid(CEI, 11, 1) <> 8 And Mid(CEI, 11, 1) <> 9 Then
        bOk = False               'atividade não confere
        VerificaCEI = bOk
        Exit Function
    End If
    
    peso(1) = 7
    peso(2) = 4
    peso(3) = 1
    peso(4) = 8
    peso(5) = 5
    peso(6) = 2
    peso(7) = 1
    peso(8) = 6
    peso(9) = 3
    peso(10) = 7
    peso(11) = 4

    For i = 1 To 11
        soma = soma + Mid(CEI, i, 1) * peso(i)
    Next i

    dezena = Right(str(soma), 2)
    dezena = (Left(dezena, 1))
    soma2 = Val(Right(soma, 1)) + dezena
    digito = 10 - Val(Right(soma2, 1))     'digito verificador
    digito_cei = Mid(CEI, 12, 1)           'digito verificador digitado
    
    If Right(CStr(digito), 1) <> (digito_cei) Then
        bOk = False                         'digito não confere
        VerificaCEI = bOk
        Exit Function
    End If

    VerificaCEI = bOk
    
    Exit Function

ERRO_VERIFICACEI:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro ao Verificar CEI.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
End Function

Public Function FormataMalote(ByVal pNumMalote As String) As String

    Dim nChars  As Integer

    If Not IsNumeric(pNumMalote) Then Exit Function
    
    
    
    
    If CStr(Left(Val(pNumMalote), 1)) <> "9" Then
        nChars = 11
        
    ElseIf CStr(Left(Val(pNumMalote), 1)) = "9" Then
        nChars = 12
    Else
        FormataMalote = ""
        Exit Function
    End If
    
    FormataMalote = Format(pNumMalote, String(nChars, "0"))

End Function


Public Sub SoNumero(ByRef KeyAscii As Integer)

    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then KeyAscii = 0

End Sub


Public Sub SelecionarTexto(ByVal pObjeto As Object)


    On Error Resume Next
    
    pObjeto.SelStart = 0
    pObjeto.SelLength = Len(pObjeto)
'    pObjeto.SetFocus
'
'    If Err <> 0 Then Err = 0
    

End Sub
Public Function VerificaDoctosExcluidosCapa(ByVal sIdCapa As Long) As Boolean

    Dim qryGetPesqDoctosOcorr As rdoQuery
    Dim qryGetMotivoExclusao  As rdoQuery
    Dim RsMotivoExclusao      As rdoResultset
    Dim TotalDoctos           As Integer
    Dim TotalDoctoComOCorr    As Integer
    Dim TotalDoctosSemCorr    As Integer

    On Error GoTo VerificaDoctosExcluidosCapa_Erro

    VerificaDoctosExcluidosCapa = False

    '* Verifica se capa possui motivo de exclusão *'
    Set qryGetMotivoExclusao = Geral.Banco.CreateQuery("", "{Call GetMotivoExclusao(?,?)}")

    With qryGetMotivoExclusao
        .rdoParameters(0).Value = Geral.DataProcessamento
        .rdoParameters(1).Value = sIdCapa
        Set RsMotivoExclusao = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With

    If Not RsMotivoExclusao.EOF Then
        VerificaDoctosExcluidosCapa = True
        Exit Function
    End If

    '* Seleciona o total de documentos que possuem ocorrência da capa *'
    Set qryGetPesqDoctosOcorr = Geral.Banco.CreateQuery("", "{Call GetPesqDoctosOcorr(?,?,?,?,?)}")

    With qryGetPesqDoctosOcorr
        .rdoParameters(0).Value = Geral.DataProcessamento
        .rdoParameters(1).Value = sIdCapa
        .rdoParameters(2).Direction = rdParamOutput
        .rdoParameters(3).Direction = rdParamOutput
        .rdoParameters(4).Direction = rdParamOutput
        .Execute
    End With

    TotalDoctos = qryGetPesqDoctosOcorr.rdoParameters(2)
    TotalDoctoComOCorr = qryGetPesqDoctosOcorr.rdoParameters(3)
    TotalDoctosSemCorr = qryGetPesqDoctosOcorr.rdoParameters(4)

    If TotalDoctos = TotalDoctoComOCorr And TotalDoctos <> 0 Then Exit Function

    VerificaDoctosExcluidosCapa = True

    Exit Function

VerificaDoctosExcluidosCapa_Erro:
  Select Case TratamentoErro("Erro ao verificar documentos com ocorrência.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Public Function G_AtualizaCamposDocumento(ByRef bDuplicidade As Boolean, ByVal lIdDocto As Long, Optional lIdCapa As Long = 0, Optional iTipoDocto As Integer = 0, Optional sLeitura As String = "", Optional sStatus As String = "", Optional sFrente As String = "", Optional sVerso As String = "", Optional dValor As Double = 0) As Boolean
 
Dim qryAlteraCamposDocumento As rdoQuery         'Altera Campos da tabela documento

On Error GoTo Err_AtualizaCamposDocumento

    'Altera Campos da tabela Documento
    Set qryAlteraCamposDocumento = Geral.Banco.CreateQuery("", "{? = call AlteraCamposDocumento (?,?,?,?,?,?,?,?,?,?)}")
        qryAlteraCamposDocumento.rdoParameters(0).Direction = rdParamReturnValue
        qryAlteraCamposDocumento.rdoParameters(10).Direction = rdParamOutput
        'Parâmetros (1)-Data (2)-IdDocto (3)-IdCapa (4)-Tipo Docto (5)-Leitura (6)-Status (7)-Frente (8)-Verso (9)-Valor (10)-Duplicidade

    G_AtualizaCamposDocumento = False
    bDuplicidade = False
    
    With qryAlteraCamposDocumento
        .rdoParameters(3) = Null        'IdCapa
        .rdoParameters(4) = Null        'Tipo Docto
        .rdoParameters(5) = Null        'Leitura
        .rdoParameters(6) = Null        'Status
        .rdoParameters(7) = Null        'Imagem Frente
        .rdoParameters(8) = Null        'Imagem Verso
        .rdoParameters(9) = Null        'Valor
            
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = lIdDocto
        If lIdCapa <> 0 Then .rdoParameters(3) = lIdCapa
        If iTipoDocto <> 0 Then .rdoParameters(4) = iTipoDocto
        If sLeitura <> "" Then .rdoParameters(5) = sLeitura
        If sStatus <> "" Then .rdoParameters(6) = sStatus
        If sFrente <> "" Then .rdoParameters(7) = sFrente
        If sVerso <> "" Then .rdoParameters(8) = sVerso
        If dValor <> 0 Then .rdoParameters(9) = dValor
        .Execute
        
        'Verifica se houve erro na atualização
        If .rdoParameters(0).Value <> 0 Then Exit Function
        
        'Verifica se houve ocorrencia de duplicidade para o campo Leitura
        If .rdoParameters("@Duplicidade") <> 0 Then bDuplicidade = True
        
        G_AtualizaCamposDocumento = True
        
    End With

    qryAlteraCamposDocumento.Close
    
    'Sai da função
     Exit Function
    
Err_AtualizaCamposDocumento:
    
    qryAlteraCamposDocumento.Close
    Select Case TratamentoErro("Não foi possível atualizar documento!", Err, rdoErrors, False)
'        Case vbCancel
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, "Atenção"
    End Select

End Function

Public Function VerificaDataMMAAAA(ByVal pviData As String) As Boolean
    
    '''''''''''''''''''''''''''''''''''''''''''''
    ' Retorna True se a data é válida           '
    ' Data deve ser informada no formato MMAAAA '
    '''''''''''''''''''''''''''''''''''''''''''''
    
    Dim iMes As String
    Dim iAno As String
    Dim sData As String
    Dim bOk As Boolean
    
    bOk = True
    
    sData = pviData
    
    iMes = Mid(sData, 1, 2)
    iAno = Right(sData, 4)
    
    ''''''''''''''''''''''''''''''''''''''
    ' Verifica se mês está entre 01 e 12 '
    ''''''''''''''''''''''''''''''''''''''
    If Val(iMes) < 1 Or Val(iMes) > 12 Then
        Beep
        MsgBox "Data inválida. Digite novamente.", vbExclamation + vbOKOnly
        bOk = False
    End If
    
    If Val(iAno) < 1950 Then
        Beep
        MsgBox "O ano não pode ser menor do que 1950.", vbExclamation
        bOk = False
    End If
    If Val(iAno) > 2051 Then
        Beep
        MsgBox "O ano não pode ser maior do que 2051.", vbExclamation
        bOk = False
    End If
        
    VerificaDataMMAAAA = bOk

End Function

Public Function Analisa_Ocor(Ocorrencia) As String
    
    Dim TbOcorrencia As rdoResultset
    Dim sSql As String
    
    If Len(Ocorrencia) > 5 Then
        Ocorrencia = Left(Ocorrencia, 3) & Right(Ocorrencia, 2)
    End If
        
    If Mid(Ocorrencia, 1, 3) <> "999" And Ocorrencia <> Space(3) And _
       Val(Mid(Ocorrencia, 1, 3)) >= 0 And Val(Mid(Ocorrencia, 1, 3)) < 999 Then
        
        Ocorrencia = Mid(Ocorrencia, 1, 3)
        
        ''''''''''''''''''''''''''''''''''''''''''''
        ' Alteração p/ verificar as ocorrências da '
        ' tabela e não mais nos fontes do sistema  '
        ' Alterado by Leda - 28/03/2000            '
        ''''''''''''''''''''''''''''''''''''''''''''
        If IsNumeric(Ocorrencia) Then
        
            sSql = "Select * From Ocorrencia "
            sSql = sSql & "Where Ocorrencia = " & Val(Ocorrencia)
            
            Set TbOcorrencia = Geral.Banco.OpenResultset(sSql, rdOpenKeyset, rdConcurReadOnly)
            
            If Not TbOcorrencia.EOF Then
                Analisa_Ocor = TbOcorrencia!Descricao
            Else
                Analisa_Ocor = Ocorrencia & " - Codigo de Ocorrencia nao tratado"
            End If
            TbOcorrencia.Close
        Else
            Analisa_Ocor = ""
        End If
        
    Else
    
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Versão 3.3 C/E -                               '
        ' Mostrar ocorrencia não tratada pelo UBB        '
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        If (Mid(Ocorrencia, 1, 3) = "999") Then
            Select Case Mid$(Ocorrencia, 4, 2)
                Case "01"
                    Analisa_Ocor = "Erro Operacional"
                Case "41"
                    Analisa_Ocor = "Arrecadacao nao Conveniada"
                Case "42"
                    Analisa_Ocor = "Envelope recebido para Processamento"
                Case "43"
                    Analisa_Ocor = "Pagamento com cheque Roxo"
                Case "44"
                    Analisa_Ocor = "Agencia nao Cadastrada"
                Case "45"
                   Analisa_Ocor = "Ficha de Deposito já utilizada"
                Case "46"
                   Analisa_Ocor = "Conta Poupança nao Encontrada"
                Case "47"
                   Analisa_Ocor = "Agencia nao Cadastrada"
                Case "48"
                   Analisa_Ocor = "Conta Corrente nao Encontrada"
                Case "49"
                   Analisa_Ocor = "Codigo de Barras Zerado"
                Case "50"
                   Analisa_Ocor = "Erro no envio da BHS1"
                Case "51"
                   Analisa_Ocor = "Retorno de Mensagem nao Tratado"
                Case "52"
                   Analisa_Ocor = "Erro no Vinculo (Cheque x Titulo)"
                Case "53"
                   Analisa_Ocor = "Valor dos cheques diferente do Informado"
                Case "54"
                   Analisa_Ocor = "Excluido pelo Supervisor"
                Case "55"
                   Analisa_Ocor = "Conta nao encontrada"
                Case "56"
                   Analisa_Ocor = "Conta Unibanco nao Existe"
                Case Else
                   Analisa_Ocor = Mid$(Ocorrencia, 4, 2) & " - Retorno de Mensagem nao Tratado"
            End Select
        Else
            Analisa_Ocor = ""
        End If
    End If

End Function

Public Function VerificaAtraso(ByVal Mov_Ant As String, ByVal Mov_At As String, ByVal Data_Vencimento As String) As Boolean
    Dim tb As rdoResultset
    Dim sSql As String
    Dim week_day, week_day_at, hoje As Integer
    Dim dia_ant As Date
    Dim myd_at, myd, dt_pr As String
    Dim resp As Boolean
    Dim total_dias, feriado As Integer
    Dim dias As Integer
        
    VerificaAtraso = False
                
    '************************************************'
    '* Verifica o WEEKDAY da data digitada ddmmaaaa *'
    '************************************************'
    myd = Right$(Data_Vencimento, 2) + "/" + Mid$(Data_Vencimento, 5, 2) + "/" + Left(Data_Vencimento, 4)
    week_day = Weekday(myd)
    
    ' Sunday    1     Thursday 5
    ' Monday    2     Friday   6
    ' Tuesday   3     Saturday 7
    ' wednesday 4
    
    ''''''''''''''''''''''''''''''
    ' Verifica diferença de dias '
    ''''''''''''''''''''''''''''''
    'A2_OK-130
    dt_pr = Mid$(Geral.DataProcessamento, 7, 2) + "/" + Mid$(Geral.DataProcessamento, 5, 2) + "/" + Mid$(Geral.DataProcessamento, 1, 4)
    hoje = Weekday(dt_pr)
    dias = DateDiff("d", myd, dt_pr)
    total_dias = dias
    
    If (dias > 0) Then
        If week_day = 1 Then
            total_dias = total_dias - 1
        ElseIf week_day = 7 Then
            total_dias = total_dias - 2
        Else
            If Data_Vencimento > Mov_Ant Then
                feriado = DateDiff("d", Mov_Ant, Data_Vencimento)
                total_dias = total_dias - feriado
            End If
        End If
    End If

    If total_dias > 0 Then
        'Retorna a qtde de dias aceito pelo Unibanco sem ter que consultar no Host
        sSql = ""
        sSql = sSql & "select distinct a.PrazoVencimento "
        sSql = sSql & "from parametro a "
        sSql = sSql & "where a.dataprocessamento = " & Geral.DataProcessamento
        
        Set tb = Geral.Banco.OpenResultset(sSql, rdOpenKeyset, rdConcurReadOnly)
        
        If Not tb.EOF Then
            If Not IsNull(tb!PrazoVencimento) Then
                If total_dias >= tb!PrazoVencimento Then
                    'Grava "s" na tabela de documentos
                    VerificaAtraso = True
                Else
                    'Grava'N' na tabela de documentos
                    VerificaAtraso = False
                End If
            End If
        Else
            MsgBox "Verifiqueos parâmetros pois não foi possível consultar o prazo de vencimento do documento.", vbInformation, "Prazo Vencimento"
            VerificaAtraso = False
            Exit Function
        End If
    End If
End Function

Public Function Desformata_Valor(Valor As String) As String
    ''''''''''''''''''''''''''''''''''''''
    ' Versã0 3.3 (C/E - Item 20)         '
    ' desformatar campo p/ gravar        '
    ''''''''''''''''''''''''''''''''''''''
        
    Dim nInd As Integer
    Dim sTexto As String
    
    For nInd = 1 To Len(Valor)
        If IsNumeric(Mid(Valor, nInd, 1)) Then
            sTexto = sTexto & Mid(Valor, nInd, 1)
        End If
    Next nInd
    
    Desformata_Valor = sTexto

End Function

Public Function Formata_Valor(Valor As String) As String
    ''''''''''''''''''''''''''''''''''''''
    ' Versão 3.3  (20)                   '
    ' Formatar todos os campos de valor  '
    ' em formato moeda                   '
    ''''''''''''''''''''''''''''''''''''''
    
    Dim nInd As Integer
    Dim sTexto As String
    Dim sTexto2 As String
    
    sTexto = Valor
    sTexto2 = ""
    
    For nInd = 1 To Len(sTexto)
        If IsNumeric(Mid(sTexto, nInd, 1)) Then
            sTexto2 = sTexto2 & Mid(sTexto, nInd, 1)
        End If
    Next nInd
    
    DoEvents
    
    If Val(sTexto2) = 0 Then
        sTexto = ""
    ElseIf Len(Trim(sTexto2)) > 2 Then
        sTexto2 = Format(Mid(sTexto2, 1, (Len(sTexto2) - 2)), "###,###,###") & "," & Right(sTexto2, 2)
    End If
    Formata_Valor = sTexto2

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''
' Inibir a digitação de letras em campos numéricos '
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InibirTeclaAlfaCompensa(ByRef priTecla As Integer, Optional ByVal pvbSalta As Boolean)
    If (priTecla < &H30 Or priTecla > &H39) And priTecla <> &H20 And priTecla <> &H8 And priTecla <> &HD And priTecla <> &H1B Or priTecla = 46 Then
        priTecla = 0
    ElseIf priTecla = &HD Then
        If pvbSalta Then
            SendKeys "{TAB}"
        End If
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Inibir a digitação de letras em campos numéricos '
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InibirTeclaAlfaValor(ByRef priTecla As Integer, Optional ByVal pvbSalta As Boolean)
    'versão 3.3
    If (priTecla < &H30 Or priTecla > &H39) And priTecla <> &H2E And priTecla <> &H2C And priTecla <> &H8 And priTecla <> &HD And priTecla <> &H1B Then
        priTecla = 0
    ElseIf priTecla = &HD Then
        If pvbSalta Then
            SendKeys "{TAB}"
        End If
    End If
End Sub

Function OldSplitEnvelope(ByVal pviDataProcessamento As Long, ByVal pviIdEnvelopePai As Long, ByVal pviIdEnvelopeFilho As Long, ByVal pviIdLote As Long, ByVal pviIdDocto As Long, ByVal pviEnvelope As Long) As Boolean
    Dim sSql As String
    Dim tb As rdoResultset
    Dim iTotalDoctosAntes As Long
    Dim iTotalDoctosDepois As Long
    Dim iIdLote As Long
    Dim iDoctos As Long
    
    OldSplitEnvelope = False
    
    On Error GoTo ErroSplit
    
    sSql = ""
    sSql = sSql & "select QtdDoctos "
    sSql = sSql & "from envelope "
    'A2_OK-77
    sSql = sSql & "where DataProcessamento = " & pviDataProcessamento
    sSql = sSql & "  and IdEnvelope        = " & pviIdEnvelopePai
    
    Set tb = Geral.Banco.OpenResultset(sSql, rdOpenKeyset, rdConcurReadOnly)
    
    If Not tb.EOF Then
        iTotalDoctosAntes = tb!QtdDoctos
        tb.Close
        
        Geral.Banco.BeginTrans
        
        sSql = ""
        'A2_OK-78
        sSql = sSql & "exec SetarEnvelope " & Geral.DataProcessamento & "," & pviIdLote & "," & pviIdDocto & "," & pviEnvelope
        Geral.Banco.Execute sSql
        
        sSql = ""
        sSql = sSql & "update Documento "
        sSql = sSql & "Set IdEnvelope = " & pviIdEnvelopeFilho & " "
        'A2_OK-79
        sSql = sSql & "where DataProcessamento = " & pviDataProcessamento
        sSql = sSql & "  and idenvelope        = " & pviIdEnvelopePai & " "
        sSql = sSql & "  and status            = '0' "
        sSql = sSql & "  and tipodocto    not in (32,33,34) "
        Geral.Banco.Execute sSql
        
        iDoctos = Geral.Banco.RowsAffected + 1
        
        sSql = ""
        sSql = sSql & "update Envelope "
        sSql = sSql & "set QtdDoctos = " & iDoctos & " "
        'A2_OK-80
        sSql = sSql & "where DataProcessamento = " & pviDataProcessamento
        sSql = sSql & "  and IdEnvelope        = " & pviIdEnvelopeFilho & " "
        Geral.Banco.Execute sSql
        
        sSql = ""
        sSql = sSql & "select count(*) as Conta "
        sSql = sSql & "from Documento "
        'A2_OK-81
        sSql = sSql & "where DataProcessamento = " & pviDataProcessamento
        sSql = sSql & "  and idenvelope        = " & pviIdEnvelopePai & " "
        sSql = sSql & "  and status            = '1' "
        Set tb = Geral.Banco.OpenResultset(sSql, rdOpenKeyset, rdConcurReadOnly)
        
        iDoctos = tb!Conta
        
        tb.Close
        
        sSql = ""
        sSql = sSql & "update Envelope "
        sSql = sSql & "set QtdDoctos = " & iDoctos & " "
        'A2_OK-82
        sSql = sSql & "where DataProcessamento = " & pviDataProcessamento
        sSql = sSql & "  and IdEnvelope        = " & pviIdEnvelopePai & " "
        Geral.Banco.Execute sSql
        
        sSql = ""
        'A2_OK-83
        sSql = sSql & "exec ChecarStatusEnvelope " & pviDataProcessamento & "," & pviIdEnvelopePai & ",0,0 "
        Geral.Banco.Execute sSql
        
        sSql = ""
        'A2_OK-84
        sSql = sSql & "exec ChecarStatusEnvelope " & pviDataProcessamento & "," & pviIdEnvelopeFilho & ",0,0 "
        Geral.Banco.Execute sSql
        
        Geral.Banco.CommitTrans
        
        sSql = ""
        sSql = sSql & "select QtdDoctos "
        sSql = sSql & "from envelope "
        'A2_OK-85
        sSql = sSql & "where DataProcessamento = " & pviDataProcessamento
        sSql = sSql & "  and IdEnvelope        = " & pviIdEnvelopePai
        
        Set tb = Geral.Banco.OpenResultset(sSql, rdOpenKeyset, rdConcurReadOnly)
                
        If Not tb.EOF Then
            iTotalDoctosDepois = tb!QtdDoctos
            
            If iTotalDoctosAntes <> iTotalDoctosDepois Then
                OldSplitEnvelope = True
            End If
        End If
        
        tb.Close
    Else
        tb.Close
    End If
    
    Exit Function
ErroSplit:
    Geral.Banco.RollbackTrans
    Select Case TratamentoErro("Não foi possível realizar o Split deste Envelope!", Err, rdoErrors, False)
        Case vbCancel
            OldSplitEnvelope = False
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, "Atenção"
            OldSplitEnvelope = False
    End Select
End Function

Function SplitEnvelope(ByVal pviDataProcessamento As Long, ByVal pviIdEnvelopePai As Long, ByVal pviIdEnvelopeFilho As Long, ByVal pviIdLote As Long, ByVal pviIdDocto As Long, ByVal pviEnvelope As Long) As Boolean
    
    Dim sSql As String
    Dim tb As rdoResultset
    Dim iTotalDoctosAntes As Long
    Dim iTotalDoctosDepois As Long
    Dim iIdLote As Long
    Dim iDoctos As Long
    Dim qry As rdoQuery
    
    SplitEnvelope = False
    
    On Error GoTo ErroSplit
    
    sSql = ""
    sSql = sSql & "select QtdDoctos "
    sSql = sSql & "from envelope "
    'A2_OK-86
    sSql = sSql & "where DataProcessamento = " & pviDataProcessamento
    sSql = sSql & "  and IdEnvelope        = " & pviIdEnvelopePai
    
    Set tb = Geral.Banco.OpenResultset(sSql, rdOpenKeyset, rdConcurReadOnly)
    
    If Not tb.EOF Then
        iTotalDoctosAntes = tb!QtdDoctos
        tb.Close
        
        sSql = "{? = call SplitEnvelope2(?,?,?,?,?,?)}"
        
        Set qry = Geral.Banco.CreateQuery("", sSql)
        
        With qry
            'A2_OK-87
            .rdoParameters(1) = pviDataProcessamento
            .rdoParameters(2) = pviIdEnvelopePai
            .rdoParameters(3) = pviIdEnvelopeFilho
            .rdoParameters(4) = pviIdLote
            .rdoParameters(5) = pviIdDocto
            .rdoParameters(6) = pviEnvelope
            .Execute
            
            Do While Geral.Banco.StillExecuting
                DoEvents
            Loop
            
            If .rdoParameters(0) <> 1 Then
                qry.Close
                GoTo ErroSplit
            End If
        End With
        qry.Close
        
        sSql = ""
        sSql = sSql & "select QtdDoctos "
        sSql = sSql & "from envelope "
        'A2_OK-88
        sSql = sSql & "where DataProcessamento = " & pviDataProcessamento
        sSql = sSql & "  and IdEnvelope        = " & pviIdEnvelopePai
        
        Set tb = Geral.Banco.OpenResultset(sSql, rdOpenKeyset, rdConcurReadOnly)
                
        If Not tb.EOF Then
            iTotalDoctosDepois = tb!QtdDoctos
            
            If iTotalDoctosAntes <> iTotalDoctosDepois Then
                SplitEnvelope = True
            End If
        End If
        
        tb.Close
    Else
        tb.Close
    End If
    
    Exit Function
ErroSplit:
    Geral.Banco.RollbackTrans
    Select Case TratamentoErro("Não foi possível realizar o Split deste Envelope!", Err, rdoErrors, False)
        Case vbCancel
            SplitEnvelope = False
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, "Atenção"
            SplitEnvelope = False
    End Select
End Function

Public Function VerificarArrecadacaoConvencional(ByVal pCodigoBarras As String) As Boolean

  On Error GoTo ERRO_VERIFICA_ARREC

  Dim sSql As String
  Dim RsCONAX As rdoResultset
  Dim qryGetCONAX As rdoQuery

  If Mid(pCodigoBarras, 1, 1) <> "8" Then
    VerificarArrecadacaoConvencional = False
    Exit Function
  End If

  'Validar Código do Produto (tabela : 'CONAX')
  sSql = Mid(pCodigoBarras, 16, 4) & ","                            'Código do Produto
  sSql = sSql & Mid(pCodigoBarras, 2, 1) & ","                      'Código do Segmento
  sSql = sSql & "'',"                                               'Descrição do Produto
  sSql = sSql & Geral.AgenciaCentral & ","                          'Agencia Central
  sSql = sSql & "4"                                                 'Tipo de Consulta

  Set qryGetCONAX = Geral.Banco.CreateQuery("", "{call GetCONAX (" & sSql & ")}")
  
  Set RsCONAX = qryGetCONAX.OpenResultset(rdOpenStatic, rdConcurReadOnly)
  If RsCONAX.EOF Then
    VerificarArrecadacaoConvencional = True
  Else
    VerificarArrecadacaoConvencional = False
  End If

  Exit Function

ERRO_VERIFICA_ARREC:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Verificar se o Documento é uma Arrecadação Convencional.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Public Function ChecarDiretorio(ByVal pvsDir As String, pvsMsgErro As String) As Boolean
    On Error Resume Next
    If Len(Trim(Dir(pvsDir, vbDirectory))) <> 0 Then
        ChecarDiretorio = True
    Else
        If MsgBox(pvsMsgErro & vbCr & vbCr & "Deseja criá-lo?", vbQuestion + vbYesNo, "Validação dos Parâmetros") = vbYes Then
        
            Err.Clear
            MkDir pvsDir
            
            If Err <> 0 Then
                MsgBox "Não foi possível criar o diretório " & pvsDir & "!", vbCritical + vbOKOnly, "Validação do Parâmetros"
                ChecarDiretorio = False
            Else
                ChecarDiretorio = True
            End If
        Else
            ChecarDiretorio = False
        End If
    End If
End Function

Public Function PegarOpcaoINI(ByVal pvsSecao As String, pvsItem As String, ByVal pvsDefault As String) As String
    Dim iRet As Long
    Dim sDado As String
    Dim sDadoAux As String
    Dim i As Integer
    Dim sArquivoINI As String
    
    'Abrir arquivo INI conforme opção
    If pvsSecao = "Conexao" And (pvsItem = "Senha" Or pvsItem = "Usuario") Then
        sArquivoINI = App.path & "\MDI_Conexao.INI"
    Else
        sArquivoINI = App.path & "\MDI_UBB.INI"
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

Public Sub GravarOpcaoINI(ByVal pvsSecao As String, pvsItem As String, ByVal pvsValor As String)
    Dim iRet As Long
    
    iRet = WritePrivateProfileString(pvsSecao, pvsItem, pvsValor, App.path & "\MDI_UBB.INI")
End Sub

Public Function TransID(pvsCodigoBarras As String) As Byte
   
   Dim TipoDocto As Byte
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' Esta função retorna o tipo de transação encontrada  '
   ' 20 - Água
   ' 21 - Gás
   ' 22 - Luz
   ' 23 - Telefone
   ' 24 - Tributos Municipais
   ' 25 - Tributos Estaduais
   ' 26 - Tributos Federais
   ' 27 - Arrecadação Convencional
   ' 28 - Unicobrança Unibanco
   ' 29 - Cobrança Imediata Unibanco
   ' 30 - Cobrança Especial Unibanco
   ' 31 - Cobrança Terceiros
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
   TipoDocto = 0    'default - não tem codigo de transação
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' Verifica se é Concessionaria ou Ficha de Compensação '
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   If (Len(Trim(pvsCodigoBarras) = 44)) Then
      
      ' se concessionária
      If (Mid(pvsCodigoBarras, 1, 1) = "8") Then
         Select Case (Mid(pvsCodigoBarras, 2, 1))
            Case "1"
               TipoDocto = 24           'TRIBUTOS MUNICIPAIS
               'If Mid(pvsCodigoBarras, 16, 4) = "0000" Then
               '   Valida_Vencto_PMSP
               'End If
            Case "2"
               TipoDocto = 20           'ÁGUA
            Case "3"
               If (Mid(pvsCodigoBarras, 17, 3) = "056") Or (Mid(pvsCodigoBarras, 17, 3) = "057") Then
                  TipoDocto = 21        'GÁS
               Else
                  TipoDocto = 22        'LUZ
               End If
            Case "4"
                  TipoDocto = 23        'TELEFONE
            Case "5"
               If Val(Mid(pvsCodigoBarras, 17, 3) <= 27) Then
                  TipoDocto = 25        'TRIBUTOS ESTADUAIS
                  'If Mid(pvsCodigoBarras, 17, 3) = "025" Then
                  '   reg_ind.Cod_trans = "0384"
                  'End If
               Else
                  TipoDocto = 26        'TRIBUTOS FEDERAIS
               End If
            Case "6"
               TipoDocto = 27           'ARRECADAÇÃO CONVENCIONAL
            Case Else
               TipoDocto = 0            'não tem codigo de transação
         End Select
      Else
         ' se ficha de compensação
         If (Mid(pvsCodigoBarras, 1, 3) = "409") Then
            If (Mid(pvsCodigoBarras, 20, 2) = "04") Then
               TipoDocto = 28           'UNICOBRANÇA
            End If
            If (Mid(pvsCodigoBarras, 20, 1) = "6") Then
               TipoDocto = 29           'COBRANÇA IMEDIATA UNIBANCO
            End If
            If (Val((Mid(pvsCodigoBarras, 20, 1))) >= 1) And (Val((Mid(pvsCodigoBarras, 20, 1))) <= 5) Then
               TipoDocto = 30           '1,2,3,4,5 COBRANÇA ESPECIAL UNIBANCO
            End If
            
            'caso alguns dos números acima não tenham sido identificados,
            'o codigo de barras deverá ser redigitado
            If (Val((Mid(pvsCodigoBarras, 20, 1))) > 6) Then
               TipoDocto = 0            'não tem cod_trans
            End If
         Else
            TipoDocto = 31              'COBRANÇA DE TERCEIROS
         End If
      End If
   End If
      
   TransID = TipoDocto
   
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
   
   
  ' Verifica digitos do CGC CNPJ 00.000.000 a 99.999.999
  
  If Mid(CGC, 2, 8) = "00000000" Or _
     Mid(CGC, 2, 8) = "11111111" Or _
     Mid(CGC, 2, 8) = "22222222" Or _
     Mid(CGC, 2, 8) = "33333333" Or _
     Mid(CGC, 2, 8) = "44444444" Or _
     Mid(CGC, 2, 8) = "55555555" Or _
     Mid(CGC, 2, 8) = "66666666" Or _
     Mid(CGC, 2, 8) = "77777777" Or _
     Mid(CGC, 2, 8) = "88888888" Or _
     Mid(CGC, 2, 8) = "99999999" Then
    
        bOk = False               'digito não confere
        VerificaCGC = bOk
        Exit Function
     
  End If
  
  '------------------------
   
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
   
   
   'Verifica CPF 111.111.111-11 a 999.999.999-99
  If CPF = "11111111111" Or _
    CPF = "22222222222" Or _
    CPF = "33333333333" Or _
    CPF = "44444444444" Or _
    CPF = "55555555555" Or _
    CPF = "66666666666" Or _
    CPF = "77777777777" Or _
    CPF = "88888888888" Or _
    CPF = "99999999999" Or _
    CPF = "00000000000" Then
    
    bOk = False               'digito não confere
    VerificaCPF = bOk
    Exit Function
    
  End If
   
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

Public Function TratarCamposCMC7(ByVal CMC7 As String, ByRef Campo1 As String, _
                            ByRef Campo2 As String, ByRef Campo3 As String, _
                            ByRef Valor As String) As Boolean
    Dim Pos As Integer
    Dim buffer As String
    Dim Banco As String * 3
    Dim Agencia As String * 4
    Dim DV2 As String * 1
    Dim Compe As String * 3
    Dim Cheque As String * 6
    Dim Tipif As String * 1
    Dim DV1 As String * 1
    Dim Conta As String * 10
    Dim DV3 As String * 1
    Dim Aux As String
    Dim Count As Integer
    
    ' Inicializando as variaveis
    Banco = String(3, "0")
    Agencia = String(4, "0")
    DV2 = "0"
    Compe = String(3, "0")
    Cheque = String(6, "0")
    Tipif = "0"
    DV1 = "0"
    Conta = String(10, "0")
    DV3 = "0"
    Aux = String(3, "0")
    Valor = String(3, "0")
    
    CMC7 = Trim(CMC7)
    
    ' Montar o buffer sem o caracteres delimitadores dos campos
    buffer = ""
    For Count = 1 To Len(CMC7)
        If Mid(CMC7, Count, 1) <> "<" And _
           Mid(CMC7, Count, 1) <> ">" And _
           Mid(CMC7, Count, 1) <> ":" And _
           Mid(CMC7, Count, 1) <> ";" Then
            buffer = buffer & Mid(CMC7, Count, 1)
        End If
    Next
    
    ' jogar nas variaveis pela posicao no buffer
    If Len(buffer) >= 3 Then
        Banco = Mid(buffer, 1, 3)
    End If
    If Len(buffer) >= 7 Then
        Agencia = Mid(buffer, 4, 4)
    End If
    If Len(buffer) >= 8 Then
        DV2 = Mid(buffer, 8, 1)
    End If
    If Len(buffer) >= 11 Then
        Compe = Mid(buffer, 9, 3)
    End If
    If Len(buffer) >= 17 Then
        Cheque = Mid(buffer, 12, 6)
    End If
    If Len(buffer) >= 18 Then
        Tipif = Mid(buffer, 18, 1)
    End If
    If Len(buffer) >= 19 Then
        DV1 = Mid(buffer, 19, 1)
    End If
    If Len(buffer) >= 29 Then
        Conta = Mid(buffer, 20, 10)
    End If
    If Len(buffer) >= 30 Then
        DV3 = Mid(buffer, 30, 1)
    End If
    Pos = InStr(1, CMC7, ":", vbTextCompare)
    If Pos > 0 Then
        Aux = Mid(CMC7, Pos + 1, Len(CMC7) - Pos)
        Valor = ""
        For Count = 1 To Len(Aux)
            If IsNumeric(Mid(Aux, Count, 1)) Then
                Valor = Valor & Mid(Aux, Count, 1)
            End If
        Next
        If Len(Valor) < 3 Then
            Valor = Format(Val(Valor), "000")
        End If
    End If
    
    ' Verifica se valores sao numericos
    If Not IsNumeric(Banco) Or Not IsNumeric(Agencia) Then
        Banco = String(3, "0")
        Agencia = String(4, "0")
    End If
    If Not IsNumeric(Compe) Or Not IsNumeric(Cheque) Or Not IsNumeric(Tipif) Then
        Compe = String(3, "0")
        Cheque = String(6, "0")
        Tipif = "0"
    End If
    If Not IsNumeric(Conta) Then
        Conta = String(10, "0")
    End If
    
    ' verifica se eh possivel calcular os DVs
    If Val(Banco & Agencia) > 0 And IsNumeric(DV1) Then
        If DV10(Banco & Agencia) <> DV1 Then
            Campo1 = String(8, "0")
        Else
            Campo1 = Banco & Agencia & DV2
        End If
    Else
        Campo1 = String(8, "0")
    End If
    If Val(Compe & Cheque & Tipif) > 0 And IsNumeric(DV2) Then
        If DV10(Compe & Cheque & Tipif) <> DV2 Then
            Campo2 = String(10, "0")
        Else
            Campo2 = Compe & Cheque & Tipif
        End If
    Else
        Campo2 = String(10, "0")
    End If
    If Val(Conta) > 0 And IsNumeric(DV3) Then
        If DV10(Conta) <> DV3 Then
            Campo3 = String(12, "0")
        Else
            Campo3 = DV1 & Conta & DV3
        End If
    Else
        Campo3 = String(12, "0")
    End If
    
    ' retorna True se todos os campos batem o dv
    If Val(Campo1) > 0 And Val(Campo2) > 0 And Val(Campo3) > 0 Then
        TratarCamposCMC7 = True
    Else
        TratarCamposCMC7 = False
    End If
                            
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Converter data do formato AAAAMMDD para DDMMAAAA '
''''''''''''''''''''''''''''''''''''''''''''''''''''
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

Sub Main()
    Dim iRet As Long
    Dim tb As rdoResultset
    Dim tb1 As rdoResultset
    Dim Data As SYSTEMTIME
    Dim ScannerOk As Boolean
    Dim NumBoxes As Long
    Dim MaxDocBox As Long
    Dim BoxDefault As Long
    Dim Threshold As Long
    Dim Compress As Long
    Dim Resolution As Long
    
    ''''''''''''''''''''''''''''''''''''''
    ' Chave do algorito de criptografia  '
    ''''''''''''''''''''''''''''''''''''''
    Key(0) = 97
    Key(1) = 150
    Key(2) = 127
    Key(3) = 254
    
    '''''''''''''''''''''''''''''''''''''''''
    ' Definir rotina de tratamento de erros '
    '''''''''''''''''''''''''''''''''''''''''
    On Error GoTo ErroMain
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Verificar se o programa foi aberto mais de 1 vez '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    If App.PrevInstance Then
        MsgBox "Programa já esta sendo executado, não é possível executar outra cópia.", vbExclamation + vbOKOnly, App.Title
        End
    End If
    
    ' Set o diretorio corrente para a VipsDll encontrar os arquivos
    Call SetCurrentDirectory(App.path)
    
    Geral.DataProcessamento = Val(Format(Now, "yyyymmdd"))
    
    '''''''''''''''''''''''''''
    ' Rotina de inicialização '
    '''''''''''''''''''''''''''
    Load Password
    
    While Not Password.SenhaOk
        Password.Show vbModal, Principal
        If Password.Cancelou Then
            Unload Principal
            End
        End If
    Wend
    
    Unload Password

    ''''''''''''''''''''''''''''''''''
    ' Ajustar Data e Hora da máquina '
    ''''''''''''''''''''''''''''''''''
    Set tb = Geral.Banco.OpenResultset("select getdate()")
    GetLocalTime Data
    With Data
        .wDay = Day(tb(0))
        .wDayOfWeek = Weekday(tb(0), vbSunday) - 1
        .wMonth = Month(tb(0))
        .wYear = Year(tb(0))
        .wHour = Hour(tb(0))
        .wMinute = Minute(tb(0))
        .wSecond = Second(tb(0))
        .wMilliseconds = 0
    End With
    SetLocalTime Data
    tb.Close
    
    '''''''''''''''''''''''''''
    ' Apresentar SplashScreen '
    '''''''''''''''''''''''''''
    SplashScreen.Show
    DoEvents
    
    ''''''''''''''''''''''''''''''''''''''''''
    ' Cria query para criar tabela parametro '
    ''''''''''''''''''''''''''''''''''''''''''
    Set Geral.qryCriarParametro = Geral.Banco.CreateQuery("", "{? = call CriarParametro (?)}")
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Cria query para a leitura da tabela parametros '
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    Set Geral.qryLeituraParametro = Geral.Banco.CreateQuery("", "{call LerParametro(?)}")
    MensagemSplash "Inicializando parametros do sistema..."
    Load Principal
    
    On Error GoTo ErroMain

    Unload SplashScreen

    With AguardarRobo
        .Show vbModal, Principal
        If .Cancelou Then
            Unload AguardarRobo
            Unload Principal
            End
        End If
        Unload AguardarRobo
    End With

    ''''''''''''''''''''''''''''''''''''''''
    ' Leitura da tabela parametros '
    ''''''''''''''''''''''''''''''''''''''''
    With Geral.qryLeituraParametro
        .rdoParameters(0) = Geral.DataProcessamento
        Set tb1 = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    If tb1.EOF Then
        MsgBox "Não foi possível ler informações da tabela Parâmetro.", vbCritical + vbOKOnly, App.Title
        Geral.qryLeituraParametro.Close
        GoTo ErroMain
    End If

    Geral.DiretorioImagens = tb1!Dir_Imagens & "\" & Geral.DataProcessamento & "\"
    Geral.AgenciaCentral = Format(tb1!AgenciaCentral, "0000")
    Geral.Intervalo = tb1!TM_Pendente
    Geral.Atualizacao = tb1!TM_Atualizacao
    Geral.ValorChqInferior = tb1!ValorInferior
    Geral.DiretorioDados = tb1!Dir_Dados & "\"
    Geral.DiretorioTrabalho = tb1!Dir_Trabalho & "\"
    Geral.ValorMaxADCC = tb1!ValorMaxADCC
    Geral.ValorAlcadaCoord_Mal = IIf(IsNull(tb1!ValorAlcadaCoord_Mal), 0, tb1!ValorAlcadaCoord_Mal)

    Set Autentica = Nothing
    
    '''''''''''''''''''''''''''''''''''''
    ' Inicializar parametros do sistema '
    '''''''''''''''''''''''''''''''''''''
    With Geral
        .Estacao = Val(PegarOpcaoINI("Diversos", "Estacao", "1"))
        .Scanner = Val(PegarOpcaoINI("Diversos", "Scanner", "0"))
        .autenticadora = Val(PegarOpcaoINI("Diversos", "Autenticadora", "0"))
        .VIPSDLL = Val(PegarOpcaoINI("Diversos", "VipsDll", "0"))
    End With
    NumBoxes = Val(PegarOpcaoINI("Diversos", "NumBoxes", "1"))
    MaxDocBox = Val(PegarOpcaoINI("Diversos", "MaxDocBox", "200"))
    BoxDefault = Val(PegarOpcaoINI("Diversos", "BoxDefault", "0"))
    Threshold = Val(PegarOpcaoINI("Diversos", "CutBords", "50"))
    Compress = Val(PegarOpcaoINI("Diversos", "Compress_JPG", "30"))
    Resolution = Val(PegarOpcaoINI("Diversos", "Resolution", "100"))
    
    Load Principal
    
    Set Autentica = Nothing
    
    If Geral.autenticadora = 1 Then
        Set Autentica = New Autentica_IBM
    ElseIf Geral.autenticadora = 2 Then
        Set Autentica = New Autentica_Procomp
    End If
    
    '''''''''''''''''''''''
    ' Inicializar Scanner '
    '''''''''''''''''''''''
    MensagemSplash "Inicializando scanner..."
    ScannerOk = False
    
    iRet = 1
    If Geral.Scanner = escnVIPS Then
    
        If Geral.VIPSDLL = eDllProservi Then
            iRet = MC93_SetImagem(3)
            If iRet = 1 Then
              iRet = MC93_SetLeitora(3)
              If iRet = 1 Then
                iRet = MC93_SetDPI(100)
                If iRet = 1 Then
                  iRet = MC93_SetAltura(420)
                  If iRet = 1 Then
                    iRet = MC93_SetComPort(1)
                    If iRet = 1 Then
                      iRet = MC93_SetImageDirectory(Geral.DiretorioImagens)
                      If iRet = 1 Then
                        iRet = MC93_CutBords(1)
                        If iRet = 1 Then
                          iRet = MC93_Init()
                          If iRet = 1 Then
                              ScannerOk = True
                          End If
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            End If
            If Not ScannerOk Then
              MsgBox "Não foi possível inicializar a VIPS." & vbCr & "Erro: " & iRet, vbExclamation + vbOKOnly, App.Title
            End If
        ElseIf Geral.VIPSDLL = eDllNovaUBB Then 'VipsDll (Nova Versão)
            tSC_ParamDLL.BoxDefault = BoxDefault
            tSC_ParamDLL.MaxDocBox = MaxDocBox
            tSC_ParamDLL.NumBoxes = NumBoxes
            
            If InicializarVips Then
                ScannerOk = True
                bInicializou = True
            End If
            
        Else ' VipsDll do Unibanco
            VIPS_SetBoxes (NumBoxes)
            VIPS_SetMaxDocBox (MaxDocBox)
            VIPS_SetBoxDefault (BoxDefault)
            VIPS_SetCompress (Compress)
            VIPS_SetCutBords (Threshold)
            VIPS_SetCameraFile ("Doc100.cpf")
            VIPS_SetImageDirectory (Geral.DiretorioImagens)
            VIPS_SetResolution (Resolution)
            iRet = VIPS_Init()
            If iRet <> 0 Then
                MsgBox "Não foi possível inicializar a VIPS." & vbCr & "Erro: " & iRet, vbExclamation + vbOKOnly, App.Title
            Else
                ScannerOk = True
            End If
        End If

    ElseIf Geral.Scanner = escnCanonLS500 Then
      ' Inicializção da LS500 e Canon
      iRet = LS_ProcuraLS500(string1, string2, string3)
      If iRet = 0 Then
          MsgBox "Não foi possível localizar o scanner.", vbExclamation + vbOKOnly, App.Title
      Else
          iRet = LS_SetNumGauges(1)
          iRet = LS_Lapso(30)           'SCSI antiga/nova
          iRet = LS_SetSepara(0)        '1- separa
                                        '0- não separa
          iRet = LS_SetTimeOut(500)     '1/2 segundo
          iRet = LS_SetImage(3)         '(1) digitaliza só frente
                                        '(2) digitaliza só verso
                                        '(3) digitaliza frente e verso
          ScannerOk = True
      End If
    End If
    
    If ScannerOk Then
      ' habilitar menu no form principal
      Principal.mnuCapCaptura(5).Enabled = True
    Else
      ' desabilitar menu no form principal
      Principal.mnuCapCaptura(5).Enabled = False
    End If
    
    ''''''''''''''''''''''''
    ' Fim da inicialização '
    ''''''''''''''''''''''''
    Principal.Show
    DoEvents
    
    Unload SplashScreen
    
    tb1.Close
    Geral.qryLeituraParametro.Close

    If Not ChecarParametros(Geral) Then
        MsgBox "Não foi possível inicializar o Sistema.", vbExclamation + vbOKOnly, App.Title
        Geral.Banco.Close
        End
    End If
    Principal.Caption = App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision & "  [" & Left(Geral.NomeUsuario, 15) & "]  [" & Format(DataDDMMAAAA(Geral.DataProcessamento), "00/00/0000") & "]" & " Ag. Proc.: [" & Geral.AgenciaCentral & "]"
    
    Principal.Refresh

    Exit Sub

ErroMain:
    If Err.Number = 53 And Err.Description = "File not found: VIPSDLL.DLL" Then
        MsgBox "Erro na tentativa de inicializar a VIPS, favor contatar o suporte.", vbCritical, App.Title
        End
    End If
    Select Case TratamentoErro("Não foi possível inicializar o Sistema.", Err, rdoErrors)
        Case vbCancel
            End
        Case vbRetry
            Resume
    End Select

End Sub

Public Function ChecarParametros(ByRef pvrParametro As tpGlobais) As Boolean
    Dim bRet As Boolean
    
    bRet = True
    
    If Not ChecarDiretorio(pvrParametro.DiretorioDados, "Diretório de Dados não existe!") Then
        bRet = False
    ElseIf Not ChecarDiretorio(pvrParametro.DiretorioImagens, "Diretório de Imagens não existe!") Then
        bRet = False
    ElseIf Not ChecarDiretorio(pvrParametro.DiretorioTrabalho, "Diretório de Trabalho não existe!") Then
        bRet = False
    ElseIf (pvrParametro.Scanner < 0 Or pvrParametro.Scanner > 2) And pvrParametro.Scanner <> escnDummy Then
        MsgBox "Seleção de scanner inválida!", vbExclamation + vbOKOnly, App.Title
        bRet = False
    End If
    ChecarParametros = bRet
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''''
' Inibir a digitação de letras em campos numéricos '
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InibirTeclaAlfa(ByRef priTecla As Integer, Optional ByVal pvbSalta As Boolean)
    If (priTecla < &H30 Or priTecla > &H39) And priTecla <> &H8 And priTecla <> &HD And priTecla <> &H1B Or priTecla = 46 Then
        priTecla = 0
    ElseIf priTecla = &HD Then
        If pvbSalta Then
            SendKeys "{TAB}"
        End If
    End If
End Sub

Sub MensagemSplash(ByVal pvsTexto As String)
    SplashScreen.Mensagem.Caption = pvsTexto
    DoEvents
End Sub

Public Function AntigoTratamentoErro(ByVal pvsTexto As String, pvoErro As ErrObject, ByRef pvoRDOErrors As rdoErrors) As VbMsgBoxResult
    Dim sMens As String
    Dim sErro As String
    Dim sRdo As String
    Dim oErro As rdoError
    Dim Retorno As VbMsgBoxResult
    
    If pvoErro.Number <> 0 And InStr(pvoErro.Description, "ODBC") = 0 And InStr(pvoErro.Description, "SQL") = 0 Then
        sErro = pvoErro.Number & " - " & pvoErro.Description
    Else
        sErro = ""
    End If
    
    sRdo = ""
    
    For Each oErro In pvoRDOErrors
        With oErro
'            If .Number = 1205 Then
'                TratamentoErro = vbRetry
'                Sleep 500
'                Exit Function
'            ElseIf .SQLState <> "01000" And .Source = "ODBC" Then
                sRdo = sRdo & .Number & " - " & ChecarRDOError(.Description) & " (SQL State=" & .SQLState & ")" & vbCr
'            End If
        End With
    Next
    
    rdoErrors.Clear
    
    Load MensagemErro

    With MensagemErro
        .Texto = Trim(pvsTexto)
        .Erro = Trim(sErro)
        .ErroBanco = Trim(sRdo)
        .Mostrar
        .Show vbModal, Principal

        If .Retorno = 0 Then
            AntigoTratamentoErro = vbRetry
        Else
            AntigoTratamentoErro = vbCancel
        End If
    End With

    Unload MensagemErro
End Function
Private Function TratarStringErro(ByVal pvsTexto As String) As String
    Dim i As Long
    Dim sAux As String
    
    sAux = ""
    For i = 1 To Len(pvsTexto)
        If Mid(pvsTexto, i, 1) <> "'" Then
            sAux = sAux & Mid(pvsTexto, i, 1)
        End If
    Next
    
    TratarStringErro = sAux
End Function

Public Function TratamentoErro(ByVal pvsTexto As String, pvoErro As ErrObject, ByRef pvoRDOErrors As rdoErrors, Optional pvbMostrar As Boolean = True) As VbMsgBoxResult
    Dim sMens As String
    Dim sErro As String
    Dim sRdo As String
    Dim oErro As rdoError
    Dim Retorno As VbMsgBoxResult
    
    GravarErro pvsDescricao:=pvsTexto
    
    If pvoErro.Number <> 0 And InStr(pvoErro.Description, "ODBC") = 0 And InStr(pvoErro.Description, "SQL") = 0 Then
        sErro = pvoErro.Number & " - " & pvoErro.Description
        GravarErro pvoErro.Number, pvoErro.Description
    Else
        sErro = ""
    End If
    
    sRdo = ""
    
    For Each oErro In pvoRDOErrors
        With oErro
            GravarErro .Number, .Description
            If .Number = 1205 Then
                TratamentoErro = vbRetry
                Sleep 500
                Exit Function
            'ElseIf .SQLState <> "01000" And .Source = "ODBC" Then
             Else
                sRdo = sRdo & .Number & " - " & ChecarRDOError(.Description) & " (SQL State=" & .SQLState & ")" & vbCr
            End If
        End With
    Next
    
    rdoErrors.Clear
    
    If pvbMostrar Then
        Load MensagemErro
        
        With MensagemErro
            .Texto = Trim(pvsTexto)
            .Erro = Trim(sErro)
            .ErroBanco = Trim(sRdo)
            .Mostrar
            .Show vbModal, Principal
            
            If .Retorno = 0 Then
                TratamentoErro = vbRetry
            Else
                TratamentoErro = vbCancel
            End If
        End With
        
        Unload MensagemErro
    Else
        TratamentoErro = vbCancel
    End If

End Function

Private Sub GravarErro(Optional ByVal pviErro As Long = 0, Optional ByVal pvsDescricao As String = "")
    Dim sSql As String
    Dim qryInsereLogErro As rdoQuery
    
    On Error Resume Next
           
    Set qryInsereLogErro = Geral.Banco.CreateQuery("", "{ call InsereLogErro( ?,?,?,?,? ) }")
    With qryInsereLogErro
        .rdoParameters(0) = Geral.DataProcessamento
        .rdoParameters(1) = Geral.Estacao
        .rdoParameters(2) = Geral.Usuario
        .rdoParameters(3) = pviErro
        .rdoParameters(4) = TratarStringErro(pvsDescricao)
        .Execute
        .Close
    End With
    
End Sub

Public Function ChecarRDOError(ByVal pvsTexto As String) As String
    Dim sTexto As String
    Dim i As Long
    
    sTexto = ""
    
    For i = Len(pvsTexto) To 1 Step -1
        If Mid(pvsTexto, i, 1) = "]" Then
            Exit For
        End If
        
        sTexto = Mid(pvsTexto, i, 1) & sTexto
    Next
    
    ChecarRDOError = sTexto
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

Public Function InserePonto(ByVal Valor As String) As String
    Valor = Format(Valor, "000")
    If Val(Valor) = 0 Then
        InserePonto = "0.00"
    Else
        InserePonto = Left(Valor, Len(Valor) - 2) & "." & Right(Valor, 2)
    End If
End Function

Public Function RPad(ByVal str As String, ByVal Tam As Integer) As String
    If str = "" Then
        RPad = String(Tam, " ")
    Else
        If Len(str) >= Tam Then
            RPad = Mid(str, 1, Tam)
        Else
            RPad = str & Space(Tam - Len(str))
        End If
    End If
End Function

Public Function LPad(ByVal str As String, ByVal Tam As Integer) As String
    If str = "" Then
        LPad = String(Tam, " ")
    Else
        If Len(str) >= Tam Then
            LPad = Mid(str, 1, Tam)
        Else
            LPad = Space(Tam - Len(str)) & str
        End If
    End If
End Function

Public Function FormataValor(ByVal Valor As Currency, ByVal Tam As Integer) As String
    Dim strValor As String
    Dim strDecimal As String
    Dim strInteiro As String
    Dim strResult As String
    Dim Count As Integer
    
    strValor = Trim(str(Valor))
    If InStr(1, strValor, ".", 1) = 0 And InStr(1, strValor, ",", 1) = 0 Then
        strInteiro = strValor
        strDecimal = "00"
    Else
        Count = 1
        While Mid(strValor, Count, 1) <> "," And Mid(strValor, Count, 1) <> "."
            strInteiro = strInteiro & Mid(strValor, Count, 1)
            Count = Count + 1
        Wend
        strDecimal = Mid(strValor, Count + 1, 2)
        If Len(strDecimal) = 1 Then
            strDecimal = strDecimal & "0"
        End If
    End If
    
    For Count = 1 To Len(strInteiro)
        strResult = Mid(strInteiro, Len(strInteiro) - Count + 1, 1) & strResult
        If (Count Mod 3 = 0) And (Count < Len(strInteiro)) Then
            If Mid(strInteiro, Len(strInteiro) - Count, 1) <> "-" Then
                strResult = "." & strResult
            End If
        End If
    Next
    If Len(strResult) = 0 Then
        strResult = "0"
    End If
    strResult = strResult & "," & strDecimal
    strResult = LPad(strResult, Tam)
    FormataValor = strResult
End Function

Public Function FormataConta(ByVal Conta As Long) As String
    Dim strConta As String
    
    strConta = Format(Conta, "0000000")
    FormataConta = Mid(strConta, 1, 3) & "." & Mid(strConta, 4, 3) & "-" & Right(strConta, 1)
End Function

Public Sub GravaLog(ByVal IdCapa As Long, _
                         ByVal IdDocto As Long, _
                         ByVal Acao As Integer)
    
    Dim qryInserirLog As rdoQuery

    On Error GoTo ErroGravaLog

    Set qryInserirLog = Geral.Banco.CreateQuery("", "{call InsereLog (?,?,?,?,?)}")
 
    With qryInserirLog
        .rdoParameters(0) = Geral.DataProcessamento
        .rdoParameters(1) = IdCapa
        .rdoParameters(2) = IdDocto
        .rdoParameters(3) = Geral.Usuario
        .rdoParameters(4) = Acao
        .Execute
    End With

    Set qryInserirLog = Nothing

    Exit Sub

ErroGravaLog:
  Select Case TratamentoErro("Erro ao Gravar Log do Usuário.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub

Public Function ValidaCodigoBanco(ByVal sCodigo As String) As Boolean

  On Error GoTo ERRO_VALIDACODIGOBANCO

  Dim RsBanco As rdoResultset
  Dim qryGetTFSBanco As rdoQuery

  ValidaCodigoBanco = False

  'Pesquisar na tabela TFSBanco
  Set qryGetTFSBanco = Geral.Banco.CreateQuery("", "{call GetTFSBanco (" & sCodigo & ")}")

  Set RsBanco = qryGetTFSBanco.OpenResultset(rdOpenStatic, rdConcurReadOnly)
  If Not RsBanco.EOF Then
    'Encontrou o Código do Banco
    ValidaCodigoBanco = True
  End If

  Exit Function

ERRO_VALIDACODIGOBANCO:
  Select Case TratamentoErro("Erro ao Pesquisar Código de Banco.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function

'limpa todos os campos da tabela
Public Sub LimpaTela(Janela As Form)
  Dim i As Long
  Dim Ctrl As Object
  Dim Mask As String
  Dim campo As String
  On Error Resume Next
  For Each Ctrl In Janela.Controls
   'If TypeOf Ctrl Is DataCombo Then Ctrl.BoundText = ""
   'If TypeOf Ctrl Is DataCombo Then Ctrl = ""
   If (TypeOf Ctrl Is TextBox) Then Ctrl.Text = ""
   If (TypeOf Ctrl Is CurrencyEdit) Then Ctrl.Text = ""
   'If (TypeOf Ctrl Is MaskEdBox) Then
   ' Mask = Ctrl.Mask
   ' Ctrl.Mask = ""
   ' Ctrl.Text = ""
   ' Ctrl.Mask = Mask
   'End If
   If TypeOf Ctrl Is ComboBox Then Ctrl.ListIndex = -1
   If TypeOf Ctrl Is CheckBox Then Ctrl.Value = Unchecked
   'If TypeOf Ctrl Is DTPicker Then Ctrl.Value = Date
  Next
End Sub

Function ValidaAgencia(ByVal CodigoAgencia As Integer, ByVal sVencimento As String, ByVal ValidaData As Boolean, Optional ByVal bCarregaGeralCapa As Boolean = False) As Integer

' Parâmetro:    bCarregaGeralCapa - Quando (TRUE) será carregado os campos no type GERAL.CAPA,
'                                   esta opção será utilizada na complementação de Malote/Envelope

  Dim RsAgenf As rdoResultset
  Dim qryGetAgenf As rdoQuery

  'Código de Retorno
  '0 - Data de Vencimento OK
  '1 - Documento Vencido
  '2 - Agencia em Feriado
  '3 - Agencia Fechada
  '4 - Agencia não cadastrada
  '5 - Data não Verificada

  ValidaAgencia = 5

  'Verificar o Status da Agencia
  Set qryGetAgenf = Geral.Banco.CreateQuery("", "{call GetAgenf (" & CodigoAgencia & ")}")

  Set RsAgenf = qryGetAgenf.OpenResultset(rdOpenStatic, rdConcurReadOnly)
  If Not RsAgenf.EOF Then
    'A Agencia está cadastrada -> Verificar o Status
    If RsAgenf!agefsstmovi = 9 Then
      'Feriado
      ValidaAgencia = 2
      Exit Function

      'MsgBox "A Agência de Origem está em Feriado.", vbInformation, App.Title
    ElseIf RsAgenf!agefsstmovi = 0 Then
      'Agencia Fechada
      ValidaAgencia = 3
      Exit Function

      'MsgBox "A Agência de Origem está Fechada.", vbInformation, App.Title
    ElseIf RsAgenf!agefsstmovi = 2 Then
      'Agencia Aberta -> Verificar data do Movimento Anterior
      If ValidaData Then
        If DataAAAAMMDD(sVencimento) <= TransformaDataAAAAMMDD(RsAgenf!agefsdtmvan) Then
          'A Data de Vencimento é menor ou igual à data do Movimento Anterior -> Não Aceitar
          ValidaAgencia = 1
          Exit Function

          'MsgBox "A Data de Vencimento deve ser maior que a Data do Movimento Anterior.", vbInformation, App.Title
        End If
      End If
    End If
  Else
    ValidaAgencia = 4
    Exit Function

    'MsgBox "A Agência de Origem não está Cadastrada.", vbInformation, App.Title
  End If

  ValidaAgencia = 0
  
  'Verifica se Carrega dados da AGENF em type Geral.capa
  If bCarregaGeralCapa Then
        With Geral.Capa
            .agefsdtmvan = RsAgenf!agefsdtmvan
            .agefsdtmvat = RsAgenf!agefsdtmvat
            .agefsestado = RsAgenf!agefsestado
            .agefsstmovi = RsAgenf!agefsstmovi
        End With
  End If
  
End Function
Function CarregaAGENF(ByVal CodigoAgencia As Integer) As Boolean
  
Dim RsAgenf As rdoResultset
Dim qryGetAgenf As rdoQuery
On Error GoTo err_CarregaAGENF

    CarregaAGENF = False
    
    'Verificar o Status da Agencia
    Set qryGetAgenf = Geral.Banco.CreateQuery("", "{call GetAgenf (" & CodigoAgencia & ")}")
    
    Set RsAgenf = qryGetAgenf.OpenResultset(rdOpenStatic, rdConcurReadOnly)
  
    If RsAgenf.EOF Then Exit Function
    
    With Geral.Capa
        .agefsdtmvan = RsAgenf!agefsdtmvan
        .agefsdtmvat = RsAgenf!agefsdtmvat
        .agefsestado = RsAgenf!agefsestado
        .agefsstmovi = RsAgenf!agefsstmovi
    End With

    Set RsAgenf = Nothing
    CarregaAGENF = True
    Exit Function

err_CarregaAGENF:
    Set RsAgenf = Nothing
    MsgBox "Não foi possível ler informações da agência (AGENF)", vbCritical + vbOKOnly, App.Title
    
End Function
Function TransformaDataAAAAMMDD(ByVal sData As String) As Long

  'Formata a data para 6 bytes
  sData = Format(sData, "000000")

  'Acrescenta o século
  If Val(Right(sData, 2)) > 50 Then
    sData = Mid(sData, 1, 4) & "19" & Mid(sData, 5, 2)
  Else
    sData = Mid(sData, 1, 4) & "20" & Mid(sData, 5, 2)
  End If

  'Formatar para AAAAMMDD
  TransformaDataAAAAMMDD = Mid(sData, 5, 4) & Mid(sData, 3, 2) & Mid(sData, 1, 2)
End Function
Function VerificaPreenchimentoCMC7(verForm As Variant) As Boolean

    VerificaPreenchimentoCMC7 = False

    With verForm
        
        'Valida CMC7 - campo 1
        If .txtCMC71.MaxLength <> Len(Trim(.txtCMC71.Text)) Then
            MsgBox "CMC7 inválido.", vbInformation + vbOKOnly, App.Title
           .txtCMC71.SetFocus
            Exit Function
        End If
         
        'Valida CMC7 - campo 2
        If .txtCMC72.MaxLength <> Len(Trim(.txtCMC72.Text)) Then
            MsgBox "CMC7 inválido.", vbInformation + vbOKOnly, App.Title
           .txtCMC72.SetFocus
            Exit Function
        End If
        
        'Valida CMC7 - campo 3
        If .txtCMC73.MaxLength <> Len(Trim(.txtCMC73.Text)) Then
            MsgBox "CMC7 inválido.", vbInformation + vbOKOnly, App.Title
           .txtCMC73.SetFocus
            Exit Function
        End If
        
    End With

    VerificaPreenchimentoCMC7 = True

End Function
Public Function VerificaUsuario(ByRef pTbUsuario As rdoResultset, ByVal pUsuario As String, ByVal pSenha As String, ByVal pBaseBackup As Boolean) As enumRetornoUsuario
Dim Control As Control

    ' ***********************************
    ' * Verificação do Login do Usuário *
    ' ***********************************
    
    
    If UCase(Trim(pUsuario)) = "DESENV" And UCase(Trim(pSenha)) = "DIAMANTE" Then

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' * Se for Usuário for = desenv habilita todos Menus * '
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For Each Control In Principal.Controls
            If TypeName(Control) = "Menu" Then
                Control.Enabled = True
            End If
        Next
        
        ''''''''''''''''''''''''''''''
        'Retorno de usuario supervisor'
        ''''''''''''''''''''''''''''''
        VerificaUsuario = eSUPERVISOR
        
    ElseIf pTbUsuario.EOF Then
        VerificaUsuario = eNAO_EXISTENTE
    ElseIf UCase(Decript(Trim(pTbUsuario!Senha))) <> UCase(Trim(pSenha)) Then
        VerificaUsuario = eSENHA_INCORRETA
    Else
        VerificaUsuario = eOK
    End If

End Function
Public Function VerificaAcessoUsuario(ByVal pIdUsuario As Long, ByVal pIdModulo As Integer) As Boolean

    Dim qryGetModuloUsuario As rdoQuery
    Dim rsUsuario As rdoResultset

    On Error GoTo VerificaAcessoUsuario_Err

    VerificaAcessoUsuario = False

    Set qryGetModuloUsuario = Geral.Banco.CreateQuery("", "{Call GetModuloUsuario(?,?)}")

    qryGetModuloUsuario.rdoParameters(0).Value = pIdUsuario      'IdUsuario passado por parametro
    qryGetModuloUsuario.rdoParameters(1).Value = pIdModulo       'Modulo

    Set rsUsuario = qryGetModuloUsuario.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

    If Not rsUsuario.EOF Then
        If rsUsuario!Total = 0 Then Exit Function
    End If

    VerificaAcessoUsuario = True

    Exit Function

VerificaAcessoUsuario_Err:
    Select Case TratamentoErro("Erro ao verificar se usuário tem permissão ao módulo.", Err, rdoErrors)
        Case vbCancel, vbRetry
    End Select
End Function
Public Function DataDD_MM_AAAA(ByVal pviData As Long) As String
    Dim sData As String
    
    sData = CStr(pviData)
    
    DataDD_MM_AAAA = Right(sData, 2) & "/" & Mid(sData, 5, 2) & "/" & Left(sData, 4)
End Function

Public Function Formata(vVariavel As Variant, sTipoConversao As String) As String

    If vVariavel = 0 Then
        Formata = "-"
    Else
        If sTipoConversao = "I" Then
            Formata = Format(vVariavel, "###,###,##0")
        Else
            Formata = Format(vVariavel, "##0")
        End If
    End If
    
End Function

Public Function CarregaAgenciaColetaEmCombo(objNomeCombo As Object) As Boolean

Dim qryLerAgenf     As rdoQuery
Dim RsLerAgenf      As rdoResultset

On Error GoTo Err_CarregaAgenciaColeta
    
    CarregaAgenciaColetaEmCombo = False
    
    Set qryLerAgenf = Geral.Banco.CreateQuery("", "{call GetAllAgenf}")
    Set RsLerAgenf = qryLerAgenf.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

    If RsLerAgenf.EOF Then
        Beep
        MsgBox "Não existem agências de coleta cadastradas, favor verificar!", vbInformation, App.Title
        GoTo Exit_CarregaAgenciaColeta
    End If
    
    objNomeCombo.Clear
    objNomeCombo.AddItem " Todas"
    objNomeCombo.ItemData(0) = 0
    
    While Not RsLerAgenf.EOF
            objNomeCombo.AddItem Format(RsLerAgenf(0).Value, "0000") & " - " & RsLerAgenf(1).Value
            objNomeCombo.ItemData(objNomeCombo.NewIndex) = RsLerAgenf(0).Value
            RsLerAgenf.MoveNext
    Wend
    'Posiciona em todas agências
    objNomeCombo.ListIndex = 0
    
    CarregaAgenciaColetaEmCombo = True

Exit_CarregaAgenciaColeta:
    qryLerAgenf.Close
    If Not (RsLerAgenf Is Nothing) Then Set RsLerAgenf = Nothing
    Exit Function

Err_CarregaAgenciaColeta:
    Beep
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, App.Title
    GoTo Exit_CarregaAgenciaColeta

End Function

Public Sub Auditoria()

On Error GoTo Err_Auditoria
    
    FrmLogUsuario.Show vbModal

    Exit Sub
    
Err_Auditoria:
    MsgBox "Não foi possível abrir a tela de auditoria !", vbExclamation + vbOKOnly, App.Title

End Sub

Public Function G_EncriptaBO(ByVal intTipoDocto As Integer, ByVal strCampoChave As String, Optional ByVal lngIdDocto As Long = 0) As String
'-----------------------------------------------------------------------------------------------
'   DLL:        Encripta
'
'   Parâmetros: intTipoDocto    - Tipo do documento.
'               strCampoChave   - Informação do documento para transformação da chave de encriptação.
'               lngIdDocto      - IdDocto do documento (Opcional), Se Informado Gravar Autenticação Digital
'                                 na tabela documento.
'
'   Retorno:    String encriptografada
'   Regra:      Segue abaixo relação do Documento, TipoDocto e Campos para formação da chave;
'               ADCC(4)                         TipoDocto + Conta
'               Ajuste Credito(32,34)           TipoDocto + Conta
'               Ajuste Debito(33,38)            TipoDocto + Conta
'               Arrec.Convencional(27)          TipoDocto + Produto
'               Arrec.Eletrônica(20,21,22,23)   TipoDocto + Subs(CdBarras,1,4) + Subs(CdBarras,11,3)
'               Cartão Avulso(36)               TipoDocto + Left(Cartão,7)
'               CBIndex(24,25,26)               TipoDocto + Subs(CdBarras,1,4) + Subs(CdBarras,11,3)
'               Cheque(5,6,7)                   TipoDocto + Subs(CMC7,23,7)
'               Cobrança Especial(14)           TipoDocto + Cedente
'               Cobrança Registrada(13)         TipoDocto + Subs(NossoNumero,1,7)
'               DarfPreto(16)                   TipoDocto + Subs(CPFCGC,8,7)
'               DarfSimples(17)                 TipoDocto + Subs(CPFCGC,8,7)
'               Darm(15)                        TipoDocto + Subs(CCM,1,7)
'               Depósito(2,3)                   TipoDocto + Conta
'               FGTS(40)                        TipoDocto + Left(CNPJ_Empresa,7)
'               Ficha Compensação(28,29,30,31)  TipoDocto + Subs(CdBarras,1,7)
'               GARE(18)                        TipoDocto + ( Se CPFCGC <> nulo então Left(CPFCGC,7) senão Left(InscricaoEstadual,7) )
'               GPS(35)                         TipoDocto + Subs(Identificador,1,7)
'               Lançamento Interno(41)          TipoDocto + Subs(ControleBanco,1,7)
'               OCT(37)                         TipoDocto + ContaCredito
'               Título(12)                      TipoDocto + Banco
'-----------------------------------------------------------------------------------------------

On Error GoTo Err_EncriptaBO

Dim strEncripta As String
Dim strChave    As String

    strEncripta = Space(16)
    G_EncriptaBO = ""
    
    'Verifica se Informado Tipo do Documento
    If intTipoDocto <= 0 Then GoTo Err_MontaChave
    
    strCampoChave = CStr(Trim(strCampoChave))
    
    Select Case intTipoDocto
        Case 2, 3, 4, 12, 14, 27, 32, 33, 34, 37, 38   'Diversos
            strChave = Right(String(8, "0") & CStr(intTipoDocto) & strCampoChave, 8)
        
        Case 20, 21, 22, 23, 24, 25, 26 'Arrec.eletrônica/ CBIndex
            strChave = Right(String(8, "0") & CStr(intTipoDocto) & Mid(strCampoChave, 1, 4) & Mid(strCampoChave, 11, 3), 8)
        
        Case 5, 6, 7            'Cheque (Força tipodocto =5)
            strChave = Right(String(8, "0") & CStr(5) & Mid(strCampoChave, 23, 7), 8)
        
        Case 13, 15, 28, 29, 30, 31, 35, 41     'Diversos
            strChave = Right(String(8, "0") & CStr(intTipoDocto) & Mid(strCampoChave, 1, 7), 8)
        
        Case 16, 17, 18         'DARF PRETO/SIMPLES e GARE
            strChave = Right(String(8, "0") & CStr(intTipoDocto) & Left(strCampoChave, 7), 8)
        
        Case 36, 40                 'FGTS / Cartão Avulso
            strChave = Right(String(8, "0") & CStr(intTipoDocto) & Left(strCampoChave, 7), 8)
    End Select
    
    If strChave = "" Then GoTo Err_MontaChave
    
    Encripta CLng(Val((strChave))), strEncripta
    
    'Verifica se grava em documento a Autenticação Digital
    If lngIdDocto > 0 Then
        If Not G_GravaEncriptarBO(lngIdDocto, strEncripta) Then GoTo Exit_EncriptaBO
    End If
    
    G_EncriptaBO = strEncripta
    

Exit_EncriptaBO:
    Exit Function

Err_MontaChave:
    MsgBox "Chave para Autenticação Digital fora do padrão, favor verificar!", vbCritical + vbOKOnly, App.Title
    GoTo Exit_EncriptaBO
    
Err_EncriptaBO:
    MsgBox Err.Description & Err.Source, vbCritical + vbOKOnly, App.Title
    GoTo Exit_EncriptaBO
    
End Function


Public Function G_GravaEncriptarBO(ByVal lngIdDocto As Long, ByVal strAutenticacao As String) As Boolean
    
Dim qryEncripta As rdoQuery

On Error GoTo Err_GravaEncriptarBO

    G_GravaEncriptarBO = False
    
    Set qryEncripta = Geral.Banco.CreateQuery("", "{? = call GravaAutenticacaoDigital (?,?,?)}")
 
    With qryEncripta
        .rdoParameters(0).Direction = rdParamReturnValue
        
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = lngIdDocto
        .rdoParameters(3) = strAutenticacao
        .Execute
    
        If .rdoParameters(0).Value <> 0 Then GoTo Err_GravaEncriptarBO
    End With

    G_GravaEncriptarBO = True
    
Exit_GravaEncriptarBO:
    Set qryEncripta = Nothing
    Exit Function
    
Err_GravaEncriptarBO:
    MsgBox "Erro ao tentar gravar Autenticação Digital no Documento", vbCritical + vbOKOnly, App.Title
    GoTo Exit_GravaEncriptarBO

End Function
Public Function GravaComplementoOcorrencia(ByVal lngIdDocto As Long, ByVal strAcao As String, ByRef strDescricao As String) As Boolean
' strAcao   -   (G) Inserir / Alterar
'               (C) Consultar
'               (E) Excluir

Dim qryComplemento As rdoQuery
Dim rsComplemento As rdoResultset

On Error GoTo Err_GravaComplementoOcorrencia
    
    GravaComplementoOcorrencia = False

    Set qryComplemento = Geral.Banco.CreateQuery("", "{? = call GravaComplementoOcorrencia(?,?,?,?) }")
    With qryComplemento
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = lngIdDocto
        .rdoParameters(3) = strAcao
        .rdoParameters(4) = strDescricao
        Set rsComplemento = .OpenResultset(rdOpenStatic, rdConcurReadOnly)
        
        If .rdoParameters(0).Value <> 0 Then GoTo Err_GravaComplementoOcorrencia
        'Verifica se existe retorno para ação de Consulta
        If strAcao = "C" Then
            If Not rsComplemento.EOF Then strDescricao = Trim(rsComplemento!Descricao & "")
        End If
        
    End With

    GravaComplementoOcorrencia = True

Exit_GravaComplementoOcorrencia:
    Set rsComplemento = Nothing
    Set qryComplemento = Nothing
    Exit Function
    
Err_GravaComplementoOcorrencia:
    MsgBox "Erro na atualização do Complemento de Ocorrência", vbCritical, App.Title
    GoTo Exit_GravaComplementoOcorrencia

End Function

Public Function ValidaAgenciaPorDocto(ByVal CodigoAgencia As Integer, ByVal sVencimento As String, ByVal ValidaData As Boolean, Optional ByVal bExibeMsgDoctoVencido As Boolean = False) As Boolean
'---------------------------------------------------------------------------------------------
'   Função de verificação da agência (AGENF) apenas para determinados forms
'
'     Parâmetros:
'       CodigoAgencia:          Entra como opcional caso a variavel Geral.Capa.AgOrig
'                               não contenha informação
'                               Obs.: Geral.Capa.AgOrig é carregada na complementação de capa
'
'       sVencimento:            Data de vencto à ser consistida no formato (DDMMAAAA)
'
'       ValidaData:             Identificador de consitência da Data de
'                               Vencto (sVencimento qdo informada)
'
'       bExibeMsgDoctoVencido:  A verificação de docto vencido está liberado para
'                               alguns dos doctos, salvos excessão de alguns onde será
'                               forçada apresentação da mensagem através deste parâmetro
'
'--------------------------------------------------------------------------------------------

Dim RetAgencia As Integer
    
ValidaAgenciaPorDocto = False

    'Validar Agencia
    If Geral.Capa.agefsestado = "" Then
        'Se não existe informação da agência carregada, verificar (ValidaAgencia)
        If ValidaData Then
            RetAgencia = ValidaAgencia(CodigoAgencia, sVencimento, True)
        Else
            RetAgencia = ValidaAgencia(CodigoAgencia, "", False)
        End If
        
        If RetAgencia = 0 Then ValidaAgenciaPorDocto = True
        'Dispensa verificação com mensagem abaixo devido função (ValidaAgencia) conter tratamento
        Exit Function

    Else
        If Geral.Capa.agefsstmovi = 9 Then
              RetAgencia = 2
        ElseIf Geral.Capa.agefsstmovi = 0 Then
              RetAgencia = 3
        ElseIf Geral.Capa.agefsstmovi = 2 Then
            'Agencia Aberta -> Verificar data do Movimento Anterior
            If ValidaData Then
                If DataAAAAMMDD(sVencimento) <= TransformaDataAAAAMMDD(Geral.Capa.agefsdtmvan) Then
                    'A Data de Vencimento é menor ou igual à data do Movimento Anterior -> Não Aceitar
                    RetAgencia = 1
                Else
                    RetAgencia = 0
                End If
            Else
                RetAgencia = 0
            End If
        Else
            RetAgencia = 0
        End If
    End If

    'Verificar Retorno da Função
    Select Case RetAgencia
        '08/05/2001''''''''''''''''''''''''''''''''''''''''
        'Pode aceitar docto vencido - Cobrança Registrada
        '                           - Cobrança Especial
        '''''''''''''''''''''''''''''''''''''''''''''''''''
      Case 1
        'Documento Vencido
        If bExibeMsgDoctoVencido Then
            MsgBox "Documento vencido não aceito na regra de caixa expresso.", vbInformation + vbOKOnly, App.Title
            Exit Function
        End If

'        'Documento Vencido
'        If Geral.capa.IdEnv_Mal = "E" Then
'            'Envelope -> Não Aceitar
'            MsgBox "Documento vencido não aceito na regra de caixa expresso.", vbInformation + vbOKOnly, App.Title
'            TxtVencimento.SetFocus
'            Exit Function
'        ElseIf Geral.capa.IdEnv_Mal = "M" Then
'            'Malote -> Pedir Confirmação
'            If MsgBox("Este documento pertence a um Malote e está vencido. Confirma ?", vbYesNo + vbInformation, App.Title) = vbNo Then
'              TxtVencimento.SetFocus
'              Exit Function
'            End If
'        Else
'            'Tipo Indefinido
'            MsgBox "Não foi possível definir se o documento pertence a um Envelope ou Malote " & Chr(13) & _
'            "para aplicar regra de validação de Data de Vencimento.", vbInformation + vbOKOnly, App.Title
'            Exit Function
'        End If
      Case 2
        'Agencia em Feriado
        MsgBox "A agência de origem está em feriado.", vbInformation + vbOKOnly, App.Title
        Exit Function
      Case 3
        'Agencia Fechada
        MsgBox "A agência de origem está fechada.", vbInformation + vbOKOnly, App.Title
        Exit Function
      Case 4
        'Agencia não Cadastrada
        MsgBox "A agência de origem não está cadastrada.", vbInformation + vbOKOnly, App.Title
        Exit Function
    End Select

ValidaAgenciaPorDocto = True

End Function
Public Sub InicializaDLLsVIPS()

    iVIPSDLL = LoadLibrary("VIPSDLL.DLL")
    iVIPSGRIM = LoadLibrary("VIPSGRIM.DLL")
    iVIPSDRV = LoadLibrary("VIPSDRV.DLL")
    iVIPSSERIE = LoadLibrary("VIPSSERIE.DLL")
    iVIPSPROD = LoadLibrary("VIPSPROD.DLL")
    iVIPSCODE = LoadLibrary("VIPSCODE.DLL")
    iVIPSXPGMK = LoadLibrary("VIPSXPGMK.DLL")

End Sub

Public Sub FinalizaDLLsVIPS()

Dim iRet As Long

    If Geral.VIPSDLL = eDllNovaUBB Then
        If bInicializou = True Then
            iRet = SC_DeInit()
            bInicializou = False
        End If
    Else
        If iVIPSXPGMK <> 0 Then FreeLibrary (iVIPSXPGMK)
        If iVIPSCODE <> 0 Then FreeLibrary (iVIPSCODE)
        If iVIPSPROD <> 0 Then FreeLibrary (iVIPSPROD)
        If iVIPSSERIE <> 0 Then FreeLibrary (iVIPSSERIE)
        If iVIPSDRV <> 0 Then FreeLibrary (iVIPSDRV)
        If iVIPSGRIM <> 0 Then FreeLibrary (iVIPSGRIM)
        If iVIPSDLL <> 0 Then FreeLibrary (iVIPSDLL)
    End If

End Sub

