Attribute VB_Name = "GLOBAIS"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''
' Definição do tipo de variáveis globais '
''''''''''''''''''''''''''''''''''''''''''

Type tpUsuario
    Login                       As String * 10
    Nome                        As String * 30
End Type

Type tpCDR
    Drive                       As String * 3
    DiretorioImagens            As String
    DiretorioDados              As String
End Type

Type TpCapa
    AgOrig As Integer
    Capa As Double
    Duplicidade As Integer
    IdCapa As Long
    IdLote As Long
    IdEnv_Mal As String
    Num_Malote As Double
    Status As String
End Type
 
Type tpGlobais
    Autenticadora               As enumAutentica
    Atualizacao                 As Integer          'usado no timer de atualizacao dos forms
    AgenciaCentral              As String
    AgenciaApresentante         As String
    Banco                       As rdoConnection
    BancoCaixa                  As rdoConnection
    CDR                         As tpCDR
    Capa                        As TpCapa
    DataProcessamento           As Long
    DiretorioDados              As String
    DiretorioImagens            As String
    DiretorioTrabalho           As String
    Intervalo                   As Integer          'usado no timer para atualizar DataAtual da capa
    qryLeituraParametro         As rdoQuery
    StringConexao               As String
    Scanner                     As enumScanner
    Usuario                     As tpUsuario
End Type


'''''''''''''''''''''''''''''''''''''''''
' Estrutura do arquivo de Retorno VIPs  '
'''''''''''''''''''''''''''''''''''''''''
Type tpRetornoVips
    Tipo        As String * 1
    Leitura     As String * 63
    Frente      As String * 19
    Verso       As String * 19
    origem      As String * 1
    CrLf        As String * 2
End Type

'''''''''''''''''''''''''''''
' Principal variável global '
'''''''''''''''''''''''''''''
Public Geral As tpGlobais

'''''''''''''''''''''''
' Tratar Arquivos INI '
'''''''''''''''''''''''
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

Public Const FO_DELETE = &H3
Public Const FO_COPY = &H2
Public Const FOF_MULTIDESTFILES = &H1
Public Const FOF_NOCONFIRMATION = &H10             '  Don't prompt the user.
Public Const FOF_NOCONFIRMMKDIR = &H200            '  don't confirm making any needed dirs
Public Const FOF_ALLOWUNDO = &H40
Public Const FOF_SILENT = &H4
Public Const FOF_FILESONLY = &H80                  '  on *.*, do only files
Public Const FOF_SIMPLEPROGRESS = &H100

Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_FIXED = 3
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_CDROM = 5
Public Const DRIVE_RAMDISK = 6

Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
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
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long


' Variáveis Globais para manipulação de Janelas (LeadTools)
Global hCtl                         As OLE_HANDLE
Global IsMove                       As Boolean
Global Xold, Yold, Xatual, Yatual   As Single
Global Atualiza                     As Integer
Global Autentica                    As Object
Global ObScanner                    As Object

'Constantes Globais para manipulação de diretórios
Global Const gstrSEP_URLDIR$ = "/"                      ' Separator for dividing directories in URL addresses.
Global Const gstrSEP_DIR$ = "\"                         ' Directory separator character


' Variáveis e funções para manipulação de Data e Hora
Type SYSTEMTIME
        wYear           As Integer
        wMonth          As Integer
        wDayOfWeek      As Integer
        wDay            As Integer
        wHour           As Integer
        wMinute         As Integer
        wSecond         As Integer
        wMilliseconds   As Integer
End Type

Private Type SHFILEOPSTRUCT
        hwnd As Long
        wFunc As Long
        pFrom As String
        pTo As String
        fFlags As Integer
        fAborted As Boolean
        hNameMaps As Long
        sProgress As String
End Type

Declare Function SetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long
Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Declare Function SetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long
Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function SHFileOperation Lib "shell32.dll" Alias _
   "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long


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

Public Sub LimpaTela(Janela As Form)
    Dim i       As Long
    Dim Ctrl    As Object
    Dim Mask    As String
    Dim campo   As String
    
    On Error Resume Next
    
    For Each Ctrl In Janela.Controls
        'If TypeOf Ctrl Is DataCombo Then Ctrl.BoundText = ""
        'If TypeOf Ctrl Is DataCombo Then Ctrl = ""
        If (TypeOf Ctrl Is TextBox) Then Ctrl.Text = ""
        If TypeName(Ctrl) = "DCurrencyEdit" Then Ctrl.Text = ""
        'If (TypeOf Ctrl Is MaskEdBox) Then
        ' Mask = Ctrl.Mask
        ' Ctrl.Mask = ""
        ' Ctrl.Text = ""
        ' Ctrl.Mask = Mask
        'End If
        If TypeOf Ctrl Is ComboBox Then Ctrl.ListIndex = -1
        If TypeOf Ctrl Is CheckBox Then Ctrl.Checked = False
        'If TypeOf Ctrl Is DTPicker Then Ctrl.Value = Date
    Next
End Sub

Public Function FileExist(ByVal pFileName As String) As Boolean

    Dim lclFileNum As Integer

    On Error Resume Next
    
    lclFileNum = FreeFile
    
    Open pFileName For Input As lclFileNum

    FileExist = IIf(Err = 0, True, False)

    Close lclFileNum

    Err = 0

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
    pObjeto.SetFocus
    
    If Err <> 0 Then Err = 0
    

End Sub

Public Function VerificaDataMMAAAA(ByVal pviData As String) As Boolean
    
    '''''''''''''''''''''''''''''''''''''''''''''
    ' Retorna True se a data é válida           '
    ' Data deve ser informada no formato MMAAAA '
    '''''''''''''''''''''''''''''''''''''''''''''
    
    Dim iMes    As String
    Dim iAno    As String
    Dim sData   As String
    Dim bOk     As Boolean
    
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

Public Function ChecarDiretorio(ByVal pvsDir As String, pvsMsgErro As String) As Boolean
    On Error Resume Next
    If Len(Trim(Dir(pvsDir, vbDirectory))) <> 0 Then
        ChecarDiretorio = True
    Else
        If MsgBox(pvsMsgErro & vbCr & vbCr & "Deseja criá-lo?", vbQuestion + vbYesNo, "Validação dos Parâmetros") = vbYes Then
        
            Err.Clear
            'MkDir pvsDir
            CriaDir pvsDir
            
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
    Dim iRet        As Long
    Dim sDado       As String
    Dim sDadoAux    As String
    Dim i           As Integer
    
    sDado = String(255, " ")
    iRet = GetPrivateProfileString(pvsSecao, pvsItem, pvsDefault, sDado, 255, App.Path & "\MDI_UBB.INI")
    
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
    
    iRet = WritePrivateProfileString(pvsSecao, pvsItem, pvsValor, App.Path & "\MDI_UBB.INI")
End Sub

Public Function VerificaCGC(ByVal CGC As String) As Boolean
   
   '------------------------------------------------
   '--------- MODULO 11 (2 BASE 9) -----------------
   ' Esta rotina serve para conferir o CGC: tam = 15
   '------------------------------------------------
   
   Dim soma         As Integer
   Dim resto        As Integer
   Dim digito_11    As Integer
   Dim p            As Integer
   Dim peso         As Integer
   Dim digito_rv    As String
   Dim bOk          As Boolean
   
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
   
   Dim soma         As Integer
   Dim resto        As Integer
   Dim digito_11    As Integer
   Dim p            As Integer
   Dim peso         As Integer
   Dim digito_rv    As String
   Dim bOk          As Boolean
   
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

Public Function TratarCamposCMC7(ByVal CMC7 As String, ByRef Campo1 As String, _
                            ByRef Campo2 As String, ByRef Campo3 As String, _
                            ByRef Valor As String) As Boolean
    Dim Pos         As Integer
    Dim buffer      As String
    Dim Banco       As String * 3
    Dim agencia     As String * 4
    Dim DV2         As String * 1
    Dim Compe       As String * 3
    Dim cheque      As String * 6
    Dim Tipif       As String * 1
    Dim DV1         As String * 1
    Dim Conta       As String * 10
    Dim DV3         As String * 1
    Dim Aux         As String
    Dim Count       As Integer
    
    ' Inicializando as variaveis
    Banco = String(3, "0")
    agencia = String(4, "0")
    DV2 = "0"
    Compe = String(3, "0")
    cheque = String(6, "0")
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
        agencia = Mid(buffer, 4, 4)
    End If
    If Len(buffer) >= 8 Then
        DV2 = Mid(buffer, 8, 1)
    End If
    If Len(buffer) >= 11 Then
        Compe = Mid(buffer, 9, 3)
    End If
    If Len(buffer) >= 17 Then
        cheque = Mid(buffer, 12, 6)
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
    If Not IsNumeric(Banco) Or Not IsNumeric(agencia) Then
        Banco = String(3, "0")
        agencia = String(4, "0")
    End If
    If Not IsNumeric(Compe) Or Not IsNumeric(cheque) Or Not IsNumeric(Tipif) Then
        Compe = String(3, "0")
        cheque = String(6, "0")
        Tipif = "0"
    End If
    If Not IsNumeric(Conta) Then
        Conta = String(10, "0")
    End If
    
    ' verifica se eh possivel calcular os DVs
    If Val(Banco & agencia) > 0 And IsNumeric(DV1) Then
        If DV10(Banco & agencia) <> DV1 Then
            Campo1 = String(8, "0")
        Else
            Campo1 = Banco & agencia & DV2
        End If
    Else
        Campo1 = String(8, "0")
    End If
    If Val(Compe & cheque & Tipif) > 0 And IsNumeric(DV2) Then
        If DV10(Compe & cheque & Tipif) <> DV2 Then
            Campo2 = String(10, "0")
        Else
            Campo2 = Compe & cheque & Tipif
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
    Dim iDia        As Byte
    Dim iMes        As Byte
    Dim iAno        As Integer
    Dim sData       As String
    Dim iUltimoDia  As Byte
    Dim bOk         As Boolean
    
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

Private Function TratarStringErro(ByVal pvsTexto As String) As String
    Dim i       As Long
    Dim sAux    As String
    
    sAux = ""
    For i = 1 To Len(pvsTexto)
        If Mid(pvsTexto, i, 1) <> "'" Then
            sAux = sAux & Mid(pvsTexto, i, 1)
        End If
    Next
    
    TratarStringErro = sAux
End Function

Public Function TratamentoErro(ByRef pConnection As rdo.rdoConnection, _
                               ByVal pvsTexto As String, _
                               pvoErro As ErrObject, _
                               ByRef pvoRDOErrors As rdoErrors, _
                               Optional pvbMostrar As Boolean = True) As VbMsgBoxResult
    Dim sMens       As String
    Dim sErro       As String
    Dim sRdo        As String
    Dim oErro       As rdoError
    Dim Retorno     As VbMsgBoxResult
    
    GravarErro pConnection, pvsDescricao:=pvsTexto
    
    If pvoErro.Number <> 0 And InStr(pvoErro.Description, "ODBC") = 0 And InStr(pvoErro.Description, "SQL") = 0 Then
        sErro = pvoErro.Number & " - " & pvoErro.Description
        GravarErro pConnection, pvoErro.Number, pvoErro.Description
    Else
        sErro = ""
    End If
    
    sRdo = ""
    
    For Each oErro In pvoRDOErrors
        With oErro
            GravarErro pConnection, .Number, .Description
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
            .Show vbModal ', Principal
            
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

Private Sub GravarErro(ByRef pConnection As rdo.rdoConnection, _
            Optional ByVal pviErro As Long = 0, Optional ByVal pvsDescricao As String = "")
    Dim sSql                As String
    Dim qryInsereLogErro    As rdoQuery
    
    On Error Resume Next
    
    Set qryInsereLogErro = pConnection.CreateQuery("", "{ call MDIAG_InsereLogErro( ?,? ) }")
    With qryInsereLogErro
        .rdoParameters(0) = pviErro
        .rdoParameters(1) = TratarStringErro(pvsDescricao)
        .Execute
        .Close
    End With
    
End Sub

Public Function ChecarRDOError(ByVal pvsTexto As String) As String
    Dim sTexto  As String
    Dim i       As Long
    
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
    Dim Result  As String
    Dim Count   As Integer
    
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
    Dim strValor    As String
    Dim strDecimal  As String
    Dim strInteiro  As String
    Dim strResult   As String
    Dim Count       As Integer
    
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

Public Sub GravaLog(ByRef pConnection As rdo.rdoConnection, _
                        ByVal DataProcessamento As Long, _
                        ByVal IdCapa As Long, _
                        ByVal IdDocto As Long, _
                        ByVal Login As String, _
                        ByVal Acao As Byte)
    
Dim qryInserirLog As rdoQuery

On Error GoTo ErroGravaLog

Set qryInserirLog = pConnection.CreateQuery("", "{call MDIAG_InsereLog (?,?,?,?,?)}")
    
With qryInserirLog
    .rdoParameters(0) = DataProcessamento
    .rdoParameters(1) = IdCapa
    .rdoParameters(2) = IdDocto
    .rdoParameters(3) = Login
    .rdoParameters(4) = Acao
    .Execute
End With

qryInserirLog.Close
    
ErroGravaLog:
On Error GoTo 0

End Sub

Function ValidaAgencia(ByRef pConnection As rdo.rdoConnection, _
            ByVal CodigoAgencia As Integer, _
            ByVal sVencimento As String, _
            ByVal ValidaData As Boolean) As Integer

  Dim RsAgenf       As rdoResultset
  Dim qryGetAgenf   As rdoQuery

  'Código de Retorno
  '0 - Data de Vencimento OK
  '1 - Documento Vencido
  '2 - Agencia em Feriado
  '3 - Agencia Fechada
  '4 - Agencia não cadastrada
  '5 - Data não Verificada

  ValidaAgencia = 5

  'Verificar o Status da Agencia
  Set qryGetAgenf = pConnection.CreateQuery("", "{call MDIAG_GetAgenf (" & CodigoAgencia & ")}")

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
End Function

Function TransformaDataAAAAMMDD(sData As String) As Long

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

Public Function ShellDelete(ByVal pPath As String) As Boolean
   Dim SHFileOp As SHFILEOPSTRUCT
   Dim i As Long

   With SHFileOp
      .wFunc = FO_DELETE
      .pFrom = pPath
      .fFlags = FOF_NOCONFIRMATION
   End With

   i = SHFileOperation(SHFileOp)
   
   If i = 0 Then
      ShellDelete = True
   Else
      ShellDelete = False
   End If
End Function

Public Function ShellCopy(ByVal pPathOrigem As String, ByVal pPathDestino As String) As Boolean
   Dim SHFileOp As SHFILEOPSTRUCT
   Dim i As Long

   With SHFileOp
      .wFunc = FO_COPY
      .pFrom = pPathOrigem
      .pTo = pPathDestino
      .fFlags = FOF_NOCONFIRMATION + FOF_NOCONFIRMMKDIR
   End With

   i = SHFileOperation(SHFileOp)
   
   If i = 0 Then
      ShellCopy = True
   Else
      ShellCopy = False
   End If
End Function


