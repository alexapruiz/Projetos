Attribute VB_Name = "GLOBAIS"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''
' Definição do tipo de variáveis globais '
''''''''''''''''''''''''''''''''''''''''''
'-------------------------------------
'           TIPO DE CAPA
'-------------------------------------
Type TpCapa
    IdCapa As Long
    IdLote As Long
    IdEnv_Mal As String
    capa As Double
    Num_Malote As Double
    AgOrig As Integer
    Duplicidade As Integer
    Status As String
    
End Type
'-------------------------------------
'           TIPO DE DOCUMENTO
'-------------------------------------
Type TpDocumento
'    DataProcessamento As Long
    IdDocto As Long
    IdCapa As Long
    TipoDocto As Integer
    Leitura As String
    Frente As String
    Verso As String
    Status As String
    Alcada As String
    Autenticado As String
    Ocorrencia As Long
    OcorrenciaOK As String
    Ordem As String
    ValorTotal As Currency
    NSU As String
    Terminal As Integer
    Vinculo As Long
    CMC7Associado As String
    Duplicidade As Integer
    Atualizacao As Long
    Transacao As Integer
    Efetivado As Boolean
    PagtoTerceiro As String
    TotalVinculado As Currency
    Excluido As Boolean
    AjusteInterno As Boolean
    Agencia As Integer
    Conta As Long
    AgenciaVinculo As Integer
    ContaVinculo As Long
End Type

''''''''''''''''''''''''''''''''''''''''''
' Definição do tipo de variáveis globais '
''''''''''''''''''''''''''''''''''''''''''
Type tpGlobais
    capa                        As TpCapa
    Documento                   As TpDocumento
    qryLeituraParametro         As rdoQuery
    qryCriarParametro           As rdoQuery
    AgenciaCentral              As String
    Banco                       As rdoConnection
    DataProcessamento           As Long
    DiretorioDados              As String
    DiretorioImagens            As String
    DiretorioTrabalho           As String
    Estacao                     As Integer
    Scanner                     As enumScanner
    autenticadora               As enumAutentica    'alteração versão 3.3 (67)
    VIPSDLL                     As enumVipsDll      ' Unibanco ou Proservi
    Usuario                     As String
    RetornoFinal                As String
    Intervalo                   As Integer          'usado no timer para atualizar DataAtual da capa
    Atualizacao                 As Integer          'usado no timer de atualizacao dos forms
    ValorChqInferior            As Currency
    ValorMaxADCC                As Currency
    StringConexao               As String
    DataFinalRegraAntiga_Mal    As Long             'Data Limite para aceitar Malotes na regra antiga
    LimiteMaxDifLancto_Mal      As Currency         'Valor máximo permitido para diferencas a Debito ou _
                                                     Credito em Lancamentos Internos
End Type
'''''''''''''''''''''''''''''''''''''''''
' Estrutura do arquivo de Retorno Final '
'''''''''''''''''''''''''''''''''''''''''
Type tpRetornoFinal
    Status As String * 2
    Tipo As String * 1
    Leitura As String * 63
    Frente As String * 12
    Verso As String * 12
    Origem As String * 1
    Estacao As String * 2
    CrLf As String * 2
End Type

Type tpRetornoVips
    Tipo As String * 1
    Leitura As String * 63
    Frente As String * 19
    Verso As String * 19
    Origem As String * 1
    CrLf As String * 2
End Type

'''''''''''''''''''''''''''''
' Principal variável global '
'''''''''''''''''''''''''''''
Global Geral As tpGlobais

'Constante contendo a cor de fundo do objeto desabilitado
Public Const G_ColorGray = &H8000000F
Public Const G_ColorBlue = &H800000    '&HC00000


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

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
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


' Variáveis Globais para manipulação de Janelas (LeadTools)
Global hCtl As OLE_HANDLE
Global IsMove As Boolean
Global Xold, Yold, Xatual, Yatual As Single
Global Atualiza As Integer
Global Autentica As Object

' Variáveis e funções para manipulação de Data e Hora
Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Declare Function SetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long
Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Declare Function SetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long
Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

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
    
    If CStr(digito) <> (digito_cei) Then
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
    pObjeto.SetFocus
    
    If Err <> 0 Then Err = 0
    

End Sub
Public Function VerificaDoctosExcluidosCapa(ByVal sIdCapa As Integer) As Boolean
'* Verifica se todos os documentos da capa possuem ocorrência, se a capa não possuir Motivo
'  de exclusão *'

Dim qryGetPesqDoctosOcorr As rdoQuery
Dim qryGetMotivoExclusao  As rdoQuery
Dim RsDoctosOcorr         As rdoResultset
Dim RsMotivoExclusao      As rdoResultset
Dim TotalDoctos           As Integer
Dim TotalDoctoComOCorr    As Integer
Dim TotalDoctosSemCorr    As Integer

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
  Set qryGetPesqDoctosOcorr = Geral.Banco.CreateQuery("", "{Call GetPesqDoctosOcorr(?,?)}")
  
  With qryGetPesqDoctosOcorr
    .rdoParameters(0).Value = Geral.DataProcessamento
    .rdoParameters(1).Value = sIdCapa
     Set RsDoctosOcorr = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
  End With

    If RsDoctosOcorr.RowCount <> 0 Then
        TotalDoctos = RsDoctosOcorr!TotalDoctos
        TotalDoctoComOCorr = RsDoctosOcorr!ComOcorrencia
        TotalDoctosSemCorr = RsDoctosOcorr!SemOcorrencia
    End If

    If TotalDoctos = TotalDoctoComOCorr And TotalDoctos <> 0 Then Exit Function

    VerificaDoctosExcluidosCapa = True

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

Public Function CalculaDataCodigoBarras(pvnIndice As Integer) As String
    Dim nDia As Integer
    Dim nMes As Integer
    Dim nAno As Integer
    Dim nInd As Integer
    Dim nUltimoDia As Integer
    
    If pvnIndice < 1001 Then
        CalculaDataCodigoBarras = "00000000"
        Exit Function
    End If
    
    nDia = 1
    nMes = 6
    nAno = 2000
    
    If pvnIndice <> 1001 Then
        For nInd = 1 To (pvnIndice - 1001)
            Select Case nMes
                Case 1, 3, 5, 7, 8, 10, 12
                    nUltimoDia = 31
                Case 2
                    nUltimoDia = IIf((nAno Mod 4) > 0, 28, 29)
                Case Else
                    nUltimoDia = 30
            End Select
            
            nDia = nDia + 1
            
            If nDia > nUltimoDia Then
                nDia = 1
                nMes = nMes + 1
                If nMes > 12 Then
                    nMes = 1
                    nAno = nAno + 1
                End If
            End If
        Next nInd
    End If
    
    CalculaDataCodigoBarras = Format(nAno, "0000") & Format(nMes, "00") & Format(nDia, "00")
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

  If Mid(pCodigoBarras, 2, 1) = "6" Then
    VerificarArrecadacaoConvencional = True
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
    
    DoEvents
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Cria query para a leitura da tabela parametros '
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    Set Geral.qryLeituraParametro = Geral.Banco.CreateQuery("", "{call LerParametro(?)}")
    Load Principal
    
    On Error GoTo ErroMain

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
    
    Set Autentica = Nothing
    
    ''''''''''''''''''''''''
    ' Fim da inicialização '
    ''''''''''''''''''''''''
    Principal.Show
    
    tb1.Close
    Geral.qryLeituraParametro.Close

    If Not ChecarParametros(Geral) Then
        MsgBox "Não foi possível inicializar o Sistema.", vbExclamation + vbOKOnly, App.Title
        Geral.Banco.Close
        End
    End If
    
    Principal.Caption = App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision & "  [" & Geral.Usuario & "]  [" & Format(DataDDMMAAAA(Geral.DataProcessamento), "00/00/0000") & "]" & " Ag. Proc.: [" & Geral.AgenciaCentral & "]"
    Principal.Refresh

    Exit Sub

ErroMain:
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
                         ByVal Acao As Byte)
    
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

qryInserirLog.Close
    
ErroGravaLog:
On Error GoTo 0

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

Function ValidaAgencia(ByVal CodigoAgencia As Integer, ByVal sVencimento As String, ByVal ValidaData As Boolean) As Integer

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

