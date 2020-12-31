Attribute VB_Name = "SubMain"
Option Explicit
Sub Main()
    Dim iRet                    As Long
    Dim tb                      As rdoResultset
    Dim tb1                     As rdoResultset
    Dim Data                    As SYSTEMTIME
    Dim ScannerOk               As Boolean
    Dim NumBoxes                As Long
    Dim MaxDocBox               As Long
    Dim BoxDefault              As Long
    Dim Threshold               As Long
    Dim Compress                As Long
    Dim Resolution              As Long
    Dim qryLimpaMovimento       As rdoQuery
    Dim iFile                   As Integer
    Dim sFileName               As String
    Dim qryAtualizaAgencia      As rdoQuery
    Dim i                       As Integer
    
    
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
    
    Splash.Show
    
    ' Set o diretorio corrente para a VipsDll encontrar os arquivos
    Call SetCurrentDirectory(App.Path)
    
    Geral.DataProcessamento = Val(Format(Now, "yyyymmdd"))
    
    ''''''''''''''''''''''''''''''''''''''''''''
    ' Rotina de Informar data do Processamento '
    '''''''''''''''''''''''''''''''''''''''' '''
    If Not DataMovimento.ShowModal(Geral.DataProcessamento) Then End
    
    Screen.MousePointer = vbHourglass

    ''''''''''''''''''''''''''''''''''''''''''''
    ' Rotina de Conectar na base MDI_AG        '
    ''''''''''''''''''''''''''''''''''''''''''''
    Set Geral.Banco = New rdo.rdoConnection
    Geral.Banco.Connect = "DSN=MSSQLSERVER;DataBase=UbbMDI;UID=mdi;PWD=mdi;"
    Geral.Banco.CursorDriver = rdUseOdbc
    Geral.Banco.EstablishConnection True
    
    
    ''''''''''''''''''''''''''''''''''''''''''''
    ' Rotina de Conectar na base AGENCIA       '
    ''''''''''''''''''''''''''''''''''''''''''''
    Set Geral.BancoCaixa = New rdo.rdoConnection
    Geral.BancoCaixa.Connect = "DSN=MSSQLSERVER;DataBase=UbbDB;UID=mdi;PWD=mdi;"
    Geral.BancoCaixa.CursorDriver = rdUseOdbc
    Geral.BancoCaixa.EstablishConnection True

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
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Atualização da tabela parametro na obtenção da agencia'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Geral.AgenciaApresentante = Mid(Environ("CNAME"), 2, 4)
    Set qryAtualizaAgencia = Geral.Banco.CreateQuery("", "{ ? = Call MDIAG_AtualizaAgencia(?)}")
    With qryAtualizaAgencia
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1).Value = Geral.AgenciaApresentante
        .Execute
        If .rdoParameters(0).Value <> 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "Não foi possível inicializar o sistema.", vbExclamation
            End
        End If
    End With
    qryAtualizaAgencia.Close

    
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Cria query para a leitura da tabela parametros '
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    Set Geral.qryLeituraParametro = Geral.Banco.CreateQuery("", "{call MDIAG_LerParametro}")
    On Error GoTo ErroMain

    ''''''''''''''''''''''''''''''''
    ' Leitura da tabela parametros '
    ''''''''''''''''''''''''''''''''
    With Geral.qryLeituraParametro
        Set tb1 = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    If tb1.EOF Then
        Screen.MousePointer = vbDefault
        MsgBox "Não foi possível ler informações da tabela Parâmetro.", vbCritical + vbOKOnly, App.Title
        Geral.qryLeituraParametro.Close
        Geral.Banco.Close
        End
    End If
    
    '''''''''''''''''''''''''''''''''
    'Abrir tela de seleção de drives'
    '''''''''''''''''''''''''''''''''
    MsgBox "Voltar ao que estava antes"
    For i = Asc("A") To Asc("Z")
        If GetDriveType(Chr(i) & ":") = DRIVE_REMOVABLE Then
            Geral.CDR.Drive = Chr(i) & ":\"
            Geral.CDR.DiretorioImagens = "IMAGENS\"
            Geral.CDR.DiretorioDados = "DADOS\"
            Exit For
        End If
    Next i
    
    If Left(Geral.CDR.Drive, 1) = Chr(0) Then
        MsgBox "Não foi possível localizar nenhuma unidade de disco removível. " & vbCrLf & _
            "Ligue a unidade de disco removível e reinicie a estação.", vbOKOnly + vbExclamation, App.Title
        End
    End If
    ''''''''''''''''''''''''''''
    'Verificacao dos parametros
    ''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Usuario informou uma data de processamento já existente'
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Geral.DataProcessamento = tb1!DataProcessamento Then
        ''''''''''''''''''''''''''''''''''''''''''''
        'Caso o sistema fora finalizado normalmente'
        ''''''''''''''''''''''''''''''''''''''''''''
        If Not IsNull(tb1!Hm_Fechamento) Then
            ''''''''''''''''''''''''''''
            'Habilitar controle Geração'
            ''''''''''''''''''''''''''''
            Principal.mnuRecepcao.Enabled = False
            Principal.mnuCaptura.Enabled = False
            'Principal.mnuEstatistica.Enabled = False
            Principal.mnuControleGeracao.Enabled = True
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Caso contrário é a mesma data e não foi finalizado, executar procedimento normal'
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Else
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        'Se data diferente e o sistema não foi finalizado'
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        If IsNull(tb1!Hm_Fechamento) Then
            MsgBox "O movimento do dia '" & Format(Format(tb1!DataProcessamento, "0000/00/00"), "dd/mm/yyyy") & "' não foi encerrado." & _
                   "Será necessário" & Chr(10) & "encerrar o movimento para que se possa processar um novo dia." & Chr(10) & _
                   "O sistema assumirá a data do dia '" & Format(Format(tb1!DataProcessamento, "0000/00/00"), "dd/mm/yyyy") & "'.", vbExclamation
            
            Geral.DataProcessamento = tb1!DataProcessamento
        Else
            ''''''''''''''''''''''''''''''''''''
            'Limpar movimento e criar novo lote'
            ''''''''''''''''''''''''''''''''''''
            If MsgBox("Confirma a limpeza do Banco de Dados e das Imagens?", vbQuestion + vbYesNo) = vbYes Then
            
                ' Chamar limpeza das imagens
                If Not ShellDelete(IIf(Right(tb1!Dir_Imagens, 1) = "\", tb1!Dir_Imagens, tb1!Dir_Imagens & "\") & "*.*") Or _
                   Not ShellDelete(IIf(Right(tb1!Dir_Dados, 1) = "\", tb1!Dir_Dados, tb1!Dir_Dados & "\") & "*.*") Or _
                   Not ShellDelete(IIf(Right(tb1!Dir_Trabalho, 1) = "\", tb1!Dir_Trabalho, tb1!Dir_Trabalho & "\") & "*.*") Then
                    MsgBox "Atenção! Erro ao limpar as imagens dos movimentos anteriores.", vbOKOnly + vbCritical, App.Title
                End If
'                ''''''''''''''''''''''''''''''''''
'                'Verifica se o CD está na unidade'
'                ''''''''''''''''''''''''''''''''''
'                Do While (DirExists(Geral.CDR.Drive) = 0)
'                    If MsgBox("Favor inserir o CD na unidade " & Left(Geral.CDR.Drive, 2) & ".", vbExclamation + vbOKCancel) = vbCancel Then
'                        MsgBox "Não será possível inicializar o sistema.", vbExclamation
'                        End
'                    End If
'                Loop
                ''''''''''''''''''''
                'Limpar o CD tambem'
                ''''''''''''''''''''
                If Not ShellDelete(Geral.CDR.Drive & "*.*") Then
                    MsgBox "Atenção! Erro ao limpar as imagens dos movimentos anteriores no CD.", vbCritical
                    MsgBox "Não será possível inicializar o sistema.", vbExclamation
                    End
                End If
            
                Set qryLimpaMovimento = Geral.Banco.CreateQuery("", "{? = call MDIAG_LimpaMovimento(?)}")
                With qryLimpaMovimento
                    .rdoParameters(0).Direction = rdParamReturnValue
                    .rdoParameters(1).Value = Geral.DataProcessamento
                    .Execute
                    If .rdoParameters(0).Value <> 0 Then
                        Screen.MousePointer = vbDefault
                        MsgBox "Não foi possível limpar a base.", vbCritical
                        End
                    End If
                    .Close
                    With Principal
                        .mnuRecepcao.Enabled = True
                        .mnuCaptura.Enabled = True
                        .mnuEstatistica.Enabled = True
                        .mnuControleGeracao.Enabled = True
                    End With
                End With
            Else
                Screen.MousePointer = vbDefault
                MsgBox "Não será possível inicializar o sistema.", vbExclamation
                End
            End If
        End If
    End If
    
    Geral.DiretorioImagens = IIf(Right(tb1!Dir_Imagens, 1) = "\", tb1!Dir_Imagens, tb1!Dir_Imagens & "\") & Geral.DataProcessamento & "\"
    Geral.AgenciaCentral = Format(tb1!AgenciaCentral, "0000")
    Geral.AgenciaApresentante = Format(tb1!AgenciaApresentante, "0000")
    Geral.Intervalo = tb1!TM_Pendente
    Geral.Atualizacao = tb1!TM_Atualizacao
    Geral.DiretorioDados = IIf(Right(tb1!Dir_Dados, 1) = "\", tb1!Dir_Dados, tb1!Dir_Dados & "\")
    Geral.DiretorioTrabalho = IIf(Right(tb1!Dir_Trabalho, 1) = "\", tb1!Dir_Trabalho, tb1!Dir_Trabalho & "\")
    
    '''''''''''''''''''''''''''''''''''''
    ' Inicializar parametros do sistema '
    '''''''''''''''''''''''''''''''''''''
    With Geral
        .Scanner = Val(PegarOpcaoINI("Diversos", "Scanner", "0"))
        .Autenticadora = Val(PegarOpcaoINI("Diversos", "Autenticadora", "0"))
    End With
    NumBoxes = Val(PegarOpcaoINI("Diversos", "NumBoxes", "1"))
    MaxDocBox = Val(PegarOpcaoINI("Diversos", "MaxDocBox", "200"))
    BoxDefault = Val(PegarOpcaoINI("Diversos", "BoxDefault", "0"))
    Threshold = Val(PegarOpcaoINI("Diversos", "CutBords", "50"))
    Compress = Val(PegarOpcaoINI("Diversos", "Compress_JPG", "30"))
    Resolution = Val(PegarOpcaoINI("Diversos", "Resolution", "100"))
    
    Load Principal
    
    Set Autentica = Nothing
    
    If (Geral.Autenticadora = 1) Or (Geral.Autenticadora = 2) Then
        Set Autentica = New Autenticadora
    End If
'        Set Autentica = New Autenticadora
'    ElseIf Geral.Autenticadora = 2 Then
'        Set Autentica = New Autentica_Procomp
'    End If
    
    '''''''''''''''''''''''
    ' Inicializar Scanner '
    '''''''''''''''''''''''
    ScannerOk = False
    rdoErrors.Clear
    
    iRet = 1
    If Geral.Scanner = escnVIPS Then
    
        Set ObScanner = New Scanner
    
        ObScanner.SetBoxes (NumBoxes)
        ObScanner.SetMaxDocBox (MaxDocBox)
        ObScanner.SetBoxDefault (BoxDefault)
        ObScanner.SetCompress (Compress)
        ObScanner.SetCutBords (Threshold)
        ObScanner.SetCameraFile ("Doc100.cpf")
        ObScanner.SetImageDirectory (Geral.DiretorioImagens)
        ObScanner.SetResolution (Resolution)
        iRet = ObScanner.Init()
        If iRet <> 0 Then
            MsgBox "Não foi possível inicializar o Scanner." & vbCr & "Erro: " & iRet, vbExclamation + vbOKOnly, App.Title
        Else
            ScannerOk = True
        End If
    End If
    
    If ScannerOk Then
      ' habilitar menu no form principal
      Principal.mnuCapCaptura.Enabled = True
    Else
      ' desabilitar menu no form principal
      Principal.mnuCapCaptura.Enabled = False
    End If
    
    '''''''''''''''''''''''''''
    'Carregar arquivo de login'
    '''''''''''''''''''''''''''
    iFile = FreeFile
    sFileName = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\") & "LOGIN.TXT"
    
    If FileExist(sFileName) Then
        Open sFileName For Binary As iFile Len = Len(Geral.Usuario)
            Input #iFile, Geral.Usuario.Login
            Input #iFile, Geral.Usuario.Nome
        Close iFile
    Else
        Geral.Usuario.Login = String(Len(Geral.Usuario.Login), " ")
        Geral.Usuario.Nome = String(Len(Geral.Usuario.Nome), " ")
        MsgBox "O arquivo '" & StrConv(sFileName, vbProperCase) & "' não existe.", vbExclamation
    End If
    
    With Geral.Usuario
        .Login = Trim(.Login)
        .Nome = Trim(.Nome)
    End With
    
    ''''''''''''''''''''''''
    ' Fim da inicialização '
    ''''''''''''''''''''''''
    Unload Splash
    Principal.Show
    Screen.MousePointer = vbDefault
    'Unload SplashScreen
    
    tb1.Close
    Geral.qryLeituraParametro.Close

    If Not ChecarParametros(Geral) Then
        MsgBox "Não foi possível inicializar o Sistema.", vbExclamation + vbOKOnly, App.Title
        Geral.Banco.Close
        End
    End If
    
    Principal.Caption = App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision & " [" & Trim(Geral.Usuario.Login) & "] [" & Format(DataDDMMAAAA(Geral.DataProcessamento), "00/00/0000") & "]" & " Agência: [" & Geral.AgenciaApresentante & "]"
    
    Principal.Refresh

    Exit Sub

ErroMain:
    Screen.MousePointer = vbDefault
    Close iFile
    Select Case TratamentoErro(Geral.Banco, "Não foi possível inicializar o Sistema.", Err, rdoErrors)
        Case vbCancel
            End
        Case vbRetry
            End
    End Select
    Unload Splash
End Sub



