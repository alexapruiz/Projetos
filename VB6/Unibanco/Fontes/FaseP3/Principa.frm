VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form Principal 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8544
   ClientLeft      =   48
   ClientTop       =   636
   ClientWidth     =   12228
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8544
   ScaleWidth      =   12228
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport RptGeral2 
      Left            =   11340
      Top             =   660
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport RptGeral 
      Left            =   11328
      Top             =   264
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ComctlLib.StatusBar BarMain 
      Align           =   2  'Align Bottom
      Height          =   336
      Left            =   0
      TabIndex        =   0
      Top             =   8208
      Width           =   12228
      _ExtentX        =   21569
      _ExtentY        =   593
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   17622
            MinWidth        =   17622
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2011
            MinWidth        =   2011
            TextSave        =   "03/06/2003"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1834
            MinWidth        =   1834
            TextSave        =   "12:16"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuRecepcao 
      Caption         =   "&Recepção"
      Index           =   0
      Begin VB.Menu mnuRecRecepcao 
         Caption         =   "&Recepção..."
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu mnuRecCadCapas 
         Caption         =   "&Cadastro"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu mnuRecRegistroOcorrencia 
         Caption         =   "Registro de &Ocorrência..."
         Enabled         =   0   'False
         Index           =   4
      End
   End
   Begin VB.Menu MnuCaptura 
      Caption         =   "&Captura"
      Index           =   0
      Begin VB.Menu mnuCapCaptura 
         Caption         =   "&Captura..."
         Enabled         =   0   'False
         Index           =   5
      End
      Begin VB.Menu mnuCapCtrQualidade 
         Caption         =   "Controle de &Qualidade..."
         Enabled         =   0   'False
         Index           =   6
      End
      Begin VB.Menu MnuCapRecaptura 
         Caption         =   "&Recaptura"
         Enabled         =   0   'False
         Index           =   7
      End
   End
   Begin VB.Menu mnuComplementacao 
      Caption         =   "Com&plementação..."
      Enabled         =   0   'False
      Index           =   8
   End
   Begin VB.Menu mnuIlegiveis 
      Caption         =   "&Ilegíveis..."
      Enabled         =   0   'False
      Index           =   9
   End
   Begin VB.Menu mnuProvaZero 
      Caption         =   "Prova &Zero..."
      Enabled         =   0   'False
      Index           =   10
   End
   Begin VB.Menu mnuExpedicao 
      Caption         =   "&Expedição..."
      Enabled         =   0   'False
      Index           =   11
   End
   Begin VB.Menu mnuSupervisao 
      Caption         =   "S&upervisão"
      Index           =   0
      Begin VB.Menu mnuSupAcompanhamento 
         Caption         =   "Acompanhamento &Produção..."
         Enabled         =   0   'False
         Index           =   12
      End
      Begin VB.Menu mnuSupMensagensRobo 
         Caption         =   "Mensagens do Robô"
         Enabled         =   0   'False
         Index           =   36
      End
      Begin VB.Menu MnuSupAcompAtividade 
         Caption         =   "Acompanhamento de A&tividades"
         Enabled         =   0   'False
         Index           =   13
      End
      Begin VB.Menu MnuSupAcompRecepcao 
         Caption         =   "Acompanhamento de &Recepção"
         Enabled         =   0   'False
         Index           =   14
      End
      Begin VB.Menu mnusupacompexped 
         Caption         =   "Acompanhamento de &Expedição"
         Enabled         =   0   'False
         Index           =   27
      End
      Begin VB.Menu mnuSupAcompUsers 
         Caption         =   "Acompanhamento de &Usuários"
         Enabled         =   0   'False
         Index           =   15
      End
      Begin VB.Menu mnuSupAlcada 
         Caption         =   "&Alçada..."
         Enabled         =   0   'False
         Index           =   16
      End
      Begin VB.Menu mnuSupVinculo 
         Caption         =   "&Vínculo..."
         Enabled         =   0   'False
         Index           =   17
      End
      Begin VB.Menu mnuSupAuditoria 
         Caption         =   "Au&ditoria..."
         Enabled         =   0   'False
         Index           =   18
      End
      Begin VB.Menu mnuSupCSP 
         Caption         =   "C.S.P."
         Enabled         =   0   'False
         Index           =   31
      End
      Begin VB.Menu Mnusupfinalcapas 
         Caption         =   "&Finalização de Capas"
         Enabled         =   0   'False
         Index           =   26
      End
      Begin VB.Menu MnuSupEstorno 
         Caption         =   "Estor&no"
         Enabled         =   0   'False
         Index           =   28
      End
      Begin VB.Menu MnuSupConfAgCc 
         Caption         =   "Conf&irmação de Ag/Conta"
         Enabled         =   0   'False
         Index           =   19
      End
      Begin VB.Menu MnuSupCorrecaoAgCc 
         Caption         =   "Correção de A&g/Conta"
         Enabled         =   0   'False
         Index           =   25
      End
      Begin VB.Menu mnuSupTrocarOrdem 
         Caption         =   "Trocar &Ordem de Documento"
         Enabled         =   0   'False
         Index           =   20
      End
      Begin VB.Menu mnuSupExclusao 
         Caption         =   "E&xclusão..."
         Enabled         =   0   'False
         Index           =   21
      End
      Begin VB.Menu mnuSupParametros 
         Caption         =   "Parâmetros do &Sistema..."
         Enabled         =   0   'False
         Index           =   22
      End
      Begin VB.Menu mnuSupCadastroUsuario 
         Caption         =   "&Cadastro de Usuário..."
         Enabled         =   0   'False
         Index           =   23
      End
      Begin VB.Menu mnuSupAlteraSenha 
         Caption         =   "Alteração de &Senha"
         Index           =   0
      End
      Begin VB.Menu mnuSupGrpMod 
         Caption         =   "Cadastro de &Grupo de Módulos"
         Enabled         =   0   'False
         Index           =   29
      End
   End
   Begin VB.Menu mnuConsultasRelatorios 
      Caption         =   "C&onsultas e Relatórios"
      Index           =   0
      Begin VB.Menu mnuConConsulta 
         Caption         =   "&Consulta..."
         Enabled         =   0   'False
         Index           =   24
      End
      Begin VB.Menu mnuConRelatorio 
         Caption         =   "&Relatório"
         Enabled         =   0   'False
         Index           =   30
         Begin VB.Menu mnuConRelRecepcao 
            Caption         =   "R&ecepção..."
            Index           =   0
         End
         Begin VB.Menu MnuConRelDifRecep 
            Caption         =   "Inconsistência de Recepção"
            Index           =   0
         End
         Begin VB.Menu mnuRelOpcoesTotais 
            Caption         =   "&Totais"
            Index           =   0
            Begin VB.Menu RelTotais 
               Caption         =   "&Estatística de Documentos"
               Index           =   0
            End
            Begin VB.Menu mnuRelTotalConsolidado 
               Caption         =   "&Estatística de Documentos Consolidado"
               Index           =   0
            End
            Begin VB.Menu mnuRelTotalPorTipoDocto 
               Caption         =   "Totais Aberto por Tipo de &Documento"
               Index           =   0
            End
            Begin VB.Menu mnuRelTotDocumentoPorCliente 
               Caption         =   "Totais de Documentos por &Cliente"
               Index           =   0
            End
         End
         Begin VB.Menu mnuRelPercDoctoModulo 
            Caption         =   "&Percentual de Documentos por Módulo"
            Index           =   0
         End
         Begin VB.Menu mnuControleExpedidos 
            Caption         =   "&Controle de Expedidos..."
            Index           =   0
         End
         Begin VB.Menu mnuCapasNaoFinalizadas 
            Caption         =   "Capas não &Finalizadas"
            Index           =   0
         End
         Begin VB.Menu mnuConRelProcAnalitico 
            Caption         =   "Processamento &Analítico..."
            Index           =   0
         End
         Begin VB.Menu mnuConRelProcConsolidado 
            Caption         =   "Processamento &Consolidado..."
            Index           =   0
         End
         Begin VB.Menu mnuConsRelEstorno 
            Caption         =   "Documentos Es&tornados"
            Index           =   0
         End
         Begin VB.Menu mnuRelEstDocAgencia 
            Caption         =   "Estatistica de Documentos por Agência"
            Index           =   0
         End
         Begin VB.Menu MnuRelSegDoctos 
            Caption         =   "Se&gmentação de Documentos"
            Index           =   0
         End
         Begin VB.Menu MnuRelMovComp 
            Caption         =   "M&ovimento de Comp - Analítico"
            Index           =   0
         End
         Begin VB.Menu mnuRelChequesUnibanco 
            Caption         =   "Cheques &Unibanco"
            Index           =   0
            Begin VB.Menu mnuRelMovtoChqUBBComp 
               Caption         =   "Cheques &UBB para compensação"
               Index           =   0
            End
            Begin VB.Menu mnuRelChqUBBCompValor 
               Caption         =   "Cheques UBB para compensação por &Valor"
               Index           =   0
            End
         End
         Begin VB.Menu mnuRelRelacao_TC_AR 
            Caption         =   "Relatório de &TC´s e AR´s"
            Index           =   0
         End
         Begin VB.Menu MnuCaixaExpresso 
            Caption         =   "Caixa &Expresso com Ocorrência"
            Index           =   0
         End
         Begin VB.Menu MnuCaixaExpressoChequeSuperior 
            Caption         =   "Caixa Expresso com c&heque superior"
            Index           =   0
         End
         Begin VB.Menu MnuConRelValeTransporte 
            Caption         =   "Resumo &Mensal - Vale Transporte"
            Index           =   0
         End
         Begin VB.Menu MnuRelLancamentoInterno 
            Caption         =   "&Lançamento Interno"
            Index           =   0
         End
         Begin VB.Menu mnuRelDiferencasCaixa 
            Caption         =   "&Diferenças no Fechamento do Caixa"
            Index           =   0
         End
         Begin VB.Menu MnuEnvMalEx 
            Caption         =   "Envelopes / Malote Excluídos"
            Index           =   0
         End
         Begin VB.Menu mnuEnvelopeFininvest 
            Caption         =   "Envelope &Fininvest"
            Index           =   0
            Begin VB.Menu MnuEnvelopesFininvestProc 
               Caption         =   "Envelopes Fininvest &Recepcionados"
               Index           =   0
            End
            Begin VB.Menu MnuRelEnvFininvestDevolvido 
               Caption         =   "Envelopes Fininvest &Devolvidos"
               Index           =   0
            End
            Begin VB.Menu mnuRelEnvFininvestPorAgencia 
               Caption         =   "Envelopes Fininvest por &Agência"
               Index           =   0
            End
         End
         Begin VB.Menu mnuRelConcessionarias 
            Caption         =   "Relação de &Concessionárias"
            Index           =   0
         End
         Begin VB.Menu mnuMotivoDoctosIlegiveis 
            Caption         =   "Relação de Doctos. Enviados para Ilegíveis"
            Index           =   0
         End
         Begin VB.Menu mnuRelAcompProducao 
            Caption         =   "Acompanhamento da Produção"
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuAltDataMovimento 
      Caption         =   "&Trocar Data"
      Index           =   0
   End
   Begin VB.Menu mnuSupBloquearAplicacao 
      Caption         =   "&Bloquear"
      Index           =   0
   End
   Begin VB.Menu mnuSobre 
      Caption         =   "&Info"
      Index           =   0
   End
   Begin VB.Menu mnuSair 
      Caption         =   "&Sair"
      Index           =   0
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()

Dim Control As Control

    If CBool(Val(mnuSupBloquearAplicacao(0).Tag)) Then Exit Sub

    Call AtualizaAtividade(1)

    RptGeral.WindowLeft = 0
    RptGeral.WindowTop = 0

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                    * Define quais Môdulos o Usuário Corrente poderá Usar *                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If Geral.idUsuario = 0 Then
        'Desabilita opção do menu de alteração da senha para usuário 'desenv'
        For Each Control In Me.Controls
            If LCase(Trim(Control.Name)) = "mnusupalterasenha" Then
                Control.Enabled = False
                Exit For
            End If
        Next
        
        Exit Sub
    End If

    If Not PreparaMenu(Me) Then
        Unload Me
    End If
End Sub


Private Sub Form_Load()

    'Inicialização das DLLs da VIPS
    Call InicializaDLLsVIPS
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim iRet As Long
    On Error Resume Next
    If Geral.Scanner = escnVIPS Then
        Call FinalizaDLLsVIPS
        'If Geral.VIPSDLL = eDllProservi Then
        '    iRet = MC93_DeInit
        'Else
        '    VIPS_Done
        'End If
    End If

    Geral.Banco.Close
End Sub

Private Sub mnuAltDataMovimento_Click(Index As Integer)

    Dim sDataProcessamento      As String
    Dim sDataProcessamentoOld   As String
    Dim tb1                     As RDO.rdoResultset


    '''''''''''''''''''''''''''''''''''''''''
    'Guarda a data de processamento corrente'
    '''''''''''''''''''''''''''''''''''''''''
    sDataProcessamentoOld = Geral.DataProcessamento
    
    If AlteraDataMovimento.ShowModal(sDataProcessamento) Then
        ''''''''''''''''''''''''''''''''
        'Altera a data de processamento'
        ''''''''''''''''''''''''''''''''
        Geral.DataProcessamento = sDataProcessamento
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Cria query para a leitura da tabela parametros '
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        Set Geral.qryLeituraParametro = Geral.Banco.CreateQuery("", "{call LerParametro(?)}")
        
        '''''''''''''''''''''''''''''''''''''''''''''
        'Carrega parametros da data de processamento'
        '''''''''''''''''''''''''''''''''''''''''''''
        With Geral.qryLeituraParametro
            .rdoParameters(0) = Geral.DataProcessamento
            Set tb1 = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
        End With
        If tb1.EOF Then
            MsgBox "Não foi possível carregar os parâmetros do sistema do dia " & Format(Format(sDataProcessamento, "0000/00/00"), "dd/mm/yyyy"), vbCritical
            Geral.qryLeituraParametro.Close
            ''''''''''''''''''''''''''''''''''''''''
            'Volta a data de processamento anterior'
            ''''''''''''''''''''''''''''''''''''''''
            Geral.DataProcessamento = sDataProcessamentoOld
            Exit Sub
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
        'Fecha query de leitura de parametro'
        '''''''''''''''''''''''''''''''''''''
        Geral.qryLeituraParametro.Close
        
        '''''''''''''''''''''''''
        'Altera o caption do MDI'
        '''''''''''''''''''''''''
        Principal.Caption = App.Title & " " & _
                            App.Major & "." & _
                            App.Minor & "." & _
                            App.Revision & "  [" & _
                            Left(Geral.NomeUsuario, 15) & "]  [" & _
                            Format(DataDDMMAAAA(Geral.DataProcessamento), "00/00/0000") & "]" & _
                            " Ag. Proc.: [" & Geral.AgenciaCentral & "]"
        Principal.Refresh

    End If

End Sub
Private Sub MnuCaixaExpresso_Click(Index As Integer)

    Load FrmExpedidos
    
    FrmExpedidos.m_TipoRelatorio = eCaixaExpresso_MaloteEmpresa_Ocorrencia
    
    FrmExpedidos.Show vbModal, Me
  
End Sub
Private Sub MnuCaixaExpressoChequeSuperior_Click(Index As Integer)
    Load FrmExpedidos
    
    FrmExpedidos.m_TipoRelatorio = eCaixaExpresso_MaloteEmpresa_Cheque
    
    FrmExpedidos.Show vbModal, Me

End Sub


Private Sub mnuCapasNaoFinalizadas_Click(Index As Integer)

    Dim qryCapasNaoFinal  As rdoQuery
    Dim RsCapasNaoFinal   As rdoResultset
    Dim sTitulo As String
    
    sTitulo = "Relatório de Capas não finalizadas"
    Set qryCapasNaoFinal = Geral.Banco.CreateQuery("", "{call GetCapasNaoFinalizadas(?)}")
    
    Screen.MousePointer = vbHourglass
    With qryCapasNaoFinal
        .rdoParameters(0).Value = Geral.DataProcessamento
        Set RsCapasNaoFinal = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsCapasNaoFinal.RowCount <> 0 Then
        RptGeral.Connect = Geral.StringConexao
        
        If Geral.Backup Then
            RptGeral.ReportFileName = App.path & "\CapasNaoFinalizadasBk.rpt"
        Else
            RptGeral.ReportFileName = App.path & "\CapasNaoFinalizadasProd.rpt"
        End If
        
        RptGeral.Formulas(0) = "AgenciaCentral    = '" & Geral.AgenciaCentral & "'"
        RptGeral.Formulas(1) = "DataMovto = '" & Right(Geral.DataProcessamento, 2) & "/" & Mid(Geral.DataProcessamento, 5, 2) & "/" & Left(Geral.DataProcessamento, 4) & "'"
        
        RptGeral.StoredProcParam(0) = Geral.DataProcessamento
        RptGeral.Destination = crptToWindow
        RptGeral.WindowState = crptMaximized
        RptGeral.WindowTitle = sTitulo
        RptGeral.Action = 1
    Else
        Screen.MousePointer = vbDefault
        MsgBox "Não existem informações suficientes para emissão deste Relatório!" & vbCr & sTitulo, vbInformation, App.Title
    End If
    
    RptGeral.ReportFileName = Empty
    RptGeral.Formulas(0) = Empty
    RptGeral.Formulas(1) = Empty
    RptGeral.StoredProcParam(0) = Empty
    Screen.MousePointer = vbDefault

End Sub

Private Sub MnuCapRecaptura_Click(Index As Integer)

    Recaptura.Show vbModal, Me
End Sub

Private Sub MnuConRelDifRecep_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    
    Dim qryDifRecepcao As rdoQuery
    Dim RsDifRecepcao  As rdoResultset
    
    Screen.MousePointer = vbHourglass
    Set qryDifRecepcao = Geral.Banco.CreateQuery("", "{call GetInconsistenciaRecepcao(?)}")
    
    With qryDifRecepcao
        .rdoParameters(0).Value = Geral.DataProcessamento
        Set RsDifRecepcao = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsDifRecepcao.RowCount <> 0 Then
                
        RptGeral2.Connect = Geral.StringConexao
        
        If Geral.Backup Then
            RptGeral2.ReportFileName = App.path & "\IncRecepBk.rpt"
        Else
            RptGeral2.ReportFileName = App.path & "\IncRecepProd.rpt"
        End If
        
        RptGeral2.Formulas(0) = "AgenciaCentral    = '" & Geral.AgenciaCentral & "'"
        RptGeral2.StoredProcParam(0) = Geral.DataProcessamento
        RptGeral2.WindowTitle = "Relatório de Inconsistências na Recepção de Caixa Expresso / Malote Empresa"
        RptGeral2.Destination = crptToWindow
        RptGeral2.WindowState = crptMaximized
        RptGeral2.Action = 1
    Else
        Screen.MousePointer = vbDefault
        MsgBox "Não existe movimento para o período atual!", vbInformation, App.Title
    End If
    
    RptGeral2.ReportFileName = Empty
    RptGeral2.StoredProcParam(0) = Empty
    RptGeral2.Formulas(0) = Empty
    
    Screen.MousePointer = vbDefault
    
End Sub
Private Sub mnuConRelProcConsolidado_Click(Index As Integer)

    Dim qryVerDadosRelcons As rdoQuery
    Dim RsVerDadosRelcons  As rdoResultset
    
    Set qryVerDadosRelcons = Geral.Banco.CreateQuery("", "{call relcons(?,?)}")
    
    Screen.MousePointer = vbHourglass
    
    With qryVerDadosRelcons
        .rdoParameters(0).Value = Geral.DataProcessamento
        .rdoParameters(1).Value = InserePonto(Geral.ValorChqInferior * 100)
        Set RsVerDadosRelcons = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsVerDadosRelcons.RowCount <> 0 Then
        RptGeral.Connect = Geral.StringConexao
        
        If Geral.Backup Then
            RptGeral.ReportFileName = App.path & "\RelConsBk.rpt"
        Else
            RptGeral.ReportFileName = App.path & "\RelConsProd.rpt"
        End If
        
        RptGeral.Formulas(0) = "AgenciaCentral    = '" & Geral.AgenciaCentral & "'"
        RptGeral.StoredProcParam(0) = Geral.DataProcessamento
        RptGeral.StoredProcParam(1) = InserePonto(Geral.ValorChqInferior * 100)
        RptGeral.WindowTitle = "Relatório de Processamento Consolidado"
        RptGeral.Destination = crptToWindow
        RptGeral.WindowState = crptMaximized
        RptGeral.Action = 1
    Else
        Screen.MousePointer = vbDefault
        MsgBox "Não existe movimento para o período atual!", vbInformation, App.Title
    End If
    
    RptGeral.ReportFileName = Empty
    RptGeral.Formulas(0) = Empty
    RptGeral.StoredProcParam(0) = Empty
    RptGeral.StoredProcParam(1) = Empty
    
    Screen.MousePointer = vbDefault
  
End Sub
Private Sub mnuConsRelEstorno_Click(Index As Integer)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                         *  Relatório de Documentos Estornados *                             '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim qryGetRelEstorno   As rdoQuery
Dim rsRelEstorno       As rdoResultset
Dim SDtaProcessamento  As String

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                         '* Formatação da Data de Processamento *'                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     SDtaProcessamento = Mid(Geral.DataProcessamento, 7, 2) & "/" & _
                         Mid(Geral.DataProcessamento, 5, 2) & "/" & _
                         Mid(Geral.DataProcessamento, 1, 4)
    
    
    Set qryGetRelEstorno = Geral.Banco.CreateQuery("", "{call GetRelEstorno(?)}")
    
    Screen.MousePointer = vbHourglass
    
    With qryGetRelEstorno
        .rdoParameters(0).Value = Geral.DataProcessamento
        Set rsRelEstorno = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If rsRelEstorno.RowCount <> 0 Then
        RptGeral.Connect = Geral.StringConexao
        
        If Geral.Backup Then
            RptGeral.ReportFileName = App.path & "\RelEstornadosBk.rpt"
        Else
            RptGeral.ReportFileName = App.path & "\RelEstornadosProd.rpt"
        End If
        
        RptGeral.Formulas(0) = "AgProcessadora = '" & Geral.AgenciaCentral & "'"
        RptGeral.Formulas(1) = "DataProcessamento   = '" & SDtaProcessamento & "'"
        RptGeral.StoredProcParam(0) = Geral.DataProcessamento
        RptGeral.WindowTitle = "Relatório de Documentos Estonados"
        RptGeral.Destination = crptToWindow
        RptGeral.WindowState = crptMaximized
        RptGeral.Action = 1
    Else
        Screen.MousePointer = vbDefault
        MsgBox "Não existe movimento para o período atual!", vbInformation, App.Title
    End If
    
    Set qryGetRelEstorno = Nothing
    Set rsRelEstorno = Nothing
    
    RptGeral.ReportFileName = Empty
    RptGeral.Formulas(0) = Empty
    RptGeral.Formulas(1) = Empty
    RptGeral.StoredProcParam(0) = Empty

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub mnuControleExpedidos_Click(Index As Integer)
    
    FrmRelExpedidos.Show vbModal, Me
    
End Sub
Private Sub mnuconrelprocanalitico_Click(Index As Integer)

    Dim qryVerDadosRelanali As rdoQuery
    Dim RsVerDadosRelanali  As rdoResultset
    
    Set qryVerDadosRelanali = Geral.Banco.CreateQuery("", "{call relanali(?,?)}")
    
    Screen.MousePointer = vbHourglass
    
    With qryVerDadosRelanali
        .rdoParameters(0).Value = Geral.DataProcessamento
        .rdoParameters(1).Value = InserePonto(Geral.ValorChqInferior * 100)
        Set RsVerDadosRelanali = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsVerDadosRelanali.RowCount <> 0 Then
        RptGeral.Connect = Geral.StringConexao
        
        If Geral.Backup Then
            RptGeral.ReportFileName = App.path & "\RelAnaliBk.rpt"
        Else
            RptGeral.ReportFileName = App.path & "\RelAnaliProd.rpt"
        End If
        
        RptGeral.Formulas(0) = "AgenciaCentral    = '" & Geral.AgenciaCentral & "'"
        RptGeral.StoredProcParam(0) = Geral.DataProcessamento
        RptGeral.StoredProcParam(1) = InserePonto(Geral.ValorChqInferior * 100)
        RptGeral.WindowTitle = "Relatório de Processamento Analítico"
        RptGeral.Destination = crptToWindow
        RptGeral.WindowState = crptMaximized
        RptGeral.Action = 1
    Else
        Screen.MousePointer = vbDefault
        MsgBox "Não existe movimento para o período atual!", vbInformation, App.Title
    End If
    
    Set qryVerDadosRelanali = Nothing
    Set RsVerDadosRelanali = Nothing
    
    RptGeral.ReportFileName = Empty
    RptGeral.Formulas(0) = Empty
    RptGeral.StoredProcParam(0) = Empty
    RptGeral.StoredProcParam(1) = Empty
    
    Screen.MousePointer = vbDefault
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
End Sub
Private Sub mnuCapCaptura_Click(Index As Integer)
    Captura.Show vbModal, Me
End Sub

Private Sub mnuCapCtrQualidade_Click(Index As Integer)
    ControleQualidade.Show vbModal, Me
End Sub

Private Sub mnuComplementacao_Click(Index As Integer)
    Complementacao.Show vbModal, Me
End Sub

Private Sub mnuConConsulta_Click(Index As Integer)
    Consulta.Show vbModal, Me
End Sub
Private Sub mnuConRelRecepcao_Click(Index As Integer)

    Dim qryVerDadosrecep  As rdoQuery
    Dim RsVerDadosRecep   As rdoResultset
    
    Set qryVerDadosrecep = Geral.Banco.CreateQuery("", "{call relrecep(?)}")
    
    Screen.MousePointer = vbHourglass
    
    With qryVerDadosrecep
        .rdoParameters(0).Value = Geral.DataProcessamento
        Set RsVerDadosRecep = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsVerDadosRecep.RowCount <> 0 Then
        RptGeral.Connect = Geral.StringConexao
        
        If Geral.Backup Then
            RptGeral.ReportFileName = App.path & "\RelRecepBk.rpt"
        Else
            RptGeral.ReportFileName = App.path & "\RelRecepProd.rpt"
        End If
        
        RptGeral.Formulas(0) = "AgenciaCentral    = '" & Geral.AgenciaCentral & "'"
        RptGeral.StoredProcParam(0) = Geral.DataProcessamento
        RptGeral.WindowTitle = "Relatório de Recepção"
        RptGeral.Destination = crptToWindow
        RptGeral.WindowState = crptMaximized
        RptGeral.PrintReport
    Else
        Screen.MousePointer = vbDefault
        MsgBox "Não existem capas recepcionadas para período atual!", vbInformation, App.Title
    End If
    
    RptGeral.ReportFileName = Empty
    RptGeral.Formulas(0) = Empty
    RptGeral.StoredProcParam(0) = Empty
    RptGeral.StoredProcParam(1) = Empty
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub MnuEnvelopesFininvestProc_Click(Index As Integer)

    Dim qryEnvelopesFininvest  As rdoQuery
    Dim RsEnvelopesFininvest   As rdoResultset
    Dim sTitulo As String
    
    sTitulo = "Relatório de Envelopes Fininvest Recepcionados"
    Set qryEnvelopesFininvest = Geral.Banco.CreateQuery("", "{call GetEnvelopeFininvest(?)}")
    
    Screen.MousePointer = vbHourglass
    With qryEnvelopesFininvest
        .rdoParameters(0).Value = Geral.DataProcessamento
        Set RsEnvelopesFininvest = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsEnvelopesFininvest.RowCount <> 0 Then
        RptGeral.Connect = Geral.StringConexao
        
        If Geral.Backup Then
            RptGeral.ReportFileName = App.path & "\RelEnvelopesFininvestBK.rpt"
        Else
            RptGeral.ReportFileName = App.path & "\RelEnvelopesFininvestProd.rpt"
        End If
        
        RptGeral.Formulas(0) = "AgenciaCentral    = '" & Geral.AgenciaCentral & "'"
        RptGeral.Formulas(1) = "DataMovto = '" & Right(Geral.DataProcessamento, 2) & "/" & Mid(Geral.DataProcessamento, 5, 2) & "/" & Left(Geral.DataProcessamento, 4) & "'"
        
        RptGeral.StoredProcParam(0) = Geral.DataProcessamento
        RptGeral.Destination = crptToWindow
        RptGeral.WindowState = crptMaximized
        RptGeral.WindowTitle = sTitulo
        RptGeral.Action = 1
    Else
        Screen.MousePointer = vbDefault
        MsgBox "Não existem informações suficientes para emissão deste Relatório!" & vbCr & sTitulo, vbInformation, App.Title
    End If
    
    RptGeral.ReportFileName = Empty
    RptGeral.Formulas(0) = Empty
    RptGeral.Formulas(1) = Empty
    RptGeral.StoredProcParam(0) = Empty
    Screen.MousePointer = vbDefault

End Sub

Private Sub MnuEnvMalEx_Click(Index As Integer)

    Dim qryVerDadosMalEnvDel  As rdoQuery
    Dim RsVerDadosMalEnvDel   As rdoResultset
    
    Set qryVerDadosMalEnvDel = Geral.Banco.CreateQuery("", "{call GetCapasDeletadas(?)}")
    
    Screen.MousePointer = vbHourglass
    
    With qryVerDadosMalEnvDel
        .rdoParameters(0).Value = Geral.DataProcessamento
        Set RsVerDadosMalEnvDel = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsVerDadosMalEnvDel.RowCount <> 0 Then
        RptGeral.Connect = Geral.StringConexao
        
        If Geral.Backup Then
            RptGeral.ReportFileName = App.path & "\RelEnvMalExcBk.rpt"
        Else
            RptGeral.ReportFileName = App.path & "\RelEnvMalExcProd.rpt"
        End If
        
        RptGeral.Formulas(0) = "AgenciaCentral    = '" & Geral.AgenciaCentral & "'"
        RptGeral.StoredProcParam(0) = Geral.DataProcessamento
        RptGeral.Destination = crptToWindow
        RptGeral.WindowState = crptMaximized
        RptGeral.WindowTitle = "Relatório de Capa de Envelope/Malote Excluidos"
        RptGeral.Action = 1
    Else
        Screen.MousePointer = vbDefault
        MsgBox "Não existem capas excluidas para o período atual!", vbInformation, App.Title
    End If
    
    RptGeral.ReportFileName = Empty
    RptGeral.Formulas(0) = Empty
    RptGeral.StoredProcParam(0) = Empty
    
    Screen.MousePointer = vbDefault
  
End Sub
Private Sub mnuExpedicao_Click(Index As Integer)
    Dim ret_imp As Integer

    'Inicia autenticadora
    ''''''''''''''''''''''''''''''''''''
    ' Verifica se é impressora IBM (1) '
    ' ou PROCOMP (2)                   '
    ''''''''''''''''''''''''''''''''''''
    If Geral.autenticadora <> 0 Then
        On Error GoTo ErroAutentica
        
        Screen.MousePointer = vbHourglass
        ret_imp = Autentica.Inicia()
        Screen.MousePointer = vbDefault
        
        On Error GoTo 0
        
        If (ret_imp <> 0) Then
            MsgBox "A Autenticadora não está respondendo. Verifique se ela está ligada!", vbExclamation + vbOKOnly, App.Title
        End If
    End If
    
    Expedicao.Show vbModal, Me
    
    If Geral.autenticadora <> 0 Then
        Autentica.Finaliza
    End If
    
    Exit Sub
    
ErroAutentica:
    Screen.MousePointer = vbDefault
    MsgBox "Não foi possível iniciar a Autenticadora. Verifique se o arquivo .DLL da autenticadora se encontra no diretório do Windows.", vbInformation + vbOKOnly, App.Title
    
End Sub
Private Sub mnuIlegiveis_Click(Index As Integer)
    Ilegiveis.Show vbModal, Me
End Sub

Private Sub mnuSupMensagensRobo_Click(Index As Integer)

    Dim qryGetMensagemRobo      As RDO.rdoQuery
    Dim rstMensagemRobo         As RDO.rdoResultset
    
    Set qryGetMensagemRobo = Geral.Banco.CreateQuery("", "{call GetMensagemRobo(?,?)}")
    '''''''''''''''''''
    'Configura a Query'
    '''''''''''''''''''
    qryGetMensagemRobo.rdoParameters(0) = Geral.DataProcessamento
    qryGetMensagemRobo.rdoParameters(1) = Geral.idUsuario
    
    '''''''''
    'Executa'
    '''''''''
    Set rstMensagemRobo = qryGetMensagemRobo.OpenResultset(rdOpenStatic, rdConcurReadOnly)
    ''''''''''''''''''''''''''''''
    'Se tem mensagem, abre a tela'
    ''''''''''''''''''''''''''''''
    If Not rstMensagemRobo.EOF() Then
        Call MensagemRobo.ShowModal(rstMensagemRobo)
    Else
        MsgBox "Não há mensagem de retorno enviada pelo robô.", vbInformation, App.Title
    End If
    
    rstMensagemRobo.Close

End Sub

Private Sub mnuMotivoDoctosIlegiveis_Click(Index As Integer)

On Error GoTo Err_mnuMotivoDoctosIlegiveis
    'Aumenta timeout devido ao processamento demorado da Procedure
    
    With RptGeral

        .Connect = Geral.Banco.Connect
        If Geral.Backup Then
            .ReportFileName = App.path & "\RelMotivoDoctosIlegiveisBK.rpt"
        Else
            .ReportFileName = App.path & "\RelMotivoDoctosIlegiveisProd.rpt"
        End If
        
        .Formulas(0) = "AgenciaCentral = '" & Geral.AgenciaCentral & "'"
        .Formulas(1) = "DataMovimento  = '" & DataDD_MM_AAAA(Geral.DataProcessamento) & "'"

        .StoredProcParam(0) = Geral.DataProcessamento
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = "Relação de Documentos Enviados para Ilegíveis"
        .Action = 1
        
        .ReportFileName = Empty
        .Connect = Empty
        .Formulas(0) = Empty
        .Formulas(1) = Empty
        .StoredProcParam(0) = Empty
        .Destination = Empty
        .WindowState = Empty
        .WindowTitle = Empty
        
        Screen.MousePointer = vbDefault
    End With
    
Exit_mnuMotivoDoctosIlegiveis:
    Screen.MousePointer = vbDefault
    Exit Sub

Err_mnuMotivoDoctosIlegiveis:
    Beep
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, App.Title
    GoTo Exit_mnuMotivoDoctosIlegiveis

End Sub

Private Sub mnuRelAcompProducao_Click(Index As Integer)
    Dim qryVerificarMovto   As rdoQuery
    Dim RsVerificarMovto    As rdoResultset
    Dim QryTimeOut          As Variant
    
On Error GoTo Err_mnuRelAcompProducao
    'Aumenta timeout devido ao processamento demorado da Procedure
    QryTimeOut = Geral.Banco.QueryTimeout
    Geral.Banco.QueryTimeout = 300
    
    Set qryVerificarMovto = Geral.Banco.CreateQuery("", "{call GetAcompanhamentoProducao(?,?,?)}")
    
    Screen.MousePointer = vbHourglass
    With qryVerificarMovto
        .rdoParameters(0).Value = Geral.DataProcessamento
        .rdoParameters(1).Value = 0     'Relatório de meia em meia hora (0)=Não  (1)=Sim
        .rdoParameters(2).Value = 1     'Somente verificação de movto   (0)=Não  (1)=Sim
        Set RsVerificarMovto = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    With RptGeral

        If qryVerificarMovto.rdoParameters(2).Value > 0 Then
            .Connect = Geral.Banco.Connect
            If Geral.Backup Then
                .ReportFileName = App.path & "\RelAcompProducaoBK.rpt"
            Else
                .ReportFileName = App.path & "\RelAcompProducao.rpt"
            End If
            
            .Formulas(0) = "AgProcessadora    = '" & Geral.AgenciaCentral & "'"
            .Formulas(1) = "DataMovto    = '" & DataDD_MM_AAAA(Geral.DataProcessamento) & "'"
    
            .StoredProcParam(0) = Geral.DataProcessamento
            .StoredProcParam(1) = 0     'Relatório de meia em meia hora (0)=Não  (1)=Sim
            .StoredProcParam(2) = 0     'Somente verificação de movto   (0)=Não  (1)=Sim
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .WindowTitle = "Relatório de Acompanhamento da Produção"
            .Action = 1
        Else
            Screen.MousePointer = vbDefault

            MsgBox "Não existe movimento na data " & DataDD_MM_AAAA(Geral.DataProcessamento) & _
            " para emissão do Relatório!", vbInformation, App.Title & " ( " & "Relatório de Acompanhamento da Produção" & " )"
        End If
        
        .ReportFileName = Empty
        .Connect = Empty
        .Formulas(0) = Empty
        .Formulas(1) = Empty
        .StoredProcParam(0) = Empty
        .StoredProcParam(1) = Empty
        .StoredProcParam(2) = Empty
        .Destination = Empty
        .WindowState = Empty
        .WindowTitle = Empty
        
        Screen.MousePointer = vbDefault
    End With
    
Exit_mnuRelAcompProducao:
    Screen.MousePointer = vbDefault
    Geral.Banco.QueryTimeout = QryTimeOut
    qryVerificarMovto.Close
    If Not (RsVerificarMovto Is Nothing) Then Set RsVerificarMovto = Nothing
    Exit Sub

Err_mnuRelAcompProducao:
    Beep
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, App.Title
    GoTo Exit_mnuRelAcompProducao

End Sub

Private Sub mnuRelChqUBBCompValor_Click(Index As Integer)

    Dim qryChqCompValor      As rdoQuery
    Dim RsChqCompValor       As rdoResultset
    Dim sTitulo As String
    
    sTitulo = "Relatório de Cheques UBB para Compensação por Valor"
    Set qryChqCompValor = Geral.Banco.CreateQuery("", "{call GetChequeUBBCompensado(?)}")
    
    Screen.MousePointer = vbHourglass
    With qryChqCompValor
        .rdoParameters(0).Value = Geral.DataProcessamento
        Set RsChqCompValor = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsChqCompValor.RowCount <> 0 Then
        RptGeral.Connect = Geral.StringConexao
        
        If Geral.Backup Then
            RptGeral.ReportFileName = App.path & "\ChequesUBBCompValorBk.rpt"
        Else
            RptGeral.ReportFileName = App.path & "\ChequesUBBCompValorProd.rpt"
        End If
        
        RptGeral.Formulas(0) = "AgenciaCentral    = '" & Geral.AgenciaCentral & "'"
        RptGeral.Formulas(1) = "DataMovto = '" & Right(Geral.DataProcessamento, 2) & "/" & Mid(Geral.DataProcessamento, 5, 2) & "/" & Left(Geral.DataProcessamento, 4) & "'"
        
        RptGeral.StoredProcParam(0) = Geral.DataProcessamento
        RptGeral.Destination = crptToWindow
        RptGeral.WindowState = crptMaximized
        RptGeral.WindowTitle = sTitulo
        RptGeral.Action = 1
    Else
        Screen.MousePointer = vbDefault
        MsgBox "Não existem informações suficientes para emissão deste Relatório!" & vbCr & sTitulo, vbInformation, App.Title
    End If
    
    RptGeral.ReportFileName = Empty
    RptGeral.Formulas(0) = Empty
    RptGeral.Formulas(1) = Empty
    RptGeral.StoredProcParam(0) = Empty
    Screen.MousePointer = vbDefault


End Sub

Private Sub mnuRelConcessionarias_Click(Index As Integer)
    
    frmRelConcessionarias.Show vbModal, Me

End Sub

Private Sub mnuRelDiferencasCaixa_Click(Index As Integer)

Dim qryDiferencaCaixa  As rdoQuery
Dim RsDiferencaCaixa   As rdoResultset
    
    Set qryDiferencaCaixa = Geral.Banco.CreateQuery("", "{call GetRelDiferencaCaixa(?)}")
          
    Screen.MousePointer = vbHourglass
    
    With qryDiferencaCaixa
        .rdoParameters(0).Value = Geral.DataProcessamento
        Set RsDiferencaCaixa = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If Not RsDiferencaCaixa.EOF Then
        RptGeral.Connect = Geral.StringConexao
        
        If Geral.Backup Then
           RptGeral.ReportFileName = App.path & "\RelDiferencaCaixaBk.rpt"
        Else
           RptGeral.ReportFileName = App.path & "\RelDiferencaCaixa.rpt"
        End If
                
        RptGeral.Formulas(0) = "DataProcessamento  = '" & Right(Geral.DataProcessamento, 2) & "/" & _
                                                      Mid(Geral.DataProcessamento, 5, 2) & "/" & _
                                                      Left(Geral.DataProcessamento, 4) & "'"
        RptGeral.StoredProcParam(0) = Geral.DataProcessamento
        RptGeral.Destination = crptToWindow
        RptGeral.WindowState = crptMaximized
        RptGeral.WindowTitle = "Relatório de Diferenças no Fechamento do Caixa"
        RptGeral.Action = 1
    Else
        Screen.MousePointer = vbDefault
        MsgBox "Não existem diferenças no período atual.", vbInformation, App.Title
    End If
    
    RptGeral.ReportFileName = Empty
    RptGeral.Formulas(0) = Empty
    RptGeral.StoredProcParam(0) = Empty
        
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub MnuRelEnvFininvestDevolvido_Click(Index As Integer)
    
    Dim qryEnvelopesFininvest  As rdoQuery
    Dim RsEnvelopesFininvest   As rdoResultset
    Dim sTitulo As String
    
    sTitulo = "Relatório de Envelopes Fininvest Devolvidos"
    Set qryEnvelopesFininvest = Geral.Banco.CreateQuery("", "{call GetEnvelopeFininvestDevolvido(?,?)}")
    
    Screen.MousePointer = vbHourglass
    With qryEnvelopesFininvest
        .rdoParameters(0).Value = Geral.DataProcessamento
        .rdoParameters(1).Value = 1         'Somente verificação de movto   (0)=Não  (1)=Sim
        Set RsEnvelopesFininvest = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If qryEnvelopesFininvest.rdoParameters(1).Value > 0 Then
        RptGeral.Connect = Geral.StringConexao
        
        If Geral.Backup Then
            RptGeral.ReportFileName = App.path & "\RelEnvFininvestDevolvidosBK.rpt"
        Else
            RptGeral.ReportFileName = App.path & "\RelEnvFininvestDevolvidosProd.rpt"
        End If
        
        RptGeral.Formulas(0) = "AgenciaCentral    = '" & Geral.AgenciaCentral & "'"
        RptGeral.Formulas(1) = "DataMovto = '" & Right(Geral.DataProcessamento, 2) & "/" & Mid(Geral.DataProcessamento, 5, 2) & "/" & Left(Geral.DataProcessamento, 4) & "'"
        
        RptGeral.StoredProcParam(0) = Geral.DataProcessamento
        RptGeral.StoredProcParam(1) = 0     'Somente verificação de movto   (0)=Não  (1)=Sim
        RptGeral.Destination = crptToWindow
        RptGeral.WindowState = crptMaximized
        RptGeral.WindowTitle = sTitulo
        RptGeral.Action = 1
    Else
        Screen.MousePointer = vbDefault
        MsgBox "Não existem informações suficientes para emissão deste Relatório!" & vbCr & sTitulo, vbInformation, App.Title
    End If
    
    RptGeral.ReportFileName = Empty
    RptGeral.Formulas(0) = Empty
    RptGeral.Formulas(1) = Empty
    RptGeral.StoredProcParam(0) = Empty
    RptGeral.StoredProcParam(1) = Empty
    Screen.MousePointer = vbDefault

End Sub

Private Sub mnuRelEnvFininvestPorAgencia_Click(Index As Integer)

    Dim qryEnvelopesFininvest  As rdoQuery
    Dim RsEnvelopesFininvest   As rdoResultset
    Dim sTitulo As String
    
    sTitulo = "Relatório Total de Envelopes Fininvest por Agência"
    Set qryEnvelopesFininvest = Geral.Banco.CreateQuery("", "{call GetTotalEnvFininvestPorCapa(?,?)}")
    
    Screen.MousePointer = vbHourglass
    With qryEnvelopesFininvest
        .rdoParameters(0).Value = Geral.DataProcessamento
        .rdoParameters(1).Value = 1         'Somente verificação de movto   (0)=Não  (1)=Sim
        Set RsEnvelopesFininvest = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If qryEnvelopesFininvest.rdoParameters(1).Value > 0 Then
        RptGeral.Connect = Geral.StringConexao
        
        If Geral.Backup Then
            RptGeral.ReportFileName = App.path & "\RelEnvFininvestPorAgenciaBK.rpt"
        Else
            RptGeral.ReportFileName = App.path & "\RelEnvFininvestPorAgenciaProd.rpt"
        End If
        
        RptGeral.Formulas(0) = "AgenciaCentral    = '" & Geral.AgenciaCentral & "'"
        RptGeral.Formulas(1) = "DataMovto = '" & Right(Geral.DataProcessamento, 2) & "/" & Mid(Geral.DataProcessamento, 5, 2) & "/" & Left(Geral.DataProcessamento, 4) & "'"
        
        RptGeral.StoredProcParam(0) = Geral.DataProcessamento
        RptGeral.StoredProcParam(1) = 0     'Somente verificação de movto   (0)=Não  (1)=Sim
        RptGeral.Destination = crptToWindow
        RptGeral.WindowState = crptMaximized
        RptGeral.WindowTitle = sTitulo
        RptGeral.Action = 1
    Else
        Screen.MousePointer = vbDefault
        MsgBox "Não existem informações suficientes para emissão deste Relatório!" & vbCr & sTitulo, vbInformation, App.Title
    End If
    
    RptGeral.ReportFileName = Empty
    RptGeral.Formulas(0) = Empty
    RptGeral.Formulas(1) = Empty
    RptGeral.StoredProcParam(0) = Empty
    RptGeral.StoredProcParam(1) = Empty
    Screen.MousePointer = vbDefault


End Sub

Private Sub mnuRelEstDocAgencia_Click(Index As Integer)

    Dim qryEstDocAgencia     As rdoQuery
    Dim RsEstDocAgencia      As rdoResultset
    Dim sTitulo As String
    
    sTitulo = "Estatística de Documentos por Agência"
    Set qryEstDocAgencia = Geral.Banco.CreateQuery("", "{call GetAgenEstatistica(?)}")
    
    Screen.MousePointer = vbHourglass
    With qryEstDocAgencia
        .rdoParameters(0).Value = Geral.DataProcessamento
        Set RsEstDocAgencia = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsEstDocAgencia.RowCount <> 0 Then
        RptGeral.Connect = Geral.StringConexao
        
        If Geral.Backup Then
            RptGeral.ReportFileName = App.path & "\RelEstDocAgenciaBk.rpt"
        Else
            RptGeral.ReportFileName = App.path & "\RelEstDocAgenciaProd.rpt"
        End If
        
        RptGeral.Formulas(0) = "AgenciaCentral    = '" & Geral.AgenciaCentral & "'"
        RptGeral.Formulas(1) = "DataMovto = '" & Right(Geral.DataProcessamento, 2) & "/" & Mid(Geral.DataProcessamento, 5, 2) & "/" & Left(Geral.DataProcessamento, 4) & "'"
        
        RptGeral.StoredProcParam(0) = Geral.DataProcessamento
        RptGeral.Destination = crptToWindow
        RptGeral.WindowState = crptMaximized
        RptGeral.WindowTitle = sTitulo
        RptGeral.Action = 1
    Else
        Screen.MousePointer = vbDefault
        MsgBox "Não existem informações suficientes para emissão deste Relatório!" & vbCr & sTitulo, vbInformation, App.Title
    End If
    
    RptGeral.ReportFileName = Empty
    RptGeral.Formulas(0) = Empty
    RptGeral.Formulas(1) = Empty
    RptGeral.StoredProcParam(0) = Empty
    Screen.MousePointer = vbDefault

End Sub

Private Sub MnuRelLancamentoInterno_Click(Index As Integer)

    Dim qryLanctoInterno  As rdoQuery
    Dim RsLanctoInterno   As rdoResultset
    
    Set qryLanctoInterno = Geral.Banco.CreateQuery("", "{call GetRelacaoLanctoInterno(?,?)}")
          
    Screen.MousePointer = vbHourglass
    
    With qryLanctoInterno
        .rdoParameters(0).Value = Geral.DataProcessamento
        .rdoParameters(1).Value = "C"   'Somente contagem
        Set RsLanctoInterno = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsLanctoInterno(0).Value > 0 Then
        RptGeral.Connect = Geral.StringConexao
        
        If Geral.Backup Then
            RptGeral.ReportFileName = App.path & "\RelLanctoInternoBk.rpt"
        Else
            RptGeral.ReportFileName = App.path & "\RelLanctoInternoProd.rpt"
        End If
        
        RptGeral.Formulas(0) = "AgenciaCentral = '" & Geral.AgenciaCentral & "'"
        RptGeral.Formulas(1) = "DataMovimento  = '" & Right(Geral.DataProcessamento, 2) & "/" & _
                                                      Mid(Geral.DataProcessamento, 5, 2) & "/" & _
                                                      Left(Geral.DataProcessamento, 4) & "'"
        RptGeral.StoredProcParam(0) = Geral.DataProcessamento
        RptGeral.StoredProcParam(1) = "R"
        RptGeral.Destination = crptToWindow
        RptGeral.WindowState = crptMaximized
        RptGeral.WindowTitle = "Relatório de Lançamentos internos Processados"
        RptGeral.Action = 1
    Else
        Screen.MousePointer = vbDefault
        MsgBox "Não existe lançamento interno processado no período atual.", vbInformation, App.Title
    End If
    
    RptGeral.ReportFileName = Empty
    RptGeral.Formulas(0) = Empty
    RptGeral.Formulas(1) = Empty
    RptGeral.StoredProcParam(0) = Empty
    RptGeral.StoredProcParam(1) = Empty
    
    Screen.MousePointer = vbDefault

End Sub
Private Sub mnuRelMovtoChqUBBComp_Click(Index As Integer)
    
    Dim qryVerDadosCompAnal  As rdoQuery
    Dim RsVerDadosComAnal    As rdoResultset
    Dim sTitulo As String
    
    sTitulo = "Relatório de Cheques UBB para compensação"
    Set qryVerDadosCompAnal = Geral.Banco.CreateQuery("", "{call GetMovtoChqUBBCompensado(?)}")
    
    Screen.MousePointer = vbHourglass
    With qryVerDadosCompAnal
        .rdoParameters(0).Value = Geral.DataProcessamento
        Set RsVerDadosComAnal = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsVerDadosComAnal.RowCount <> 0 Then
        RptGeral.Connect = Geral.StringConexao
        
        If Geral.Backup Then
            RptGeral.ReportFileName = App.path & "\ChequesUBBCompensacaoBk.rpt"
        Else
            RptGeral.ReportFileName = App.path & "\ChequesUBBCompensacaoProd.rpt"
        End If
        
        RptGeral.Formulas(0) = "AgenciaCentral    = '" & Geral.AgenciaCentral & "'"
        RptGeral.Formulas(1) = "DataMovto = '" & Right(Geral.DataProcessamento, 2) & "/" & Mid(Geral.DataProcessamento, 5, 2) & "/" & Left(Geral.DataProcessamento, 4) & "'"
        
        RptGeral.StoredProcParam(0) = Geral.DataProcessamento
        RptGeral.Destination = crptToWindow
        RptGeral.WindowState = crptMaximized
        RptGeral.WindowTitle = sTitulo
        RptGeral.Action = 1
    Else
        Screen.MousePointer = vbDefault
        MsgBox "Não existem informações suficientes para emissão deste Relatório!" & vbCr & sTitulo, vbInformation, App.Title
    End If
    
    RptGeral.ReportFileName = Empty
    RptGeral.Formulas(0) = Empty
    RptGeral.Formulas(1) = Empty
    RptGeral.StoredProcParam(0) = Empty
    Screen.MousePointer = vbDefault

End Sub
Private Sub mnuProvaZero_Click(Index As Integer)
    ProvaZero.Show vbModal, Me
End Sub
Private Sub mnuRecCadCapas_Click(Index As Integer)
    FrmRecCapa.Show vbModal, Me
End Sub
Private Sub mnuRecRecepcao_Click(Index As Integer)
    Dim ret_imp As Integer

    'Inicia autenticadora
    ''''''''''''''''''''''''''''''''''''
    ' Verifica se é impressora IBM (1) '
    ' ou PROCOMP (2)                   '
    ''''''''''''''''''''''''''''''''''''
    If Geral.autenticadora <> 0 Then
        On Error GoTo ErroAutentica
        
        Screen.MousePointer = vbHourglass
        ret_imp = Autentica.Inicia()
        Screen.MousePointer = vbDefault
        
        On Error GoTo 0
        
        If (ret_imp <> 0) Then
            MsgBox "A Autenticadora não está respondendo. Verifique se ela está ligada!", vbExclamation + vbOKOnly, App.Title
        End If
    End If
    
    Recepcao.Show vbModal, Me
    
    If Geral.autenticadora <> 0 Then
        Autentica.Finaliza
    End If
    
    Exit Sub
    
ErroAutentica:
    Screen.MousePointer = vbDefault
    MsgBox "Não foi possível iniciar a Autenticadora. Verifique se o arquivo .DLL da autenticadora se encontra no diretório do Windows.", vbInformation + vbOKOnly, App.Title
    
End Sub
Private Sub mnuRecRegistroOcorrencia_Click(Index As Integer)
    Dim ret_imp As Integer

    'Inicia autenticadora
    ''''''''''''''''''''''''''''''''''''
    ' Verifica se é impressora IBM (1) '
    ' ou PROCOMP (2)                   '
    ''''''''''''''''''''''''''''''''''''
    If Geral.autenticadora = 0 Then
        MsgBox "Para Registro de Ocorrência é necessário que alguma autenticadora esteja ligada a estação. Verifique se existe alguma autenticadora ligada a esta estação e a selecione no módulo de Parâmetros do Sistema.", vbExclamation + vbOKOnly, App.Title
        Exit Sub
    Else
        On Error GoTo ErroAutentica
        
        Screen.MousePointer = vbHourglass
        ret_imp = Autentica.Inicia()
        Screen.MousePointer = vbDefault
        
        On Error GoTo 0
        
        If (ret_imp <> 0) Then
            MsgBox "A Autenticadora não está respondendo. Verifique se ela está ligada!", vbExclamation + vbOKOnly, App.Title
            Exit Sub
        End If
    End If
    
    FrmRegOcorr.Show vbModal, Me
    
    Autentica.Finaliza
    
    Exit Sub
    
ErroAutentica:
    Screen.MousePointer = vbDefault
    MsgBox "Não foi possível iniciar a Autenticadora. Verifique se o arquivo .DLL da autenticadora se encontra no diretório do Windows.", vbInformation + vbOKOnly, App.Title
    
End Sub

Private Sub mnurelconschqscomp_Click()
    ConsultaChequeComp.Show vbModal, Me
End Sub
Private Sub MnuRelMovComp_Click(Index As Integer)

    Dim qryVerDadosCompAnal  As rdoQuery
    Dim RsVerDadosComAnal    As rdoResultset
    
    Set qryVerDadosCompAnal = Geral.Banco.CreateQuery("", "{call GetMovCompAnalitico(?)}")
    
    Screen.MousePointer = vbHourglass
    With qryVerDadosCompAnal
        .rdoParameters(0).Value = Geral.DataProcessamento
        Set RsVerDadosComAnal = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsVerDadosComAnal.RowCount <> 0 Then
        RptGeral.Connect = Geral.StringConexao
        
        If Geral.Backup Then
            RptGeral.ReportFileName = App.path & "\MovCompBk.rpt"
        Else
            RptGeral.ReportFileName = App.path & "\MovCompProd.rpt"
        End If
        
        RptGeral.Formulas(0) = "AgenciaCentral    = '" & Geral.AgenciaCentral & "'"
        
        RptGeral.StoredProcParam(0) = Geral.DataProcessamento
        RptGeral.Destination = crptToWindow
        RptGeral.WindowState = crptMaximized
        RptGeral.WindowTitle = "Relatório de Compensação - Analítico"
        RptGeral.Action = 1
    Else
        Screen.MousePointer = vbDefault
        MsgBox "Não existem informações suficientes para emissão deste Relatório!" & vbCr & " Relatório de Compensação - Analítico", vbInformation, App.Title
    End If
    
    RptGeral.ReportFileName = Empty
    RptGeral.Formulas(0) = Empty
    RptGeral.StoredProcParam(0) = Empty
    Screen.MousePointer = vbDefault
  
End Sub

Private Sub mnuRelPercDoctoModulo_Click(Index As Integer)

    frmRelPercDoctosPorModulo.Show vbModal, Me

End Sub

Private Sub MnuRelSegDoctos_Click(Index As Integer)

    Dim qryVerDadosSegDoctos As rdoQuery
    Dim RsVerDadosSegDoctos  As rdoResultset
    Dim QryTimeOut          As Variant
    
    On Error GoTo Erro_AbreRelatorio
    
    'Aumenta timeout devido ao processamento demorado da Procedure
    QryTimeOut = Geral.Banco.QueryTimeout
    Geral.Banco.QueryTimeout = 300
    
    Set qryVerDadosSegDoctos = Geral.Banco.CreateQuery("", "{call GetSegmentacaoDoc (?)}")
    
    Screen.MousePointer = vbHourglass
    
    With qryVerDadosSegDoctos
        .rdoParameters(0).Value = Geral.DataProcessamento
        Set RsVerDadosSegDoctos = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsVerDadosSegDoctos.RowCount <> 0 Then
        RptGeral.Connect = Geral.StringConexao
        
        If Geral.Backup Then
            RptGeral.ReportFileName = App.path & "\SegDoctosBk.rpt"
        Else
            RptGeral.ReportFileName = App.path & "\SegDoctosProd.rpt"
        End If
        
        RptGeral.Formulas(0) = "AgenciaCentral    = '" & Geral.AgenciaCentral & "'"
        RptGeral.StoredProcParam(0) = Geral.DataProcessamento
        RptGeral.Destination = crptToWindow
        RptGeral.WindowState = crptMaximized
        RptGeral.WindowTitle = "Relatório de Segmentação por Documento / Agência"
        RptGeral.PrintReport
    Else
        Screen.MousePointer = vbDefault
        MsgBox "Sistema não possui informações Suficientes para emissão deste Relatório!" & vbCr & "Relatório de Segmentação por Documento / Agência", vbInformation, App.Title
    End If
    
    RptGeral.ReportFileName = Empty
    RptGeral.Formulas(0) = Empty
    RptGeral.StoredProcParam(0) = Empty
    
    
    Screen.MousePointer = vbDefault
    
    'Retorno timeout default
    Geral.Banco.QueryTimeout = QryTimeOut
    
    Exit Sub
    
Erro_AbreRelatorio:
    
    Screen.MousePointer = vbDefault
    'Retorno timeout default
    Geral.Banco.QueryTimeout = QryTimeOut
    
    Call TratamentoErro("Não foi possível abrir o relatório.", Err, rdoErrors)
  
End Sub

Private Sub mnuRelRelacao_TC_AR_Click(Index As Integer)
    
    Dim qryRelacao  As rdoQuery
    Dim rsRelacao   As rdoResultset
    Dim sTitulo As String
    
    sTitulo = "Relatório de TC's e AR's"
    Set qryRelacao = Geral.Banco.CreateQuery("", "{call GetRelacao_TC_AR(?,?)}")

    Screen.MousePointer = vbHourglass
    With qryRelacao
        .rdoParameters(0).Value = Geral.DataProcessamento
        .rdoParameters(1).Value = 1         'Somente verificação de movto   (0)=Não  (1)=Sim
        Set rsRelacao = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If qryRelacao.rdoParameters(1).Value > 0 Then
        RptGeral.Connect = Geral.StringConexao
        
        If Geral.Backup Then
            RptGeral.ReportFileName = App.path & "\RelRelacao_TC_AR_Bk.rpt"
        Else
            RptGeral.ReportFileName = App.path & "\RelRelacao_TC_AR_Prod.rpt"
        End If
        
        RptGeral.Formulas(0) = "AgenciaCentral    = '" & Geral.AgenciaCentral & "'"
        RptGeral.Formulas(1) = "DataMovto = '" & Right(Geral.DataProcessamento, 2) & "/" & Mid(Geral.DataProcessamento, 5, 2) & "/" & Left(Geral.DataProcessamento, 4) & "'"
        
        RptGeral.StoredProcParam(0) = Geral.DataProcessamento
        RptGeral.StoredProcParam(1) = 0         'Somente verificação de movto   (0)=Não  (1)=Sim
        RptGeral.Destination = crptToWindow
        RptGeral.WindowState = crptMaximized
        RptGeral.WindowTitle = sTitulo
        RptGeral.Action = 1
    Else
        Screen.MousePointer = vbDefault
        MsgBox "Não existem informações suficientes para emissão deste Relatório!" & vbCr & sTitulo, vbInformation, App.Title
    End If
    
    RptGeral.ReportFileName = Empty
    RptGeral.Formulas(0) = Empty
    RptGeral.Formulas(1) = Empty
    RptGeral.StoredProcParam(0) = Empty
    RptGeral.StoredProcParam(1) = Empty
    Screen.MousePointer = vbDefault

End Sub

Private Sub mnuRelTotalConsolidado_Click(Index As Integer)
    
    frmRelTotalConsolidado.Show vbModal, Me
    
End Sub

Private Sub mnuRelTotalPorTipoDocto_Click(Index As Integer)

    frmRelTotalPorDocto.Show vbModal, Me

End Sub

Private Sub mnuRelTotDocumentoPorCliente_Click(Index As Integer)

    frmRelTotalDoctoPorCliente.Show vbModal, Me

End Sub

Friend Sub mnuSair_Click(Index As Integer)
    Call RemoveAtividade
    Unload Me
End Sub

Private Sub mnuSobre_Click(Index As Integer)
    frmSobre.Show vbModal, Me
End Sub

Private Sub mnuSupAcompanhamento_Click(Index As Integer)
    Estatistica.Show vbModal, Me
End Sub
Private Sub MnuSupAcompAtividade_Click(Index As Integer)
    'Inclusão de chamada a rotina AtualizaAtividade
    Call AtualizaAtividade(13)
    FrmAcompAtividade.Show vbModal, Me
End Sub
Private Sub mnusupacompexped_Click(Index As Integer)
    AcompExpedicao.Show vbModal, Me
End Sub
Private Sub MnuSupAcompRecepcao_Click(Index As Integer)
    FrmAcompRec.Show vbModal, Me
End Sub
Private Sub mnuSupAcompUsers_Click(Index As Integer)
    FrmAcompUsers.Show vbModal, Me
End Sub
Private Sub mnuSupAlcada_Click(Index As Integer)
    Alcada.Show vbModal, Me
End Sub

Private Sub mnuSupAlteraSenha_Click(Index As Integer)
    
    AlteraSenha.Show vbModal, Me
    
End Sub

Private Sub mnuSupAuditoria_Click(Index As Integer)
  FrmLogUsuario.Show vbModal, Me
End Sub
Private Sub mnuSupBloquearAplicacao_Click(Index As Integer)

    Dim bBloquear       As Boolean
    Dim qryUsuario      As rdoQuery
    Dim tbUsuario       As rdoResultset
    Dim eRetorno        As enumRetornoUsuario
    Dim strctUsuario    As TpUsuario
    Dim iRetornoLogin   As Integer
    
    bBloquear = CBool(Val(mnuSupBloquearAplicacao(0).Tag))

    'TAG = 0 -> Bloquear
    'TAG = 1 -> Desbloquear

    If bBloquear = False Then
        mnuSupBloquearAplicacao(0).Caption = "Desbloquear"
        mnuSupBloquearAplicacao(0).Tag = 1

        '''''''''''''''''''''''''''''''''
        '*     Desabilita os menus     *'
        '''''''''''''''''''''''''''''''''
        For Each Control In Principal.Controls
            If TypeName(Control) = "Menu" Then
                If Control.Index <> 0 Then
                    Control.Enabled = bBloquear
                End If
                'Desabilita opção de Alteração de Senha
                If LCase(Control.Name) = "mnusupalterasenha" Then Control.Enabled = bBloquear
             End If
            
        Next
        mnuSupBloquearAplicacao(0).Enabled = True
        
    Else
        
        Set qryUsuario = Geral.Banco.CreateQuery("", "{call GetUsuario (?)}")

        With qryUsuario
            .rdoParameters(0).Value = Trim(Geral.Usuario)
            Set tbUsuario = .OpenResultset(rdConcurReadOnly)
        End With

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Atribui o nome do usuario e senha para a tela de re-login'
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        strctUsuario.Usuario = Geral.Usuario
        
        If Not tbUsuario.EOF Then
            strctUsuario.Senha = tbUsuario!Senha
        ElseIf UCase(Trim(Geral.Usuario)) = "DESENV" Then
            strctUsuario.Senha = "DIAMANTE"
        End If
        
        iRetornoLogin = frmLogin.Senha_Ok(strctUsuario)
        
        'iRetornoLogin = 0 'Cancelar
        'iRetornoLogin = 1 'Senha ok
        'iRetornoLogin = 2 'Fechar aplicacao
       
        If iRetornoLogin = 0 Then
            'MsgBox "Não foi possível desbloquear a Aplicação.", vbExclamation
            Exit Sub
        ElseIf iRetornoLogin = 2 Then
            Exit Sub
        End If

        eRetorno = VerificaUsuario(tbUsuario, strctUsuario.Usuario, strctUsuario.Senha, Geral.Backup)

        If eRetorno = eSUPERVISOR Then
            Geral.idUsuario = 0
            SenhaOk = True
        ElseIf eRetorno = eNAO_EXISTENTE Then
            Beep
            MsgBox "Usuário não Cadastrado !", vbExclamation + vbOKOnly, App.Title
            With txtUsuario
                .SelStart = 0
                .SelLength = Len(Trim(.Text))
                .SetFocus
            End With
        ElseIf eRetorno = eSENHA_INCORRETA Then
            Beep
            MsgBox "Senha não Confere !", vbExclamation + vbOKOnly, App.Title
            With txtSenha
                .SelStart = 0
                .SelLength = Len(Trim(.Text))
                .SetFocus
            End With
        Else
            SenhaOk = True
            
            ' *******************************************************
            ' * Ajustando Menu de Opções para Permissões do Usuário *
            ' *******************************************************
            While Not tbUsuario.EOF
                '* Grava Informação de IdUsuario *'
                Geral.idUsuario = tbUsuario!idUsuario
                tbUsuario.MoveNext
            Wend
            
        End If
        
        tbUsuario.Close

        mnuSupBloquearAplicacao(0).Caption = "Bloquear"
        mnuSupBloquearAplicacao(0).Tag = 0
        
        DoEvents
        
    End If

End Sub
Private Sub mnuSupCadastroUsuario_Click(Index As Integer)
    CadUsuario.Show vbModal, Me
End Sub
Private Sub MnuSupConfAgCc_Click(Index As Integer)
    ConfirmaAgenciaConta.Show vbModal
End Sub
Private Sub MnuSupCorrecaoAgCc_Click(Index As Integer)
    CorrecaoAgenciaConta.Show vbModal, Me
End Sub

Private Sub mnuSupCSP_Click(Index As Integer)
    frmCSP.Show vbModal, Me
End Sub

Private Sub MnuSupEstorno_Click(Index As Integer)
    Estorno.Show vbModal, Me
End Sub
Private Sub mnuSupExclusao_Click(Index As Integer)
    FrmDelEnvMal.Show vbModal, Me
End Sub
Private Sub Mnusupfinalcapas_Click(Index As Integer)
    FinalizacaoCapa.Show vbModal, Me
End Sub
Private Sub mnuSupGrpMod_Click(Index As Integer)
    GrupoModulo.Show vbModal, Me
End Sub
Private Sub mnuSupParametros_Click(Index As Integer)
    Parametros.Show vbModal
End Sub
Private Sub mnuSupTrocarOrdem_Click(Index As Integer)
    frmTrocarOrdemDocumento.Show vbModal, Me
End Sub
Private Sub mnuSupVinculo_Click(Index As Integer)
    VinculoManual.Show vbModal, Me
End Sub
Private Sub MnuConRelValeTransporte_Click(Index As Integer)
  FrmRelPerArrecadacao.Show vbModal, Me
End Sub
Private Sub RelTotais_Click(Index As Integer)
  
CppagQtdeCapMal = 0:        CppagQtdePagMal = 0:    CppagQtdeChPagMal = 0
CppagQtdeDepMal = 0:        CppagQtdeChDepMal = 0
            
CpDepQtdeCapMal = 0:        CpDepQtdePagMal = 0:    CpDepQtdeChPagMal = 0
CpDepQtdeDepMal = 0:        CpDepQtdeChDepMal = 0
            
CpPagDepQtdeCapMal = 0:     CpPagDepQtdePagMal = 0: CpPagDepQtdeChPagMal = 0
CpPagDepQtdeDepMal = 0:     CpPagDepQtdeChDepMal = 0
            
CppagQtdeCapEnv = 0:        CppagQtdePagEnv = 0:    CppagQtdeChPagEnv = 0
CppagQtdeDepEnv = 0:        CppagQtdeChDepEnv = 0
            
CpDepQtdeCapEnv = 0:        CpDepQtdePagEnv = 0:    CpDepQtdeChPagEnv = 0
CpDepQtdeDepEnv = 0:        CpDepQtdeChDepEnv = 0
            
CpPagDepQtdeCapEnv = 0:     CpPagDepQtdePagEnv = 0: CpPagDepQtdeChPagEnv = 0
CpPagDepQtdeDepEnv = 0:     CpPagDepQtdeChDepEnv = 0
DataFormatada = 0

CpAjPagtoM = 0:             CpAjDeptoM = 0:         CpAjPagDepM = 0
CpAjDeptoE = 0:             CpAjPagtoE = 0:         CpAjPagDepE = 0
            
AgTransRobo0 = 0:           AgTransRobo1 = 0:       AgTransRobo2 = 0
AgTransRobo3 = 0:           AgTransRobo4 = 0
            
UltTransRobo0 = 0:          UltTransRobo1 = 0:      UltTransRobo2 = 0
UltTransRobo3 = 0:          UltTransRobo4 = 0:      HrUltTransRobo0 = 0
HrUltTransRobo1 = 0:        HrUltTransRobo2 = 0:    HrUltTransRobo3 = 0
HrUltTransRobo4 = 0
  
Call RelEstatistica

End Sub

