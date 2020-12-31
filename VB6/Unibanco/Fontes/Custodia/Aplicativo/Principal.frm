VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "*\A..\Scanner\SCANNER.vbp"
Begin VB.Form Principal 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Unibanco - Sistema de Captura - Menu Principal"
   ClientHeight    =   7290
   ClientLeft      =   300
   ClientTop       =   720
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   Picture         =   "Principal.frx":0000
   ScaleHeight     =   7290
   ScaleWidth      =   11610
   WindowState     =   2  'Maximized
   Begin Scanner_Control.Scanner Scanner 
      Left            =   10560
      Top             =   5400
      _ExtentX        =   979
      _ExtentY        =   873
      wordLenght      =   7
   End
   Begin VB.FileListBox filArquivosRecepcao 
      Archive         =   0   'False
      Enabled         =   0   'False
      Height          =   285
      Left            =   10590
      Normal          =   0   'False
      ReadOnly        =   0   'False
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   396
   End
   Begin Crystal.CrystalReport CrystalReport 
      Left            =   10560
      Top             =   6315
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   324
      Left            =   3504
      TabIndex        =   1
      Top             =   7656
      Width           =   4908
      _ExtentX        =   8652
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Min             =   1e-4
      Max             =   10
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar StatusBarPrincipal 
      Align           =   2  'Align Bottom
      Height          =   372
      Left            =   0
      TabIndex        =   0
      Top             =   6912
      Width           =   11616
      _ExtentX        =   20479
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5997
            MinWidth        =   5997
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8731
            MinWidth        =   8731
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2469
            MinWidth        =   2469
            Text            =   "Usuário"
            TextSave        =   "Usuário"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2187
            MinWidth        =   2187
            TextSave        =   "2/9/2002"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblRecepcaoArquivo 
      Alignment       =   2  'Center
      Caption         =   "Nome do arquivo em recepção"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   3480
      TabIndex        =   4
      Top             =   6408
      Visible         =   0   'False
      Width           =   4812
   End
   Begin VB.Label lblRecepcao 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Recepcionando arquivo "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   3456
      TabIndex        =   3
      Top             =   6048
      Visible         =   0   'False
      Width           =   4860
   End
   Begin VB.Menu MnuComplementacao 
      Caption         =   "&Digitação"
   End
   Begin VB.Menu MnuProvaZero 
      Caption         =   "Prova &Zero"
   End
   Begin VB.Menu MnuRecepcao 
      Caption         =   "Rece&pção"
      Begin VB.Menu MnuRecAvisoDiferenca 
         Caption         =   "&Aviso de Diferença"
      End
      Begin VB.Menu MnuRecConfRemessa 
         Caption         =   "Co&nfirmação de Remessa"
      End
      Begin VB.Menu MnuRecRejeitados 
         Caption         =   "Re&jeitados"
      End
      Begin VB.Menu MnuRecDataBoa 
         Caption         =   "Movimento de Data &Boa"
      End
      Begin VB.Menu MnuRecBaixa 
         Caption         =   "Baixa de &Cheques"
      End
      Begin VB.Menu MnuRecInstrucoes 
         Caption         =   "Tabela de &Instruções"
      End
      Begin VB.Menu MnuRegraGP 
         Caption         =   "Regra do GP"
      End
   End
   Begin VB.Menu MnuGeracao 
      Caption         =   "&Geração"
      Begin VB.Menu MnuGerVC 
         Caption         =   "&Movimento para CH"
      End
      Begin VB.Menu MnuGerCEL 
         Caption         =   "Arquivo &CEL"
         Begin VB.Menu MnuGerCEL_Limite 
            Caption         =   "Cheque &Limite"
         End
         Begin VB.Menu MnuGerCEL_Superior 
            Caption         =   "Cheque &Superior"
         End
         Begin VB.Menu MnuGerCEL_Unibanco 
            Caption         =   "Cheque &Unibanco"
         End
      End
      Begin VB.Menu MnuArqGerTer 
         Caption         =   "Arquivo &TER"
         Begin VB.Menu MnuGerTer 
            Caption         =   "&Geração"
         End
         Begin VB.Menu MnuReGerter 
            Caption         =   "&Re-Geração"
         End
      End
      Begin VB.Menu MnuGerAvisoDiferenca 
         Caption         =   "&Aviso de Diferença"
      End
      Begin VB.Menu MnuGerRejeitados 
         Caption         =   "Movimento de Corri&gidos"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGerExportacao 
         Caption         =   "Exportação de Dados"
         Begin VB.Menu mnuGerExportacaoBordero 
            Caption         =   "&Borderô"
         End
         Begin VB.Menu mnuGerExportacaoAlteracaoData 
            Caption         =   "&Alteração de Data"
         End
         Begin VB.Menu mnuGerExportacaoChqBaixados 
            Caption         =   "&Cheques Baixados"
         End
         Begin VB.Menu mnuGerExportacaoDataBoa 
            Caption         =   "&Data Boa"
         End
      End
   End
   Begin VB.Menu MnuDataBoa 
      Caption         =   "&Data Boa"
      Begin VB.Menu MnuDataBoaCheques 
         Caption         =   "C&apturar Cheques"
      End
      Begin VB.Menu MnuDataBoaFusao 
         Caption         =   "F&usão Automática"
      End
   End
   Begin VB.Menu MnuSupervisao 
      Caption         =   "&Supervisão"
      Begin VB.Menu MnuSupAcompProd 
         Caption         =   "Acompanhamento &Produção"
      End
      Begin VB.Menu MnuSupSupervisor 
         Caption         =   "Super&visor"
      End
      Begin VB.Menu MnuSupParametros 
         Caption         =   "Parâme&tros"
      End
      Begin VB.Menu MnuSupCadUsuario 
         Caption         =   "&Cadastro de Usuários"
      End
   End
   Begin VB.Menu MnuConsultas 
      Caption         =   "Consul&ta"
      Begin VB.Menu MnuConsBorderoCheques 
         Caption         =   "&Borderôs/Cheques"
      End
      Begin VB.Menu MnuConsInstrucoes 
         Caption         =   "&Instruções VC"
      End
      Begin VB.Menu mnuConsultaChequesBaixados 
         Caption         =   "&Cheques Baixados"
      End
   End
   Begin VB.Menu MnuRelatorios 
      Caption         =   "Re&latórios"
      Begin VB.Menu mnuRelBordero 
         Caption         =   "&Borderôs..."
      End
      Begin VB.Menu mnuBorderoTransmissao 
         Caption         =   "Borderôs Preparados Para &Transmissão"
      End
      Begin VB.Menu mnuBorderosConfirmacao 
         Caption         =   "Borderôs Pendentes de &Confirmação"
      End
      Begin VB.Menu mnuBorderosConfirmados 
         Caption         =   "Borderôs Confirmados"
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBorderosDatasChequesRejeitados 
         Caption         =   "Borderôs, Datas e Cheques &Rejeitados"
      End
      Begin VB.Menu MnuChequesDataBoa 
         Caption         =   "Cheques da Data Boa"
      End
      Begin VB.Menu MnuChequesPendenteFusao 
         Caption         =   "Cheques Data Boa Pendente de Fusão"
      End
      Begin VB.Menu mnuChequesBaixados 
         Caption         =   "Cheques Baixados"
         Begin VB.Menu mnuChBxDataPro 
            Caption         =   "D&ata de Processamento"
         End
         Begin VB.Menu mnuChBxGeral 
            Caption         =   "T&odos os Cheques Baixados "
         End
      End
      Begin VB.Menu mnuRelAvisoDiferenca 
         Caption         =   "Aviso Diferença"
         Begin VB.Menu mnuRelAvisoGerado 
            Caption         =   "Aviso Diferença Gerado"
         End
         Begin VB.Menu mnuRelAvisoRecebido 
            Caption         =   "Aviso Diferença Recebido"
         End
      End
   End
   Begin VB.Menu MnuSair 
      Caption         =   "Sai&r"
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Proc_Selecionar As New Custodia.Selecionar
Private Proc_Inserir As New Custodia.Inserir
Private Proc_Atualizar As New Custodia.atualizar
Private Proc_Excluir As New Custodia.Excluir
Private Sub Form_Activate()

    With ProgressBar1
         .Left = StatusBarPrincipal.Panels(StatusBar.Col_ProgressBar).Left + 30
         .Top = StatusBarPrincipal.Top + 50
         .Height = StatusBarPrincipal.Height - 100
         .Width = StatusBarPrincipal.Panels(StatusBar.Col_ProgressBar).Width - 50

         'Desabilita Progress Bar
         .Visible = False
         .Enabled = False
    End With

    'Apresenta em status bar os 15 primeiros dígitos do nome do usuário
    StatusBarPrincipal.Panels(StatusBar.Col_Usuario).Text = Trim(LCase(Left(Geral.UsuarioNome, 15)))

    'Exibe a data de processamento na barra de título , junto com a versão do executável
    Me.Caption = App.Title & " " & _
                 App.Major & "." & _
                 App.Minor & "." & _
                 App.Revision & " [" & _
                 Format(DataDDMMAAAA(Geral.DataProcessamento), "00/00/0000") & "]"

End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    '''''''''''''''''''''''''''''''''''''''''
    'Finaliza a conexão com o Banco de Dados'
    '''''''''''''''''''''''''''''''''''''''''
    If Not SetConnection(g_cMainConnection, False) Then
        MsgBox "Não foi possível finalizar a conexão.", vbExclamation
    End If

End Sub

Private Sub mnuChBxDataPro_Click()
    
    Dim rsChequesBaixados              As New ADODB.Recordset
    Dim SemRegistros                   As Integer
    Dim Selecao                        As New Custodia.Selecionar
       
    MousePointer = vbHourglass
   
    CrystalReport.SelectionFormula = ""
   
    Set rsChequesBaixados = g_cMainConnection.Execute(Selecao.GetDataBaixa(Geral.DataProcessamento))
    
    If rsChequesBaixados.EOF Then
       SemRegistros = MsgBox("Período Sem Cheques Baixados", vbExclamation + vbApplicationModal, App.Title)
       MousePointer = vbDefault
       Exit Sub
    End If
      
    CrystalReport.CopiesToPrinter = 1
    CrystalReport.SelectionFormula = "{ChequesBaixados.DataBaixa}=" & Trim(Geral.DataProcessamento)
    
    CrystalReport.ReportFileName = App.path & "\Reports\ChequesBaixados.rpt"
    CrystalReport.Formulas(0) = "DataProcessamento = '" & FormataData(Geral.DataProcessamento, DD_MM_AAAA) & "'"
    CrystalReport.Formulas(1) = "CNPJterceira = '" & g_Parametros.CNPJ_Terceira & "'"
    CrystalReport.WindowState = crptMaximized
    CrystalReport.WindowTitle = "Emissão do Relatório de Cheques Baixados"
    CrystalReport.Action = 0
    
    CrystalReport.SelectionFormula = ""
   
    MousePointer = vbDefault
End Sub

Private Sub mnuChBxGeral_Click()
    
    Dim rsChequesBaixados              As New ADODB.Recordset
    Dim SemRegistros                   As Integer
    Dim Selecao                        As New Custodia.Selecionar
       
    MousePointer = vbHourglass
   
    CrystalReport.SelectionFormula = ""
   
    Set rsChequesBaixados = g_cMainConnection.Execute(Selecao.GetDataBaixa())
    
    If rsChequesBaixados.EOF Then
      SemRegistros = MsgBox("Cheques Baixados Inexistentes no Banco de Dados", vbExclamation + vbApplicationModal, App.Title)
      MousePointer = vbDefault
      Exit Sub
    End If
      
    CrystalReport.CopiesToPrinter = 1
    CrystalReport.SelectionFormula = ""
    
    CrystalReport.ReportFileName = App.path & "\Reports\ChequesBaixados.rpt"
    CrystalReport.Formulas(0) = "DataProcessamento = '" & FormataData(Geral.DataProcessamento, DD_MM_AAAA) & "'"
    CrystalReport.Formulas(1) = "CNPJterceira = '" & g_Parametros.CNPJ_Terceira & "'"
    CrystalReport.WindowState = crptMaximized
    CrystalReport.WindowTitle = "Emissão do Relatório de Cheques Baixados"
    CrystalReport.Action = 0
    
    CrystalReport.SelectionFormula = ""
   
    MousePointer = vbDefault
    
End Sub

Private Sub MnuDataBoaCheques_Click()
    DataBoa.Show vbModal
End Sub

Private Sub MnuDataBoaFusao_Click()
    'Fusão automática apartir da leitura do ter.gcc
    
    If MsgBox("Deseja Iniciar o Processo de Fusão Automática?", vbQuestion + vbYesNo + vbApplicationModal, App.Title) = vbYes Then
       FusaoDialog.Show vbModal
    End If
     
End Sub

Private Sub MnuGerAvisoDiferenca_Click()
    Dim GerAviso            As New Arquivo_AvisoDiferença
    Dim lRetorno            As Long
    Dim NumeroAviso         As Integer
    
    
    On Error GoTo Erro_GeracaoAviso
    
    
    If MsgBox("Inicia geração do Arquivo de Aviso de Diferença", vbQuestion + vbYesNo, App.Title) = vbNo Then
       Exit Sub
    End If
    
    
    NumeroAviso = NumeroAD(Geral.DataProcessamento)
    
    
    ''''''''''''''''''''''''
    'Habilita o ProgressBar'
    ''''''''''''''''''''''''
    Me.ProgressBar1.Enabled = True
    Me.ProgressBar1.Visible = True


    ''''''''''''''''''''''''''''''
    'Ativa a classe de exportação'
    ''''''''''''''''''''''''''''''
    
    GerAviso.SetConnection g_cMainConnection
    GerAviso.SetProgressBar Me.ProgressBar1
    GerAviso.DiretorioSaida = g_Parametros.DiretorioTransmissao

    GerAviso.ArquivoSaida = Format(g_Parametros.Codigo_Terceira, "0000") + Format(NumeroAviso, "0000") & ".AD"
    
    GerAviso.DataProcessamento = Geral.DataProcessamento
    
    
    If Not GerAviso.Gera_AvisoDiferença() Then
        MsgBox "Não foi possível gerar o Arquivo de Aviso de Diferença!", vbCritical, Me.Caption
        GoTo Erro_GeracaoAviso
    Else
        MsgBox "Aquivo Gerado com Sucesso.", vbExclamation, Me.Caption
    End If
    '''''''''''''''''''''''''
    
    
Erro_GeracaoAviso:
    ''''''''''''''''''''''''''
    'desabilita o ProgressBar'
    ''''''''''''''''''''''''''
    Me.ProgressBar1.Enabled = False
    Me.ProgressBar1.Visible = False

End Sub

Private Sub MnuRecAvisoDiferenca_Click()

    Call recepcao.RecAvisoDiferenca
    
End Sub
Private Sub mnuBorderosConfirmacao_Click()

    Dim rsPesquisaBorderoConfirmacao   As New ADODB.Recordset
    Dim SemRegistros                   As Integer
    Dim Selecao                        As New Custodia.Selecionar
    
    MousePointer = vbHourglass
    
    Set rsPesquisaBorderoConfirmacao = g_cMainConnection.Execute(Selecao.GetBorderoConfirmacao(Geral.DataProcessamento))
    
    If rsPesquisaBorderoConfirmacao.EOF Then
       SemRegistros = MsgBox("Período Sem Borderôs Para Confirmação", vbExclamation, "Relatório de Borderôs Para Confirmação")
       MousePointer = vbDefault
       Exit Sub
    End If
    
    CrystalReport.CopiesToPrinter = 1
    CrystalReport.ReportFileName = App.path & "\Reports\RelBorderosPendentesConfirmacao.rpt"
    CrystalReport.SelectionFormula = "{Bordero.Status}='T' and {Bordero.DataProcessamento}=" & Trim(Geral.DataProcessamento)
    CrystalReport.Formulas(0) = "DataProcessamento = '" & FormataData(Geral.DataProcessamento, DD_MM_AAAA) & "'"
    CrystalReport.Formulas(1) = "Terceira = '" & g_Parametros.CNPJ_Terceira & "'"
    CrystalReport.WindowState = crptMaximized
    CrystalReport.WindowTitle = "Emissão do Relatório de Borderôs Pendentes de Confirmação"
    CrystalReport.Action = 0
    
    MousePointer = vbDefault

End Sub
Private Sub mnuBorderosConfirmados_Click()
    
    Dim rsPesquisaBorderoConfirmado    As New ADODB.Recordset

    Dim SemRegistros                   As Integer

    Dim Selecao                        As New Custodia.Selecionar
    
    MousePointer = vbHourglass
    
    
    Set rsPesquisaBorderoConfirmado = g_cMainConnection.Execute(Selecao.GetBorderoConfirmados(Geral.DataProcessamento))
    
    If rsPesquisaBorderoConfirmado.EOF Then
       SemRegistros = MsgBox("Período Sem Borderôs Confirmados Para Transmissão", vbExclamation, "Relatório de Borderôs Para Transmissão")
       MousePointer = vbDefault
       Exit Sub
    End If
    
    CrystalReport.CopiesToPrinter = 1
    CrystalReport.ReportFileName = App.path & "\Reports\RelBorderosConfirmados.rpt"
    CrystalReport.SelectionFormula = "{Bordero.Status}='E' and {Bordero.DataProcessamento}=" & Trim(Geral.DataProcessamento)
    CrystalReport.Formulas(0) = "DataProcessamento = '" & FormataData(Geral.DataProcessamento, DD_MM_AAAA) & "'"
    CrystalReport.Formulas(1) = "CNPJterceira = '" & g_Parametros.CNPJ_Terceira & "'"
    CrystalReport.WindowState = crptMaximized
    CrystalReport.WindowTitle = "Emissão do Relatório de Borderôs Confirmados"
    CrystalReport.Action = 0
    MousePointer = vbDefault
End Sub
Private Sub mnuBorderosDatasChequesRejeitados_Click()
On Error GoTo Erro:

    Dim rsCheques       As New ADODB.Recordset
    Dim rsDatasBorderos As New ADODB.Recordset
    Dim Selecao         As New Custodia.Selecionar
    Dim Impressao       As Boolean
    
    Screen.MousePointer = vbHourglass
    
    Set rsCheques = g_cMainConnection.Execute(Selecao.GetChequesRejeitados(Geral.DataProcessamento))
    Set rsDatasBorderos = g_cMainConnection.Execute(Selecao.GetDatasBorderosRejeitados(Geral.DataProcessamento))
    
    CrystalReport.CopiesToPrinter = 1
    If rsCheques.EOF And rsDatasBorderos.EOF Then
       MsgBox "Não há Datas, Cheques e Borderôs Rejeitados nesta data", vbExclamation, "Cheques Rejeitados"
       MousePointer = vbDefault
       StatusBarPrincipal.Panels(2).Text = ""
       Exit Sub
    ElseIf Not rsCheques.EOF And rsDatasBorderos.EOF Then
        Screen.MousePointer = vbHourglass
        StatusBarPrincipal.Panels(2).Text = "Imprimindo Relatórios: Cheques Rejeitados, Aguarde... "
        
        CrystalReport.ReportFileName = App.path & "\Reports\RelChequesRejeitados.rpt"
        CrystalReport.SelectionFormula = "{RelChequesRejeitados.DataProcessamento}=" & Trim(Geral.DataProcessamento)
        CrystalReport.Formulas(0) = "DataProcessamento = '" & FormataData(Geral.DataProcessamento, DD_MM_AAAA) & "'"
        CrystalReport.Formulas(1) = "Terceira = '" & g_Parametros.CNPJ_Terceira & "'"
        CrystalReport.WindowState = crptMaximized
        CrystalReport.WindowTitle = "Emissão do Relatório de Cheques Rejeitados"
        CrystalReport.Action = 0
        
        MsgBox "Não há Cheques Rejeitados nesta data", vbExclamation, "Cheques Rejeitados"
        MousePointer = vbDefault
        StatusBarPrincipal.Panels(2).Text = ""
        Exit Sub
    ElseIf rsCheques.EOF And Not rsDatasBorderos.EOF Then
        Screen.MousePointer = vbHourglass
        StatusBarPrincipal.Panels(2).Text = "Imprimindo Relatórios: Datas/Borderôs Rejeitados, Aguarde... "
        
        CrystalReport.ReportFileName = App.path & "\Reports\RelBorderosDatasRejeitados.rpt"
        CrystalReport.SelectionFormula = "{RelBorderosDatasRejeitados.DataProcessamento}=" & Trim(Geral.DataProcessamento)
        CrystalReport.Formulas(0) = "DataProcessamento = '" & FormataData(Geral.DataProcessamento, DD_MM_AAAA) & "'"
        CrystalReport.Formulas(1) = "Terceira = '" & g_Parametros.CNPJ_Terceira & "'"
        CrystalReport.WindowState = crptMaximized
        CrystalReport.WindowTitle = "Emissão do Relatório de Datas/Borderôs Rejeitados"
        CrystalReport.Action = 0
        
        MsgBox "Não há Datas/Borderos Rejeitados nesta data", vbExclamation, "Cheques Rejeitados"
        MousePointer = vbDefault
        StatusBarPrincipal.Panels(2).Text = ""
        Exit Sub
    ElseIf Not rsCheques.EOF And Not rsDatasBorderos.EOF Then
        Screen.MousePointer = vbHourglass
        StatusBarPrincipal.Panels(2).Text = "Imprimindo Relatórios: Cheques Rejeitados, Aguarde... "
        CrystalReport.ReportFileName = App.path & "\Reports\RelChequesRejeitados.rpt"
        CrystalReport.SelectionFormula = "{RelChequesRejeitados.DataProcessamento}=" & Trim(Geral.DataProcessamento)
        CrystalReport.Formulas(0) = "DataProcessamento = '" & FormataData(Geral.DataProcessamento, DD_MM_AAAA) & "'"
        CrystalReport.Formulas(1) = "Terceira = '" & g_Parametros.CNPJ_Terceira & "'"
        CrystalReport.WindowTitle = "Emissão do Relatório de Cheques Rejeitados"
        CrystalReport.WindowState = crptNormal
        CrystalReport.WindowLeft = 30
        CrystalReport.WindowTop = 30
        CrystalReport.WindowHeight = 700
        CrystalReport.WindowWidth = 950
        CrystalReport.Action = 0
        
        StatusBarPrincipal.Panels(2).Text = "Imprimindo Relatórios: Datas/Borderôs Rejeitados, Aguarde... "
        
        CrystalReport.ReportFileName = App.path & "\Reports\RelBorderosDatasRejeitados.rpt"
        CrystalReport.SelectionFormula = "{RelBorderosDatasRejeitados.DataProcessamento}=" & Trim(Geral.DataProcessamento)
        CrystalReport.Formulas(0) = "DataProcessamento = '" & FormataData(Geral.DataProcessamento, DD_MM_AAAA) & "'"
        CrystalReport.Formulas(1) = "Terceira = '" & g_Parametros.CNPJ_Terceira & "'"
        CrystalReport.WindowTitle = "Emissão do Relatório de Datas/Borderôs Rejeitados"
        CrystalReport.WindowLeft = 60
        CrystalReport.WindowTop = 60
        CrystalReport.WindowHeight = 700
        CrystalReport.WindowWidth = 950
        
        CrystalReport.Action = 0
        
        StatusBarPrincipal.Panels(2).Text = ""
    End If
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
            
Erro:
    Call TratamentoErro("Falha na Preparação para Impressão do Relatório de Rejeitados.", Err, False, False)
    

End Sub
Private Sub mnuBorderoTransmissao_Click()
Dim rsPesquisaBorderoTransmissao   As New ADODB.Recordset
Dim SemRegistros                   As Integer
Dim Selecao                        As New Custodia.Selecionar

MousePointer = vbHourglass

Set rsPesquisaBorderoTransmissao = g_cMainConnection.Execute(Selecao.GetBorderoTransmissao(Geral.DataProcessamento))

If rsPesquisaBorderoTransmissao.EOF Then
   MsgBox "Período Sem Borderôs Para Transmissão", vbExclamation, "Relatório de Borderôs Para Transmissão"
   MousePointer = vbDefault
   Exit Sub
End If

CrystalReport.CopiesToPrinter = 1
CrystalReport.ReportFileName = App.path & "\Reports\RelBorderosTransmissao.rpt"
CrystalReport.SelectionFormula = "{Bordero.Status}='R' and {Bordero.DataProcessamento}=" & Trim(Geral.DataProcessamento)
CrystalReport.Formulas(0) = "DataProcessamento = '" & FormataData(Geral.DataProcessamento, DD_MM_AAAA) & "'"
CrystalReport.Formulas(1) = "Terceira = '" & g_Parametros.CNPJ_Terceira & "'"
CrystalReport.WindowState = crptMaximized
CrystalReport.WindowTitle = "Emissão do Relatório de Borderôs Para Transmissão"
CrystalReport.Action = 0

MousePointer = vbDefault

End Sub
Private Sub Old_MnuChequesBaixados_Click()
    
    Dim rsChequesBaixados              As New ADODB.Recordset
    Dim SemRegistros                   As Integer
    Dim Selecao                        As New Custodia.Selecionar
    
       
    MousePointer = vbHourglass
   
   CrystalReport.SelectionFormula = ""
   
   If MsgBox("Deseja Imprimir todos os Cheques Baixados?", vbQuestion + vbYesNo + vbApplicationModal, App.Title) = vbNo Then
    
      Set rsChequesBaixados = g_cMainConnection.Execute(Selecao.GetDataBaixa(Geral.DataProcessamento))
    
      If rsChequesBaixados.EOF Then
         SemRegistros = MsgBox("Período Sem Cheques Baixados", vbExclamation + vbApplicationModal, App.Title)
         MousePointer = vbDefault
         Exit Sub
      End If
      
      CrystalReport.SelectionFormula = "{ChequesBaixados.DataBaixa}=" & Trim(Geral.DataProcessamento)
          
   Else
   
      Set rsChequesBaixados = g_cMainConnection.Execute(Selecao.GetDataBaixa())
    
      If rsChequesBaixados.EOF Then
         SemRegistros = MsgBox("Cheques Baixados Inexistentes no Banco de Dados", vbExclamation + vbApplicationModal, App.Title)
         MousePointer = vbDefault
         Exit Sub
      End If
      
      CrystalReport.SelectionFormula = ""
      
   End If
    
   CrystalReport.CopiesToPrinter = 1
   CrystalReport.ReportFileName = App.path & "\Reports\ChequesBaixados.rpt"
   CrystalReport.Formulas(0) = "DataProcessamento = '" & FormataData(Geral.DataProcessamento, DD_MM_AAAA) & "'"
   CrystalReport.Formulas(1) = "CNPJterceira = '" & g_Parametros.CNPJ_Terceira & "'"
   CrystalReport.WindowState = crptMaximized
   CrystalReport.WindowTitle = "Emissão do Relatório de Cheques Baixados"
   CrystalReport.Action = 0
    
   CrystalReport.SelectionFormula = ""
   
    MousePointer = vbDefault
End Sub

Private Sub MnuChequesDataBoa_Click()
    
    
    Dim rsChequesDataBoa               As New ADODB.Recordset

    Dim SemRegistros                   As Integer

    Dim Selecao                        As New Custodia.Selecionar
    
    MousePointer = vbHourglass
    
    
    Set rsChequesDataBoa = g_cMainConnection.Execute(Selecao.GetChequeDataBoa(Geral.DataProcessamento))
    
    If rsChequesDataBoa.EOF Then
       SemRegistros = MsgBox("Período sem cheques com Data Boa", vbExclamation, "Relatório de Borderôs Para Transmissão")
       MousePointer = Default
       Exit Sub
    End If
    
    CrystalReport.CopiesToPrinter = 1
    CrystalReport.ReportFileName = App.path & "\Reports\RelChequesDataBoa.rpt"
    CrystalReport.SelectionFormula = "{ChequeDataBoa.DataDeposito}=" & Trim(Geral.DataProcessamento)
    CrystalReport.Formulas(0) = "DataProcessamento = '" & FormataData(Geral.DataProcessamento, DD_MM_AAAA) & "'"
    CrystalReport.Formulas(1) = "CNPJterceira = '" & g_Parametros.CNPJ_Terceira & "'"
    CrystalReport.WindowState = crptMaximized
    CrystalReport.WindowTitle = "Emissão do Relatório de Cheques da Data Boa"
    CrystalReport.Action = 0
    
    MousePointer = vbDefault
    
End Sub

Private Sub MnuChequesPendenteFusao_Click()
    
    Dim rsChequesPendenteFusao         As New ADODB.Recordset

    Dim SemRegistros                   As Integer

    Dim Selecao                        As New Custodia.Selecionar
    
    MousePointer = vbHourglass
        
    Set rsChequesPendenteFusao = g_cMainConnection.Execute(Selecao.GetChequePendenteFusao(Geral.DataProcessamento))
    
    If rsChequesPendenteFusao.EOF Then
       SemRegistros = MsgBox("Período Sem Cheques Pendentes de Fusão", vbExclamation, "Relatório de Borderôs Para Transmissão")
       MousePointer = Default
       Exit Sub
    End If
    
    CrystalReport.CopiesToPrinter = 1
    CrystalReport.ReportFileName = App.path & "\Reports\RelChequesPendentesFusao.rpt"
    CrystalReport.SelectionFormula = "{ChequeDataBoa.Fusao}=0 and {ChequeDataBoa.DataProcessamento}=" & Trim(Geral.DataProcessamento)
    CrystalReport.Formulas(0) = "DataProcessamento = '" & FormataData(Geral.DataProcessamento, DD_MM_AAAA) & "'"
    CrystalReport.Formulas(1) = "CNPJterceira = '" & g_Parametros.CNPJ_Terceira & "'"
    CrystalReport.WindowState = crptMaximized
    CrystalReport.WindowTitle = "Emissão do Relatório de Cheques Pendetes de Fusão"
    CrystalReport.Action = 0
    MousePointer = vbDefault
    
End Sub
Private Sub MnuComplementacao_Click()

    Complementacao.Show vbModal, Me
    
End Sub
Private Sub MnuConsBorderoCheques_Click()

    Consulta.Show vbModal, Me
    
End Sub
Private Sub MnuConsInstrucoes_Click()

    Instrucoes.Show vbModal, Me
    
End Sub
Private Sub mnuConsultaChequesBaixados_Click()

    ConsultaChequesBaixados.Show vbModal, Me
    
End Sub

Private Sub MnuGerCEL_Limite_Click()

    Dim AC                  As New Arquivo_CEL
    Dim lDataTroca          As Long
    Dim Proc_Atualizar      As New Custodia.atualizar
    Dim lRetorno            As Long
    Dim iTipoCheque         As Integer
    Dim sOpcaoMenu          As String
    
    On Error GoTo Erro_GeracaoCEL_Limite
    
    sOpcaoMenu = "Geração do arquivo CEL (Cheque Limite)"
    
    '''''''''''''''''''''''''''''''''''
    'Verifica a geração do Arquivo CEL'
    '''''''''''''''''''''''''''''''''''
    If Not g_Parametros.Gerar_Arquivo_CEL Then
        MsgBox "O parâmetro Gerar Arquivo CEL não está ativado.", vbExclamation, sOpcaoMenu
        Exit Sub
    End If
    
    DataTroca.SetTipoArquivoCEL eCheque_Limite
    If DataTroca.ShowModal(lDataTroca, iTipoCheque, 0) = eRetornoCancelar Then
        Exit Sub
    End If

    
    ''''''''''''''''''''''''
    'Habilita o ProgressBar'
    ''''''''''''''''''''''''
    Me.ProgressBar1.Enabled = True
    Me.ProgressBar1.Visible = True


    ''''''''''''''''''''''''''''''
    'Ativa a classe de exportação'
    ''''''''''''''''''''''''''''''
    
    AC.SetConnection g_cMainConnection
    AC.SetProgressBar Me.ProgressBar1
    AC.DataTroca = lDataTroca
    AC.TipoCheque = iTipoCheque '- Cheque de outros bancos
    AC.TipoExportacao = Limite
    AC.DiretorioSaida = g_Parametros.DiretorioTransmissao

    AC.ArquivoSaida = FormataData(Geral.DataProcessamento, DDMM) & _
                      FormataString(g_Parametros.Num_Remessa_TER, "0", Len(g_Parametros.Num_Remessa_TER), True) & _
                      "L.CEL"

    AC.DataProcessamento = Geral.DataProcessamento
    
    AC.TipoArquivo = eCheque_Limite
    '''''''''''''''''''''''''
    'Abre transação
    g_cMainConnection.BeginTrans
    
    If Not AC.Exportar() Then
        'Cancela transação
        g_cMainConnection.RollbackTrans
        'Deleta arquivo txt
        If FileExist(AC.DiretorioSaida & AC.ArquivoSaida) Then
            Kill AC.DiretorioSaida & AC.ArquivoSaida
        End If
        
        MsgBox "Não foi possível gerar o arquivo CEL.", vbExclamation, sOpcaoMenu
        GoTo Erro_GeracaoCEL_Limite
    End If
    
    '''''''''''''''''''''''''
    With g_Parametros
       ' Call g_cMainConnection.Execute(Proc_Atualizar.AtualizaParametros( _
       '                                Geral.DataProcessamento, _
       '                                .QuantidadeCheques, _
       '                                .QuantidadeDatas, _
       '                                .DiretorioTransmissao, _
       '                                .DiretorioRecepcao, _
       '                                .Codigo_USB, _
       '                                .CodigoAgAcolhed, _
       '                                .CPD_Origem, _
       '                                .CPD_Destino, _
       '                                .Codigo_Terceira, _
       '                               .CNPJ_Terceira, _
       '                                .UF_Terceira, _
       '                               .CodigoAplicacao, _
       '                                .ValorChequeLimite, _
       '                                .HeaderAV, _
       '                                .Gerar_Arquivo_CEL, _
       '                                .Comp_Origem_CEL, _
       '                                .Numero_Versao_Inicial_CEL, _
       '                                .Numero_Versao_Final_CEL, _
       '                                .QuantidadeMinimaDias), _
       '                         lRetorno, _
       '                         adCmdText)

'        If lRetorno = 0 Then
'            'Cancela transação
'            g_cMainConnection.RollbackTrans
'            'Deleta arquivo txt
'            If FileExist(AC.DiretorioSaida & AC.ArquivoSaida) Then
'                Kill AC.DiretorioSaida & AC.ArquivoSaida
'            End If
'
'            MsgBox "Erro ao atualizar os parâmetros.", vbExclamation, sOpcaoMenu
'            GoTo Erro_GeracaoCEL_Limite
'        End If
        
    End With
    
    'Encerra transação
    g_cMainConnection.CommitTrans
    
    MsgBox "Aquivo Gerado com Sucesso.", vbExclamation, sOpcaoMenu
    
    
Erro_GeracaoCEL_Limite:
    ''''''''''''''''''''''''''
    'desabilita o ProgressBar'
    ''''''''''''''''''''''''''
    Me.ProgressBar1.Enabled = False
    Me.ProgressBar1.Visible = False

End Sub

Private Sub MnuGerCEL_Superior_Click()

    Dim AC                  As New Arquivo_CEL
    Dim lDataTroca          As Long
    Dim Proc_Atualizar      As New Custodia.atualizar
    Dim lRetorno            As Long
    Dim iTipoCheque         As Integer
    Dim sOpcaoMenu          As String
    
    On Error GoTo Erro_GeracaoCEL_Superior
    
    sOpcaoMenu = "Geração do arquivo CEL (Cheque Superior)"
    
    '''''''''''''''''''''''''''''''''''
    'Verifica a geração do Arquivo CEL'
    '''''''''''''''''''''''''''''''''''
    If Not g_Parametros.Gerar_Arquivo_CEL Then
        MsgBox "O parâmetro Gerar Arquivo CEL não está ativado.", vbExclamation, sOpcaoMenu
        Exit Sub
    End If
    
    DataTroca.SetTipoArquivoCEL eCheque_Superior
    If DataTroca.ShowModal(lDataTroca, iTipoCheque) = eRetornoCancelar Then
        Exit Sub
    End If

    
    ''''''''''''''''''''''''
    'Habilita o ProgressBar'
    ''''''''''''''''''''''''
    Me.ProgressBar1.Enabled = True
    Me.ProgressBar1.Visible = True


    ''''''''''''''''''''''''''''''
    'Ativa a classe de exportação'
    ''''''''''''''''''''''''''''''
    
    AC.SetConnection g_cMainConnection
    AC.SetProgressBar Me.ProgressBar1
    AC.DataTroca = lDataTroca
    AC.TipoCheque = iTipoCheque '- Cheque de outros bancos
    AC.TipoExportacao = Superior
    AC.DiretorioSaida = g_Parametros.DiretorioTransmissao

    AC.ArquivoSaida = FormataData(Geral.DataProcessamento, DDMM) & _
                      FormataString(g_Parametros.Num_Remessa_TER, "0", Len(g_Parametros.Num_Remessa_TER), True) & _
                      "S.CEL"

    AC.DataProcessamento = Geral.DataProcessamento
    
    AC.TipoArquivo = eCheque_Superior
    '''''''''''''''''''''''''
    'Abre transação
    g_cMainConnection.BeginTrans
    
    If Not AC.Exportar() Then
        'Cancela transação
        g_cMainConnection.RollbackTrans
        'Deleta arquivo txt
        If FileExist(AC.DiretorioSaida & AC.ArquivoSaida) Then
            Kill AC.DiretorioSaida & AC.ArquivoSaida
        End If
    
        MsgBox "Não foi possível gerar o arquivo CEL.", vbExclamation, sOpcaoMenu
        GoTo Erro_GeracaoCEL_Superior
    End If
    '''''''''''''''''''''''''
    With g_Parametros
        Call g_cMainConnection.Execute(Proc_Atualizar.AtualizaParametros( _
                                       Geral.DataProcessamento, _
                                       .QuantidadeCheques, _
                                       .QuantidadeDatas, _
                                       .DiretorioTransmissao, _
                                       .DiretorioRecepcao, _
                                       .Codigo_USB, _
                                       .CodigoAgAcolhed, _
                                       .CPD_Origem, _
                                       .CPD_Destino, _
                                       .Codigo_Terceira, _
                                       .CNPJ_Terceira, _
                                        vbNull, _
                                       .UF_Terceira, _
                                       .CodigoAplicacao, _
                                       .ValorChequeLimite, _
                                       .HeaderAV, .chkSoma, _
                                       .Gerar_Arquivo_CEL, _
                                       .Comp_Origem_CEL, _
                                       .Numero_Versao_Inicial_CEL, _
                                       .Numero_Versao_Final_CEL, _
                                       .QuantidadeMinimaDias, _
                                       .Cidade_Terceira, .Nome_Terceira), _
                                lRetorno, _
                                adCmdText)

        If lRetorno = 0 Then
            'Cancela transação
            g_cMainConnection.RollbackTrans
            'Deleta arquivo txt
            If FileExist(AC.DiretorioSaida & AC.ArquivoSaida) Then
                Kill AC.DiretorioSaida & AC.ArquivoSaida
            End If
            
            MsgBox "Erro ao atualizar os parâmetros na geração do arquivo CEL.", vbExclamation, sOpcaoMenu
            GoTo Erro_GeracaoCEL_Superior
        End If
        
    End With
    
    'Encerra transação
    g_cMainConnection.CommitTrans
    
    MsgBox "Aquivo Gerado com Sucesso.", vbExclamation, sOpcaoMenu
    
    
Erro_GeracaoCEL_Superior:
    ''''''''''''''''''''''''''
    'desabilita o ProgressBar'
    ''''''''''''''''''''''''''
    Me.ProgressBar1.Enabled = False
    Me.ProgressBar1.Visible = False

End Sub


Private Sub MnuGerCEL_Unibanco_Click()
    Dim AC                  As New Arquivo_CEL
    Dim lDataTroca          As Long
    Dim Proc_Atualizar      As New Custodia.atualizar
    Dim lRetorno            As Long
    Dim iTipoCheque         As Integer
    Dim sOpcaoMenu          As String
    
    On Error GoTo Erro_GeracaoCEL_Superior
    
    sOpcaoMenu = "Geração do arquivo CEL (Cheque Unibanco)"
    
    '''''''''''''''''''''''''''''''''''
    'Verifica a geração do Arquivo CEL'
    '''''''''''''''''''''''''''''''''''
    If Not g_Parametros.Gerar_Arquivo_CEL Then
        MsgBox "O parâmetro Gerar Arquivo CEL não está ativado.", vbExclamation, sOpcaoMenu
        Exit Sub
    End If
    
    DataTroca.SetTipoArquivoCEL eCheque_Unibanco
    If DataTroca.ShowModal(lDataTroca, iTipoCheque) = eRetornoCancelar Then
        Exit Sub
    End If

    
    ''''''''''''''''''''''''
    'Habilita o ProgressBar'
    ''''''''''''''''''''''''
    Me.ProgressBar1.Enabled = True
    Me.ProgressBar1.Visible = True


    ''''''''''''''''''''''''''''''
    'Ativa a classe de exportação'
    ''''''''''''''''''''''''''''''
    
    AC.SetConnection g_cMainConnection
    AC.SetProgressBar Me.ProgressBar1
    AC.DataTroca = lDataTroca
    AC.TipoCheque = iTipoCheque '- Cheque de outros bancos
    AC.TipoExportacao = Superior
    AC.DiretorioSaida = g_Parametros.DiretorioTransmissao

    AC.ArquivoSaida = FormataData(Geral.DataProcessamento, DDMM) & _
                      FormataString(g_Parametros.Num_Remessa_TER, "0", Len(g_Parametros.Num_Remessa_TER), True) & _
                      "U.CEL"

    AC.DataProcessamento = Geral.DataProcessamento
    
    AC.TipoArquivo = eCheque_Superior
    '''''''''''''''''''''''''
    'Abre transação
    g_cMainConnection.BeginTrans
    
    If Not AC.Exportar() Then
        'Cancela transação
        g_cMainConnection.RollbackTrans
        'Deleta arquivo txt
        If FileExist(AC.DiretorioSaida & AC.ArquivoSaida) Then
            Kill AC.DiretorioSaida & AC.ArquivoSaida
        End If
    
        MsgBox "Não foi possível gerar o arquivo CEL.", vbExclamation, sOpcaoMenu
        GoTo Erro_GeracaoCEL_Superior
    End If
    '''''''''''''''''''''''''
    
    With g_Parametros
        Call g_cMainConnection.Execute(Proc_Atualizar.AtualizaParametros( _
                                       Geral.DataProcessamento, _
                                       .QuantidadeCheques, _
                                       .QuantidadeDatas, _
                                       .DiretorioTransmissao, _
                                       .DiretorioRecepcao, _
                                       .Codigo_USB, _
                                       .CodigoAgAcolhed, _
                                       .CPD_Origem, _
                                       .CPD_Destino, _
                                       .Codigo_Terceira, _
                                       .CNPJ_Terceira, _
                                        vbNull, _
                                       .UF_Terceira, _
                                       .CodigoAplicacao, _
                                       .ValorChequeLimite, _
                                       .HeaderAV, .chkSoma, _
                                       .Gerar_Arquivo_CEL, _
                                       .Comp_Origem_CEL, _
                                       .Numero_Versao_Inicial_CEL, _
                                       .Numero_Versao_Final_CEL, _
                                       .QuantidadeMinimaDias, _
                                       .Cidade_Terceira, .Nome_Terceira), _
                                lRetorno, _
                                adCmdText)

        If lRetorno = 0 Then
            'Cancela transação
            g_cMainConnection.RollbackTrans
            'Deleta arquivo txt
            If FileExist(AC.DiretorioSaida & AC.ArquivoSaida) Then
                Kill AC.DiretorioSaida & AC.ArquivoSaida
            End If
            
            MsgBox "Erro ao atualizar os parâmetros na geração do arquivo CEL.", vbExclamation, sOpcaoMenu
            GoTo Erro_GeracaoCEL_Superior
        End If
        
    End With
    
    'Encerra transação
    g_cMainConnection.CommitTrans
  
    MsgBox "Aquivo Gerado com Sucesso.", vbExclamation, sOpcaoMenu
  
    
Erro_GeracaoCEL_Superior:
    ''''''''''''''''''''''''''
    'desabilita o ProgressBar'
    ''''''''''''''''''''''''''
    Me.ProgressBar1.Enabled = False
    Me.ProgressBar1.Visible = False

End Sub
Private Sub mnuGerExportacaoAlteracaoData_Click()
     
     ExportacaoDeDados.m_EscolhaExportacao = "D"
     ExportacaoDeDados.Show vbModal, Me

End Sub
Private Sub mnuGerExportacaoBordero_Click()

     ExportacaoDeDados.m_EscolhaExportacao = "B"
     ExportacaoDeDados.Show vbModal, Me

End Sub
Private Sub mnuGerExportacaoChqBaixados_Click()

     ExportacaoDeDados.m_EscolhaExportacao = "C"
     ExportacaoDeDados.Show vbModal, Me

End Sub

Private Sub mnuGerExportacaoDataBoa_Click()

     ExportacaoDeDados.m_EscolhaExportacao = "O"
     ExportacaoDeDados.Show vbModal, Me

End Sub

Private Sub MnuGerRejeitados_Click()
    MovimentoVc.SetStatus ("C")
End Sub

Private Sub MnuGerTER_Click()
    Dim GerTer              As New Arquivo_TERGCC
    Dim lDataTroca          As Long
    Dim lRetorno            As Long
    Dim iTipoCheque         As Integer
    Dim iNumRemessa         As Integer
    
    On Error GoTo Erro_GeracaoTERGCC
        
'    If Not g_Parametros.Gerar_Arquivo_CEL Then
'        MsgBox "O parâmetro Gerar Arquivo CEL não está ativado.", vbExclamation, Me.Caption
'        Exit Sub
'    End If
        
    DataTroca.SetTipoArquivoCEL eArquivo_TER
    DataTroca.SetTipoGer 0
    If DataTroca.ShowModal(lDataTroca, iTipoCheque, "G", 0) = eRetornoCancelar Then
        Exit Sub
    End If

    ''''''''''''''''''''''''
    'Habilita o ProgressBar'
    ''''''''''''''''''''''''
    Me.ProgressBar1.Enabled = True
    Me.ProgressBar1.Visible = True


    ''''''''''''''''''''''''''''''
    'Ativa a classe de exportação'
    ''''''''''''''''''''''''''''''
    
    GerTer.SetConnection g_cMainConnection
    
    GerTer.SetProgressBar Me.ProgressBar1
    GerTer.DataTroca = lDataTroca
    GerTer.DiretorioSaida = g_Parametros.DiretorioTransmissao

    GerTer.ArquivoSaida = Format(g_Parametros.Codigo_Terceira, "0000") + Format((g_Parametros.Num_Remessa_TER) + 1, "0000") & ".GCC"
    
    GerTer.DataProcessamento = Geral.DataProcessamento
    
    ''''''''''''''''''''''''''''''''''
    'TER.TipoArquivo = eCheque_Limite
    ''''''''''''''''''''''''''''''''''
    
    If Not GerTer.Gera_TERGCC() Then
        MsgBox "Não foi possível gerar arquivo TERGCC. Favor verificar.", vbExclamation, Me.Caption
        GoTo Erro_GeracaoTERGCC
    Else
        MsgBox "Aquivo Gerado com Sucesso.", vbExclamation, Me.Caption
    End If
    '''''''''''''''''''''''''
    
    
Erro_GeracaoTERGCC:
    ''''''''''''''''''''''''''
    'desabilita o ProgressBar'
    ''''''''''''''''''''''''''
    Me.ProgressBar1.Enabled = False
    Me.ProgressBar1.Visible = False
End Sub
Private Sub MnuGerVC_Click()

    MovimentoVc.SetStatus ("R")
    
End Sub
Private Sub MnuProvaZero_Click()

     ProvaZero.Show vbModal
     
End Sub
Private Sub MnuRecBaixa_Click()

   If Geral.DataProcessamento = Format(Date, "yyyymmdd") Then
       Call recepcao.RecConfBaixas
   Else
     MsgBox "Data de Processamento Inválida!", vbOKOnly & vbExclamation
   End If
     
End Sub

Private Sub MnuRecConfRemessa_Click()

     Call recepcao.RecConfRemessa

End Sub

Private Sub MnuRecDataBoa_Click()
     
   If Geral.DataProcessamento = Format(Date, "yyyymmdd") Then
     Call recepcao.RecChqDataBoa
   Else
     MsgBox "Data de Processamento Inválida!", vbOKOnly & vbExclamation
   End If
End Sub
Private Sub MnuRecInstrucoes_Click()

     Call recepcao.RecInstrucoes

End Sub
Private Sub MnuRecRejeitados_Click()

    Call recepcao.RecRejeitados
    
End Sub

Private Sub MnuReGerter_Click()
    Dim ReGerTer            As New Arquivo_TERGCC
    Dim lDataTroca          As Long
    Dim lRetorno            As Long
    Dim iTipoCheque         As Integer
    Dim iNumRemessa         As Integer
    Dim lNovaRemessa        As Boolean
    
    
    On Error GoTo Erro_GeracaoTERGCC
        
    
    DataTroca.SetTipoArquivoCEL eArquivo_TER
    DataTroca.SetTipoGer 1
    If DataTroca.ShowModal(lDataTroca, iTipoCheque, "R", iNumRemessa, lNovaRemessa) = eRetornoCancelar Then
        Exit Sub
    End If

    ''''''''''''''''''''''''
    'Habilita o ProgressBar'
    ''''''''''''''''''''''''
    Me.ProgressBar1.Enabled = True
    Me.ProgressBar1.Visible = True


    ''''''''''''''''''''''''''''''
    'Ativa a classe de exportação'
    ''''''''''''''''''''''''''''''
    
    
    ReGerTer.SetConnection g_cMainConnection
    
    
    ReGerTer.SetProgressBar Me.ProgressBar1
    ReGerTer.DataTroca = lDataTroca
    ReGerTer.NumRemessa = iNumRemessa
    ReGerTer.NovaRemessa = lNovaRemessa
    ReGerTer.DiretorioSaida = g_Parametros.DiretorioTransmissao

    ReGerTer.ArquivoSaida = Format(g_Parametros.Codigo_Terceira, "0000") + Format(iNumRemessa, "0000") & ".GCC"
    
    ReGerTer.DataProcessamento = Geral.DataProcessamento
    
    ''''''''''''''''''''''''''''''''''
    'TER.TipoArquivo = eCheque_Limite
    ''''''''''''''''''''''''''''''''''
    
    If Not ReGerTer.Gera_TERGCC() Then
        MsgBox "Não foi possível gerar arquivo TERGCC. Favor verificar.", vbExclamation, Me.Caption
        GoTo Erro_GeracaoTERGCC
    Else
        MsgBox "Aquivo Gerado com Sucesso.", vbExclamation, Me.Caption
    End If
    '''''''''''''''''''''''''
    
    
Erro_GeracaoTERGCC:
    ''''''''''''''''''''''''''
    'desabilita o ProgressBar'
    ''''''''''''''''''''''''''
    Me.ProgressBar1.Enabled = False
    Me.ProgressBar1.Visible = False
    
End Sub

Private Sub MnuRegraGP_Click()
    Call recepcao.RecRegraGP
End Sub

Private Sub mnuRelAvisoGerado_Click()
    RelAvisoDiferenca.Show vbModal
End Sub

Private Sub mnuRelAvisoRecebido_Click()
     Call ImpRecAvisoDif.ImpRecAvisoDiferença  ' Impressão do Aviso Recebido
End Sub

Private Sub mnuRelBordero_Click()

    RelBordero.Show vbModal
    
End Sub
Private Sub mnuSair_Click()

     End
     
End Sub
Private Sub MnuSupAcompProd_Click()

    Estatistica.Show vbModal
    
End Sub
Private Sub MnuSupCadUsuario_Click()

     CadUsuario.Show vbModal
     
End Sub
Private Sub MnuSupParametros_Click()

    Parametros.Show vbModal
    
End Sub
Private Sub MnuSupSupervisor_Click()

    Supervisor.Show vbModal
    
End Sub
Private Sub X_AtualizaParametro()

    MsgBox "xxx"
    
End Sub
Function SetScanner()
    Dim Repete As Boolean

   'Se scanner não configurado sai
    If g_Parametros.Scanner = 0 Then
        Scanner.HABILITADO = False
        Exit Function
    End If
 
   'Inicialização do Scanner
    On Error GoTo ErroScanner
    StatusBarPrincipal.Panels(2).Text = "Inicializando Scanner, Aguarde ..."
        
    Repete = True
    Scanner.HABILITADO = True
    Scanner.OS = GetWinVersion()
    Scanner.Scanner = g_Parametros.Scanner
    Scanner.CommPort = g_Parametros.PortaCom
    Scanner.BaudRate = 2400
    Scanner.WordLenght = 7
    Scanner.Parity = 0
    Scanner.StopBits = 1
    
Repetir:
    Screen.MousePointer = vbArrowHourglass
    If Not Scanner.Inicializa Then
        If Not Scanner.Erro Is Nothing Then
            Err.Raise Scanner.Erro.Number, App.Title, Scanner.Erro.Description
        End If
    End If
        
   'Grava no Ini scanner como ativo
    GravarOpcaoINI "Scanner", "EmUso", 1
    
    StatusBarPrincipal.Panels(2).Text = ""
    Screen.MousePointer = vbDefault
    
    Exit Function
    
ErroScanner:
    Screen.MousePointer = vbDefault
    StatusBarPrincipal.Panels(2).Text = ""
    Call TratamentoErro("Falha na Inicialização de Scanner, Tentar Novamente ?", Scanner.Erro, Repete, True)
    If Repete Then
       'Verifica se Ficou pendente inicializacao.
        If PegarOpcaoINI("Scanner", "EmUso", "") = 1 Then
           'Finaliza Scanner caso queda de sistema com o mesmo em uso
            Scanner.Finaliza
        End If
    
        Resume Repetir
    End If
    
    MsgBox "Não será possível Captura automática de CMC7", vbCritical + vbOKOnly

End Function
Function DelScanner()
   'Finaliza Scanner
    On Error GoTo ErroScanner
    
    Screen.MousePointer = vbArrowHourglass
    StatusBarPrincipal.Panels(2).Text = "Finalizando Scanner, Aguarde ..."
    
    If Scanner.HABILITADO Then
        Scanner.HABILITADO = False
        If Not Scanner.Finaliza Then
            If Not Scanner.Erro Is Nothing Then
                Err.Raise Scanner.Erro.Number, App.Title, Scanner.Erro.Description
            End If
        End If
    End If
    
   'Marca scanner finalizado no Ini
    GravarOpcaoINI "Scanner", "EmUso", 0
    
    StatusBarPrincipal.Panels(2).Text = ""
    Screen.MousePointer = vbDefault
    
    Exit Function

ErroScanner:
    Screen.MousePointer = vbDefault
    StatusBarPrincipal.Panels(2).Text = ""
    Call TratamentoErro("Falha na Finalização de Scanner.", Scanner.Erro, False, True)
    
End Function

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub


Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

End Sub


