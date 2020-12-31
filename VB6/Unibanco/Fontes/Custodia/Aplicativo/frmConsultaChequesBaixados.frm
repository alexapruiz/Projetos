VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{4A8E27C4-EACA-11D3-9FFC-00104BC8688C}#1.0#0"; "DateEdit.ocx"
Begin VB.Form ConsultaChequesBaixados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta Cheques Baixados"
   ClientHeight    =   5340
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   11016
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   11016
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRotacao 
      Caption         =   "Nova Seleção"
      Height          =   720
      Left            =   8400
      Picture         =   "frmConsultaChequesBaixados.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4536
      Width           =   816
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "&Confirmar"
      Height          =   720
      Left            =   9285
      Picture         =   "frmConsultaChequesBaixados.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4536
      Width           =   816
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   720
      Left            =   10170
      Picture         =   "frmConsultaChequesBaixados.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4536
      Width           =   816
   End
   Begin TabDlg.SSTab TabOpcoes 
      Height          =   4260
      Left            =   90
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   150
      Width           =   10905
      _ExtentX        =   19219
      _ExtentY        =   7514
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   420
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Argumentos de Pesquisa"
      TabPicture(0)   =   "frmConsultaChequesBaixados.frx":0A56
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frabordero"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraDeposito"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fracheque"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&Instruções"
      TabPicture(1)   =   "frmConsultaChequesBaixados.frx":0A72
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdInstrucoes"
      Tab(1).ControlCount=   1
      Begin MSFlexGridLib.MSFlexGrid GrdInstrucoes 
         Height          =   3630
         Left            =   -74880
         TabIndex        =   20
         Top             =   390
         Width           =   10695
         _ExtentX        =   18860
         _ExtentY        =   6392
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         FocusRect       =   0
         ScrollBars      =   2
      End
      Begin VB.Frame fracheque 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Identificação do Cheque"
         ForeColor       =   &H80000008&
         Height          =   2772
         Left            =   6240
         TabIndex        =   2
         Top             =   1056
         Width           =   3180
         Begin VB.TextBox txtNumCheque 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1632
            MaxLength       =   7
            TabIndex        =   9
            Top             =   2208
            Width           =   1092
         End
         Begin VB.OptionButton optcheque 
            BackColor       =   &H80000000&
            Caption         =   "Cheque"
            Height          =   348
            Left            =   300
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   312
            Width           =   1065
         End
         Begin VB.TextBox TxtcBco 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2280
            MaxLength       =   3
            TabIndex        =   6
            Top             =   432
            Width           =   444
         End
         Begin VB.TextBox txtcAgencia 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2208
            MaxLength       =   4
            TabIndex        =   7
            Top             =   1032
            Width           =   516
         End
         Begin VB.TextBox TxtCcc 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1632
            MaxLength       =   7
            TabIndex        =   8
            Top             =   1632
            Width           =   1092
         End
         Begin VB.Label Label4 
            Caption         =   "Núm. Cheque"
            Height          =   228
            Left            =   1728
            TabIndex        =   21
            Top             =   2016
            Width           =   996
         End
         Begin VB.Label Label5 
            Caption         =   "Banco"
            Height          =   228
            Left            =   2256
            TabIndex        =   19
            Top             =   240
            Width           =   492
         End
         Begin VB.Label Label6 
            Caption         =   "Agência"
            Height          =   228
            Left            =   2112
            TabIndex        =   18
            Top             =   840
            Width           =   612
         End
         Begin VB.Label Label7 
            Caption         =   "Conta Corrente"
            Height          =   228
            Left            =   1656
            TabIndex        =   17
            Top             =   1440
            Width           =   1092
         End
      End
      Begin VB.Frame fraDeposito 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Data Depósito"
         ForeColor       =   &H80000008&
         Height          =   900
         Left            =   1572
         TabIndex        =   1
         Top             =   2328
         Width           =   3948
         Begin VB.OptionButton optdteAnterior 
            BackColor       =   &H80000000&
            Caption         =   "Depósito"
            Height          =   348
            Left            =   300
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   360
            Width           =   1116
         End
         Begin DATEEDITLib.DateEdit DteAnterior 
            Height          =   348
            Left            =   1968
            TabIndex        =   5
            Top             =   384
            Width           =   1524
            _Version        =   65537
            _ExtentX        =   2688
            _ExtentY        =   614
            _StockProps     =   93
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
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Data"
            Height          =   252
            Left            =   2016
            TabIndex        =   16
            Top             =   192
            Width           =   1428
         End
      End
      Begin VB.Frame frabordero 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Identificação do Borderô"
         ForeColor       =   &H80000008&
         Height          =   900
         Left            =   1572
         TabIndex        =   0
         Top             =   1056
         Width           =   3948
         Begin VB.OptionButton optbordero 
            BackColor       =   &H80000000&
            Caption         =   "Borderô"
            Height          =   348
            Left            =   300
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   336
            Width           =   876
         End
         Begin VB.TextBox txtnumBordero 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1320
            MaxLength       =   19
            TabIndex        =   4
            Top             =   384
            Width           =   2484
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "Identificação"
            Height          =   252
            Left            =   1488
            TabIndex        =   15
            Top             =   192
            Width           =   2172
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdCheque 
         Height          =   4884
         Left            =   -74808
         TabIndex        =   14
         Top             =   528
         Width           =   9324
         _ExtentX        =   16447
         _ExtentY        =   8615
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   0
      End
   End
End
Attribute VB_Name = "ConsultaChequesBaixados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                            '* Type de Utilização de Banco *'                              '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type Procedures
    Inclusao            As New Custodia.Inserir    'Querys de Insert
    Alteracao           As New Custodia.Atualizar  'Querys de Update
    Deletacao           As New Custodia.Excluir    'Querys de Delete
    Selecao             As New Custodia.Selecionar 'Querys de Select
End Type
Private Procedures      As Procedures

Dim sDtaProcessamento   As Long       'Data de Processamento Formatada MM/DD/AAAA
Dim sDtaAnterior        As Long       'Data Anterior
Dim sDtaNova            As Long       'Data Nova
Dim sValor              As String     'Valor Formatado


Public Sub FormataDataProcessamento()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                         '* Formatação da Data de Processamento *'                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    sDtaProcessamento = Mid(Geral.DataProcessamento, 1, 4) & _
                        Mid(Geral.DataProcessamento, 5, 2) & _
                        Mid(Geral.DataProcessamento, 7, 2)

Exit Sub
TrataErro:
    Call TratamentoErro("Erro ao formatar Data de Processamento.", Err)
    Unload Me

End Sub
Private Sub FormataDataAnterior()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                         '* Formatação da Data de Anterior *'                              '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    sDtaAnterior = Mid(DteAnterior.Text, 5, 4) & _
                   Mid(DteAnterior.Text, 3, 2) & _
                   Mid(DteAnterior.Text, 1, 2)
                   
Exit Sub
TrataErro:
    Call TratamentoErro("Erro ao formatar Data Anterior.", Err)
    Unload Me
                   
End Sub

Private Sub TrataPesquisa()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                           * Tratamento de Pesquisa / Formatações  *                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    Dim rsPesquisaInstrucoes    As New ADODB.Recordset
    Dim nContador               As Long
    
    Dim NCheque                 As Long

    Dim Controle                As Control

    Screen.MousePointer = vbHourglass

    ''''''''''''''''''''''''''''''''''''
    '   * Verifica e Formata Datas *   '
    ''''''''''''''''''''''''''''''''''''
    If Trim(DteAnterior.Text) <> "" Then
        '''''''''''''''''''''''''''''
        ' * Formata Data Anterior * '
        '''''''''''''''''''''''''''''
        Call FormataDataAnterior
    End If
    
    
    ''''''''''''''''''''''''''''''''
    '* Faz Pesquisa de Instruções *'
    ''''''''''''''''''''''''''''''''
    Set rsPesquisaInstrucoes = g_cMainConnection.Execute _
                              (Procedures.Selecao.GetChequesBaixados(txtnumBordero.Text _
                             , sDtaAnterior _
                             , TxtcBco.Text _
                             , txtcAgencia.Text _
                             , TxtCcc.Text _
                             , txtNumCheque.Text))
    
    If Not rsPesquisaInstrucoes.EOF Then
    
        ''''''''''''''''''''''''''''''''''''''''
        '     * Número de Linhas da Grade *    '
        ''''''''''''''''''''''''''''''''''''''''
        GrdCheque.Rows = rsPesquisaInstrucoes.RecordCount + 1
        GrdInstrucoes.Rows = rsPesquisaInstrucoes.RecordCount + 1
        
        ''''''''''''''''''''''''''''''''''''''''
        '     * Número incial do Contador *    '
        ''''''''''''''''''''''''''''''''''''''''
        nContador = 1
        
        Do While Not rsPesquisaInstrucoes.EOF
             
            ''''''''''''''''''''''''''''''''''''''''
            '   * Preenche Grade de Instruções *   '
            ''''''''''''''''''''''''''''''''''''''''
            With GrdInstrucoes
                .TextMatrix(nContador, 0) = rsPesquisaInstrucoes!Num_Bordero
                .TextMatrix(nContador, 1) = Mid(rsPesquisaInstrucoes!DataDeposito, 7, 2) & "/" & Mid(rsPesquisaInstrucoes!DataDeposito, 5, 2) & "/" & Mid(rsPesquisaInstrucoes!DataDeposito, 1, 4)
                .TextMatrix(nContador, 2) = rsPesquisaInstrucoes!BancoEmitente
                .TextMatrix(nContador, 3) = rsPesquisaInstrucoes!AgenciaEmitente
                .TextMatrix(nContador, 4) = rsPesquisaInstrucoes!CcEmitente
                .TextMatrix(nContador, 5) = rsPesquisaInstrucoes!NrChequeEmitente
                .TextMatrix(nContador, 6) = Format(rsPesquisaInstrucoes!ValorCheque, "###,###,##0.00")
                                                                
            End With
            
            nContador = nContador + 1
            rsPesquisaInstrucoes.MoveNext
            
         Loop
         
         NCheque = GrdInstrucoes.Rows
         
        ''''''''''''''''''''''''''''''''''''''''
        '   * Focaliza Tab de Instruções *     '
        ''''''''''''''''''''''''''''''''''''''''
        TabOpcoes.TabEnabled(1) = True
        TabOpcoes.Tab = 1
         
        ''''''''''''''''''''''''''''''''''''''''
        '      * Marca 1ª Linha da Grade *     '
        ''''''''''''''''''''''''''''''''''''''''
        Call GrdInstrucoes_SelChange
        Screen.MousePointer = vbDefault
        
    Else
        MsgBox "Nenhum registro foi encontrado.", vbExclamation + vbOKOnly, App.Title
        
        For Each Controle In Me.Controls
        
            If TypeName(Controle) = "TextBox" Then
        
                If Len(Controle.Text) <> 0 Then
                   Controle.SelStart = 0
                   Controle.SelLength = Len(Controle.Text)
                   Controle.SetFocus
                   Exit For
                End If
                
            End If
                
        Next
        
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    sDtaAnterior = Empty
    sDtaNova = Empty
    
Exit Sub
TrataErro:
    Screen.MousePointer = vbDefault
    
    'If GrdInstrucoes.Rows = nContador Then
    '   Resume Next
    '   Exit Sub
    'End If
      
    Call TratamentoErro("Erro ao tratar Pesquisas.", Err)
    Unload Me
    
End Sub
Private Sub cmdConfirmar_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                           * Tratamento de Pesquisa / Formatações  *                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ''''''''''''''''''''''''''''''''
    '* Pesquisa Por Bco+Ag+Cc+Vlr *'
    ''''''''''''''''''''''''''''''''
    'If optcheque.Value = True Then
    '    If Len(Trim(TxtcBco.Text)) <> 0 Or Len(Trim(txtcAgencia.Text)) <> 0 Or _
            Len(Trim(TxtCcc.Text)) <> 0 Then
    '    Else
    '        MsgBox "Informe o Banco, Agência ou Conta a ser pesquisado.", vbInformation + vbOKOnly, App.Title
    '        TxtcBco.SetFocus
    '        Exit Sub
    '    End If
    'End If
    
    ''''''''''''''''''''''''''''''''''''
    '* Pesquisa por Numero do Bordero *'
    ''''''''''''''''''''''''''''''''''''
    'If optbordero.Value = True Then
    '    If Len(Trim(txtnumBordero.Text)) = 0 Then
    '        MsgBox "Informe o número do Borderô a ser pesquisado.", vbInformation + vbOKOnly, App.Title
    '        txtnumBordero.SetFocus
    '        Exit Sub
    '    End If
    'End If
    
       
    ''''''''''''''''''''''''''''''''
    '* Pesquisa por Data Anterior *'
    ''''''''''''''''''''''''''''''''
    'If optdteAnterior.Value = True Then
    '    If Len(Trim(DteAnterior.Text)) = 0 Then
    '        MsgBox "Informe a Data de Depósito a ser pesquisada.", vbInformation + vbOKOnly, App.Title
    '        DteAnterior.SetFocus
    '        Exit Sub
    '    End If
    'End If
    
    Call TrataPesquisa
    
    ''''''''''''''''''''''''''''''''
    '* Valor Default de Variáveis *'
    ''''''''''''''''''''''''''''''''
    'sDtaAnterior = 0
    'sDtaNova = 0
    
End Sub

Private Sub cmdRotacao_Click()

   optbordero.Value = False
   optcheque.Value = False
   optdteAnterior.Value = False
   
   LimpaTela Me
   
   TabOpcoes.Tab = 0
   
End Sub

Private Sub CmdSair_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                   *  Finaliza Tela  *                                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Unload Me
End Sub

Private Sub DteAnterior_GotFocus()

If Len(DteAnterior.Text) <> 0 Then
   DteAnterior.SelStart = 0
   DteAnterior.SelLength = Len(DteAnterior.Text)
End If

End Sub

Private Sub DteAnterior_KeyPress(KeyAscii As Integer)

    ''''''''''''''''''''''''''''''''''
    '* Este campo só aceita números *'
    ''''''''''''''''''''''''''''''''''
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack And KeyAscii <> 13 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If KeyAscii = 13 Then
        cmdConfirmar.SetFocus
        
        ''''''''''''''''''''''''
        '* Se data for válida *'
        ''''''''''''''''''''''''
        If Len(Trim(DteAnterior.Text)) <> 0 Then
            Call cmdConfirmar_Click
        End If
    End If
    
End Sub

Private Sub Form_Load()

TabOpcoes.TabEnabled(1) = False

GrdInstrucoes.Row = 0

GrdInstrucoes.Col = 0
GrdInstrucoes.ColWidth(0) = 1800
GrdInstrucoes.ColAlignment(0) = flexAlignRightCenter
GrdInstrucoes.Text = "Numero do Borderô"

GrdInstrucoes.Col = 1
GrdInstrucoes.ColWidth(1) = 1600
GrdInstrucoes.ColAlignment(1) = flexAlignRightCenter
GrdInstrucoes.Text = "Data de Depósito"

GrdInstrucoes.Col = 2
GrdInstrucoes.ColWidth(2) = 1500
GrdInstrucoes.ColAlignment(2) = flexAlignRightCenter
GrdInstrucoes.Text = "Código do Banco"

GrdInstrucoes.Col = 3
GrdInstrucoes.ColWidth(3) = 1500
GrdInstrucoes.ColAlignment(3) = flexAlignRightCenter
GrdInstrucoes.Text = "Código da Agência"

GrdInstrucoes.Col = 4
GrdInstrucoes.ColWidth(4) = 1300
GrdInstrucoes.ColAlignment(4) = flexAlignRightCenter
GrdInstrucoes.Text = "Conta Corrente"

GrdInstrucoes.Col = 5
GrdInstrucoes.ColWidth(5) = 1500
GrdInstrucoes.ColAlignment(5) = flexAlignRightCenter
GrdInstrucoes.Text = "Número do Cheque"

GrdInstrucoes.Col = 6
GrdInstrucoes.ColWidth(6) = 1400
GrdInstrucoes.ColAlignment(6) = flexAlignRightCenter
GrdInstrucoes.Text = "Valor"

'optbordero.Value = True
'optdteAnterior.Value = False
'optcheque.Value = False

txtnumBordero.Enabled = False
TxtcBco.Enabled = False
txtcAgencia.Enabled = False
TxtCcc.Enabled = False
txtNumCheque.Enabled = False
DteAnterior.Enabled = False

End Sub

Private Sub GrdInstrucoes_SelChange()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                 * Seleciona Linha Inteira *'                                 '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Static m_SelChange As Boolean
    
    'Controle de Acesso
    If m_SelChange = True Then Exit Sub
    
    m_SelChange = True
        
        GrdInstrucoes.Row = GrdInstrucoes.RowSel
        GrdInstrucoes.Col = 0
        GrdInstrucoes.ColSel = 6
        GrdInstrucoes.SetFocus
        
    m_SelChange = False

End Sub
Private Sub optbordero_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                       * Habilita Pesquisa por número de Borderô *                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''''''''
    '* Desabilita demais pesquisas *'
    '''''''''''''''''''''''''''''''''
    optcheque.Value = False
    TxtcBco.Locked = True
    txtcAgencia.Locked = True
    TxtCcc.Locked = True
    txtNumCheque.Locked = True
    optdteAnterior.Value = False
    DteAnterior.Locked = True
    
      
    '''''''''''''''''''''''''''''''''''''''''''''
    '* Habilita Pesquisa por Número de Borderô *'
    '''''''''''''''''''''''''''''''''''''''''''''
    txtnumBordero.Locked = False
        
    '''''''''''''''''''''''''''
    '* Limpa todos os campos *'
    '''''''''''''''''''''''''''
    LimpaTela Me
    
    txtnumBordero.Enabled = True
    TxtcBco.Enabled = False
    txtcAgencia.Enabled = False
    TxtCcc.Enabled = False
    txtNumCheque.Enabled = False
    DteAnterior.Enabled = False
    
    txtnumBordero.SetFocus
End Sub
Private Sub optcheque_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                         * Habilita Pesquisa por Dados do Cheque *                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''''''''
    '* Desabilita demais pesquisas *'
    '''''''''''''''''''''''''''''''''
    optbordero.Value = False
    txtnumBordero.Locked = True
    optdteAnterior.Value = False
    DteAnterior.Locked = True
    
    '''''''''''''''''''''''''''''''''''''''''''''
    ' * Habilita Pesquisa por Dados do Cheque * '
    '''''''''''''''''''''''''''''''''''''''''''''
    TxtcBco.Locked = False
    txtcAgencia.Locked = False
    TxtCcc.Locked = False
    txtNumCheque.Locked = False
        
    '''''''''''''''''''''''''''
    '* Limpa todos os campos *'
    '''''''''''''''''''''''''''
    LimpaTela Me
    
    'SendKeys "{TAB}"
    
    txtnumBordero.Enabled = False
    TxtcBco.Enabled = True
    txtcAgencia.Enabled = True
    TxtCcc.Enabled = True
    txtNumCheque.Enabled = True
    DteAnterior.Enabled = False
    
    TxtcBco.SetFocus
    
End Sub
Private Sub optdteAnterior_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                          * Habilita Pesquisa por Data Anterior *                            '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''''''''
    '* Desabilita demais pesquisas *'
    '''''''''''''''''''''''''''''''''
    optcheque.Value = False
    TxtcBco.Locked = True
    txtcAgencia.Locked = True
    TxtCcc.Locked = True
    txtNumCheque.Locked = True
    optbordero.Value = False
    txtnumBordero.Locked = True
        
    '''''''''''''''''''''''''''''''''''''''''''
    ' * Habilita Pesquisa por Data Anterior * '
    '''''''''''''''''''''''''''''''''''''''''''
    DteAnterior.Locked = False
    
    '''''''''''''''''''''''''''
    '* Limpa todos os campos *'
    '''''''''''''''''''''''''''
    LimpaTela Me
    
    'SendKeys "{TAB}"
    
    txtnumBordero.Enabled = False
    TxtcBco.Enabled = False
    txtcAgencia.Enabled = False
    TxtCcc.Enabled = False
    txtNumCheque.Enabled = False
    DteAnterior.Enabled = True
    
    DteAnterior.SetFocus
    
End Sub

Private Sub TabOpcoes_Click(PreviousTab As Integer)

If PreviousTab = 1 Then
   
   optbordero.Value = False
   optcheque.Value = False
   optdteAnterior.Value = False
   
   txtnumBordero.Enabled = False
   TxtcBco.Enabled = False
   TxtCcc.Enabled = False
   txtNumCheque.Enabled = False
   txtcAgencia.Enabled = False
        
   LimpaTela Me
   
   TabOpcoes.TabEnabled(1) = False
   
End If

End Sub

Private Sub txtcAgencia_GotFocus()

If Len(txtcAgencia.Text) <> 0 Then
   txtcAgencia.SelStart = 0
   txtcAgencia.SelLength = Len(txtcAgencia.Text)
End If

End Sub

Private Sub txtcAgencia_KeyPress(KeyAscii As Integer)

''''''''''''''''''''''''''''''''''
    '* Este campo só aceita números *'
    ''''''''''''''''''''''''''''''''''
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack And KeyAscii <> 13 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub TxtcBco_GotFocus()

If Len(TxtcBco.Text) <> 0 Then
   TxtcBco.SelStart = 0
   TxtcBco.SelLength = Len(TxtcBco.Text)
End If

End Sub

Private Sub TxtcBco_KeyPress(KeyAscii As Integer)
    
    ''''''''''''''''''''''''''''''''''
    '* Este campo só aceita números *'
    ''''''''''''''''''''''''''''''''''
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack And KeyAscii <> 13 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    
End Sub

Private Sub TxtCcc_GotFocus()

If Len(TxtCcc.Text) <> 0 Then
   TxtCcc.SelStart = 0
   TxtCcc.SelLength = Len(TxtCcc.Text)
End If

End Sub

Private Sub TxtCcc_KeyPress(KeyAscii As Integer)

    ''''''''''''''''''''''''''''''''''
    '* Este campo só aceita números *'
    ''''''''''''''''''''''''''''''''''
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack And KeyAscii <> 13 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    
End Sub

Private Sub txtnumBordero_GotFocus()

If Len(txtnumBordero.Text) <> 0 Then
   txtnumBordero.SelStart = 0
   txtnumBordero.SelLength = Len(txtnumBordero.Text)
End If

txtnumBordero.Enabled = True
TxtcBco.Enabled = False
txtcAgencia.Enabled = False
TxtCcc.Enabled = False
txtNumCheque.Enabled = False
DteAnterior.Enabled = False

End Sub

Private Sub txtnumBordero_KeyPress(KeyAscii As Integer)

    ''''''''''''''''''''''''''''''''''
    '* Este campo só aceita números *'
    ''''''''''''''''''''''''''''''''''
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack And KeyAscii <> 13 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If KeyAscii = 13 Then
        '''''''''''''''''''''''''''''''''''''''''''''''
        '* Formata número do borderô com 19 posições *'
        '''''''''''''''''''''''''''''''''''''''''''''''
        txtnumBordero = Format(txtnumBordero, String(19, "0"))
        Call cmdConfirmar_Click
    End If
    
End Sub

Private Sub txtNumCheque_GotFocus()

If Len(txtNumCheque.Text) <> 0 Then
   txtNumCheque.SelStart = 0
   txtNumCheque.SelLength = Len(txtNumCheque.Text)
End If

End Sub

Private Sub txtNumCheque_KeyPress(KeyAscii As Integer)

    ''''''''''''''''''''''''''''''''''
    '* Este campo só aceita números *'
    ''''''''''''''''''''''''''''''''''
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack And KeyAscii <> 13 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If KeyAscii = 13 Then
        '''''''''''''''''''''''''''''''''''''''''''''''
        '* Formata número do borderô com 19 posições *'
        '''''''''''''''''''''''''''''''''''''''''''''''
    '    txtnumBordero = Format(txtnumBordero, String(19, "0"))
        Call cmdConfirmar_Click
    End If

End Sub

