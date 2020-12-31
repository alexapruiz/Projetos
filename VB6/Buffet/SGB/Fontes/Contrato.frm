VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Contrato 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SGB - Contrato"
   ClientHeight    =   6084
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   11016
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6084
   ScaleWidth      =   11016
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame6 
      Caption         =   "Dados para Pesquisa"
      Height          =   924
      Left            =   0
      TabIndex        =   77
      Top             =   0
      Width           =   7956
      Begin VB.CommandButton CmdPesquisar 
         Caption         =   "&Pesquisar"
         Height          =   324
         Left            =   3720
         TabIndex        =   80
         Top             =   165
         Width           =   1428
      End
      Begin VB.TextBox TxtPesquisaNomeCliente 
         Height          =   288
         Left            =   1824
         TabIndex        =   1
         Top             =   552
         Visible         =   0   'False
         Width           =   4284
      End
      Begin VB.TextBox TxtPesquisaNumContrato 
         Height          =   288
         Left            =   1824
         TabIndex        =   0
         Top             =   216
         Width           =   1740
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Cliente"
         Height          =   192
         Left            =   528
         TabIndex        =   79
         Top             =   600
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Número do Contrato"
         Height          =   192
         Left            =   288
         TabIndex        =   78
         Top             =   288
         Width           =   1440
      End
   End
   Begin VB.CommandButton CmdSair 
      Caption         =   "Sai&r"
      Height          =   372
      Left            =   120
      TabIndex        =   76
      Top             =   5664
      Width           =   1428
   End
   Begin VB.CommandButton CmdGravarContrato 
      Caption         =   "&Gravar Contrato"
      Height          =   372
      Left            =   9480
      TabIndex        =   69
      Top             =   5664
      Width           =   1428
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4500
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   11010
      _ExtentX        =   19431
      _ExtentY        =   7938
      _Version        =   393216
      TabHeight       =   420
      TabCaption(0)   =   "Dados do Contrato"
      TabPicture(0)   =   "Contrato.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CmdPesquisarCliente"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CmdIncluirCliente"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Condições de Pagamento"
      TabPicture(1)   =   "Contrato.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Outras Informações"
      TabPicture(2)   =   "Contrato.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).ControlCount=   1
      Begin VB.CommandButton CmdIncluirCliente 
         Caption         =   "Incluir Cliente"
         Height          =   390
         Left            =   2016
         TabIndex        =   84
         Top             =   3792
         Width           =   1644
      End
      Begin VB.CommandButton CmdPesquisarCliente 
         Caption         =   "Pesquisar Cliente"
         Height          =   390
         Left            =   192
         TabIndex        =   83
         Top             =   3792
         Width           =   1644
      End
      Begin VB.Frame Frame5 
         Caption         =   "Observações"
         Height          =   3636
         Left            =   -74808
         TabIndex        =   70
         Top             =   360
         Width           =   10572
         Begin VB.ComboBox CboDoce 
            Height          =   288
            Left            =   1056
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   1296
            Visible         =   0   'False
            Width           =   2340
         End
         Begin VB.ComboBox CboSalgado 
            Height          =   288
            Left            =   1056
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   960
            Visible         =   0   'False
            Width           =   2340
         End
         Begin VB.ComboBox CboDecor 
            Height          =   315
            Left            =   7440
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   624
            Visible         =   0   'False
            Width           =   1932
         End
         Begin VB.ComboBox CboBolo 
            Height          =   315
            Left            =   7440
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   288
            Visible         =   0   'False
            Width           =   1932
         End
         Begin VB.TextBox TxtOBS 
            Height          =   288
            Left            =   1056
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   44
            Top             =   2256
            Width           =   9300
         End
         Begin VB.TextBox TxtPais 
            Height          =   288
            Left            =   1056
            MaxLength       =   30
            TabIndex        =   43
            Top             =   1944
            Width           =   5100
         End
         Begin VB.TextBox TxtBebida 
            Height          =   288
            Left            =   1056
            LinkTimeout     =   30
            MaxLength       =   50
            TabIndex        =   42
            Top             =   1620
            Width           =   5100
         End
         Begin VB.TextBox TxtDecoracao 
            Height          =   288
            Left            =   1056
            LinkTimeout     =   30
            MaxLength       =   80
            TabIndex        =   38
            Top             =   630
            Width           =   6315
         End
         Begin VB.TextBox TxtBolo 
            Height          =   288
            Left            =   1056
            MaxLength       =   80
            TabIndex        =   36
            Top             =   288
            Width           =   6315
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Doce"
            Height          =   192
            Left            =   552
            TabIndex        =   88
            Top             =   1344
            Visible         =   0   'False
            Width           =   396
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Salgado"
            Height          =   192
            Left            =   312
            TabIndex        =   87
            Top             =   1008
            Visible         =   0   'False
            Width           =   624
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "OBS"
            Height          =   192
            Left            =   600
            TabIndex        =   75
            Top             =   2304
            Width           =   336
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Pais"
            Height          =   192
            Left            =   612
            TabIndex        =   74
            Top             =   1992
            Width           =   324
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Bebida"
            Height          =   192
            Left            =   408
            TabIndex        =   73
            Top             =   1680
            Width           =   528
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Decoração"
            Height          =   192
            Left            =   120
            TabIndex        =   72
            Top             =   672
            Width           =   816
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Bolo"
            Height          =   192
            Left            =   600
            TabIndex        =   71
            Top             =   312
            Width           =   336
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Dados de Parcelamento"
         Height          =   3204
         Left            =   -74880
         TabIndex        =   64
         Top             =   1080
         Width           =   10740
         Begin VB.ComboBox CboFormaPagto 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   525
            Width           =   2415
         End
         Begin VB.CommandButton CmdLimpar 
            Caption         =   "Limpar Campos"
            Height          =   324
            Left            =   8640
            TabIndex        =   35
            Top             =   2640
            Width           =   1590
         End
         Begin VB.CommandButton CmdExcluir 
            Caption         =   "Excluir"
            Height          =   324
            Left            =   8640
            TabIndex        =   34
            Top             =   1800
            Width           =   1590
         End
         Begin VB.CommandButton CmdIncluir 
            Caption         =   "Incluir"
            Height          =   324
            Left            =   8640
            TabIndex        =   33
            Top             =   960
            Width           =   1590
         End
         Begin VB.TextBox TxtValorParcela 
            Height          =   288
            Left            =   6810
            TabIndex        =   32
            Top             =   528
            Width           =   1590
         End
         Begin VB.TextBox TxtNumChequeParcela 
            Height          =   288
            Left            =   4995
            TabIndex        =   31
            Top             =   528
            Width           =   1530
         End
         Begin VB.TextBox TxtDataDepositoParcela 
            Height          =   288
            Left            =   768
            TabIndex        =   29
            Top             =   528
            Width           =   1530
         End
         Begin MSFlexGridLib.MSFlexGrid GrdParcelas 
            Height          =   2175
            Left            =   270
            TabIndex        =   65
            Top             =   870
            Width           =   8145
            _ExtentX        =   14351
            _ExtentY        =   3831
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
            SelectionMode   =   1
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Forma Pagto"
            Height          =   195
            Left            =   2565
            TabIndex        =   90
            Top             =   255
            Width           =   900
         End
         Begin VB.Label LblNumParcela 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Height          =   192
            Left            =   264
            TabIndex        =   86
            Top             =   576
            Width           =   468
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Num."
            Height          =   192
            Left            =   264
            TabIndex        =   85
            Top             =   336
            Width           =   372
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Valor R$"
            Height          =   195
            Left            =   6870
            TabIndex        =   68
            Top             =   255
            Width           =   630
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Nº Cheque"
            Height          =   195
            Left            =   5010
            TabIndex        =   67
            Top             =   255
            Width           =   780
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Data Depósito"
            Height          =   195
            Left            =   810
            TabIndex        =   66
            Top             =   250
            Width           =   1050
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Condições e Valores"
         Height          =   705
         Left            =   -74880
         TabIndex        =   60
         Top             =   312
         Width           =   10764
         Begin VB.TextBox TxtAdicionalPessoa 
            Height          =   300
            Left            =   9072
            MaxLength       =   5
            TabIndex        =   28
            Top             =   216
            Width           =   1596
         End
         Begin VB.TextBox TxtCustoTotal 
            Height          =   300
            Left            =   4560
            MaxLength       =   8
            TabIndex        =   27
            Top             =   216
            Width           =   1596
         End
         Begin VB.TextBox TxtDataAniversario 
            Height          =   300
            Left            =   1608
            MaxLength       =   10
            TabIndex        =   26
            Top             =   216
            Width           =   1596
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Valor Total"
            Height          =   195
            Left            =   3675
            TabIndex        =   63
            Top             =   270
            Width           =   765
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Custo Adicional por pessoa R$"
            Height          =   192
            Left            =   6816
            TabIndex        =   62
            Top             =   264
            Width           =   2220
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Data do Aniversário"
            Height          =   192
            Left            =   120
            TabIndex        =   61
            Top             =   264
            Width           =   1428
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados do Cliente"
         Height          =   2052
         Left            =   192
         TabIndex        =   48
         Top             =   1632
         Width           =   10572
         Begin VB.TextBox TxtCodCliente 
            Height          =   288
            Left            =   792
            MaxLength       =   4
            TabIndex        =   10
            Top             =   240
            Width           =   924
         End
         Begin VB.TextBox TxtRGCliente 
            Height          =   324
            Left            =   9096
            MaxLength       =   10
            TabIndex        =   21
            Top             =   1632
            Width           =   1380
         End
         Begin VB.TextBox TxtCidadeCliente 
            Height          =   324
            Left            =   4536
            MaxLength       =   30
            TabIndex        =   20
            Top             =   1632
            Width           =   2532
         End
         Begin VB.TextBox TxtBairroCliente 
            Height          =   324
            Left            =   792
            MaxLength       =   30
            TabIndex        =   19
            Top             =   1632
            Width           =   2676
         End
         Begin VB.TextBox TxtTELRESCliente 
            Height          =   324
            Left            =   4536
            MaxLength       =   10
            TabIndex        =   17
            Top             =   1272
            Width           =   1428
         End
         Begin VB.TextBox TxtTELCOMCliente 
            Height          =   324
            Left            =   792
            MaxLength       =   10
            TabIndex        =   16
            Top             =   1272
            Width           =   1428
         End
         Begin VB.TextBox TxtNumCliente 
            Height          =   324
            Left            =   7416
            MaxLength       =   5
            TabIndex        =   14
            Top             =   912
            Width           =   732
         End
         Begin VB.TextBox TxtEndCliente 
            Height          =   288
            Left            =   792
            MaxLength       =   50
            TabIndex        =   13
            Top             =   912
            Width           =   6276
         End
         Begin VB.TextBox TxtNomeCliente 
            Height          =   288
            Left            =   792
            MaxLength       =   60
            TabIndex        =   11
            Top             =   576
            Width           =   6276
         End
         Begin VB.TextBox TxtCPFCliente 
            Height          =   324
            Left            =   9096
            MaxLength       =   11
            TabIndex        =   12
            Top             =   576
            Width           =   1380
         End
         Begin VB.TextBox TxtTELCELCliente 
            Height          =   324
            Left            =   9096
            MaxLength       =   10
            TabIndex        =   18
            Top             =   1272
            Width           =   1380
         End
         Begin VB.TextBox TxtCEPCliente 
            Height          =   324
            Left            =   9096
            MaxLength       =   8
            TabIndex        =   15
            Top             =   912
            Width           =   1380
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   192
            Left            =   144
            TabIndex        =   82
            Top             =   264
            Width           =   528
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Cidade"
            Height          =   192
            Left            =   3960
            TabIndex        =   59
            Top             =   1680
            Width           =   528
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "CEP"
            Height          =   192
            Left            =   8676
            TabIndex        =   58
            Top             =   984
            Width           =   324
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "nº"
            Height          =   192
            Left            =   7176
            TabIndex        =   57
            Top             =   984
            Width           =   144
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "RG"
            Height          =   192
            Left            =   8760
            TabIndex        =   56
            Top             =   1680
            Width           =   240
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   192
            Left            =   288
            TabIndex        =   55
            Top             =   1680
            Width           =   432
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   192
            Left            =   276
            TabIndex        =   54
            Top             =   624
            Width           =   444
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Telefone"
            Height          =   192
            Left            =   72
            TabIndex        =   53
            Top             =   1320
            Width           =   648
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Telefone (res.)"
            Height          =   192
            Left            =   3432
            TabIndex        =   52
            Top             =   1320
            Width           =   1044
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Telefone (cel.)"
            Height          =   192
            Left            =   7968
            TabIndex        =   51
            Top             =   1344
            Width           =   1032
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "CPF"
            Height          =   192
            Left            =   8688
            TabIndex        =   50
            Top             =   624
            Width           =   312
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "End."
            Height          =   192
            Left            =   396
            TabIndex        =   49
            Top             =   984
            Width           =   324
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dados do Evento"
         Height          =   924
         Left            =   192
         TabIndex        =   22
         Top             =   408
         Width           =   10572
         Begin VB.TextBox TxtDataContrato 
            Height          =   288
            Left            =   9270
            MaxLength       =   10
            TabIndex        =   9
            Top             =   555
            Width           =   1215
         End
         Begin VB.TextBox TxtAniversariante 
            Height          =   288
            Left            =   1272
            MaxLength       =   30
            TabIndex        =   2
            Top             =   216
            Width           =   4116
         End
         Begin VB.TextBox TxtDiaFesta 
            Height          =   288
            Left            =   1272
            MaxLength       =   10
            TabIndex        =   6
            Top             =   552
            Width           =   1476
         End
         Begin VB.TextBox TxtIdade 
            Height          =   288
            Left            =   4416
            MaxLength       =   7
            TabIndex        =   7
            Top             =   552
            Width           =   972
         End
         Begin VB.TextBox TxtQtdeConvidados 
            Height          =   288
            Left            =   6888
            MaxLength       =   3
            TabIndex        =   8
            Top             =   552
            Width           =   684
         End
         Begin VB.ComboBox CboHoraInicio 
            Height          =   288
            Left            =   6888
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   216
            Width           =   1236
         End
         Begin VB.ComboBox CboHoraFim 
            Height          =   288
            Left            =   9264
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   192
            Width           =   1260
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Data do Contrato"
            Height          =   195
            Left            =   7980
            TabIndex        =   89
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Aniversariante"
            Height          =   192
            Left            =   144
            TabIndex        =   47
            Top             =   264
            Width           =   1032
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Data"
            Height          =   192
            Left            =   804
            TabIndex        =   46
            Top             =   600
            Width           =   348
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Idade"
            Height          =   192
            Left            =   3936
            TabIndex        =   45
            Top             =   600
            Width           =   420
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Horário Início"
            Height          =   192
            Left            =   5856
            TabIndex        =   25
            Top             =   264
            Width           =   948
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Horário Final"
            Height          =   192
            Left            =   8280
            TabIndex        =   24
            Top             =   264
            Width           =   924
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Num. Convidados"
            Height          =   192
            Left            =   5520
            TabIndex        =   23
            Top             =   600
            Width           =   1284
         End
      End
   End
   Begin VB.Label LblNumeroContrato 
      AutoSize        =   -1  'True
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   16.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   8310
      TabIndex        =   81
      Top             =   315
      Width           =   2400
   End
End
Attribute VB_Name = "Contrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CboFormaPagto_LostFocus()

    If CboFormaPagto.ListIndex <> 0 Then
        TxtValorParcela.SetFocus
    End If

End Sub

Private Sub CmdExcluir_Click()

    GrdParcelas.Col = 1
    x = GrdParcelas.Row
    If Val(GrdParcelas.Text) <> 0 Then
        If GrdParcelas.Rows > 2 Then
            GrdParcelas.RemoveItem x
        Else
            GrdParcelas.Rows = 1
        End If
    End If

    Call CmdLimpar_Click
End Sub
Private Sub CmdGravarContrato_Click()

    Dim Cliente As New ClsCliente
    Dim Contrato As New ClsContrato
    Dim Parcela As New ClsParcela
    Dim ContasAPagar As New ClsContasaPagar

    Dim ID_PAR      As String
    Dim DATA_PAR    As String
    Dim VALOR_PAR   As String
    Dim NUM_DOC_PAR As String
    Dim FORMA_PAGTO As String

    On Error GoTo Erro_Contrato

    If DadosContrato() = True Then
        If Val(LblNumeroContrato.Caption) > 0 Then
            Call Contrato.Excluir(Val(LblNumeroContrato.Caption), Db)
        End If

        'Gravar dados do Contrato
        Call Contrato.Incluir(LblNumeroContrato.Caption, Db)

        'Excluir Parcelas do contrato para regravação
        Call Parcela.Excluir(LblNumeroContrato.Caption, Db)

        'Gravar os dados das Parcelas
        For x = 1 To GrdParcelas.Rows - 1
            GrdParcelas.Row = x

            GrdParcelas.Col = 0
            ID_PAR = GrdParcelas.Text

            GrdParcelas.Col = 1
            DATA_PAR = GrdParcelas.Text

            GrdParcelas.Col = 3
            FORMA_PAGTO = GrdParcelas.Text

            GrdParcelas.Col = 2
            NUM_DOC_PAR = GrdParcelas.Text

            GrdParcelas.Col = 5
            VALOR_PAR = GrdParcelas.Text

            Call Parcela.Incluir(LblNumeroContrato.Caption, ID_PAR, DATA_PAR, VALOR_PAR, NUM_DOC_PAR, FORMA_PAGTO, Db)
        Next x

        'Excluir todas as contas do contrato e gravar novamente
        If ContasAPagar.Excluir(LblNumeroContrato.Caption, Db) = False Then
            MsgBox "Erro ao Excluir Contas do Contrato para atualização", vbExclamation, "SGB"
            Exit Sub
        End If

        'Criar registros de Contas a Pagar - Fornecedor de Decoração
        'If ContasAPagar.Incluir(LblNumeroContrato.Caption, CboDecor.ItemData(CboDecor.ListIndex), 120, "") = False Then
        '    Exit Sub
        'End If

        'Criar registros de Contas a Pagar - Fornecedor de Bolo
        'If ContasAPagar.Incluir(LblNumeroContrato.Caption, CboBolo.ItemData(CboBolo.ListIndex), 120, "") = False Then
        '    Exit Sub
        'End If

        'Criar registros de Contas a Pagar - Fornecedor de Salgado
        'If ContasAPagar.Incluir(LblNumeroContrato.Caption, CboSalgado.ItemData(CboSalgado.ListIndex), 144, "") = False Then
        '    Exit Sub
        'End If

        'Criar registros de Contas a Pagar - Fornecedor de Doce
        'If ContasAPagar.Incluir(LblNumeroContrato.Caption, CboDoce.ItemData(CboDoce.ListIndex), 100, "") = False Then
        '    Exit Sub
        'End If
    End If
    
    Exit Sub

Erro_Contrato:
    MsgBox Error
    MsgBox "Ocorreu um erro e o contrato não pode ser gravado. Feche a tela de CONTRATOS e reinicie a operação.", vbOKOnly, "SGB"
    Exit Sub
End Sub
Private Sub CmdIncluir_Click()

    If (IsDate(TxtDataDepositoParcela.Text) = True) And _
        (Val(TxtValorParcela.Text) > 0 And _
        CboFormaPagto.ListIndex <> -1) Then
        'Inserir nova parcela
        x = GrdParcelas.Rows
        GrdParcelas.AddItem x & Chr(9) & TxtDataDepositoParcela.Text & _
                                Chr(9) & TxtNumChequeParcela.Text & _
                                Chr(9) & CboFormaPagto.ItemData(CboFormaPagto.ListIndex) & _
                                Chr(9) & CboFormaPagto.Text & _
                                Chr(9) & Format(TxtValorParcela.Text, "0.00"), x
        Call CmdLimpar_Click
        TxtDataDepositoParcela.SetFocus
    Else
        MsgBox "Informe corretamente os dados da parcela", vbExclamation, "SGB"
    End If
End Sub
Private Sub CmdIncluirCliente_Click()

    Dim Cliente As New ClsCliente

    'Verificar se os dados do cliente foram informados
    If TxtNomeCliente.Text = "" Then
        MsgBox "É necessário informar os dados do Cliente", vbOKOnly, "SGB"
        Exit Sub
    End If

    If TxtCodCliente.Text <> "" Then
        If MsgBox("Cliente já Cadastrado, deseja incluir?", vbYesNo, "SGB") = vbYes Then
            TxtCodCliente.Text = Cliente.Incluir(TxtNomeCliente.Text, TxtEndCliente.Text, TxtCPFCliente.Text, TxtNumCliente.Text, TxtCEPCliente.Text, TxtTELRESCliente.Text, TxtTELCOMCliente.Text, TxtTELCELCliente.Text, TxtCidadeCliente.Text, TxtBairroCliente.Text, TxtRGCliente.Text, Db)
        Else
            Exit Sub
        End If
    Else
        TxtCodCliente.Text = Cliente.Incluir(TxtNomeCliente.Text, TxtEndCliente.Text, TxtCPFCliente.Text, TxtNumCliente.Text, TxtCEPCliente.Text, TxtTELRESCliente.Text, TxtTELCOMCliente.Text, TxtTELCELCliente.Text, TxtCidadeCliente.Text, TxtBairroCliente.Text, TxtRGCliente.Text, Db)
    End If
End Sub
Private Sub CmdLimpar_Click()

    LblNumParcela.Caption = ""
    TxtDataDepositoParcela.Text = ""
    TxtNumChequeParcela.Text = ""
    TxtValorParcela.Text = ""
End Sub
Private Sub CmdPesquisar_Click()

    If TxtPesquisaNumContrato.Text <> "" Then
        Call PesquisaContrato(TxtPesquisaNumContrato.Text)
    End If
End Sub
Private Sub CmdPesquisarCliente_Click()

    Cliente.Show 1
End Sub
Private Sub CmdSair_Click()

    Unload Me
End Sub
Private Sub Form_Load()

    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

    'Carregar os combos auxiliares
    Call CarregaCombo(CboFormaPagto, "FORMA_PAGTO", "ID_FORMA", "DSC_FORMA")

    'Carregar combos de horário
    Call CarregaCombo(CboHoraInicio, "HORARIO_FESTA", "ID_HORARIO", "DSC_HORARIO")
    Call CarregaCombo(CboHoraFim, "HORARIO_FESTA", "ID_HORARIO", "DSC_HORARIO")

    'Carregar os Combos de Fornecedores
    Call CarregaCombo(CboSalgado, "FORNECEDORES", "ID_FOR", "NOME_FOR", " WHERE TIPO_FOR = 1")
    Call CarregaCombo(CboBolo, "FORNECEDORES", "ID_FOR", "NOME_FOR", " WHERE TIPO_FOR = 2")
    Call CarregaCombo(CboDoce, "FORNECEDORES", "ID_FOR", "NOME_FOR", " WHERE TIPO_FOR = 3")
    Call CarregaCombo(CboDecor, "FORNECEDORES", "ID_FOR", "NOME_FOR", " WHERE TIPO_FOR = 4")

    'Se o contrato já possui parcelas deve exibir
    'If Val(LblNumeroContrato.Caption) > 0 Then
        Call CarregaGrid
    'End If
    
    'TxtAniversariante.SetFocus
    LblNumeroContrato.Caption = "Novo Contrato"
    
    CboBolo.ListIndex = 0
    CboSalgado.ListIndex = 1
    CboDoce.ListIndex = 0
    
End Sub
Private Function DadosContrato() As Integer

    Dim Reserva As New ClsReserva

    'Verificar se os dados do evento foram informados
    If TxtAniversariante.Text = "" Or CboHoraInicio.Text = "" Or CboHoraFim.Text = "" Or _
        TxtDiaFesta.Text = "" Or TxtIdade.Text = "" Or TxtQtdeConvidados.Text = "" Then
        DadosContrato = False
        MsgBox "É necessário informar os dados do Evento", vbOKOnly, "SGB"
        Exit Function
    End If

    'Verificar se os dados do cliente foram informados
    If TxtCodCliente.Text = "" Or TxtNomeCliente.Text = "" Then
        DadosContrato = False
        MsgBox "É necessário informar os dados do Cliente", vbOKOnly, "SGB"
        Exit Function
    End If

    'Verificar se os dados das condições de pagto foram informados
    If TxtDataAniversario.Text = "" Or TxtCustoTotal.Text = "" Or _
       TxtAdicionalPessoa.Text = "" Then
        DadosContrato = False
        MsgBox "É necessário informar os dados da Condição de Pagamento", vbOKOnly, "SGB"
        Exit Function
    End If

    'Verificar se os dados das parcelas foram informados

    'Verificar se os dados das OBS foram informados
    If TxtBolo.Text = "" Or TxtDecoracao.Text = "" Or TxtBebida.Text = "" Or TxtPais.Text = "" Or TxtOBS.Text = "" Then
        DadosContrato = False
        MsgBox "É necessário informar os dados de OBS", vbOKOnly, "SGB"
        Exit Function
    End If

    'Verificar se existe uma reserva para esta data e solicitar a confirmação
    'If Reserva.Selecionar(TxtDataAniversario.Text, CboHoraInicio.ListIndex, Db) = True Then
    '    If MsgBox("Existe uma reserva para esta data/hora. Confirma a inclusão deste contrato?", vbYesNo, "SGB") = vbNo Then
    '        'If MsgBox("O Cliente deste contrato é o mesmo que efetuou a reserva?", vbYesNo, "SGB") = vbYes Then
    '        DadosContrato = False
    '        Exit Function
    '    End If
    'End If

    DadosContrato = True
End Function

Private Sub TxtAdicionalPessoa_GotFocus()

    If Len(Trim(TxtAdicionalPessoa.Text)) > 0 Then
        TxtAdicionalPessoa.SelStart = 0
        TxtAdicionalPessoa.SelLength = Len(TxtAdicionalPessoa.Text)
    End If
End Sub
Private Sub TxtAdicionalPessoa_LostFocus()

    If Val(TxtAdicionalPessoa.Text) > 0 Then
        TxtAdicionalPessoa.Text = Format(TxtAdicionalPessoa.Text, "0.00")
    End If
End Sub

Private Sub TxtAniversariante_GotFocus()

    If Len(Trim(TxtAniversariante.Text)) > 0 Then
        TxtAniversariante.SelStart = 0
        TxtAniversariante.SelLength = Len(TxtAniversariante.Text)
    End If
End Sub
Private Sub TxtBebida_GotFocus()

    If Len(Trim(TxtBebida.Text)) > 0 Then
        TxtBebida.SelStart = 0
        TxtBebida.SelLength = Len(TxtBebida.Text)
    End If
End Sub
Private Sub TxtBolo_GotFocus()

    If Len(Trim(TxtBolo.Text)) > 0 Then
        TxtBolo.SelStart = 0
        TxtBolo.SelLength = Len(TxtBolo.Text)
    End If
End Sub
Private Sub TxtCodCliente_GotFocus()

    If Len(Trim(TxtCodCliente.Text)) > 0 Then
        TxtCodCliente.SelStart = 0
        TxtCodCliente.SelLength = Len(TxtCodCliente.Text)
    End If
End Sub
Private Sub TxtCustoTotal_GotFocus()

    If Len(Trim(TxtCustoTotal.Text)) > 0 Then
        TxtCustoTotal.SelStart = 0
        TxtCustoTotal.SelLength = Len(TxtCustoTotal.Text)
    End If
End Sub

Private Sub TxtCustoTotal_LostFocus()

    If Val(TxtCustoTotal.Text) > 0 Then
        TxtCustoTotal.Text = Format(TxtCustoTotal.Text, "0.00")
    End If
End Sub
Private Sub CarregaGrid()

    Dim Rs As New ADODB.Recordset
    Dim sSql As String
    Dim x As Integer
    
    GrdParcelas.Clear
    GrdParcelas.Rows = 1
    GrdParcelas.Cols = 6

    GrdParcelas.Row = 0

    GrdParcelas.Col = 0
    GrdParcelas.ColWidth(0) = 500
    GrdParcelas.Text = "Parcela"

    GrdParcelas.Col = 1
    GrdParcelas.ColWidth(1) = 1200
    GrdParcelas.Text = "Data"

    GrdParcelas.Col = 2
    GrdParcelas.ColWidth(2) = 1800
    GrdParcelas.Text = "Num. Doc."

    GrdParcelas.Col = 3
    GrdParcelas.ColWidth(3) = 1
    GrdParcelas.Text = ""

    GrdParcelas.Col = 4
    GrdParcelas.ColWidth(4) = 2200
    GrdParcelas.Text = "Forma Pagto"

    GrdParcelas.Col = 5
    GrdParcelas.ColWidth(5) = 1000
    GrdParcelas.Text = "Valor"

    sSql = "Select * from PARCELA_CONTRATO where ID_CNT = " & Val(LblNumeroContrato.Caption)
    Rs.Open sSql, Db, adOpenDynamic, adLockOptimistic

    x = 1
    Do Until Rs.EOF
        GrdParcelas.AddItem x & Chr(9) & Rs("DATA_PAR") & Chr(9) & Rs("NUM_DOC_PAR") & Chr(9) & Format(Rs("VALOR_PAR"), "0.00"), x
        x = x + 1
        Rs.MoveNext
    Loop
End Sub
Private Sub PesquisaContrato(ByVal ID_CNT As Integer)

    Dim Rs As New ADODB.Recordset
    Dim Rs2 As New ADODB.Recordset
    Dim Rs3 As New ADODB.Recordset

    Dim Clientes As New ClsCliente

    sSql = "SELECT c.* , cl.nom_cli FROM CONTRATOS c , CLIENTES cl WHERE c.ID_CNT = " & ID_CNT
    sSql = sSql & " AND c.cod_cli = cl.cod_cli"

    Rs.Open sSql, Db, adOpenDynamic, adLockOptimistic

    If Not Rs.EOF Then
        'Preenche os campos da tela com os dados encontrados no banco de dados
        LblNumeroContrato.Caption = Format(Rs("ID_CNT").Value, "00000")
        TxtAniversariante.Text = Rs("NOME_ANIV").Value
        TxtDiaFesta.Text = Rs("DATA_FESTA").Value
        TxtIdade.Text = Rs("IDADE_ANIV").Value
        TxtQtdeConvidados.Text = Rs("QTDE_CONV").Value
        TxtDataContrato.Text = Rs("DATA_CNT").Value

        Call PesquisaItemCombo(CboHoraInicio, Rs("HR_INI").Value)
        Call PesquisaItemCombo(CboHoraFim, Rs("HR_FIM").Value)

        'Pesquisa os dados do cliente
        Set Rs3 = Clientes.Consultar(Rs("COD_CLI").Value, Db)

        TxtCodCliente.Text = Rs3("COD_CLI").Value
        TxtNomeCliente.Text = Rs3("NOM_CLI").Value
        TxtCPFCliente.Text = Rs3("CPF_CLI").Value
        TxtEndCliente.Text = Rs3("END_CLI").Value
        TxtNumCliente.Text = Rs3("NUM_CLI").Value
        TxtCEPCliente.Text = Rs3("CEP_CLI").Value
        TxtTELRESCliente.Text = Rs3("TEL1_CLI").Value & ""
        TxtTELCOMCliente.Text = Rs3("TEL2_CLI").Value & ""
        TxtTELCELCliente.Text = Rs3("TEL3_CLI").Value & ""
        TxtBairroCliente.Text = Rs3("BAI_CLI").Value & ""
        TxtCidadeCliente.Text = Rs3("CID_CLI").Value & ""
        TxtRGCliente.Text = Rs3("RG_CLI").Value & ""
        

        TxtDataAniversario.Text = Rs("DATA_ANIV").Value
        TxtCustoTotal.Text = Format(Rs("VALOR_TOTAL").Value, "0.00")
        TxtAdicionalPessoa.Text = Format(Rs("CUSTO_ADIC").Value, "0.00")

        'Call PesquisaItemCombo(CboFormaPagto, Rs("ID_FORMA").Value)
        'Call PesquisaItemCombo(CboBolo, Rs("ID_BOLO").Value)
        'Call PesquisaItemCombo(CboDecor, Rs("ID_DECOR").Value)
        'Call PesquisaItemCombo(CboDoce, Rs("ID_DOCE").Value)
        'Call PesquisaItemCombo(CboSalgado, Rs("ID_SALGADO").Value)

        GrdParcelas.Rows = 1
        'Recupera os dados das parcelas
        sSql = "select p.* , f.DSC_FORMA from PARCELA_CONTRATO p , FORMA_PAGTO f "
        sSql = sSql & " where ID_CNT = " & Rs("ID_CNT").Value
        sSql = sSql & " and f.ID_FORMA = p.ID_FORMA "
        sSql = sSql & " ORDER BY ID_PAR"

        Rs2.Open sSql, Db, adOpenDynamic, adLockOptimistic
    
        If Not Rs2.EOF Then
            Do Until Rs2.EOF
                GrdParcelas.AddItem Rs2("ID_PAR").Value & Chr(9) & _
                                    Rs2("DATA_PAR").Value & Chr(9) & _
                                    Rs2("NUM_DOC_PAR").Value & "" & Chr(9) & _
                                    Rs2("ID_FORMA").Value & Chr(9) & _
                                    Rs2("DSC_FORMA").Value & "" & Chr(9) & _
                                    Format(Rs2("VALOR_PAR").Value, "0.00")
                Rs2.MoveNext
            Loop
        End If

        TxtBolo.Text = Rs("DSC_BOLO").Value
        TxtDecoracao.Text = Rs("DSC_DECOR").Value
        TxtBebida.Text = Rs("OBS_BEBIDA").Value
        TxtPais.Text = Rs("NOM_PAIS").Value
        TxtOBS.Text = Rs("OBS").Value
    Else
        MsgBox "Contrato não Encontrado.", vbExclamation, "SGB"
        Call LimpaCampos
    End If
End Sub
Private Sub TxtDataAniversario_GotFocus()

    If Len(Trim(TxtDataAniversario.Text)) > 0 Then
        TxtDataAniversario.SelStart = 0
        TxtDataAniversario.SelLength = Len(TxtDataAniversario.Text)
    End If
End Sub
Private Sub TxtDataAniversario_LostFocus()

    TxtDataAniversario.Text = Format(TxtDataAniversario.Text, "00/00/0000")
End Sub
Private Sub TxtDataContrato_GotFocus()

    If Len(Trim(TxtDataContrato.Text)) > 0 Then
        TxtDataContrato.SelStart = 0
        TxtDataContrato.SelLength = Len(TxtDataContrato.Text)
    End If
End Sub
Private Sub TxtDataContrato_LostFocus()

    TxtDataContrato.Text = Format(TxtDataContrato.Text, "00/00/0000")
End Sub

Private Sub TxtDataDepositoParcela_LostFocus()

    TxtDataDepositoParcela.Text = Format(TxtDataDepositoParcela.Text, "00/00/0000")
End Sub
Private Sub TxtDecoracao_GotFocus()

    If Len(Trim(TxtDecoracao.Text)) > 0 Then
        TxtDecoracao.SelStart = 0
        TxtDecoracao.SelLength = Len(TxtDecoracao.Text)
    End If
End Sub
Private Sub TxtDiaFesta_GotFocus()

    If Len(Trim(TxtDiaFesta.Text)) > 0 Then
        TxtDiaFesta.SelStart = 0
        TxtDiaFesta.SelLength = Len(TxtDiaFesta.Text)
    End If
End Sub
Private Sub TxtDiaFesta_LostFocus()

    TxtDiaFesta.Text = Format(TxtDiaFesta.Text, "00/00/0000")
End Sub
Private Sub TxtIdade_GotFocus()

    If Len(Trim(TxtIdade.Text)) > 0 Then
        TxtIdade.SelStart = 0
        TxtIdade.SelLength = Len(TxtIdade.Text)
    End If
End Sub
Private Sub TxtNomeCliente_GotFocus()

    If Len(Trim(TxtNomeCliente.Text)) > 0 Then
        TxtNomeCliente.SelStart = 0
        TxtNomeCliente.SelLength = Len(TxtNomeCliente.Text)
    End If
End Sub
Private Sub TxtOBS_GotFocus()

    If Len(Trim(TxtOBS.Text)) > 0 Then
        TxtOBS.SelStart = 0
        TxtOBS.SelLength = Len(TxtOBS.Text)
    End If
End Sub
Private Sub TxtPais_GotFocus()

    If Len(Trim(TxtPais.Text)) > 0 Then
        TxtPais.SelStart = 0
        TxtPais.SelLength = Len(TxtPais.Text)
    End If
End Sub

Private Sub TxtQtdeConvidados_GotFocus()

    If Len(Trim(TxtQtdeConvidados.Text)) > 0 Then
        TxtQtdeConvidados.SelStart = 0
        TxtQtdeConvidados.SelLength = Len(TxtQtdeConvidados.Text)
    End If
End Sub
Private Sub TxtValorParcela_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call CmdIncluir_Click
    End If
End Sub
Private Sub LimpaCampos()

    LblNumeroContrato.Caption = "Novo Contrato"
    TxtAniversariante.Text = ""
    TxtDiaFesta.Text = ""
    TxtIdade.Text = ""
    TxtQtdeConvidados.Text = ""
    CboHoraInicio.ListIndex = -1
    CboHoraFim.ListIndex = -1

    TxtCodCliente.Text = ""
    TxtNomeCliente.Text = ""
    TxtCPFCliente.Text = ""
    TxtEndCliente.Text = ""
    TxtNumCliente.Text = ""
    TxtCEPCliente.Text = ""
    TxtTELRESCliente.Text = ""
    TxtTELCOMCliente.Text = ""
    TxtTELCELCliente.Text = ""
    TxtBairroCliente.Text = ""
    TxtCidadeCliente.Text = ""
    TxtRGCliente.Text = ""

    TxtDataAniversario.Text = ""
    TxtCustoTotal.Text = ""
    TxtAdicionalPessoa.Text = ""
    TxtDataContrato.Text = ""

    CboFormaPagto.ListIndex = -1
    CboBolo.ListIndex = -1
    CboDecor.ListIndex = -1
    CboDoce.ListIndex = -1
    CboSalgado.ListIndex = -1

    GrdParcelas.Rows = 1
    TxtBolo.Text = ""
    TxtDecoracao.Text = ""
    TxtBebida.Text = ""
    TxtPais.Text = ""
    TxtOBS.Text = ""
End Sub
