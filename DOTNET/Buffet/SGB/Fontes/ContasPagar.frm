VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form ContasPagar 
   Caption         =   "SGB - Contas a Pagar"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Height          =   348
      Left            =   8145
      TabIndex        =   10
      Top             =   6750
      Width           =   1188
   End
   Begin VB.CommandButton CmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sai&r"
      Height          =   348
      Left            =   195
      TabIndex        =   9
      Top             =   6750
      Width           =   1188
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Contrato"
      Height          =   1095
      Left            =   210
      TabIndex        =   1
      Top             =   120
      Width           =   9090
      Begin VB.TextBox TxtTipoFesta 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   7470
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   630
         Width           =   1470
      End
      Begin VB.CommandButton CmdPesquisar 
         Caption         =   "&Pesquisar"
         Default         =   -1  'True
         Height          =   300
         Left            =   2520
         TabIndex        =   11
         Top             =   264
         Width           =   948
      End
      Begin VB.TextBox TxtQtde 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   645
         Width           =   1110
      End
      Begin VB.TextBox TxtHoraFesta 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   624
         Width           =   1380
      End
      Begin VB.TextBox TxtDataFesta 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   888
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   624
         Width           =   1116
      End
      Begin VB.TextBox TxtContrato 
         Height          =   300
         Left            =   888
         TabIndex        =   0
         Top             =   264
         Width           =   1548
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Festa"
         Height          =   195
         Left            =   6615
         TabIndex        =   12
         Top             =   690
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Qtde Pessoas"
         Height          =   195
         Left            =   4320
         TabIndex        =   5
         Top             =   690
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hora"
         Height          =   195
         Left            =   2235
         TabIndex        =   4
         Top             =   675
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   192
         Left            =   384
         TabIndex        =   3
         Top             =   672
         Width           =   348
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Contrato"
         Height          =   192
         Left            =   192
         TabIndex        =   2
         Top             =   336
         Width           =   600
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   180
      TabIndex        =   14
      Top             =   1305
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   5
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   420
      Enabled         =   0   'False
      TabCaption(0)   =   "Fornecedores"
      TabPicture(0)   =   "ContasPagar.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Funcionários"
      TabPicture(1)   =   "ContasPagar.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label22"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label20"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label19"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label18"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Grid"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "CmdRemover"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "CmdAdicionar"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "LstFuncoes"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "LstFuncionarios"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Despesas Fixas"
      TabPicture(2)   =   "ContasPagar.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label21"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label17"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "GrdDespesas"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Command2"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Command1"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "CboDespesa"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Text4"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Pedido Salgados"
      TabPicture(3)   =   "ContasPagar.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label41"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label40"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label38"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label35"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label33"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Label30"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "GridPaes"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "GridSalgados"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "CmdExcluiPao"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "CmdExcluiSalgado"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "CboFornecSalgados"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "CboFornecPaes"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "CmdAdicionaPaes"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "CboPaes"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "TxtQtdePaes"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "CmdAdicionaSalgado"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "CboSalgados"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "TxtQtdeSalgados"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).ControlCount=   18
      TabCaption(4)   =   "Pedido Doces"
      TabPicture(4)   =   "ContasPagar.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label28"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label27"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label25"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "MSFlexGrid1"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Text5"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Combo2"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Command4"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "CboPedidoDoces"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "Command3"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).ControlCount=   9
      Begin VB.Frame Frame2 
         Caption         =   "Fornecedores"
         Height          =   4095
         Left            =   -74055
         TabIndex        =   41
         Top             =   810
         Width           =   7740
         Begin VB.ComboBox CboBolo 
            Height          =   315
            Left            =   1230
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   315
            Width           =   1668
         End
         Begin VB.ComboBox CboDoce 
            Height          =   315
            Left            =   1230
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   1635
            Width           =   1710
         End
         Begin VB.ComboBox CboMesa 
            Height          =   315
            Left            =   1230
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   2010
            Width           =   1710
         End
         Begin VB.ComboBox CboAnimacao 
            Height          =   315
            Left            =   1215
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   3330
            Visible         =   0   'False
            Width           =   1668
         End
         Begin VB.ComboBox CboDJ 
            Height          =   315
            Left            =   1230
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   3690
            Visible         =   0   'False
            Width           =   1668
         End
         Begin VB.TextBox TxtValorBolo 
            Height          =   285
            Left            =   6090
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   315
            Width           =   1332
         End
         Begin VB.TextBox TxtValorSalgado1 
            Height          =   285
            Left            =   6090
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   900
            Width           =   1332
         End
         Begin VB.TextBox TxtValorDoce 
            Height          =   285
            Left            =   6090
            Locked          =   -1  'True
            TabIndex        =   51
            Top             =   1635
            Width           =   1332
         End
         Begin VB.TextBox TxtValorMesa 
            Height          =   285
            Left            =   6090
            TabIndex        =   50
            Top             =   1965
            Width           =   1332
         End
         Begin VB.TextBox TxtValorAnimacao 
            Height          =   285
            Left            =   6090
            TabIndex        =   49
            Top             =   3390
            Visible         =   0   'False
            Width           =   1332
         End
         Begin VB.TextBox TxtValorDJ 
            Height          =   285
            Left            =   6090
            TabIndex        =   48
            Top             =   3735
            Visible         =   0   'False
            Width           =   1332
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   3840
            TabIndex        =   47
            Top             =   315
            Width           =   885
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   3840
            TabIndex        =   46
            Top             =   1635
            Width           =   885
         End
         Begin VB.TextBox TxtValorPaes1 
            Height          =   285
            Left            =   6090
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   1260
            Width           =   1332
         End
         Begin VB.TextBox TxtQtdePaes1 
            Height          =   330
            Left            =   3840
            TabIndex        =   44
            Top             =   1260
            Width           =   885
         End
         Begin VB.TextBox TxtPaes1 
            Height          =   285
            Left            =   1230
            Locked          =   -1  'True
            TabIndex        =   43
            Text            =   "TOTAL FORNEC"
            Top             =   1260
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1230
            Locked          =   -1  'True
            TabIndex        =   42
            Text            =   "TOTAL FORNEC"
            Top             =   900
            Width           =   1695
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Animação"
            Height          =   195
            Left            =   405
            TabIndex        =   75
            Top             =   3390
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Mesa"
            Height          =   195
            Left            =   720
            TabIndex        =   74
            Top             =   2070
            Width           =   405
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Salgados"
            Height          =   195
            Left            =   420
            TabIndex        =   73
            Top             =   900
            Width           =   705
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "DJ's"
            Height          =   195
            Left            =   810
            TabIndex        =   72
            Top             =   3720
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Bolo"
            Height          =   195
            Left            =   795
            TabIndex        =   71
            Top             =   360
            Width           =   330
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Doces"
            Height          =   195
            Left            =   645
            TabIndex        =   70
            Top             =   1680
            Width           =   480
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            Height          =   195
            Left            =   5625
            TabIndex        =   69
            Top             =   360
            Width           =   390
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            Height          =   195
            Left            =   5625
            TabIndex        =   68
            Top             =   915
            Width           =   390
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            Height          =   195
            Left            =   5625
            TabIndex        =   67
            Top             =   1680
            Width           =   390
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            Height          =   195
            Left            =   5625
            TabIndex        =   66
            Top             =   2025
            Width           =   390
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            Height          =   195
            Left            =   5625
            TabIndex        =   65
            Top             =   3435
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            Height          =   195
            Left            =   5625
            TabIndex        =   64
            Top             =   3780
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Kg"
            Height          =   195
            Left            =   3555
            TabIndex        =   63
            Top             =   360
            Width           =   195
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Qtde"
            Height          =   195
            Left            =   3405
            TabIndex        =   62
            Top             =   1680
            Width           =   345
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            Height          =   195
            Left            =   5625
            TabIndex        =   61
            Top             =   1305
            Width           =   390
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Pães"
            Height          =   195
            Left            =   765
            TabIndex        =   60
            Top             =   1305
            Width           =   360
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "Qtde"
            Height          =   240
            Left            =   3405
            TabIndex        =   59
            Top             =   1305
            Width           =   345
         End
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00FFFFFF&
         Height          =   288
         Left            =   -71448
         TabIndex        =   39
         Top             =   360
         Width           =   1188
      End
      Begin VB.ComboBox CboDespesa 
         Height          =   315
         Left            =   -74088
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   360
         Width           =   2004
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Excluir"
         Height          =   276
         Left            =   -70104
         TabIndex        =   37
         Top             =   1728
         Width           =   1020
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Inserir"
         Height          =   276
         Left            =   -70104
         TabIndex        =   36
         Top             =   1032
         Width           =   1020
      End
      Begin VB.ListBox LstFuncionarios 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2370
         Left            =   -74520
         MultiSelect     =   2  'Extended
         TabIndex        =   35
         Top             =   720
         Width           =   2295
      End
      Begin VB.ListBox LstFuncoes 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         Left            =   -74520
         TabIndex        =   34
         Top             =   3585
         Width           =   2295
      End
      Begin VB.CommandButton CmdAdicionar 
         Height          =   1095
         Left            =   -71625
         Picture         =   "ContasPagar.frx":008C
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1530
         Width           =   975
      End
      Begin VB.CommandButton CmdRemover 
         Height          =   1095
         Left            =   -71640
         Picture         =   "ContasPagar.frx":04CE
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox TxtQtdeSalgados 
         Height          =   330
         Left            =   2790
         TabIndex        =   29
         Top             =   1485
         Width           =   1005
      End
      Begin VB.ComboBox CboSalgados 
         Height          =   315
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1485
         Width           =   2085
      End
      Begin VB.CommandButton CmdAdicionaSalgado 
         Caption         =   "OK"
         Height          =   330
         Left            =   3825
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1485
         Width           =   465
      End
      Begin VB.TextBox TxtQtdePaes 
         Height          =   330
         Left            =   7335
         TabIndex        =   26
         Top             =   1485
         Width           =   1005
      End
      Begin VB.ComboBox CboPaes 
         Height          =   315
         Left            =   4770
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1485
         Width           =   2085
      End
      Begin VB.CommandButton CmdAdicionaPaes 
         Caption         =   "OK"
         Height          =   330
         Left            =   8370
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1485
         Width           =   375
      End
      Begin VB.ComboBox CboFornecPaes 
         Height          =   315
         Left            =   4770
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   765
         Width           =   3960
      End
      Begin VB.ComboBox CboFornecSalgados 
         Height          =   315
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   765
         Width           =   3960
      End
      Begin VB.CommandButton CmdExcluiSalgado 
         Height          =   555
         Left            =   1845
         Picture         =   "ContasPagar.frx":0910
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3780
         Width           =   555
      End
      Begin VB.CommandButton CmdExcluiPao 
         Height          =   555
         Left            =   6390
         Picture         =   "ContasPagar.frx":0D52
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3780
         Width           =   555
      End
      Begin VB.CommandButton Command3 
         Height          =   555
         Left            =   -73020
         Picture         =   "ContasPagar.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3735
         Width           =   555
      End
      Begin VB.ComboBox CboPedidoDoces 
         Height          =   315
         Left            =   -74775
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   720
         Width           =   3960
      End
      Begin VB.CommandButton Command4 
         Caption         =   "OK"
         Height          =   330
         Left            =   -71040
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1440
         Width           =   465
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   -74775
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1440
         Width           =   2085
      End
      Begin VB.TextBox Text5 
         Height          =   330
         Left            =   -72075
         TabIndex        =   15
         Top             =   1440
         Width           =   1005
      End
      Begin MSFlexGridLib.MSFlexGrid GridSalgados 
         Height          =   1950
         Left            =   90
         TabIndex        =   30
         Top             =   1845
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   3440
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedCols       =   0
         ScrollBars      =   0
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   4350
         Left            =   -69960
         TabIndex        =   31
         Top             =   720
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   7673
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         FocusRect       =   2
         HighLight       =   2
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid GrdDespesas 
         Height          =   1620
         Left            =   -74880
         TabIndex        =   40
         Top             =   792
         Width           =   4644
         _ExtentX        =   8202
         _ExtentY        =   2858
         _Version        =   393216
         FixedCols       =   0
         SelectionMode   =   2
      End
      Begin MSFlexGridLib.MSFlexGrid GridPaes 
         Height          =   1950
         Left            =   4770
         TabIndex        =   76
         Top             =   1845
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   3440
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1950
         Left            =   -74775
         TabIndex        =   77
         Top             =   1800
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   3440
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedCols       =   0
         ScrollBars      =   0
         SelectionMode   =   1
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   192
         Left            =   -71904
         TabIndex        =   92
         Top             =   408
         Width           =   384
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Despesa"
         Height          =   192
         Left            =   -74832
         TabIndex        =   91
         Top             =   408
         Width           =   672
      End
      Begin VB.Label Label18 
         Caption         =   "Label18"
         Height          =   15
         Left            =   -74400
         TabIndex        =   90
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Funcionários"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74520
         TabIndex        =   89
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Funções"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74520
         TabIndex        =   88
         Top             =   3345
         Width           =   735
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Escala"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -69960
         TabIndex        =   87
         Top             =   480
         Width           =   585
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Pães"
         Height          =   195
         Left            =   4815
         TabIndex        =   86
         Top             =   1260
         Width           =   360
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "Salgados"
         Height          =   195
         Left            =   90
         TabIndex        =   85
         Top             =   1260
         Width           =   705
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Qtde"
         Height          =   195
         Left            =   2385
         TabIndex        =   84
         Top             =   1530
         Width           =   345
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "Qtde"
         Height          =   195
         Left            =   6930
         TabIndex        =   83
         Top             =   1530
         Width           =   345
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor"
         Height          =   195
         Left            =   90
         TabIndex        =   82
         Top             =   495
         Width           =   810
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor"
         Height          =   195
         Left            =   4770
         TabIndex        =   81
         Top             =   540
         Width           =   810
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor"
         Height          =   195
         Left            =   -74775
         TabIndex        =   80
         Top             =   450
         Width           =   810
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Qtde"
         Height          =   195
         Left            =   -72480
         TabIndex        =   79
         Top             =   1485
         Width           =   345
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Doces"
         Height          =   195
         Left            =   -74775
         TabIndex        =   78
         Top             =   1215
         Width           =   465
      End
   End
End
Attribute VB_Name = "ContasPagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FormataGrid()

    'Formatar o grid de Salgados
    GridSalgados.Cols = 5

    GridSalgados.Row = 0

    GridSalgados.Col = 1
    GridSalgados.Text = "Fornec"
    GridSalgados.Col = 3
    GridSalgados.Text = "Salgado"
    GridSalgados.Col = 4
    GridSalgados.Text = "Qtde"

    GridSalgados.ColWidth(0) = 1
    GridSalgados.ColWidth(1) = 800
    GridSalgados.ColWidth(2) = 1
    GridSalgados.ColWidth(3) = 2000
    GridSalgados.ColWidth(4) = 1000


    'Formatar o grid de Paes
    GridPaes.Cols = 5

    GridPaes.Row = 0

    GridPaes.Col = 1
    GridPaes.Text = "Fornec"
    GridPaes.Col = 3
    GridPaes.Text = "Pão"
    GridPaes.Col = 4
    GridPaes.Text = "Qtde"

    GridPaes.ColWidth(0) = 1
    GridPaes.ColWidth(1) = 800
    GridPaes.ColWidth(2) = 1
    GridPaes.ColWidth(3) = 2000
    GridPaes.ColWidth(4) = 1000
End Sub
Private Sub CmdExcluir_Click()

    Dim x As Integer

    GrdFunc.Col = 0
    x = GrdFunc.Row
    If Val(GrdFunc.Text) <> 0 Then
        GrdFunc.Col = 4
        If InStr(LCase(GrdFunc.Text), "monit") <> 0 Then
            LblMonit.Text = Val(LblMonit.Text) - 1
        End If
        If InStr(LCase(GrdFunc.Text), "garç") <> 0 Then
            LblGarcom.Text = Val(LblGarcom.Text) - 1
        End If

        If GrdFunc.Rows > 2 Then
            GrdFunc.RemoveItem x
        Else
            GrdFunc.Rows = 1
        End If
    End If

    CboFuncao.ListIndex = -1
    CboColab.ListIndex = -1
End Sub
Private Sub CmdInserir_Click()

    x = GrdFunc.Rows
    If CboColab.ListIndex <> -1 And CboFuncao.ListIndex <> -1 Then
        GrdFunc.AddItem x & Chr(9) & CboColab.ItemData(CboColab.ListIndex) & Chr(9) & _
                        CboColab.Text & Chr(9) & _
                        CboFuncao.ItemData(CboFuncao.ListIndex) & Chr(9) & _
                        CboFuncao.Text, x
    End If

    'Atualizar os contadores
    If InStr(LCase(CboFuncao.Text), "monit") <> 0 Then
        LblMonit.Text = Val(LblMonit.Text) + 1
    End If

    If InStr(LCase(CboFuncao.Text), "garç") <> 0 Then
        LblGarcom.Text = Val(LblGarcom.Text) + 1
    End If

    CboFuncao.ListIndex = -1
    CboColab.ListIndex = -1
End Sub

Private Sub CboBolo_Click()

    'Pesquisar na tabela FORNECEDORES o valor que este fornecedor pratica
    
    
    'Multiplicar o valor pela quantidade/kg que o usuário informou
    
    'Salvar na tabela CONTAS_A_PAGAR
    
End Sub

Private Sub CmdAdicionaPaes_Click()

    'Verificar se foram informados o Fornecedor, o tipo de Salgado e a quantidade
    If CboFornecPaes.ListIndex <> -1 And CboPaes.ListIndex <> -1 And Val(TxtQtdePaes.Text) > 0 Then
        'Incluir no grid as informações digitadas
        GridPaes.AddItem CboFornecPaes.ItemData(CboFornecPaes.ListIndex) & Chr(9) & _
                             CboFornecPaes.List(CboFornecPaes.ListIndex) & Chr(9) & _
                             CboPaes.ItemData(CboPaes.ListIndex) & Chr(9) & _
                             CboPaes.List(CboPaes.ListIndex) & Chr(9) & _
                             Val(TxtQtdePaes.Text)
    Else
        MsgBox "As informações não estão corretas!", vbOKOnly, "SGB"
        Exit Sub
    End If
End Sub

Private Sub CmdAdicionar_Click()

    'Incluir os itens selecionados na lista de escala

    'Verificar se foi selecionado ao menos 1 funcionário e 1 função
    If LstFuncionarios.ListCount = 0 Or LstFuncoes.ListIndex = -1 Then
        MsgBox "É necessário selecionar ao menos 1 funcionário e 1 função", vbOKOnly, "SGB"
        Exit Sub
    End If

    For x = 0 To LstFuncionarios.ListCount - 1
        If LstFuncionarios.Selected(x) = True Then
            Grid.AddItem LstFuncionarios.ItemData(x) & Chr(9) & LstFuncionarios.List(x) & Chr(9) & LstFuncoes.ItemData(LstFuncoes.ListIndex) & Chr(9) & LstFuncoes.Text
            'LstEscala.AddItem LstFuncionarios.List(x) & Space(20 - Len(LstFuncionarios.List(x))) & LstFuncoes.Text
        End If
    Next x

End Sub

Private Sub CmdAdicionaSalgado_Click()

    'Verificar se foram informados o Fornecedor, o tipo de Salgado e a quantidade
    If CboFornecSalgados.ListIndex <> -1 And CboSalgados.ListIndex <> -1 And Val(TxtQtdeSalgados.Text) > 0 Then
        'Incluir no grid as informações digitadas
        GridSalgados.AddItem CboFornecSalgados.ItemData(CboFornecSalgados.ListIndex) & Chr(9) & _
                             CboFornecSalgados.List(CboFornecSalgados.ListIndex) & Chr(9) & _
                             CboSalgados.ItemData(CboSalgados.ListIndex) & Chr(9) & _
                             CboSalgados.List(CboSalgados.ListIndex) & Chr(9) & _
                             Val(TxtQtdeSalgados.Text)
    Else
        MsgBox "As informações não estão corretas!", vbOKOnly, "SGB"
        Exit Sub
    End If
End Sub
Private Sub CmdOK_Click()

    Dim Contas As New ClsContasaPagar

    If SSTab1.Enabled = False Then Exit Sub

    'Excluir os dados dos fornecedores
    Call Contas.Excluir(TxtContrato.Text, Db)

    'Criar registros de Contas a Pagar - Fornecedor de Bolo
    If CboBolo.ListIndex <> -1 And Val(TxtValorBolo.Text) > 0 Then
        If Contas.Incluir(TxtContrato.Text, CboBolo.ItemData(CboBolo.ListIndex), 0, 0, TxtValorBolo.Text) = False Then
            Exit Sub
        End If
    End If

    'Criar registros de Contas a Pagar - Fornecedor de Decoração
    If CboMesa.ListIndex <> -1 And Val(TxtValorMesa.Text) > 0 Then
        If Contas.Incluir(TxtContrato.Text, CboMesa.ItemData(CboMesa.ListIndex), 0, 0, TxtValorMesa.Text) = False Then
            Exit Sub
        End If
    End If

    'Criar registros de Contas a Pagar - Fornecedor de Doce
    If CboDoce.ListIndex <> -1 And Val(TxtValorDoce.Text) > 0 Then
        If Contas.Incluir(TxtContrato.Text, CboDoce.ItemData(CboDoce.ListIndex), 0, 0, TxtValorDoce.Text) = False Then
            Exit Sub
        End If
    End If

    'Criar registros de Contas a Pagar - Fornecedor de Salgados
    'Verificar se o pedido de pães foi informado
    If GridSalgados.Rows > 1 Then
        For x = 1 To GridSalgados.Rows - 1
            GridSalgados.Col = 0
            FORNEC = GridSalgados.Text
            GridSalgados.Col = 2
            ID_PRD = GridSalgados.Text
            GridSalgados.Col = 4
            QTDE = GridSalgados.Text

            If Contas.Incluir(TxtContrato.Text, FORNEC, ID_PRD, GridSalgados.Text, VALOR) = False Then Exit Sub
        Next x
    End If

    'Criar registros de Contas a Pagar - Fornecedor de Pães
    Call Contas.ExcluirEscala(TxtContrato.Text)

    'Inserir os dados da escala (funcionários)
    For x = 1 To Grid.Rows - 1
        Grid.Row = x
        Grid.Col = 0
        ID_COL = Grid.Text
        Grid.Col = 2
        ID_FUNC = Grid.Text
        If Contas.IncluirEscala(TxtContrato.Text, ID_COL, ID_FUNC) = False Then
            Exit Sub
        End If
    Next x

    'Inserir dados de pedidos de salgados
    'For x = 1 To GridSalgados.Rows - 1
    '    GridSalgados.Row = x
    '    GridSalgados.Col = 0
    '    ID_FOR = GridSalgados.Text
    '    Grid.Col = 2
    '    ID_FUNC = GridSalgados.Text
    '    If Contas.IncluirEscala(TxtContrato.Text, ID_COL, ID_FUNC) = False Then
    '        Exit Sub
    '    End If
    'Next x

    MsgBox "Inclusão Realizada com Sucesso", vbOKOnly, "SGB"
End Sub
Private Sub CmdPesquisar_Click()

    Dim Contrato As New ClsContrato
    Dim Rs As New ADODB.Recordset
    Dim Rs2 As New ADODB.Recordset
    Dim sSql As String

    If Len(Trim(TxtContrato.Text)) > 0 Then
        'Pesquisar os dados do contrato
        sSql = " select CT.DATA_FESTA , H.DSC_HORARIO, CT.QTDE_CONV , CT.TIPO_FESTA "
        sSql = sSql & " from CONTRATOS CT , HORARIO_FESTA H"
        sSql = sSql & " WHERE CT.ID_CNT = " & TxtContrato.Text
        sSql = sSql & " AND CT.HR_INI = H.ID_HORARIO "

        Rs.Open sSql, Db, adOpenDynamic, adLockReadOnly
        
        If Not Rs.EOF Then
            TxtDataFesta.Text = Rs("DATA_FESTA").Value
            TxtHoraFesta.Text = Rs("DSC_HORARIO").Value
            TxtQtde.Text = Rs("QTDE_CONV").Value
            
            If Rs("TIPO_FESTA").Value = "C" Then
                TxtTipoFesta.Text = "COMPLETA"
            ElseIf Rs("TIPO_FESTA").Value = "E" Then
                TxtTipoFesta.Text = "ECONÔMICA"
            Else
                TxtTipoFesta.Text = "Não Informado"
            End If
        Else
            MsgBox "Contrato não encontrado", vbInformation, "SGB"
            Call LimpaCampos

            Grid.Rows = 1
            Exit Sub
        End If
        
        Set Rs = Nothing
        sSql = " select C.* , F.TIPO_FOR , CT.DATA_FESTA , H.DSC_HORARIO, CT.QTDE_CONV "
        sSql = sSql & " from CONTAS_A_PAGAR C , FORNECEDORES F , CONTRATOS CT , HORARIO_FESTA H"
        sSql = sSql & " WHERE CT.ID_CNT = " & TxtContrato.Text
        sSql = sSql & " AND F.ID_FOR = C.ID_FOR "
        sSql = sSql & " AND CT.ID_CNT = C.ID_CNT "
        sSql = sSql & " AND CT.HR_INI = H.ID_HORARIO "

        Rs.Open sSql, Db, adOpenDynamic, adLockReadOnly

        If Not Rs.EOF Then
            'Preencher os dados do contrato
            TxtDataFesta.Text = Rs("DATA_FESTA").Value
            TxtHoraFesta.Text = Rs("DSC_HORARIO").Value
            TxtQtde.Text = Rs("QTDE_CONV").Value
            
            Do Until Rs.EOF
                Select Case Rs("TIPO_FOR")
                Case 1 'Salgado
                    'Call PesquisaItemCombo(Cbo, Rs("ID_FOR"))
                    'TxtValorSalgado.Text = Format(Rs("VALOR").Value, "0.00")
       
                Case 2 'Bolo
                    Call PesquisaItemCombo(CboBolo, Rs("ID_FOR"))
                    TxtValorBolo.Text = Format(Rs("VALOR").Value, "0.00")
       
                Case 3 'Doce
                    Call PesquisaItemCombo(CboDoce, Rs("ID_FOR"))
                    TxtValorDoce.Text = Format(Rs("VALOR").Value, "0.00")
       
                Case 4 'Decoracao
                    Call PesquisaItemCombo(CboMesa, Rs("ID_FOR"))
                    TxtValorMesa.Text = Format(Rs("VALOR").Value, "0.00")
                End Select
                Rs.MoveNext
            Loop
        Else
            'MsgBox "Contrato não encontrado", vbInformation, "SGB"
            'Call LimpaCampos
            'Exit Sub
        End If
        
        Set Rs = Nothing
        'Escala
        sSql = "SELECT E.ID_COL , C.NOME_COL , E.ID_FUNC , F.DSC_FUNC "
        sSql = sSql & " FROM ESCALA E , COLABORADORES C , FUNCOES F "
        sSql = sSql & " WHERE ID_CNT = " & TxtContrato.Text
        sSql = sSql & " AND E.ID_COL = C.ID_COL "
        sSql = sSql & " AND E.ID_FUNC = F.ID_FUNC "
        sSql = sSql & " ORDER BY E.ID_FUNC , C.NOME_COL"
        
        Rs2.Open sSql, Db, adOpenDynamic, adLockOptimistic
        
        If Not Rs2.EOF Then
            Call CarregaGridEscala(Rs2)
        Else
            Rs2.Close
            sSql = "SELECT E.ID_COL , C.NOME_COL , E.ID_FUNC , F.DSC_FUNC"
            sSql = sSql & " FROM ESCALA E , COLABORADORES C , FUNCOES F"
            sSql = sSql & " WHERE ID_CNT = 0"
            sSql = sSql & " AND E.ID_COL = C.ID_COL"
            sSql = sSql & " AND E.ID_FUNC = F.ID_FUNC"
            sSql = sSql & " ORDER BY C.NOME_COL"

            Rs2.Open sSql, Db, adOpenDynamic, adLockOptimistic
            Call CarregaGridEscala(Rs2)
        End If

        SSTab1.Enabled = True
    Else
        MsgBox "Informe o Contrato.", vbExclamation, "SGB"
        TxtContrato.SetFocus
        Exit Sub
    End If

End Sub
Public Sub PesquisaItemCombo(ByRef Combo As ComboBox, Item As String)

    For x = 0 To Combo.ListCount - 1
        If Combo.ItemData(x) = Item Then
            Combo.ListIndex = x
            Exit For
        End If
    Next x
End Sub
Private Sub CmdRemover_Click()

    If Grid.Rows <= 2 Then
        Grid.Rows = 1
    Else
        Grid.RemoveItem Grid.Row
    End If
End Sub

Private Sub CmdSair_Click()

    Unload Me
End Sub

Private Sub Form_Load()

    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

    Call FormataGrid
    Call FormataGridEscala

    'Carregar os Combos de Fornecedores
    Call CarregaCombo(CboFornecSalgados, "FORNECEDORES", "ID_FOR", "NOME_FOR", " WHERE TIPO_FOR = 1")
    Call CarregaCombo(CboFornecPaes, "FORNECEDORES", "ID_FOR", "NOME_FOR", " WHERE TIPO_FOR = 6")

    Call CarregaCombo(CboSalgados, "PRODUTOS", "ID_PRD", "DSC_PRD", " WHERE TIPO_PRD = 1")
    Call CarregaCombo(CboPaes, "PRODUTOS", "ID_PRD", "DSC_PRD", " WHERE TIPO_PRD = 2")

    Call CarregaCombo(CboBolo, "FORNECEDORES", "ID_FOR", "NOME_FOR", " WHERE TIPO_FOR = 2")
    Call CarregaCombo(CboDoce, "FORNECEDORES", "ID_FOR", "NOME_FOR", " WHERE TIPO_FOR = 3")
    Call CarregaCombo(CboMesa, "FORNECEDORES", "ID_FOR", "NOME_FOR", " WHERE TIPO_FOR = 4")

    'Carregar combos de colaboradores e funções
    Call CarregaCombo(LstFuncionarios, "COLABORADORES", "ID_COL", "NOME_COL")
    Call CarregaCombo(LstFuncoes, "FUNCOES", "ID_FUNC", "DSC_FUNC")
End Sub
Private Sub LimpaCampos()

    TxtContrato.Text = ""

    CboBolo.ListIndex = -1
    TxtValorBolo.Text = ""

    CboSalgado.ListIndex = -1
    TxtValorSalgado.Text = ""

    CboDoce.ListIndex = -1
    TxtValorDoce.Text = ""

    CboMesa.ListIndex = -1
    TxtValorMesa.Text = ""

    CboAnimacao.ListIndex = -1
    TxtValorAnimacao.Text = ""

    CboDJ.ListIndex = -1
    TxtValorDJ.Text = ""
    
    TxtDataFesta.Text = ""
    TxtHoraFesta.Text = ""
    TxtQtde.Text = ""
    
    TxtContrato.SetFocus
End Sub
Private Sub CarregaGridColab()

    GrdFunc.Row = 0

    GrdFunc.Col = 0
    GrdFunc.ColWidth(0) = 1

    GrdFunc.Col = 1
    GrdFunc.ColWidth(1) = 1

    GrdFunc.Col = 2
    GrdFunc.ColWidth(2) = 1500

    GrdFunc.Col = 3
    GrdFunc.ColWidth(3) = 1

    GrdFunc.Col = 4
    GrdFunc.ColWidth(4) = 1500
End Sub
Private Sub CarregaGridEscala(ByVal Rs As Recordset)

    Grid.Rows = 1
    'Grid.Clear
    x = 1
    Do Until Rs.EOF
        Grid.AddItem Rs("ID_COL").Value & Chr(9) & _
                    Rs("NOME_COL").Value & Chr(9) & _
                    Rs("ID_FUNC").Value & Chr(9) & _
                    Rs("DSC_FUNC").Value, x

        If Rs("ID_FUNC") = 3 Then
            'LblGarcom.Text = Val(LblGarcom.Text) + 1
        ElseIf Rs("ID_FUNC") = 4 Then
            'LblMonit.Text = Val(LblMonit.Text) + 1
        End If
        x = x + 1
        Rs.MoveNext
    Loop
End Sub
Private Sub FormataGridEscala()

    Grid.Rows = 1
    Grid.Cols = 4
    Grid.Row = 0

    Grid.Col = 0
    Grid.ColWidth(0) = 1

    Grid.Col = 1
    Grid.ColWidth(1) = 2000
    Grid.Text = "Funcionário"

    Grid.Col = 2
    Grid.ColWidth(2) = 1

    Grid.Col = 3
    Grid.ColWidth(3) = 1500
    Grid.Text = "Função"
End Sub

Private Sub Text1_LostFocus()

    Dim sSql As String
    Dim Rec As New ADODB.Recordset

    If Val(Text1.Text) > 0 Then
        'Pesquisar o valor do fornecedor e multiplicar pela quantidade informada
        sSql = "SELECT * FROM FORNECEDORES WHERE ID_FOR = " & CboBolo.ItemData(CboBolo.ListIndex)
    
        Rec.Open sSql, Db, adOpenDynamic, adLockOptimistic
    
        If Not Rec.EOF Then
            VAL_FOR = Rec("VALOR_FOR").Value
    
            TxtValorBolo.Text = Format(VAL_FOR * Text1.Text, ".00")
        End If
    End If
End Sub
Private Sub Text3_LostFocus()

    Dim sSql As String
    Dim Rec As New ADODB.Recordset

    If Val(Text3.Text) > 0 Then
        'Pesquisar o valor do fornecedor e multiplicar pela quantidade informada
        sSql = "SELECT * FROM FORNECEDORES WHERE ID_FOR = " & CboDoce.ItemData(CboDoce.ListIndex)
    
        Rec.Open sSql, Db, adOpenDynamic, adLockOptimistic
    
        If Not Rec.EOF Then
            VAL_FOR = Rec("VALOR_FOR").Value
    
            TxtValorDoce.Text = Format(VAL_FOR * Text1.Text, ".00")
        End If
    End If
End Sub

Private Sub TxtContrato_Change()

    SSTab1.Enabled = False
End Sub
