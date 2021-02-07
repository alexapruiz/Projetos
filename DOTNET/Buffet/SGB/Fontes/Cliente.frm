VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Cliente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SGB - Cadastro de Clientes"
   ClientHeight    =   6585
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   10650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   3804
      Left            =   24
      TabIndex        =   29
      Top             =   2184
      Width           =   10596
      _ExtentX        =   18680
      _ExtentY        =   6694
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   324
      Left            =   9264
      TabIndex        =   28
      Top             =   6120
      Width           =   1260
   End
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "Excluir"
      Height          =   324
      Left            =   6216
      TabIndex        =   27
      Top             =   6120
      Width           =   1260
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   324
      Left            =   3168
      TabIndex        =   26
      Top             =   6120
      Width           =   1260
   End
   Begin VB.CommandButton CmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sai&r"
      Height          =   324
      Left            =   120
      TabIndex        =   25
      Top             =   6120
      Width           =   1260
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados do Cliente"
      Height          =   2052
      Left            =   24
      TabIndex        =   0
      Top             =   72
      Width           =   10596
      Begin VB.TextBox TxtCEPCliente 
         Height          =   324
         Left            =   9096
         MaxLength       =   8
         TabIndex        =   12
         Top             =   912
         Width           =   1380
      End
      Begin VB.TextBox TxtTELCELCliente 
         Height          =   324
         Left            =   9096
         MaxLength       =   10
         TabIndex        =   11
         Top             =   1272
         Width           =   1380
      End
      Begin VB.TextBox TxtCPFCliente 
         Height          =   324
         Left            =   9096
         MaxLength       =   11
         TabIndex        =   10
         Top             =   576
         Width           =   1380
      End
      Begin VB.TextBox TxtNomeCliente 
         Height          =   288
         Left            =   792
         MaxLength       =   60
         TabIndex        =   9
         Top             =   576
         Width           =   6276
      End
      Begin VB.TextBox TxtEndCliente 
         Height          =   288
         Left            =   792
         MaxLength       =   50
         TabIndex        =   8
         Top             =   912
         Width           =   6276
      End
      Begin VB.TextBox TxtNumCliente 
         Height          =   324
         Left            =   7416
         MaxLength       =   5
         TabIndex        =   7
         Top             =   912
         Width           =   732
      End
      Begin VB.TextBox TxtTELCOMCliente 
         Height          =   324
         Left            =   792
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1272
         Width           =   1428
      End
      Begin VB.TextBox TxtTELRESCliente 
         Height          =   324
         Left            =   4536
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1272
         Width           =   1428
      End
      Begin VB.TextBox TxtBairroCliente 
         Height          =   324
         Left            =   792
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1632
         Width           =   2676
      End
      Begin VB.TextBox TxtCidadeCliente 
         Height          =   324
         Left            =   4536
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1632
         Width           =   2532
      End
      Begin VB.TextBox TxtRGCliente 
         Height          =   324
         Left            =   9096
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1632
         Width           =   1380
      End
      Begin VB.TextBox TxtCodCliente 
         Height          =   288
         Left            =   792
         MaxLength       =   4
         TabIndex        =   1
         Top             =   240
         Width           =   924
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "End."
         Height          =   192
         Left            =   396
         TabIndex        =   24
         Top             =   984
         Width           =   324
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "CPF"
         Height          =   192
         Left            =   8688
         TabIndex        =   23
         Top             =   624
         Width           =   312
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Telefone (cel.)"
         Height          =   192
         Left            =   7968
         TabIndex        =   22
         Top             =   1344
         Width           =   1032
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Telefone (res.)"
         Height          =   192
         Left            =   3432
         TabIndex        =   21
         Top             =   1320
         Width           =   1044
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Telefone"
         Height          =   192
         Left            =   72
         TabIndex        =   20
         Top             =   1320
         Width           =   648
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   192
         Left            =   276
         TabIndex        =   19
         Top             =   624
         Width           =   444
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Bairro"
         Height          =   192
         Left            =   288
         TabIndex        =   18
         Top             =   1680
         Width           =   432
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "RG"
         Height          =   192
         Left            =   8760
         TabIndex        =   17
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "nº"
         Height          =   192
         Left            =   7176
         TabIndex        =   16
         Top             =   984
         Width           =   144
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "CEP"
         Height          =   192
         Left            =   8676
         TabIndex        =   15
         Top             =   984
         Width           =   324
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
         Height          =   192
         Left            =   3960
         TabIndex        =   14
         Top             =   1680
         Width           =   528
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   192
         Left            =   144
         TabIndex        =   13
         Top             =   264
         Width           =   528
      End
   End
End
Attribute VB_Name = "Cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdSair_Click()

    Unload Me
End Sub

Private Sub Form_Load()

    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
    Call DefineGrid
    Call CarregGrid
End Sub
Private Sub DefineGrid()

    Grid.Cols = 12

    Grid.Row = 0

    'Codigo do Cliente
    Grid.Col = 0
    Grid.ColWidth(0) = 1
    Grid.Text = ""

    'Nome do cliente
    Grid.Col = 1
    Grid.ColWidth(1) = 5000
    Grid.Text = "Nome"
    
    'Tel res
    Grid.Col = 2
    Grid.ColWidth(2) = 1500
    Grid.Text = "Tel. Res."
    
    'tel com
    Grid.Col = 3
    Grid.ColWidth(3) = 1500
    Grid.Text = "Tel. Com."
    
    'tel cel
    Grid.Col = 4
    Grid.ColWidth(4) = 1500
    Grid.Text = "Tel. Cel."

    'Demais campos sem exibição
    Grid.ColWidth(5) = 1
    Grid.ColWidth(6) = 1
    Grid.ColWidth(7) = 1
    Grid.ColWidth(8) = 1
    Grid.ColWidth(9) = 1
    Grid.ColWidth(10) = 1
    Grid.ColWidth(11) = 1
End Sub
Private Sub CarregGrid()

    Dim Rs As New ADODB.Recordset
    Dim sSql As String

    sSql = "select * from CLIENTES order by NOM_CLI, COD_CLI "

    Rs.Open sSql, Db, adOpenDynamic, 1

    Grid.Rows = 1
    x = 1
    If Not Rs.EOF Then
        Do Until Rs.EOF
            Grid.Rows = Grid.Rows + 1
            Grid.Row = x

            'Codigo do Cliente
            Grid.Col = 0
            Grid.Text = Rs("COD_CLI")

            'Nome do Cliente
            Grid.Col = 1
            Grid.Text = Rs("NOM_CLI") & ""

            'Telefone Residencial
            Grid.Col = 2
            Grid.Text = Rs("TEL1_CLI") & ""

            'Telefone Comercial
            Grid.Col = 3
            Grid.Text = Rs("TEL2_CLI") & ""

            'Telefone Celular
            Grid.Col = 4
            Grid.Text = Rs("TEL3_CLI") & ""

            'CPF
            Grid.Col = 5
            Grid.Text = Rs("CPF_CLI") & ""

            'Endereco
            Grid.Col = 6
            Grid.Text = Rs("END_CLI") & ""

            'Numero
            Grid.Col = 7
            Grid.Text = Rs("NUM_CLI") & ""

            'CEP
            Grid.Col = 8
            Grid.Text = Rs("CEP_CLI") & ""

            'Bairro
            Grid.Col = 9
            Grid.Text = Rs("BAI_CLI") & ""

            'Cidade
            Grid.Col = 10
            Grid.Text = Rs("CID_CLI") & ""

            'RG
            Grid.Col = 11
            Grid.Text = Rs("RG_CLI") & ""

            x = x + 1
            Rs.MoveNext
        Loop
    End If
End Sub
Private Sub Grid_Click()

    If Grid.Row > 0 Then
        'Codigo
        Grid.Col = 0
        TxtCodCliente.Text = Grid.Text
    
        'Nome
        Grid.Col = 1
        TxtNomeCliente.Text = Grid.Text
    
        'Tel res
        Grid.Col = 2
        TxtTELRESCliente.Text = Grid.Text
    
        'tel com
        Grid.Col = 3
        TxtTELCOMCliente.Text = Grid.Text
    
        'tel cel
        Grid.Col = 4
        TxtTELCELCliente.Text = Grid.Text
    
        'cpf
        Grid.Col = 5
        TxtCPFCliente.Text = Grid.Text
    
        'end
        Grid.Col = 6
        TxtEndCliente.Text = Grid.Text
    
        'numero
        Grid.Col = 7
        TxtNumCliente.Text = Grid.Text
    
        'cep
        Grid.Col = 8
        TxtCEPCliente.Text = Grid.Text
    
        'bairro
        Grid.Col = 9
        TxtBairroCliente.Text = Grid.Text
    
        'cidade
        Grid.Col = 10
        TxtCidadeCliente.Text = Grid.Text
    
        'rg
        Grid.Col = 11
        TxtRGCliente.Text = Grid.Text
    End If
End Sub
