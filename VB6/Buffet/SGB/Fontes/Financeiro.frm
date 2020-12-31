VERSION 5.00
Begin VB.Form Financeiro 
   Caption         =   "SGB - Módulo Financeiro"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4995
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Situação Financeira no Período"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   75
      TabIndex        =   36
      Top             =   3555
      Width           =   8955
      Begin VB.TextBox TxtSaldo 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5452
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   300
         Width           =   1635
      End
      Begin VB.TextBox TxtResultado 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2872
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   285
         Width           =   1635
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Saldo"
         Height          =   195
         Left            =   4957
         TabIndex        =   39
         Top             =   360
         Width           =   405
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Resultado"
         Height          =   195
         Left            =   1987
         TabIndex        =   37
         Top             =   345
         Width           =   720
      End
   End
   Begin VB.CommandButton CmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sai&r"
      Height          =   480
      Left            =   3772
      TabIndex        =   35
      Top             =   4410
      Width           =   1530
   End
   Begin VB.Frame Frame2 
      Caption         =   "Contas a Receber"
      ForeColor       =   &H00FF0000&
      Height          =   2760
      Left            =   4590
      TabIndex        =   13
      Top             =   720
      Width           =   4425
      Begin VB.CommandButton Command1 
         Caption         =   "+"
         Height          =   315
         Left            =   3285
         TabIndex        =   34
         Top             =   330
         Width           =   300
      End
      Begin VB.TextBox TxtDinheiroCaixa 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "0,00"
         Top             =   345
         Width           =   1425
      End
      Begin VB.TextBox TxtTOTALReceitas 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1785
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   2295
         Width           =   1425
      End
      Begin VB.CommandButton CmdOutrasReceitas 
         Caption         =   "+"
         Height          =   315
         Left            =   3285
         TabIndex        =   27
         Top             =   1350
         Width           =   300
      End
      Begin VB.CommandButton CmdDepositos 
         Caption         =   "+"
         Height          =   315
         Left            =   3285
         TabIndex        =   26
         Top             =   1020
         Width           =   300
      End
      Begin VB.CommandButton CmdChequesReceber 
         Caption         =   "+"
         Height          =   315
         Left            =   3285
         TabIndex        =   25
         Top             =   675
         Width           =   300
      End
      Begin VB.TextBox TxtChequesDepositar 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   690
         Width           =   1425
      End
      Begin VB.TextBox TxtDepositos 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "0,00"
         Top             =   1035
         Width           =   1425
      End
      Begin VB.TextBox TxtOutrasReceitas 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "0,00"
         Top             =   1365
         Width           =   1425
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Dinheiro em Caixa"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   390
         TabIndex        =   33
         Top             =   390
         Width           =   1275
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1050
         TabIndex        =   30
         Top             =   2340
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Outras Receitas"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   525
         TabIndex        =   19
         Top             =   1395
         Width           =   1140
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Depósitos"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   960
         TabIndex        =   18
         Top             =   1080
         Width           =   705
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cheques"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1035
         TabIndex        =   17
         Top             =   735
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contas a Pagar"
      ForeColor       =   &H000000FF&
      Height          =   2760
      Left            =   75
      TabIndex        =   2
      Top             =   720
      Width           =   4425
      Begin VB.TextBox TxtTOTALDespesas 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   2295
         Width           =   1425
      End
      Begin VB.CommandButton CmdDespesasEventuais 
         Caption         =   "+"
         Height          =   315
         Left            =   3285
         TabIndex        =   24
         Top             =   1770
         Width           =   300
      End
      Begin VB.CommandButton CmdDespesasFixas 
         Caption         =   "+"
         Height          =   315
         Left            =   3285
         TabIndex        =   23
         Top             =   1428
         Width           =   300
      End
      Begin VB.CommandButton CmdChequesPagar 
         Caption         =   "+"
         Height          =   315
         Left            =   3285
         TabIndex        =   22
         Top             =   1087
         Width           =   300
      End
      Begin VB.CommandButton CmdFuncionarios 
         Caption         =   "+"
         Height          =   315
         Left            =   3285
         TabIndex        =   21
         Top             =   746
         Width           =   300
      End
      Begin VB.CommandButton CmdFornecedores 
         Caption         =   "+"
         Height          =   315
         Left            =   3285
         TabIndex        =   20
         Top             =   405
         Width           =   300
      End
      Begin VB.TextBox TxtDespesasEventuais 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0,00"
         Top             =   1770
         Width           =   1425
      End
      Begin VB.TextBox TxtDespesasFixas 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1428
         Width           =   1425
      End
      Begin VB.TextBox TxtCheques 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0,00"
         Top             =   1087
         Width           =   1425
      End
      Begin VB.TextBox TxtFuncionarios 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   746
         Width           =   1425
      End
      Begin VB.TextBox TxtFornecedores 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   405
         Width           =   1425
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1005
         TabIndex        =   28
         Top             =   2340
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Desp. Eventuais"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   375
         TabIndex        =   7
         Top             =   1830
         Width           =   1170
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Despesas Fixas"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   375
         TabIndex        =   6
         Top             =   1470
         Width           =   1125
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cheques"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   375
         TabIndex        =   5
         Top             =   1125
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Funcionários"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   375
         TabIndex        =   4
         Top             =   780
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedores"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   375
         TabIndex        =   3
         Top             =   420
         Width           =   975
      End
   End
   Begin VB.ComboBox CboFechamento 
      Height          =   315
      Left            =   3397
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   210
      Width           =   3165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fechamento :"
      Height          =   195
      Left            =   2272
      TabIndex        =   0
      Top             =   270
      Width           =   975
   End
End
Attribute VB_Name = "Financeiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DATA_1 As String
Public DATA_2 As String

Private Sub CboFechamento_Click()

    Call CalculaValores
End Sub

Private Sub CmdChequesReceber_Click()

    If CboFechamento.ListIndex <> -1 Then
        If CDbl(TxtChequesDepositar.Text) <= 0 Then
            MsgBox "Não há Lançamentos para esta conta", vbExclamation, "SGB"
            Exit Sub
        End If

        Load Pagamento
        Pagamento.Tag = "Cheques"
        Pagamento.DATA_1 = DATA_1
        Pagamento.DATA_2 = DATA_2
        Pagamento.Show 1
        Call CalculaValores
    Else
        MsgBox "Selecione um período de Fechamento", vbExclamation, "SGB"
    End If
End Sub

Private Sub CmdDespesasFixas_Click()

    If CboFechamento.ListIndex <> -1 Then
        If Val(TxtDespesasFixas.Text) <= 0 Then
            MsgBox "Não há Lançamentos para esta conta", vbExclamation, "SGB"
            Exit Sub
        End If

        Load Pagamento
        Pagamento.Tag = "Despesas"
        Pagamento.DATA_1 = Left(CboFechamento.Text, 10)
        Pagamento.DATA_2 = Right(CboFechamento.Text, 10)
        Pagamento.ID_FECHA = CboFechamento.ItemData(CboFechamento.ListIndex)
        Pagamento.Show 1
        Call CalculaValores
    Else
        MsgBox "Selecione um período de Fechamento", vbExclamation, "SGB"
    End If
End Sub

Private Sub CmdFornecedores_Click()

    If CboFechamento.ListIndex <> -1 Then
        If Val(TxtFornecedores.Text) <= 0 Then
            MsgBox "Não há Lançamentos para esta conta", vbExclamation, "SGB"
            Exit Sub
        End If

        Load Pagamento
        Pagamento.Tag = "Fornecedor"
        Pagamento.DATA_1 = Left(CboFechamento.Text, 10)
        Pagamento.DATA_2 = Right(CboFechamento.Text, 10)
        Pagamento.Show 1
        Call CalculaValores
    Else
        MsgBox "Selecione um período de Fechamento", vbExclamation, "SGB"
    End If
End Sub

Private Sub CmdFuncionarios_Click()

    If CboFechamento.ListIndex <> -1 Then
        If Val(TxtFuncionarios.Text) <= 0 Then
            MsgBox "Não há Lançamentos para esta conta", vbExclamation, "SGB"
            Exit Sub
        End If

        Load Pagamento
        Pagamento.Tag = "Funcionario"
        Pagamento.DATA_1 = Left(CboFechamento.Text, 10)
        Pagamento.DATA_2 = Right(CboFechamento.Text, 10)
        Pagamento.Show 1
        Call CalculaValores
    Else
        MsgBox "Selecione um período de Fechamento", vbExclamation, "SGB"
    End If
End Sub
Private Sub CmdSair_Click()

    Unload Me
End Sub
Private Sub Form_Load()

    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

    Call CarregaComboFechamento
    
    CboFechamento.ListIndex = CboFechamento.ListCount - 1
End Sub
Private Sub CarregaComboFechamento()

    Dim x As Integer
    Dim Rs As New ADODB.Recordset

    sSql = "select * from FECHAMENTO"

    Rs.Open sSql, Db, adOpenDynamic, adLockReadOnly

    Do Until Rs.EOF
        CboFechamento.AddItem Format(Rs(1).Value, "dd/mm/yyyy") & " - " & Format(Rs(2).Value, "dd/mm/yyyy")
        CboFechamento.ItemData(CboFechamento.NewIndex) = Rs(0).Value
        Rs.MoveNext
    Loop
End Sub
Private Sub CalculaValores()

    Dim Rs As New ADODB.Recordset
    Dim sSql As String

    DATA_1 = Left(CboFechamento.List(CboFechamento.ListIndex), 10)
    DATA_2 = Right(CboFechamento.List(CboFechamento.ListIndex), 10)
    
    DATA_1 = Format(DATA_1, "mm-dd-yyyy")
    DATA_2 = Format(DATA_2, "mm-dd-yyyy")

    '**********************************************************************************************************
    'Fornecedores
    '**********************************************************************************************************
    sSql = "SELECT  SUM(VALOR) AS TOTAL_FORNECEDORES "
    sSql = sSql & " FROM CONTAS_A_PAGAR C , FORNECEDORES F "
    sSql = sSql & " WHERE ID_CNT IN (SELECT ID_CNT FROM CONTRATOS WHERE DATA_FESTA BETWEEN #"
    sSql = sSql & DATA_1 & "# AND #" & DATA_2 & "#) "
    sSql = sSql & " AND C.ID_FOR = F.ID_FOR"
    sSql = sSql & " AND C.PAGO = 'N'"

    Rs.Open sSql, Db, adOpenDynamic, adLockReadOnly

    If Not Rs.EOF Then
        If IsNull(Rs("TOTAL_FORNECEDORES").Value) Then
            TxtFornecedores.Text = 0
        Else
            TxtFornecedores.Text = Format(Rs("TOTAL_FORNECEDORES").Value & "", ".00")
        End If
    End If
    Rs.Close

    '**********************************************************************************************************
    'Funcionarios
    '**********************************************************************************************************
    sSql = "SELECT E.ID_COL , C.NOME_COL , COUNT(0) AS QTDE  , VAL_COL * QTDE  AS TOTAL_FUNC "
    sSql = sSql & " FROM ESCALA E , COLABORADORES C "
    sSql = sSql & " WHERE ID_CNT IN (SELECT ID_CNT FROM CONTRATOS WHERE DATA_FESTA BETWEEN #"
    sSql = sSql & DATA_1 & "# AND #" & DATA_2 & "#)"
    sSql = sSql & " AND E.ID_COL = C.ID_COL"
    sSql = sSql & " AND E.PAGO = 'N'"
    sSql = sSql & " GROUP BY E.ID_COL , C.VAL_COL , C.NOME_COL "

    Rs.Open sSql, Db, adOpenDynamic, adLockReadOnly

    TxtFuncionarios.Text = "0,00"
    If Not Rs.EOF Then
        Do Until Rs.EOF
            If IsNull(Rs("TOTAL_FUNC").Value) Then
                TxtFuncionarios.Text = "0,00"
            Else
                TxtFuncionarios.Text = Val(TxtFuncionarios.Text) + Val(Format(Rs("TOTAL_FUNC").Value & "", "0.00"))
            End If
            Rs.MoveNext
        Loop
    End If
    TxtFuncionarios.Text = Format(TxtFuncionarios.Text, "0.00")
    Rs.Close

    '**********************************************************************************************************
    'Cheques a Receber
    '**********************************************************************************************************
    sSql = "SELECT Sum(VALOR_PAR) AS CHEQUES_RECEBER "
    sSql = sSql & " FROM PARCELA_CONTRATO "
    sSql = sSql & " WHERE (((PARCELA_CONTRATO.DATA_PAR) Between #" & Format(DATA_1, "dd/mm/yyyy") & "# And #" & Format(DATA_2, "dd/mm/yyyy") & "#)) "
    sSql = sSql & " AND COMPENSADO = 'N'"

    Rs.Open sSql, Db, adOpenDynamic, adLockReadOnly

    If Not Rs.EOF Then
        If IsNull(Rs("CHEQUES_RECEBER").Value) Then
            TxtChequesDepositar.Text = "0,00"
        Else
            TxtChequesDepositar.Text = Format(Rs("CHEQUES_RECEBER").Value & "", "0.00")
        End If
    End If
    Rs.Close

    '**********************************************************************************************************
    'Despesas Fixas
    '**********************************************************************************************************
    sSql = "select * from FECHAMENTO WHERE ID_FECHA = " & CboFechamento.ItemData(CboFechamento.ListIndex)

    Rs.Open sSql, Db, adOpenDynamic, adLockReadOnly
    
    If Not Rs.EOF Then
        If Rs("INC_DESP") = "1" Then
            TxtDespesasFixas.Enabled = True
            CmdDespesasFixas.Enabled = True

            Rs.Close
            sSql = "SELECT SUM(VALOR_DESP) AS VALOR_DESPESAS FROM DESPESAS_FIXAS "
            sSql = sSql & " WHERE ID_DESP NOT IN (SELECT ID_DESP FROM PAGTO_DESP_FIXA WHERE ID_FECHA = " & CboFechamento.ItemData(CboFechamento.ListIndex) & ")"

            Rs.Open sSql, Db, adOpenDynamic, adLockReadOnly

            If Not Rs.EOF Then
                TxtDespesasFixas.Text = Format(Rs("VALOR_DESPESAS").Value, "0.00")
            End If
        Else
            TxtDespesasFixas.Text = "0,00"
            TxtDespesasFixas.Enabled = False
            CmdDespesasFixas.Enabled = False
        End If
    End If
    Rs.Close

    '**********************************************************************************************************
    'Calcula o Total de Despesas
    '**********************************************************************************************************
    TxtTotalDespesas.Text = Format(CCur(TxtFornecedores.Text) + CCur(TxtFuncionarios.Text) + CCur(TxtCheques.Text) + _
                            CCur(TxtDespesasFixas.Text) + CCur(TxtDespesasEventuais.Text), "0.00")

    'Calcula o Total de Despesas
    TxtTotalReceitas.Text = Format(CCur(TxtDinheiroCaixa.Text) + CCur(TxtChequesDepositar.Text) + _
                            CCur(TxtDepositos.Text) + CCur(TxtOutrasReceitas.Text), "0.00")

    'APURA O RESULTADO DO PERÍODO
    If CCur(TxtTotalReceitas.Text) > CCur(TxtTotalDespesas.Text) Then
        TxtResultado.Text = "LUCRO"
        TxtResultado.ForeColor = "&H00FF0000"
    ElseIf CCur(TxtTotalReceitas.Text) < CCur(TxtTotalDespesas.Text) Then
        TxtResultado.Text = "PREJUÍZO"
        TxtResultado.ForeColor = "&H000000FF"
    Else
        TxtResultado.Text = "EMPATE"
        TxtResultado.ForeColor = "&H80000012"
    End If

    TxtSaldo.Text = Format(Val(TxtTotalReceitas.Text) - Val(TxtTotalDespesas.Text), "0.00")
    If TxtResultado.Text = "LUCRO" Then
        TxtSaldo.ForeColor = "&H00FF0000"
    ElseIf TxtResultado.Text = "PREJUÍZO" Then
        TxtSaldo.ForeColor = "&H000000FF"
    Else
        TxtSaldo.ForeColor = "&H80000012"
    End If
End Sub
