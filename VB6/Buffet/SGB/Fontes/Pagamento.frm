VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Pagamento 
   Caption         =   "SGB - Módulo de Pagamentos"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5415
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sai&r"
      Height          =   375
      Left            =   3157
      TabIndex        =   7
      Top             =   4890
      Width           =   1800
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4725
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8334
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Fornecedores"
      TabPicture(0)   =   "Pagamento.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "g"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CmdTotal"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CmdRegistrarPagto"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "TxtTotal"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Funcionários"
      TabPicture(1)   =   "Pagamento.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TxtTotalFunc"
      Tab(1).Control(1)=   "CmdRegPagtoFunc"
      Tab(1).Control(2)=   "g2"
      Tab(1).Control(3)=   "Label1"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Cheques"
      TabPicture(2)   =   "Pagamento.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "CmdRegComp"
      Tab(2).Control(1)=   "TxtTotalCheques"
      Tab(2).Control(2)=   "g3"
      Tab(2).Control(3)=   "Label2"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Despesas Fixas"
      TabPicture(3)   =   "Pagamento.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "TxtTOTALDespesasFixas"
      Tab(3).Control(1)=   "CmdCompensadoDespesas"
      Tab(3).Control(2)=   "g4"
      Tab(3).Control(3)=   "Label3"
      Tab(3).ControlCount=   4
      Begin VB.TextBox TxtTOTALDespesasFixas 
         Height          =   285
         Left            =   -70200
         TabIndex        =   17
         Top             =   3900
         Width           =   1125
      End
      Begin VB.CommandButton CmdCompensadoDespesas 
         Caption         =   "Registrar Pagamento"
         Height          =   585
         Left            =   -68550
         TabIndex        =   16
         Top             =   735
         Width           =   1170
      End
      Begin VB.CommandButton CmdRegComp 
         Caption         =   "Compensado"
         Height          =   585
         Left            =   -68550
         TabIndex        =   15
         Top             =   735
         Width           =   1170
      End
      Begin VB.TextBox TxtTotalCheques 
         Height          =   285
         Left            =   -70200
         TabIndex        =   13
         Top             =   3900
         Width           =   1125
      End
      Begin VB.TextBox TxtTotalFunc 
         Height          =   285
         Left            =   -70200
         TabIndex        =   10
         Top             =   3900
         Width           =   1125
      End
      Begin VB.CommandButton CmdRegPagtoFunc 
         Caption         =   "Registrar Pagamento"
         Height          =   585
         Left            =   -68550
         TabIndex        =   9
         Top             =   690
         Width           =   1170
      End
      Begin VB.TextBox TxtTotal 
         Height          =   285
         Left            =   4800
         TabIndex        =   6
         Top             =   3900
         Width           =   1125
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cancelar Pagamento"
         Enabled         =   0   'False
         Height          =   585
         Left            =   6450
         TabIndex        =   4
         Top             =   3195
         Width           =   1170
      End
      Begin VB.CommandButton CmdRegistrarPagto 
         Caption         =   "Registrar Pagamento"
         Enabled         =   0   'False
         Height          =   585
         Left            =   6450
         TabIndex        =   3
         Top             =   1935
         Width           =   1170
      End
      Begin VB.CommandButton CmdTotal 
         Caption         =   "Registrar Pagamento"
         Height          =   585
         Left            =   6450
         TabIndex        =   2
         Top             =   690
         Width           =   1170
      End
      Begin MSFlexGridLib.MSFlexGrid g 
         Height          =   3105
         Left            =   420
         TabIndex        =   1
         Top             =   690
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   5477
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid g2 
         Height          =   3105
         Left            =   -74580
         TabIndex        =   8
         Top             =   690
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   5477
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid g3 
         Height          =   3105
         Left            =   -74580
         TabIndex        =   12
         Top             =   690
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   5477
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid g4 
         Height          =   3105
         Left            =   -74580
         TabIndex        =   18
         Top             =   690
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   5477
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
         SelectionMode   =   1
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL"
         Height          =   195
         Left            =   -70845
         TabIndex        =   19
         Top             =   3960
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL"
         Height          =   195
         Left            =   -70845
         TabIndex        =   14
         Top             =   3960
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL"
         Height          =   195
         Left            =   -70845
         TabIndex        =   11
         Top             =   3960
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL"
         Height          =   195
         Left            =   4155
         TabIndex        =   5
         Top             =   3960
         Width           =   525
      End
   End
End
Attribute VB_Name = "Pagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ID_FECHA As String
Public DATA_1 As String
Public DATA_2 As String
Private Sub SelecionaFornecedores()

    Dim Rs As New ADODB.Recordset
    Dim sSql As String

    g.Rows = 1
    g.Clear

    'Fornecedores
    sSql = "SELECT F.ID_FOR ,  F.NOME_FOR , SUM(C.VALOR) AS TOTAL_DEVIDO"
    sSql = sSql & " FROM CONTAS_A_PAGAR C , FORNECEDORES F"
    sSql = sSql & " WHERE ID_CNT IN ("
    sSql = sSql & " SELECT ID_CNT FROM CONTRATOS WHERE DATA_FESTA BETWEEN #" & DATA_1 & "# AND #" & DATA_2 & "#)"
    sSql = sSql & " AND C.ID_FOR = F.ID_FOR"
    sSql = sSql & " AND C.PAGO = 'N'"
    sSql = sSql & " GROUP BY F.ID_FOR , F.NOME_FOR"
    sSql = sSql & " ORDER BY F.ID_FOR"

    Rs.Open sSql, Db, adOpenDynamic, 1

    If Not Rs.EOF Then
        Do Until Rs.EOF
            g.AddItem Rs("ID_FOR") & Chr(9) & Rs("NOME_FOR") & Chr(9) & Format(Val(Rs("TOTAL_DEVIDO").Value), "0.00")
            Rs.MoveNext
        Loop
    End If
    Rs.Close
End Sub
Private Sub CmdCompensadoDespesas_Click()

    Dim Despesas As New clsDespesas

    If g4.Rows <= 1 Then
        MsgBox "Nenhum item selecionado.", vbExclamation, "SGB"
        Exit Sub
    End If

    g4.Col = 0
    ID_DESP = g4.Text
    ID_FECHA = Financeiro.CboFechamento.ItemData(Financeiro.CboFechamento.ListIndex)

    '1. Atualizar os cheques compensados
    Call Despesas.RegistrarDespesaPaga(ID_FECHA, ID_DESP, Db)

    '2. Refazer o grid de Cheques
    Call SelecionaDespesasFixas

    '3. Atualizar o campo TOTAL
    Call SomaGridDespesasFixas

    Call FormataGrid
End Sub
Private Sub CmdRegComp_Click()

    Dim Contrato As New ClsContrato

    If (g3.Rows <= 1) Or (g3.Row < 1) Then
        MsgBox "Nenhum item selecionado.", vbExclamation, "SGB"
        Exit Sub
    End If

    g3.Col = 0
    ID_CNT = g3.Text
    g3.Col = 1
    ID_PAR = g3.Text

    '1. Atualizar os cheques compensados
    Call Contrato.RegistrarChequeCompensado(ID_CNT, ID_PAR, Db)

    '2. Refazer o grid de Cheques
    Call SelecionaCheques

    '3. Atualizar o campo TOTAL
    Call SomaGridCheques

    Call FormataGrid
End Sub
Private Sub CmdRegPagtoFunc_Click()

    Dim Contas As New ClsContasaPagar

    If (g2.Rows <= 1) Or (g2.Row = 0) Then
        MsgBox "Nenhum item selecionado.", vbExclamation, "SGB"
        Exit Sub
    End If

    g2.Col = 0

    '1. Atualizar as contas como pagas (CONTAS_A_PAGAR)
    Call Contas.RegistrarPagtoFuncionarios(DATA_1, DATA_2, g2.Text, Db)

    '2. Inserir os registros na tabela PAGAMENTO

    '3. Refazer o grid de Fornecedores
    Call SelecionaFuncionarios
    
    '4. Atualizar o campo TOTAL
    Call SomaGridFunc
    
    Call FormataGrid
End Sub
Private Sub CmdSair_Click()

    Unload Me
End Sub
Private Sub CmdTotal_Click()

    Dim Contas As New ClsContasaPagar

    If g.Rows <= 1 Then
        MsgBox "Nenhum item selecionado.", vbExclamation, "SGB"
        Exit Sub
    End If

    g.Col = 0

    '1. Atualizar as contas como pagas (CONTAS_A_PAGAR)
    Call Contas.RegistrarPagtoFornecedores(DATA_1, DATA_2, g.Text, Db)

    '2. Inserir os registros na tabela PAGAMENTO

    '3. Refazer o grid de Fornecedores
    Call SelecionaFornecedores
    Call FormataGrid

    '4. Atualizar o campo TOTAL
    Call SomaGridFornec

End Sub
Private Sub Form_Activate()

    Call FormataGrid

    If Me.Tag = "Fornecedor" Then
        SSTab1.Tab = 0
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(3) = False

        Call SelecionaFornecedores
        Call SomaGridFornec
    ElseIf Me.Tag = "Funcionario" Then
        SSTab1.Tab = 1
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(3) = False

        Call SelecionaFuncionarios
        Call SomaGridFunc
    ElseIf Me.Tag = "Cheques" Then
        SSTab1.Tab = 2
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = True
        SSTab1.TabEnabled(3) = False

        Call SelecionaCheques
        Call SomaGridCheques
    Else
        SSTab1.Tab = 3
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(3) = True

        Call SelecionaDespesasFixas
        Call SomaGridDespesasFixas
    End If
    Call FormataGrid
End Sub
Private Sub FormataGrid()

    Select Case Me.Tag
    Case "Fornecedor"
        '*********************************************************************************************************
        'Fornecedores
        '*********************************************************************************************************
        g.Cols = 3
        g.Row = 0
        g.Col = 1
        g.Text = "Fornecedor"
        g.Col = 2
        g.Text = "Valor Devido"
        g.ColWidth(0) = 1
        g.ColWidth(1) = 2500
        g.ColWidth(2) = 1000

    Case "Funcionario"
        '*********************************************************************************************************
        'Funcionários
        '*********************************************************************************************************
        g2.Cols = 3
        g2.Row = 0
        g2.Col = 1
        g2.Text = "Funcionário"
        g2.Col = 2
        g2.Text = "Valor Devido"
        g2.ColWidth(0) = 1
        g2.ColWidth(1) = 2500
        g2.ColWidth(2) = 1000

    Case "Cheques"
        '*********************************************************************************************************
        'Cheques
        '*********************************************************************************************************
        'Nomes das colunas
        g3.Cols = 5
        g3.Row = 0
        g3.Col = 0
        g3.Text = "Contrato"
        g3.Col = 1
        g3.Text = "Parcela"
        g3.Col = 2
        g3.Text = "Num. Docto"
        g3.Col = 3
        g3.Text = "Data Deposito"
        g3.Col = 4
        g3.Text = "Valor"
    
        'Alinhamento
        g3.ColAlignment(0) = 3
        g3.ColAlignment(1) = 3
        g3.ColAlignment(2) = 3
        g3.ColAlignment(3) = 3
        'g3.ColAlignment(4) = 1
    
        'Tamanho das colunas
        g3.ColWidth(0) = 900
        g3.ColWidth(1) = 800
        g3.ColWidth(2) = 1250
        g3.ColWidth(3) = 1200
        g3.ColWidth(4) = 1288

    Case "Despesas"
        '*********************************************************************************************************
        'Despesas Fixas
        '*********************************************************************************************************
        'Nomes das colunas
        g4.Cols = 3
        g4.Row = 0
        g4.Col = 0
        g4.Text = "Código"
        g4.Col = 1
        g4.Text = "Descrição"
        g4.Col = 2
        g4.Text = "Valor"
    
        'Tamanho das colunas
        g4.ColWidth(0) = 600
        g4.ColWidth(1) = 1900
        g4.ColWidth(2) = 1400

    End Select
End Sub
Private Sub SomaGridFornec()

    Dim x As Integer

    TxtTotal.Text = ""
    For x = 1 To g.Rows - 1
        g.Row = x
        g.Col = 0
        If Val(g.Text) <> 0 Then
            g.Col = 2
            TxtTotal.Text = Format(Val(TxtTotal.Text) + Val(g.Text), "0.00")
        End If
    Next x
End Sub
Private Sub SelecionaFuncionarios()

    Dim Rs As New ADODB.Recordset
    Dim sSql As String

    g2.Rows = 1
    g2.Clear

    'Funcionarios
    'sSql = " SELECT  C.ID_COL , C.NOME_COL , COUNT(0) * 22 AS TOTAL_DEVIDO "
    'sSql = sSql & " FROM ESCALA E , COLABORADORES C "
    'sSql = sSql & " WHERE ID_CNT IN (SELECT ID_CNT FROM CONTRATOS WHERE DATA_FESTA BETWEEN #" & DATA_1 & "# AND #" & DATA_2 & "#) "
    'sSql = sSql & " AND E.ID_COL = C.ID_COL"
    'sSql = sSql & " AND E.PAGO = 'N'"
    'sSql = sSql & " GROUP BY C.ID_COL , C.NOME_COL "
    'sSql = sSql & " ORDER BY  C.NOME_COL "

    sSql = "SELECT E.ID_COL , C.NOME_COL , COUNT(0) AS QTDE  , VAL_COL * QTDE  AS TOTAL_DEVIDO "
    sSql = sSql & " FROM ESCALA E , COLABORADORES C "
    sSql = sSql & " WHERE ID_CNT IN (SELECT ID_CNT FROM CONTRATOS WHERE DATA_FESTA BETWEEN #"
    sSql = sSql & DATA_1 & "# AND #" & DATA_2 & "#)"
    sSql = sSql & " AND E.ID_COL = C.ID_COL"
    sSql = sSql & " AND E.PAGO = 'N'"
    sSql = sSql & " GROUP BY E.ID_COL , C.VAL_COL , C.NOME_COL "

    Rs.Open sSql, Db, adOpenDynamic, 1

    If Not Rs.EOF Then
        Do Until Rs.EOF
            g2.AddItem Rs("ID_COL") & Chr(9) & Rs("NOME_COL") & Chr(9) & Format(Val(Rs("TOTAL_DEVIDO").Value), "0.00")
            Rs.MoveNext
        Loop
    End If
    Rs.Close
End Sub
Private Sub SomaGridFunc()

    Dim x As Integer

    TxtTotalFunc.Text = ""
    For x = 1 To g2.Rows - 1
        g2.Row = x
        g2.Col = 0
        If Val(g2.Text) <> 0 Then
            g2.Col = 2
            TxtTotalFunc.Text = Format(Val(TxtTotalFunc.Text) + Val(g2.Text), "0.00")
        End If
    Next x
End Sub
Private Sub Form_Load()

    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub
Private Sub SelecionaCheques()

    Dim Rs As New ADODB.Recordset
    Dim sSql As String

    g3.Rows = 1
    g3.Clear

    sSql = "SELECT * FROM PARCELA_CONTRATO "
    sSql = sSql & " WHERE DATA_PAR BETWEEN #" & DATA_1 & "# AND #" & DATA_2 & "#"
    sSql = sSql & " AND COMPENSADO = 'N'"
    sSql = sSql & " ORDER BY ID_CNT , ID_PAR , NUM_DOC_PAR"

    Rs.Open sSql, Db, adOpenDynamic, 1

    If Not Rs.EOF Then
        Do Until Rs.EOF
            g3.AddItem Rs("ID_CNT") & Chr(9) & Rs("ID_PAR") & Chr(9) & Rs("NUM_DOC_PAR") & Chr(9) & Format(Rs("DATA_PAR"), "dd/mm/yyyy") & Chr(9) & Format(Rs("VALOR_PAR").Value, "0.00")
            Rs.MoveNext
        Loop
    End If
    Rs.Close
End Sub
Private Sub SomaGridCheques()

    Dim x As Integer

    TxtTotalCheques.Text = 0
    For x = 1 To g3.Rows - 1
        g3.Row = x
        g3.Col = 0
        If Val(g3.Text) <> 0 Then
            g3.Col = 4
            TxtTotalCheques.Text = Format(CCur(TxtTotalCheques.Text) + CCur(g3.Text), "0.00")
        End If
    Next x
End Sub
Private Sub SelecionaDespesasFixas()

    Dim Rs As New ADODB.Recordset
    Dim sSql As String

    g4.Rows = 1
    g4.Clear

    sSql = "SELECT ID_DESP , DSC_DESP , VALOR_DESP FROM DESPESAS_FIXAS "
    sSql = sSql & " WHERE ID_DESP NOT IN (SELECT ID_DESP FROM PAGTO_DESP_FIXA WHERE ID_FECHA = " & ID_FECHA & ")"
    sSql = sSql & " ORDER BY ID_DESP "

    Rs.Open sSql, Db, adOpenDynamic, 1

    If Not Rs.EOF Then
        Do Until Rs.EOF
            g4.AddItem Rs("ID_DESP") & Chr(9) & Rs("DSC_DESP") & Chr(9) & Format(Val(Rs("VALOR_DESP").Value), "0.00")
            Rs.MoveNext
        Loop
    End If
    Rs.Close
End Sub
Private Sub SomaGridDespesasFixas()

    Dim x As Integer

    TxtTOTALDespesasFixas.Text = ""
    For x = 1 To g4.Rows - 1
        g4.Row = x
        g4.Col = 0
        If Val(g4.Text) <> 0 Then
            g4.Col = 2
            TxtTOTALDespesasFixas.Text = Format(Val(TxtTOTALDespesasFixas.Text) + Val(g4.Text), "0.00")
        End If
    Next x
End Sub
