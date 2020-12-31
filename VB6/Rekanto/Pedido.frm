VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Pedido 
   Caption         =   "Rekanto - Inclusão de Pedidos"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5985
   ScaleWidth      =   7980
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdFinalizar 
      Caption         =   "&Finalizar Pedido"
      Height          =   375
      Left            =   6120
      TabIndex        =   15
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "&Excluir"
      Height          =   375
      Left            =   4200
      TabIndex        =   14
      Top             =   5400
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   2655
      Left            =   360
      TabIndex        =   13
      Top             =   2520
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4683
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      SelectionMode   =   1
   End
   Begin VB.CommandButton CmdLimpar 
      Caption         =   "&Limpar"
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Sai&r"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox TxtQtde 
      Height          =   315
      Left            =   5760
      TabIndex        =   7
      Top             =   1920
      Width           =   975
   End
   Begin VB.ComboBox CboProduto 
      Height          =   315
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1920
      Width           =   5295
   End
   Begin VB.ComboBox CboGarcom 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1080
      Width           =   5655
   End
   Begin VB.ComboBox CboMesa 
      Height          =   315
      ItemData        =   "Pedido.frx":0000
      Left            =   360
      List            =   "Pedido.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton CmdIncluir 
      Caption         =   "&Incluir"
      Default         =   -1  'True
      Height          =   375
      Left            =   6840
      TabIndex        =   9
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label LblPedido 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   1800
      TabIndex        =   12
      Top             =   330
      Width           =   1575
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Nr. Pedido :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   360
      TabIndex        =   11
      Top             =   360
      Width           =   1245
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Qtde"
      Height          =   195
      Left            =   5760
      TabIndex        =   3
      Top             =   1680
      Width           =   345
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Garçom"
      Height          =   195
      Left            =   2040
      TabIndex        =   2
      Top             =   840
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Produto"
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Mesa"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   390
   End
End
Attribute VB_Name = "Pedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdCancelar_Click()

    Unload Me
End Sub
Private Sub CmdExcluir_Click()

    On Error GoTo Erro_Excluir

    If Grid.Row > 0 Then
        Grid.RemoveItem Grid.RowSel
    End If

    Exit Sub
Erro_Excluir:
    If Err = 30015 Then
        Grid.ROWS = 1
        Resume Next
    End If
End Sub
Private Sub CmdFinalizar_Click()

    'Salvar os dados do pedido
    If CboMesa.Text <> "" And CboGarcom.ListIndex <> -1 And Grid.ROWS > 1 Then
        Call InserePedido(CboMesa.Text, CboGarcom.ItemData(CboGarcom.ListIndex))

        'Salvar os itens do pedido
        Call InsereItensPedido

        'Atualizar o status da mesa
        Call AtualizaStatusMesa

        'Preparar para um novo pedido
        Call CmdLimpar_Click
    Else
        MsgBox "Informe os dados do Pedido antes de efetuar a confirmação !", vbOKOnly, "Rekanto"
        CboMesa.SetFocus
    End If
End Sub
Private Sub CmdIncluir_Click()

    Dim NrPedido As Integer

    'Incluir o pedido
    'Call InserePedido(CboMesa.ItemData(CboMesa.ListIndex), CboGarcom.ItemData(CboGarcom.ListIndex))
    'Abrir a mesa

    'Incluir o item na lista
    If CboProduto.ListIndex = -1 Or Val(TxtQtde.Text) = 0 Then
        MsgBox "É necessário selecionar o produto e quantidade para adicionar na lista", vbOKOnly, "Rekanto"
        Exit Sub
    Else
        Grid.AddItem CboProduto.ItemData(CboProduto.ListIndex) & Chr(9) & CboProduto.Text & Chr(9) & TxtQtde.Text
    End If
    
    TxtQtde.Text = ""
    CboProduto.ListIndex = -1
    CboProduto.SetFocus
End Sub
Private Sub CmdLimpar_Click()

    'Limpar todos os campos e posicionar no campo de Mesa
    CboMesa.ListIndex = -1
    CboGarcom.ListIndex = -1
    CboProduto.ListIndex = -1
    TxtQtde.Text = ""
    Grid.ROWS = 1

    'Gera o número do novo pedido
    LblPedido.Caption = SelecionaUltimoPedido

    'Posiciona o foco para seleção de mesa
    CboMesa.SetFocus
End Sub
Private Sub Form_Load()

    'Carregar os combos
    'Mesas
    Call CarregaCombo(CboMesa, "MESAS", "ID_MESA", "ID_MESA")

    'Garçom
    Call CarregaCombo(CboGarcom, "FUNCIONARIOS", "ID_FUNC", "NOME_FUNC")

    'Produto
    Call CarregaCombo(CboProduto, "PRODUTOS", "ID_PRD", "DSC_PRD")

    'Formata o grid
    Call FormataGrid

    'Gera o número do novo pedido
    LblPedido.Caption = SelecionaUltimoPedido
End Sub
Public Function CarregaCombo(ByRef Combo As ComboBox, ByVal TABELA As String, CAMPO1 As String, CAMPO2 As String, Optional WHERE As String) As Integer

    Dim x As Integer
    Dim Rs As New ADODB.Recordset

    sSql = "Select " & CAMPO1 & " as CODIGO ," & CAMPO2 & " as DESCRICAO from " & TABELA

    If Len(Trim(WHERE)) > 0 Then
        sSql = sSql & WHERE
    End If
    If Len(Trim(CAMPO2)) > 0 Then
        sSql = sSql & " ORDER BY " & CAMPO2
    End If
    Rs.Open sSql, Db, adOpenDynamic, adLockReadOnly

    Do Until Rs.EOF
        Combo.AddItem Rs(1).Value
        Combo.ItemData(Combo.NewIndex) = Rs(0).Value
        Rs.MoveNext
    Loop
End Function
Private Sub InserePedido(ByVal Mesa As Integer, ByVal Garcom As Integer)

    Dim sSql As String

    sSql = "Insert into PEDIDO (ID_PED,ID_MESA,DT_HR_PED,ID_FUNC,SIT_PED) VALUES ("
    sSql = sSql & LblPedido.Caption & "," & Mesa & ",'" & Now & "'," & Garcom & ",1)"

    Db.Execute sSql
End Sub
Private Sub InsereItensPedido()

    Dim sSql As String
    Dim Rs As New ADODB.Recordset
    
    Dim ITEM As Integer
    Dim QTDE As Integer
    Dim ROWS As Integer

    ROWS = Grid.ROWS - 1
    For x = 1 To ROWS
        'NR
        ITEM = x

        'PRODUTO
        Grid.Row = x
        Grid.Col = 0
        PRODUTO = Grid.Text

        'QUANTIDADE DO ITEM
        Grid.Col = 2
        QTDE = Grid.Text

        sSql = "insert into ITEM_PEDIDO (ID_PED,ID_ITEM,ID_PRD,QTDE,SIT_ITEM,DT_HR_PED) VALUES ("
        sSql = sSql & LblPedido.Caption & "," & ITEM & "," & PRODUTO & "," & QTDE & ",1,'" & Now & "')"

        'Rs.Open sSql, Db, adOpenDynamic, adLockOptimistic
        Db.Execute sSql
    Next x
End Sub
Private Sub FormataGrid()

    Grid.ROWS = 1
    Grid.Cols = 3
    Grid.Row = 0
    
    Grid.Col = 0
    Grid.ColWidth(0) = 1

    Grid.Col = 1
    Grid.Text = "Produto"
    Grid.ColWidth(1) = 4000
    Grid.ColAlignment(1) = 1

    Grid.Col = 2
    Grid.Text = "Qtde"
    Grid.ColWidth(2) = 800
    Grid.ColAlignment(2) = 4
End Sub
Private Function SelecionaUltimoPedido() As Integer

    Dim sSql As String
    Dim Rs As New ADODB.Recordset

    sSql = "select max(ID_PED) as NR_PED from PEDIDO "

    Rs.Open sSql, Db, adOpenDynamic, adLockOptimistic

    If Not Rs.EOF Then
        SelecionaUltimoPedido = Val(Rs("NR_PED").Value & "") + 1
    End If
End Function
Private Sub AtualizaStatusMesa()

    Dim sSql As String
    Dim Rs As New ADODB.Recordset

    sSql = "update MESAS set ID_SIT = 1 where id_mesa = " & CboMesa.Text

    'Rs.Open sSql, Db, adOpenDynamic, adLockOptimistic
    Db.Execute sSql
End Sub
