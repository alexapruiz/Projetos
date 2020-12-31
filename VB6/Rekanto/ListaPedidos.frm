VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ListaPedidos 
   Caption         =   "Rekanto - Lista de Pedidos"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14970
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   14970
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   15000
      Left            =   11520
      Top             =   2160
   End
   Begin VB.CommandButton CmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sai&r"
      Height          =   375
      Left            =   3848
      TabIndex        =   4
      Top             =   10560
      Width           =   1575
   End
   Begin VB.CommandButton CmdVolta 
      Height          =   735
      Left            =   4455
      Picture         =   "ListaPedidos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton CmdOK 
      Default         =   -1  'True
      Height          =   735
      Left            =   7695
      Picture         =   "ListaPedidos.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4455
      Left            =   2775
      TabIndex        =   0
      Top             =   120
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   7858
      _Version        =   393216
      Rows            =   5
      Cols            =   4
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   4575
      Left            =   2160
      TabIndex        =   1
      Top             =   5760
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   8070
      _Version        =   393216
      Rows            =   5
      Cols            =   4
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
   End
End
Attribute VB_Name = "ListaPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FormataGrid(ByRef Grid As MSFlexGrid)

    Grid.ROWS = 1
    Grid.Cols = 7
    Grid.Row = 0

    Grid.Col = 0
    Grid.ColWidth(0) = 1
    Grid.Text = ""

    Grid.Col = 1
    Grid.ColWidth(1) = 1
    Grid.Text = ""

    Grid.Col = 2
    Grid.ColWidth(2) = 4000
    Grid.Text = "Produto"

    Grid.Col = 3
    Grid.ColWidth(3) = 800
    Grid.Text = "Qtde"

    Grid.Col = 4
    Grid.ColWidth(4) = 700
    Grid.Text = "Mesa"

    Grid.Col = 5
    Grid.ColWidth(5) = 1500
    Grid.Text = "Garçom"

    Grid.Col = 6
    Grid.ColWidth(6) = 1200
    Grid.Text = "Tempo Pedido"
End Sub
Private Sub CmdOK_Click()

    Dim sSql As String
    Dim Rs As New ADODB.Recordset

    Dim Pedido As Integer
    Dim ITEM As Integer

    If Grid.Row > 0 Then
        Grid.Col = 0
        Pedido = Grid.Text

        Grid.Col = 1
        ITEM = Grid.Text

        If Val(Pedido) <> 0 And Val(ITEM) <> 0 Then
            'Atualizar o status do item para "Atendido" -> Ler o grid e armazenar o pedido/item selecionado
            sSql = "update ITEM_PEDIDO set SIT_ITEM = 2 , DT_HR_PRE = '" & Now
            sSql = sSql & "' where ID_PED = " & Pedido
            sSql = sSql & " and ID_ITEM = " & ITEM

            'Rs.Open sSql, Db, adOpenDynamic, adLockOptimistic
            DB.Execute sSql

            Call PreencheGrid(Grid, 1)
            Call PreencheGrid(Grid2, 2)
        End If
    End If
    
    Grid.SetFocus
End Sub

Private Sub CmdSair_Click()

    Unload Me
End Sub
Private Sub CmdVolta_Click()

    Dim sSql As String
    Dim Rs As New ADODB.Recordset

    Dim Pedido As Integer
    Dim ITEM As Integer

    If Grid2.Row > 0 Then
        Grid2.Col = 0
        Pedido = Grid2.Text

        Grid2.Col = 1
        ITEM = Grid2.Text

        If Val(Pedido) <> 0 And Val(ITEM) <> 0 Then
            'Atualizar o status do item para "Atendido" -> Ler o grid e armazenar o pedido/item selecionado
            sSql = "update ITEM_PEDIDO set SIT_ITEM = 1, DT_HR_PRE = NULL where ID_PED = " & Pedido
            sSql = sSql & " and ID_ITEM = " & ITEM

            'Rs.Open sSql, Db, adOpenDynamic, adLockOptimistic
            DB.Execute sSql

            Call PreencheGrid(Grid, 1)
            Call PreencheGrid(Grid2, 2)
        End If
    End If
    
    Grid.SetFocus
End Sub
Private Sub Form_Activate()

    Grid.SetFocus
End Sub
Private Sub Form_Load()

    Call FormataGrid(Grid)
    Call FormataGrid(Grid2)
    
    Call PreencheGrid(Grid, "1")
    Call PreencheGrid(Grid2, "2")
End Sub
Private Sub PreencheGrid(ByRef Grid As MSFlexGrid, ByVal SITUACAO As Integer)

    Dim sSql As String
    Dim Rs As New ADODB.Recordset

    Grid.Visible = False

    CategoriaProduto = RecuperaCategoriaINI
    If SITUACAO = "1" Then
        'Preencher o primeiro grid com os pedidos da cozinha/bar (ler parâmetro)
        sSql = "SELECT ITEM.ID_PED ,  ITEM.ID_ITEM , PRD.DSC_PRD , PED.ID_MESA , ITEM.QTDE , PED.DT_HR_PED , SIT.DSC_STATUS , FUNC.NOME_FUNC"
        sSql = sSql & " FROM ITEM_PEDIDO ITEM , PRODUTOS PRD , PEDIDO PED , STATUS_PEDIDO SIT , FUNCIONARIOS FUNC"
        sSql = sSql & " WHERE ITEM.ID_PRD = PRD.ID_PRD"
        sSql = sSql & " AND PED.ID_PED = ITEM.ID_PED"
        sSql = sSql & " AND ITEM.SIT_ITEM = " & SITUACAO
        sSql = sSql & " AND ITEM.SIT_ITEM = SIT.SIT_PED"
        sSql = sSql & " AND FUNC.ID_FUNC = PED.ID_FUNC"
        sSql = sSql & " AND PRD.CAT_PRD IN (" & CategoriaProduto & ")"
    Else
        'Preencher o primeiro grid com os pedidos da cozinha/bar (ler parâmetro)
        sSql = "SELECT ITEM.ID_PED ,  ITEM.ID_ITEM , PRD.DSC_PRD , PED.ID_MESA , ITEM.QTDE , ITEM.DT_HR_PED , ITEM.DT_HR_PRE , SIT.DSC_STATUS , FUNC.NOME_FUNC"
        sSql = sSql & " FROM ITEM_PEDIDO ITEM , PRODUTOS PRD , PEDIDO PED , STATUS_PEDIDO SIT , FUNCIONARIOS FUNC"
        sSql = sSql & " WHERE ITEM.ID_PRD = PRD.ID_PRD"
        sSql = sSql & " AND PED.ID_PED = ITEM.ID_PED"
        sSql = sSql & " AND ITEM.SIT_ITEM IN (2,3)"
        sSql = sSql & " AND ITEM.SIT_ITEM = SIT.SIT_PED"
        sSql = sSql & " AND FUNC.ID_FUNC = PED.ID_FUNC"
        sSql = sSql & " AND PRD.CAT_PRD IN (" & CategoriaProduto & ")"

        Grid.Row = 0
        Grid.Cols = 8

        Grid.Col = 6
        Grid.Text = "Hora Pedido"

        Grid.Col = 7
        Grid.Text = "Hora Entrega"
    End If

    Rs.Open sSql, DB, adOpenDynamic, adLockOptimistic

    x = 1
    Grid.ROWS = 1
    Do Until Rs.EOF

        Grid.ROWS = x + 1
        Grid.Row = x

        Grid.Col = 0
        Grid.Text = Rs("ID_PED").Value

        Grid.Col = 1
        Grid.Text = Rs("ID_ITEM").Value

        Grid.Col = 2
        Grid.ColWidth(2) = 4000
        Grid.Text = Rs("DSC_PRD").Value

        Grid.Col = 3
        Grid.ColWidth(3) = 800
        Grid.Text = Rs("QTDE").Value
        
        Grid.Col = 4
        Grid.ColWidth(4) = 700
        Grid.Text = Rs("ID_MESA").Value

        Grid.Col = 5
        Grid.ColWidth(5) = 1500
        Grid.Text = Rs("NOME_FUNC").Value

        If SITUACAO = "1" Then
            Grid.Col = 6
            Grid.ColWidth(6) = 1100
            'Grid.Text = Format(Rs("DT_HR_PED").Value, "hh:mm:ss")
            Grid.Text = ConverteHora(DateDiff("s", Rs("DT_HR_PED").Value, Now))
        Else
            Grid.Col = 6
            Grid.ColWidth(6) = 1100
            Grid.ColAlignment(6) = 3
            Grid.Text = Format(Rs("DT_HR_PED").Value, "hh:mm:ss")

            Grid.Col = 7
            Grid.ColWidth(7) = 1100
            Grid.ColAlignment(7) = 3
            Grid.Text = Format(Rs("DT_HR_PRE").Value, "hh:mm:ss")
        End If

        x = x + 1
        Rs.MoveNext
    Loop
    Grid.Visible = True
End Sub
Private Sub Timer1_Timer()

    Call FormataGrid(Grid)
    Call FormataGrid(Grid2)
    
    Call PreencheGrid(Grid, "1")
    Call PreencheGrid(Grid2, "2")
End Sub
Private Function ConverteHora(ByVal Tempo As Long) As String


    ConverteHora = Format(Int(Tempo / 60), "00") & ":" & Format(Tempo Mod 60, "00")
End Function
Public Function RecuperaCategoriaINI() As String

    Dim f As Integer

    f = FreeFile
    Open App.Path & "\rekanto.ini" For Input As #f
    Line Input #f, Linha
    Line Input #f, Linha
    categoria = Mid(Linha, 9)
    Close #f

    Select Case categoria
        Case "ESCRITORIO"
            RecuperaCategoriaINI = "1,2"
        Case "COZINHA"
            RecuperaCategoriaINI = "1"
        Case "BAR"
            RecuperaCategoriaINI = "2"
    End Select
End Function
