VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form DetalhesPedido 
   Caption         =   "Rekanto - Detalhamento dos Pedidos"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13065
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7395
   ScaleWidth      =   13065
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "Excluir Item"
      Height          =   495
      Left            =   7200
      TabIndex        =   7
      Top             =   6720
      Width           =   1695
   End
   Begin VB.TextBox TxtFinal 
      Height          =   375
      Left            =   11265
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox TxtTOTAL 
      Height          =   375
      Left            =   11265
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton CmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sai&r"
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   6720
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   12840
      _ExtentX        =   22648
      _ExtentY        =   7858
      _Version        =   393216
      Rows            =   5
      Cols            =   4
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   2
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.Label LblTotal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   11280
      TabIndex        =   6
      Top             =   5400
      Width           =   1635
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Valor Total (+10%)"
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
      Left            =   9240
      TabIndex        =   5
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total Mesa"
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
      Left            =   9960
      TabIndex        =   2
      Top             =   4920
      Width           =   1185
   End
End
Attribute VB_Name = "DetalhesPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Mesa As Integer

Private Sub CmdExcluir_Click()

    Dim sSql As String
    Dim Rs As New ADODB.Recordset
    Dim ID_PED As Integer
    Dim ID_ITEM As Integer

    If Grid.ROWS <= 1 Then Exit Sub

    'Expurgar o item selecionado
    If MsgBox("Confirma a exclusão ?", vbYesNo, "Rekanto") = vbYes Then
        Grid.Col = 0
        ID_PED = Grid.Text

        Grid.Col = 7
        ID_ITEM = Grid.Text

        sSql = "update ITEM_PEDIDO set SIT_ITEM = 4 "
        sSql = sSql & "where ID_PED = " & ID_PED
        sSql = sSql & " and ID_ITEM = " & ID_ITEM
        
        DB.Execute sSql
        
        Call PreencheGrid
        Call CalculaValorTotal
    End If
End Sub
Private Sub CmdSair_Click()

    Unload Me
End Sub
Private Sub Form_Activate()

    Call FormataGrid
    
    Call PreencheGrid
    
    Call CalculaValorTotal
End Sub
Private Sub FormataGrid()

    Grid.ROWS = 1
    Grid.Cols = 8
    Grid.Row = 0

    Grid.Col = 0
    Grid.ColWidth(0) = 1000
    Grid.Text = "Pedido"
    Grid.ColAlignment(0) = 3

    Grid.Col = 1
    Grid.ColWidth(1) = 4800
    Grid.Text = "Produto"

    Grid.Col = 2
    Grid.ColWidth(2) = 900
    Grid.Text = "Qtde"

    Grid.Col = 3
    Grid.ColWidth(3) = 1500
    Grid.Text = "Valor"

    Grid.Col = 4
    Grid.ColWidth(4) = 1500
    Grid.Text = "Hora Pedido"

    Grid.Col = 5
    Grid.ColWidth(5) = 1500
    Grid.Text = "Hora Preparo"

    Grid.Col = 6
    Grid.ColWidth(6) = 2200
    Grid.Text = "Garçom"
    
    Grid.Col = 7
    Grid.ColWidth(7) = 0
    Grid.Text = ""
End Sub
Private Sub PreencheGrid()

    Dim sSql As String
    Dim Rs As New ADODB.Recordset

    Grid.Visible = False

    'Preencher o primeiro grid com os pedidos da cozinha/bar (ler parâmetro)
    sSql = "select PRODUTOS.DSC_PRD , ITEM.QTDE , PRODUTOS.VLR_PRD , FUNC.NOME_FUNC, "
    sSql = sSql & " ITEM.DT_HR_PED , ITEM.DT_HR_PRE , ITEM.ID_PED , ITEM.ID_ITEM "
    sSql = sSql & " from ITEM_PEDIDO ITEM , PRODUTOS , PEDIDO , FUNCIONARIOS FUNC "
    sSql = sSql & " WHERE PEDIDO.ID_MESA = " & Mesa
    sSql = sSql & " and ITEM.ID_PRD = PRODUTOS.ID_PRD "
    sSql = sSql & " and ITEM.ID_PED = PEDIDO.ID_PED "
    sSql = sSql & " and PEDIDO.ID_FUNC = FUNC.ID_FUNC "
    sSql = sSql & " and ITEM.SIT_ITEM not in (3,4) "

    Rs.Open sSql, DB, adOpenDynamic, adLockOptimistic

    x = 1
    Grid.ROWS = 1
    Do Until Rs.EOF

        Grid.ROWS = x + 1
        Grid.Row = x

        Grid.Col = 0
        Grid.ColWidth(0) = 1000
        Grid.Text = Rs("ID_PED").Value

        Grid.Col = 1
        Grid.ColWidth(1) = 4800
        Grid.Text = Rs("DSC_PRD").Value

        Grid.Col = 2
        Grid.ColWidth(2) = 800
        Grid.Text = Rs("QTDE").Value

        Grid.Col = 3
        Grid.ColWidth(3) = 1200
        Grid.Text = Format(Rs("VLR_PRD").Value & "", ".00")

        Grid.Col = 4
        Grid.ColWidth(4) = 1200
        Grid.Text = Format(Rs("DT_HR_PED").Value & "", "hh:mm:ss")

        Grid.Col = 5
        Grid.ColWidth(5) = 1200
        Grid.Text = Format(Rs("DT_HR_PRE").Value & "", "hh:mm:ss")

        Grid.Col = 6
        Grid.ColWidth(6) = 2200
        Grid.Text = Rs("NOME_FUNC").Value & ""

        Grid.Col = 7
        Grid.ColWidth(7) = 0
        Grid.Text = Rs("ID_ITEM").Value & ""

        x = x + 1
        Rs.MoveNext
    Loop
    Grid.Visible = True
End Sub
Private Sub CalculaValorTotal()

    Total = 0
    For x = 1 To Grid.ROWS - 1
        Grid.Row = x
        Grid.Col = 2
        QTDE = Grid.Text
        Grid.Col = 3
        Total = Total + (CDbl(Grid.Text) * QTDE)
    Next x
    TxtTOTAL.Text = Format(Total, ".00")
    LblTotal.Caption = "R$ " & Format(Total, ".00") * 0.1
    TxtFinal.Text = Format(TxtTOTAL.Text * 1.1, ".00")
End Sub
