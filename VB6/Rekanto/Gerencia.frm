VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{B2948DC1-18F7-455D-BA0B-F770D0B49264}#1.0#0"; "BemaFisc.ocx"
Begin VB.Form Gerencia 
   Caption         =   "Rekanto - Gerenciamento de Mesas"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12495
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9120
   ScaleWidth      =   12495
   StartUpPosition =   1  'CenterOwner
   Begin BemaFiscalWebCtl.BemaFiscalWeb Bema 
      Height          =   1455
      Left            =   10440
      TabIndex        =   8
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton CmdTransferir 
      Caption         =   "Transferir &Mesa"
      Enabled         =   0   'False
      Height          =   495
      Left            =   10200
      TabIndex        =   7
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton CmdImprimirConta 
      Caption         =   "&Imprimir Conta"
      Height          =   495
      Left            =   10200
      TabIndex        =   6
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton CmdPagto 
      Caption         =   "Registrar &Pagto"
      Height          =   495
      Left            =   10200
      TabIndex        =   3
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton CmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sai&r"
      Height          =   495
      Left            =   10200
      TabIndex        =   2
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton CmdDetalhes 
      Caption         =   "Detalhes"
      Height          =   495
      Left            =   10200
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   7800
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   13758
      _Version        =   393216
      Rows            =   6
      Cols            =   7
      FixedRows       =   0
      FixedCols       =   0
      Redraw          =   -1  'True
      HighLight       =   2
      ScrollBars      =   0
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Left            =   3480
      TabIndex        =   5
      Top             =   360
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Mesa Selecionada :"
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
      Height          =   300
      Left            =   960
      TabIndex        =   4
      Top             =   360
      Width           =   2370
   End
End
Attribute VB_Name = "Gerencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Modo As String
Private Sub CmdDetalhes_Click()

    Mesa = Grid.Text
    Load DetalhesPedido
    DetalhesPedido.Mesa = Mesa
    DetalhesPedido.Show 1
End Sub
Private Sub CmdImprimirConta_Click()

    Dim Rs As New ADODB.Recordset
    Dim Total As Currency

    'Define o modelo como IMPRESSORA TERMICA
    iModeloImpressora = 2

    'Abre a porta de comunicacao
    iPorta = IniciaPorta("LPT1")

    'Verifica se a porta de comunicação foi aberta corretamente
    If iPorta <= 0 Then
        MsgBox "Problemas ao Abrir a Porta de Comunicação.", vbOKOnly, "Rekanto"
        Exit Sub
    End If

    Squebra = Chr(13) + Chr(10)
    sBuffer = Space(17) & "Rekanto Grill" & Squebra
    sBuffer = sBuffer & Squebra
    sBuffer = sBuffer & Squebra
    sBuffer = sBuffer & "Av. Analice Sakataukas, 54" & Squebra
    sBuffer = sBuffer & "Osasco - Centro  -  " & "tel:(11) 3682-1596"
    sBuffer = sBuffer & Squebra & "Data : " & Format(Now, "dd/mm/yyyy") & "  -  Hora: " & Format(Now, "hh:mm:ss")
    sBuffer = sBuffer & Squebra
    sBuffer = sBuffer & Squebra
    sBuffer = sBuffer & Space(10) & "Fechamento da Mesa : " & Format(Grid.Text, "00")
    sBuffer = sBuffer & Squebra
    sBuffer = sBuffer & Squebra
    sBuffer = sBuffer & "Produto" & Space(21) & " Qtd  " & " Unit  " & " Total" & Squebra

    'Seleciona os registros que serão impressos
    sSql = "select PRODUTOS.DSC_PRD , ITEM.QTDE , PRODUTOS.VLR_PRD "
    sSql = sSql & " from ITEM_PEDIDO ITEM , PRODUTOS , PEDIDO "
    sSql = sSql & " WHERE PEDIDO.ID_MESA = " & Grid.Text
    sSql = sSql & " and ITEM.ID_PRD = PRODUTOS.ID_PRD "
    sSql = sSql & " and ITEM.ID_PED = PEDIDO.ID_PED "
    sSql = sSql & " and ITEM.SIT_ITEM not in (3,4) "

    Rs.Open sSql, DB, adOpenDynamic, adLockOptimistic

    Do Until Rs.EOF
        sBuffer = sBuffer & Left(Rs("DSC_PRD").Value, 28) & Space(28 - Len(Left(Rs("DSC_PRD").Value, 28))) & Space(1)
        sBuffer = sBuffer & Format(Rs("QTDE").Value, "000") & Space(2)
        sBuffer = sBuffer & Space(6 - Len(Format(Rs("VLR_PRD").Value, ".00"))) & Format(Rs("VLR_PRD").Value, ".00") & Space(1)
        sBuffer = sBuffer & Space(6 - Len(Format(Rs("VLR_PRD").Value * Rs("QTDE").Value, ".00"))) & Format(Rs("VLR_PRD").Value * Rs("QTDE").Value, ".00")

        sBuffer = sBuffer & Squebra
        Total = Total + (Rs("VLR_PRD").Value) * (Rs("QTDE").Value)
        Rs.MoveNext
    Loop

    sBuffer = sBuffer & Space(26) & "10% - Serviço : " & Format(Total * 0.1, ".00")

    sBuffer = sBuffer & Squebra
    sBuffer = sBuffer & Squebra

    sBuffer = sBuffer & Space(25) & "Total a Pagar : " & Format(Total * 1.1, ".00")

    sBuffer = sBuffer & Squebra
    sBuffer = sBuffer & Squebra
    sBuffer = sBuffer & Squebra

    'Debug.Print sBuffer
    iretorno = FormataTX(sBuffer, 2, 0, 0, 0, 1)

    'Aciona a guilhotina para o corte do papel
    iretorno = AcionaGuilhotina(1)
    iretorno = FechaPorta
End Sub
Private Sub CmdPagto_Click()

    Dim sSql As String
    Dim Rs As New ADODB.Recordset

    Mesa = Grid.Text

    If MsgBox("Confirma o fechamento ?", vbYesNo, "Rekanto") = vbYes Then
        'Verificar se a mesa possui itens pendentes
        sSql = "select  COUNT(0) as MESA_ABERTA"
        sSql = sSql & " from ITEM_PEDIDO ITEM, PEDIDO "
        sSql = sSql & " WHERE Pedido.ID_PED = ITEM.ID_PED "
        sSql = sSql & " and ITEM.SIT_ITEM = 1 "
        sSql = sSql & " and PEDIDO.ID_MESA = " & Mesa
    
        Rs.Open sSql, DB, adOpenDynamic, adLockOptimistic
    
        If Not Rs.EOF Then
            If Rs("MESA_ABERTA").Value > 0 Then
                MsgBox "Essa mesa possui itens pendentes!", vbOKOnly, "Rekanto"
                Exit Sub
            End If
        End If
        Set Rs = Nothing
    
        'Atualizar status da Mesa para LIVRE
        sSql = "update MESAS set ID_SIT = 0 where ID_MESA = " & Mesa
    
        'Rs.Open sSql, Db, adOpenDynamic, adLockOptimistic
        DB.Execute sSql
    
        Set Rs = Nothing
    
        'Atualizar os pedidos da mesa como PAGO
        sSql = "update PEDIDO set SIT_PED = 3 where ID_PED in ("
        sSql = sSql & "select ID_PED from PEDIDO where ID_MESA = " & Mesa & ")"
    
        'Rs.Open sSql, Db, adOpenDynamic, adLockOptimistic
        DB.Execute sSql
    
        'Atualizar os pedidos da mesa como PAGO
        sSql = "update ITEM_PEDIDO set SIT_ITEM = 3 where ID_PED in ("
        sSql = sSql & "select ID_PED from PEDIDO where ID_MESA = " & Mesa & ")"
    
        'Rs.Open sSql, Db, adOpenDynamic, adLockOptimistic
        DB.Execute sSql
    
        Grid.Visible = False
        Call PreparaGrid
        Call PreencheGridItensPreparados
        Call PreencheGrid
        Grid.Visible = True
        Me.Refresh
    End If
End Sub
Private Sub CmdSair_Click()

    Unload Me
End Sub
Private Sub Form_Activate()

    Call PreencheGridItensPreparados
    Call PreencheGrid
End Sub
Private Sub Form_Load()

    Modo = RecuperaModoINI
    If Modo = "COMANDA" Then
        Call PreparaGridComanda
    Else
        Call PreparaGrid
    End If
End Sub
Private Sub PreencheGrid()

    Dim sSql As String
    Dim Rs As New ADODB.Recordset
    Dim Encontrou As Integer

    Grid.Visible = False

    'Pintar as mesas que possuirem pedidos com itens em aberto
    sSql = " select ITEM.* , PEDIDO.ID_MESA "
    sSql = sSql & " From ITEM_PEDIDO ITEM , PEDIDO "
    sSql = sSql & " WHERE ITEM.dt_hr_pre Is Null "
    sSql = sSql & " AND ITEM.id_ped = PEDIDO.ID_PED "
    sSql = sSql & " AND ITEM.sit_item = 1 "

    Rs.Open sSql, DB, adOpenDynamic, adLockOptimistic

    If Not Rs.EOF Then
        Do Until Rs.EOF
            For x = 0 To 11
                Grid.Row = x
                Encontrou = False
                For y = 0 To 13
                    Grid.Col = y
                    If Val(Grid.Text) = Val(Rs("id_mesa").Value) Then
                        Grid.CellBackColor = RGB(255, 255, 0)
                        Encontrou = True
                        Exit For
                    End If
                Next y
                If Encontrou = True Then Exit For
            Next x
            Rs.MoveNext
        Loop
    End If
    Grid.Visible = True
End Sub
Private Sub PreparaGridComanda()

    Grid.Clear
    
    Grid.ROWS = 12
    Grid.Cols = 14
    Label2.Caption = ""
    Grid.Row = 0
    Grid.Col = 0
    Grid.width = (650 * 14) + 100
    Grid.Height = (650 * 12) + 100

    'Loops para redimensionar as celulas
    For x = 0 To 11
        Grid.RowHeight(x) = 650
        For y = 0 To 13
            Grid.Row = x
            Grid.Col = y
            
            Grid.ColWidth(y) = 650
            Grid.ColAlignment(y) = 4
            Grid.Text = (x * 14) + (y + 1)
        Next y
    Next x
End Sub
Private Sub Grid_Click()

    Label2.Caption = Grid.Text
End Sub
Private Sub PreencheGridItensAtrasados()

    Dim sSql As String
    Dim Rs As New ADODB.Recordset
    Dim Encontrou As Integer

    Grid.Visible = False

    'Pintar as mesas que possuirem pedidos com itens atrasados (mais de 5 minutos)
    sSql = " select ITEM.* , PEDIDO.ID_MESA "
    sSql = sSql & " From ITEM_PEDIDO ITEM , PEDIDO "
    sSql = sSql & " WHERE ITEM.dt_hr_pre Is Null "
    sSql = sSql & " AND ITEM.id_ped = PEDIDO.ID_PED "

    Rs.Open sSql, DB, adOpenDynamic, adLockOptimistic

    If Not Rs.EOF Then
        Do Until Rs.EOF
            For x = 0 To 5
                Grid.Row = x
                Encontrou = False
                For y = 0 To 6
                    Grid.Col = y
                    If Val(Grid.Text) = Val(Rs("id_mesa").Value) Then
                        Grid.CellBackColor = RGB(255, 255, 0)
                        Encontrou = True
                        Exit For
                    End If
                Next y
                If Encontrou = True Then Exit For
            Next x
            Rs.MoveNext
        Loop
    End If
    Grid.Visible = True
End Sub
Private Sub PreencheGridItensPreparados()

    Dim sSql As String
    Dim Rs As New ADODB.Recordset
    Dim Encontrou As Integer

    Grid.Visible = False

    'Pintar as mesas que possuirem pedidos com itens em aberto
    sSql = " select ITEM.* , PEDIDO.ID_MESA "
    sSql = sSql & " From ITEM_PEDIDO ITEM , PEDIDO "
    sSql = sSql & " WHERE ITEM.id_ped = PEDIDO.ID_PED "
    sSql = sSql & " AND ITEM.sit_item = 2 "

    Rs.Open sSql, DB, adOpenDynamic, adLockOptimistic

    If Not Rs.EOF Then
        Do Until Rs.EOF
            For x = 0 To 5
                Grid.Row = x
                Encontrou = False
                For y = 0 To 6
                    Grid.Col = y
                    If Val(Grid.Text) = Val(Rs("id_mesa").Value) Then
                        Grid.CellBackColor = RGB(0, 255, 0)
                        Encontrou = True
                        Exit For
                    End If
                Next y
                If Encontrou = True Then Exit For
            Next x
            Rs.MoveNext
        Loop
    End If
    Grid.Visible = True
End Sub
Private Sub PreparaGrid()

    Grid.Clear
    Label2.Caption = ""
    Grid.Row = 0
    Grid.Col = 0
    Grid.width = (1300 * 7) + 100
    Grid.Height = (1300 * 6) + 100

    'Loops para redimensionar as celulas
    For x = 0 To 5
        Grid.RowHeight(x) = 1300
        For y = 0 To 6
            Grid.Row = x
            Grid.Col = y
            
            Grid.ColWidth(y) = 1300
            Grid.ColAlignment(y) = 4
            Grid.Text = (x * 7) + (y + 1)
        Next y
    Next x
End Sub
Public Function RecuperaModoINI() As String

    Dim f As Integer

    f = FreeFile
    Open App.Path & "\rekanto.ini" For Input As #f
    Line Input #f, Linha
    Line Input #f, Linha
    Line Input #f, Linha
    Modo = UCase(Mid(Linha, 6))
    Close #f

    RecuperaModoINI = Modo
End Function

