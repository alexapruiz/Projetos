VERSION 5.00
Begin VB.Form Principal 
   Caption         =   "Sistema de Gerenciamento de Restaurantes - Rekanto Grill e Cervejaria"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   11985
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Gerenciamento de Mesas"
      Height          =   1095
      Left            =   4320
      Picture         =   "Principal.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton CmdListaPedidos 
      Caption         =   "Lista de Pedidos"
      Height          =   1095
      Left            =   2520
      Picture         =   "Principal.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton CmdIncluirPedido 
      Caption         =   "Incluir Pedido"
      Height          =   1095
      Left            =   720
      Picture         =   "Principal.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Menu MnuSistema 
      Caption         =   "Sistema"
      Begin VB.Menu MnuIniciar 
         Caption         =   "Iniciar Movimento"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSair 
         Caption         =   "Sai&r"
      End
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdIncluirPedido_Click()

    Pedido.Show 1
End Sub
Private Sub CmdListaPedidos_Click()

    ListaPedidos.Show 1
End Sub
Private Sub Command1_Click()

    Gerencia.Show 1
End Sub
Private Sub Form_Load()

    Dim f As Integer
    Dim Categoria As String

    f = FreeFile
    Open App.Path & "\rekanto.ini" For Input As #f
    Line Input #f, Linha
    Line Input #f, Linha
    Categoria = Mid(Linha, 9)
    Close #f
    If Categoria = "ESCRITORIO" Then
        Command1.Enabled = True
        CmdIncluirPedido.Enabled = True
        MnuIniciar.Enabled = True
    Else
        Command1.Enabled = False
        CmdIncluirPedido.Enabled = False
        MnuIniciar.Enabled = False
    End If
End Sub
Private Sub MnuIniciar_Click()

    Call IniciarMovimento
End Sub
Private Sub IniciarMovimento()

    Dim sSql As String
    Dim Rs As New ADODB.Recordset

    If MsgBox("Todos os itens/pedidos em aberto serão excluídos da base atual. Confirma a Inicialização do Movimento ?", vbYesNo, "SGB") = vbYes Then
        'Atualizar todos os itens e pedidos para "FECHADO"
        sSql = " update ITEM_PEDIDO set SIT_ITEM = 4 "
        sSql = sSql & " where SIT_ITEM <> 4 "
    
        Rs.Open sSql, DB, adOpenDynamic, adLockOptimistic
        
        Set Rs = Nothing
    
        sSql = " update PEDIDO set SIT_PED = 4 "
        sSql = sSql & " where SIT_PED <> 4 "
    
        Rs.Open sSql, DB, adOpenDynamic, adLockOptimistic
    
        MsgBox "O Movimento de " & Format(Now, "dd/mm/yyyy") & " pode ser iniciado.", vbOKOnly, "SGB"
    End If
End Sub
Private Sub MnuSair_Click()

    End
End Sub
