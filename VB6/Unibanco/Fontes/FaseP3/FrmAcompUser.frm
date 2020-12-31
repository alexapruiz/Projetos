VERSION 5.00
Begin VB.Form FrmAcompUsers 
   Caption         =   "Acompanhamento de Usuários"
   ClientHeight    =   1884
   ClientLeft      =   2136
   ClientTop       =   4044
   ClientWidth     =   7284
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1884
   ScaleWidth      =   7284
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   396
      Left            =   6048
      TabIndex        =   2
      Top             =   1356
      Width           =   972
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   396
      Left            =   4920
      TabIndex        =   1
      Top             =   1356
      Width           =   972
   End
   Begin VB.Frame Frame1 
      Height          =   1020
      Left            =   288
      TabIndex        =   0
      Top             =   120
      Width           =   6708
      Begin VB.ComboBox CboUsers 
         Height          =   288
         Left            =   2556
         TabIndex        =   4
         Top             =   480
         Width           =   3888
      End
      Begin VB.Label Label1 
         Caption         =   "Nome do Usuário : "
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   168
         TabIndex        =   3
         Top             =   444
         Width           =   2244
      End
   End
End
Attribute VB_Name = "FrmAcompUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private qryNomeUser As rdoQuery     'query que traz todos usuários
Private qryLogsUser As rdoQuery     'query que verifica registros de log
Private RsNomeUser  As rdoResultset 'Recordset de Usuários
Private RsLogsUser  As rdoResultset 'Recordset de Logs de Usuário
Private CountUsers  As Integer      'Conta os usuários cadastrados
Private CaputUsers  As String       'Caputura o nome do Usuário

Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub CmdImprimir_Click()

    
Screen.MousePointer = vbHourglass
    
    If CboUsers.Text = "Todos" Then
         CaputUsers = "Null"
    Else
         CaputUsers = CboUsers.Text
    End If

    'Faz pesquisa para verificar se existe registros para o usuários selecionado
    Set RsLogsUser = Nothing
    With qryLogsUser
        .rdoParameters(0) = Geral.DataProcessamento
        .rdoParameters(1) = CaputUsers
        .QueryTimeout = 300
        Set RsLogsUser = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With

    If Not RsLogsUser.EOF Then
        With Principal
            
            .RptGeral.Connect = Geral.StringConexao
            
            If Geral.Backup Then
                .RptGeral.ReportFileName = App.path & "\AcompUserBk.rpt"
            Else
                .RptGeral.ReportFileName = App.path & "\AcompUserProd.rpt"
            End If
            
            .RptGeral.WindowTop = 1
            .RptGeral.WindowLeft = 1
            .RptGeral.WindowState = crptMaximized
        
            .RptGeral.Destination = crptToWindow
            .RptGeral.WindowTitle = "Relatório de Acompanhamento de Produção por Operador"

            .RptGeral.StoredProcParam(0) = Geral.DataProcessamento
            .RptGeral.StoredProcParam(1) = CaputUsers
            .RptGeral.PrintReport

            .RptGeral.StoredProcParam(0) = Empty
            .RptGeral.StoredProcParam(1) = Empty
            .RptGeral.ReportFileName = Empty
            Screen.MousePointer = vbDefault
        End With
    Else
        Screen.MousePointer = vbDefault
        If CaputUsers = "Null" Then
            MsgBox "Nenhum usuário possui registro de Log", vbExclamation
            Exit Sub
        Else
        MsgBox "O usuário" & " " & CaputUsers & " " & " não possui registros de Log", vbExclamation
        Exit Sub
        End If
        
    End If
End Sub
Private Sub CmdSair_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
   'Inclusão de chamada a rotina AtualizaAtividade
    Call AtualizaAtividade(15)
End Sub
Private Sub Form_Load()

'Query verifica se usuário possui registros de Log
Set qryLogsUser = Geral.Banco.CreateQuery("", "{call GetAcompanhamentoUser (?,?)}")

'Query que traz todos os usuários cadastrados no Sistema
Set qryNomeUser = Geral.Banco.CreateQuery("", "{call GetTodosUsuarios}")

    With qryNomeUser
        Set RsNomeUser = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With

    If Not RsNomeUser.EOF Then
    
        For CountUsers = 0 To RsNomeUser.RowCount - 1
            CboUsers.AddItem RsNomeUser!Nome
            RsNomeUser.MoveNext
        Next
        CboUsers.AddItem "Todos"
        CboUsers.Text = "Todos"
    End If


End Sub

Private Sub Form_Unload(Cancel As Integer)

Set qryNomeUser = Nothing
Set qryLogsUser = Nothing
Set RsNomeUser = Nothing
Set RsLogsUser = Nothing

End Sub
