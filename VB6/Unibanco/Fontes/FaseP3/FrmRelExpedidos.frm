VERSION 5.00
Begin VB.Form FrmRelExpedidos 
   Caption         =   "Relatório de Capas Expedidas"
   ClientHeight    =   1512
   ClientLeft      =   3900
   ClientTop       =   4080
   ClientWidth     =   4440
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1512
   ScaleWidth      =   4440
   Begin VB.CommandButton CmdImp 
      Caption         =   "&Imprimir"
      Height          =   432
      Left            =   2208
      TabIndex        =   5
      Top             =   936
      Width           =   972
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   432
      Left            =   3300
      TabIndex        =   4
      Top             =   936
      Width           =   972
   End
   Begin VB.Frame Frame2 
      Caption         =   "Escolha uma opção :"
      Height          =   636
      Left            =   72
      TabIndex        =   0
      Top             =   144
      Width           =   4272
      Begin VB.OptionButton OptMal 
         Caption         =   "Malote"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1716
         TabIndex        =   3
         Top             =   288
         Width           =   948
      End
      Begin VB.OptionButton optEnv 
         Caption         =   "Envelope"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   408
         TabIndex        =   2
         Top             =   288
         Width           =   1164
      End
      Begin VB.OptionButton OptTodos 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2844
         TabIndex        =   1
         Top             =   288
         Value           =   -1  'True
         Width           =   948
      End
   End
End
Attribute VB_Name = "FrmRelExpedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TipoCapa   As String
Dim TituloCapa As String

Option Explicit
Private Sub CmdCancel_Click()
    Unload Me
End Sub
Private Sub CmdImp_Click()

    Dim qryVerDadosRelExpedidos As rdoQuery
    Dim RsVerDadosRelExpedidos  As rdoResultset
    
    Set qryVerDadosRelExpedidos = Geral.Banco.CreateQuery("", "{call listacapa (?,?)}")
    
    Screen.MousePointer = vbHourglass
    
    With qryVerDadosRelExpedidos
        .rdoParameters(0).Value = Geral.DataProcessamento
        .rdoParameters(1).Value = TipoCapa
        Set RsVerDadosRelExpedidos = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsVerDadosRelExpedidos.RowCount <> 0 Then
        With Principal
            .RptGeral.Connect = Geral.StringConexao
            
            If Geral.Backup Then
                .RptGeral.ReportFileName = App.path & "\RelExpedidoBk.rpt"
            Else
                If Not OptTodos.Value Then
                   .RptGeral.ReportFileName = App.path & "\RelExpedidoProd.rpt"
                Else
                   .RptGeral.ReportFileName = App.path & "\RelExpedidoProdTodos.rpt"
                End If
            End If
            
            .RptGeral.Formulas(0) = "AgenciaCentral    = '" & Geral.AgenciaCentral & "'"
            .RptGeral.StoredProcParam(0) = Geral.DataProcessamento
            .RptGeral.StoredProcParam(1) = TipoCapa
            .RptGeral.Formulas(1) = "Titulo = '" & TituloCapa & "'"
            .RptGeral.WindowTitle = "Relatório de Expedição"
            .RptGeral.Destination = crptToWindow
            .RptGeral.WindowState = crptMaximized
            .RptGeral.Action = 1
        End With
    Else
        Screen.MousePointer = vbDefault
        MsgBox "Não existem capas expedidas para o período atual", vbInformation, App.Title
    End If
    
    Principal.RptGeral.ReportFileName = Empty
    Principal.RptGeral.Formulas(0) = Empty
    
    Principal.RptGeral.StoredProcParam(0) = Empty
    Principal.RptGeral.StoredProcParam(1) = Empty
    
    qryVerDadosRelExpedidos.Close
    
    Screen.MousePointer = vbDefault
  
End Sub
Private Sub Form_Load()
    TipoCapa = "T"
    TituloCapa = "Controle de Envelope/Malote Empresa Expedidos"
End Sub
Private Sub optEnv_Click()
    TipoCapa = "E"
    TituloCapa = "Controle de Envelopes Expedidos"
End Sub
Private Sub OptMal_Click()
    TipoCapa = "M"
    TituloCapa = "Controle de Malote Empresa Expedidos"
End Sub
Private Sub OptTodos_Click()
    TipoCapa = "T"
    TituloCapa = "Controle de Envelope/Malote Empresa Expedidos"
End Sub
