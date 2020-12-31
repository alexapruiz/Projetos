VERSION 5.00
Begin VB.Form FrmExpedidos 
   Caption         =   "Relatório De Caixa Expresso/Malote Empresa"
   ClientHeight    =   1740
   ClientLeft      =   3144
   ClientTop       =   3900
   ClientWidth     =   4464
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   4464
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   432
      Left            =   3408
      TabIndex        =   4
      Top             =   1224
      Width           =   972
   End
   Begin VB.CommandButton CmdImp 
      Caption         =   "&Imprimir"
      Height          =   432
      Left            =   2316
      TabIndex        =   3
      Top             =   1224
      Width           =   972
   End
   Begin VB.Frame Frame1 
      Caption         =   "Escolha uma Opção :"
      Height          =   1032
      Left            =   96
      TabIndex        =   0
      Top             =   96
      Width           =   4272
      Begin VB.OptionButton OptNExpedidos 
         Caption         =   "Não Expedido"
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
         Left            =   2076
         TabIndex        =   2
         Top             =   504
         Width           =   1740
      End
      Begin VB.OptionButton OptExpedidos 
         Caption         =   "Expedido"
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
         TabIndex        =   1
         Top             =   504
         Width           =   1416
      End
   End
End
Attribute VB_Name = "FrmExpedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Enum enumTipoRelatorio
    eCaixaExpresso_MaloteEmpresa_Ocorrencia
    eCaixaExpresso_MaloteEmpresa_Cheque
End Enum

Private RsRelOcorCapa As rdoResultset
Private qrypesqDados  As rdoQuery
Private qryPesqDadosCheque  As rdoQuery

Public m_TipoRelatorio As enumTipoRelatorio
Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdImp_Click()

    Dim Situacao        As String
    Dim sReportFileName As String
    
    If OptExpedidos.Value = True Then Situacao = "E"
    If OptNExpedidos.Value = True Then Situacao = "T"


    If Me.m_TipoRelatorio = eCaixaExpresso_MaloteEmpresa_Ocorrencia Then
        With qrypesqDados
            .rdoParameters(0).Value = Geral.DataProcessamento
            .rdoParameters(1).Value = Situacao
            Set RsRelOcorCapa = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
        End With
        
        If Geral.Backup Then
            sReportFileName = App.path & "\RelExpedOcorBk.rpt"
        Else
            sReportFileName = App.path & "\RelExpedOcorProd.rpt"
        End If
        
    ElseIf Me.m_TipoRelatorio = eCaixaExpresso_MaloteEmpresa_Cheque Then
        With qryPesqDadosCheque
            .rdoParameters(0).Value = Geral.DataProcessamento
            .rdoParameters(1).Value = Situacao
            Set RsRelOcorCapa = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
        End With
        
        If Geral.Backup Then
            sReportFileName = App.path & "\RelCheqSupBk.rpt"
        Else
            sReportFileName = App.path & "\RelCheqSupProd.rpt"
        End If
    End If
    
    If RsRelOcorCapa.RowCount <> 0 Then
            Principal.RptGeral.Connect = Geral.StringConexao
            Principal.RptGeral.ReportFileName = sReportFileName
            Principal.RptGeral.Formulas(0) = "AgenciaCentral    = '" & Geral.AgenciaCentral & "'"
            Principal.RptGeral.StoredProcParam(0) = Geral.DataProcessamento
            Principal.RptGeral.StoredProcParam(1) = Situacao
            Principal.RptGeral.Destination = crptToWindow
            Principal.RptGeral.WindowTop = 1
            Principal.RptGeral.WindowLeft = 1
            Principal.RptGeral.WindowState = crptMaximized
            Principal.RptGeral.Action = 0
    Else
        MsgBox "Não Existem dados para emissão deste Relátorio.", vbInformation, App.Title
    End If
    
    Principal.RptGeral.ReportFileName = Empty
    Principal.RptGeral.Formulas(0) = Empty
    Principal.RptGeral.StoredProcParam(0) = Empty
    Principal.RptGeral.StoredProcParam(1) = Empty
    
End Sub

Private Sub Form_Load()

 
    'Query que verifique se exite dados para impressão
    Set qrypesqDados = Geral.Banco.CreateQuery("", "{Call GetRelExpedidos(?,?)}")
    
    '''''''''''''''''''''''''''''''''''''''''''''
    'Cria query para carregar dados do relatório'
    '''''''''''''''''''''''''''''''''''''''''''''
    Set qryPesqDadosCheque = Geral.Banco.CreateQuery("", "{Call GetRelChequeSuperior(?,?)}")
    
End Sub

