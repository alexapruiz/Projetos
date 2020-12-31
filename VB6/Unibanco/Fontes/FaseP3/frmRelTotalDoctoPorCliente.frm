VERSION 5.00
Object = "{4A8E27C4-EACA-11D3-9FFC-00104BC8688C}#1.0#0"; "dateedit.ocx"
Begin VB.Form frmRelTotalDoctoPorCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Totais de Doctos. por Cliente"
   ClientHeight    =   3156
   ClientLeft      =   2352
   ClientTop       =   3408
   ClientWidth     =   4908
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3156
   ScaleWidth      =   4908
   Begin VB.Frame fraPrincipal 
      Caption         =   "Informações de  Movimento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2652
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   4452
      Begin VB.TextBox TxtNumMalote 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   360
         Left            =   360
         MaxLength       =   12
         TabIndex        =   2
         Top             =   1440
         Width           =   2028
      End
      Begin VB.CommandButton cmdImp 
         Caption         =   "&Imprimir"
         Height          =   312
         Left            =   600
         TabIndex        =   3
         Top             =   2160
         Width           =   1572
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   312
         Left            =   2280
         TabIndex        =   4
         Top             =   2160
         Width           =   1572
      End
      Begin DATEEDITLib.DateEdit dteDataInicial 
         Height          =   384
         Left            =   360
         TabIndex        =   0
         Top             =   600
         Width           =   1584
         _Version        =   65537
         _ExtentX        =   2794
         _ExtentY        =   677
         _StockProps     =   93
         ForeColor       =   8388608
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
      End
      Begin DATEEDITLib.DateEdit dteDataFinal 
         Height          =   384
         Left            =   2520
         TabIndex        =   1
         Top             =   600
         Width           =   1584
         _Version        =   65537
         _ExtentX        =   2794
         _ExtentY        =   677
         _StockProps     =   93
         ForeColor       =   8388608
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
      End
      Begin VB.Label lblNumMalote 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número do Malote"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   192
         Left            =   360
         TabIndex        =   9
         Top             =   1200
         Width           =   1536
      End
      Begin VB.Label Label1 
         Caption         =   "à"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Left            =   2160
         TabIndex        =   8
         Top             =   720
         Width           =   252
      End
      Begin VB.Label lblDataFinal 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Data Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   192
         Left            =   2520
         TabIndex        =   7
         Top             =   360
         Width           =   864
      End
      Begin VB.Label lblDataInicial 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Data Inicial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   192
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   948
      End
   End
End
Attribute VB_Name = "frmRelTotalDoctoPorCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdImp_Click()
    
    If Trim(dteDataInicial.Text) = "" Then
        Beep
        MsgBox "Digite a data inicial do período!", vbExclamation + vbOKOnly, App.Title
        dteDataInicial.SetFocus
        Exit Sub
    End If
    
    If Trim(dteDataFinal.Text) = "" Then
        Beep
        MsgBox "Digite a data final do período!", vbExclamation + vbOKOnly, App.Title
        dteDataFinal.SetFocus
        Exit Sub
    End If
    
    If dteDataFinal.InverseText < dteDataInicial.InverseText Then
        Beep
        MsgBox "Período incorreto, favor verificar!", vbExclamation + vbOKOnly, App.Title
        dteDataInicial.SetFocus
        Exit Sub
    End If
    
    Call ImprimeRelatorio
    
End Sub

Private Sub cmdSair_Click()

    Unload Me
    
End Sub

Private Sub dteDataFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        If DataOk(Val(dteDataFinal.Text)) Then
            SendKeys "{TAB}"
        End If
    End If

End Sub

Private Sub dteDataInicial_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If DataOk(Val(dteDataInicial.Text)) Then
            SendKeys "{TAB}"
        End If
    End If

End Sub

Private Sub Form_Activate()
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
End Sub

Private Sub ImprimeRelatorio()

Dim qryVerificarMovto   As rdoQuery
Dim RsVerificarMovto    As rdoResultset
Dim QryTimeOut          As Variant
Dim i                   As Integer

On Error GoTo Err_ImprimeRelatorio

    'Aumenta timeout devido ao processamento demorado da Procedure
    QryTimeOut = Geral.Banco.QueryTimeout
    Geral.Banco.QueryTimeout = 300
    
    Set qryVerificarMovto = Geral.Banco.CreateQuery("", "{call TotalDoctosPorCliente(?,?,?,?)}")
    
    Screen.MousePointer = vbHourglass
    
    With qryVerificarMovto
        .rdoParameters(0).Value = dteDataInicial.InverseText
        .rdoParameters(1).Value = dteDataFinal.InverseText
        If TxtNumMalote.Text = "" Then
            .rdoParameters(2).Value = 0  'Numero do malote caso selecionado
        Else
            .rdoParameters(2).Value = TxtNumMalote.Text
        End If
        .rdoParameters(3).Value = 1     'Somente verificação de movto   (0)=Não  (1)=Sim
        Set RsVerificarMovto = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    With Principal.RptGeral
        If qryVerificarMovto.rdoParameters(3).Value > 0 Then
            .Connect = Geral.Banco.Connect
            If Geral.Backup Then
                .ReportFileName = App.path & "\RelTotalDoctoPorClienteBk.rpt"
            Else
                .ReportFileName = App.path & "\RelTotalDoctoPorClienteProd.rpt"
            End If
            .Formulas(0) = "AgenciaCentral    = '" & Geral.AgenciaCentral & "'"
            .Formulas(1) = "DataInicial = '" & dteDataInicial.MaskText & "'"
            .Formulas(2) = "DataFinal = '" & dteDataFinal.MaskText & "'"

            .StoredProcParam(0) = dteDataInicial.InverseText
            .StoredProcParam(1) = dteDataFinal.InverseText
            If TxtNumMalote.Text = "" Then
                .StoredProcParam(2) = 0
            Else
                .StoredProcParam(2) = TxtNumMalote.Text
            End If
            .StoredProcParam(3) = 0     'Somente verificação de movto   (0)=Não  (1)=Sim
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .WindowTitle = "Relatório de Totais de Documentos por Cliente"
            .Action = 1
        Else
            Screen.MousePointer = vbDefault
            MsgBox "Não existe movimento para emissão do Relatório!", vbInformation + vbOKOnly, App.Title
            cmdSair.SetFocus
        End If
    End With
    
Exit_ImprimeRelatorio:

    With Principal.RptGeral
        .ReportFileName = Empty
        .Connect = Empty
        .Destination = Empty
        .WindowState = Empty
        .WindowTitle = Empty
        .Connect = Empty
        For i = 0 To 2: .Formulas(i) = Empty: Next
        .StoredProcParam(0) = Empty
        .StoredProcParam(1) = Empty
        .StoredProcParam(2) = Empty
        .StoredProcParam(3) = Empty
    End With
    Geral.Banco.QueryTimeout = QryTimeOut
    Screen.MousePointer = vbDefault
    qryVerificarMovto.Close
    If Not (RsVerificarMovto Is Nothing) Then Set RsVerificarMovto = Nothing
    
    Exit Sub

Err_ImprimeRelatorio:
    Beep
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, App.Title
    GoTo Exit_ImprimeRelatorio

End Sub

Private Sub TxtNumMalote_GotFocus()
    
    TxtNumMalote.SelStart = 0
    TxtNumMalote.SelLength = TxtNumMalote.MaxLength
    
End Sub

Private Sub TxtNumMalote_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub
