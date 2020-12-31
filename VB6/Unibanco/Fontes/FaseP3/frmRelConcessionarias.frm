VERSION 5.00
Begin VB.Form frmRelConcessionarias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Concessionárias"
   ClientHeight    =   2304
   ClientLeft      =   1056
   ClientTop       =   2640
   ClientWidth     =   4908
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2304
   ScaleWidth      =   4908
   Begin VB.Frame fraPrincipal 
      Caption         =   "Agência de coleta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1812
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   4452
      Begin VB.ComboBox cmbAgenf_Agencia 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   720
         Width           =   3708
      End
      Begin VB.CommandButton cmdImp 
         Caption         =   "&Imprimir"
         Height          =   312
         Left            =   600
         TabIndex        =   1
         Top             =   1200
         Width           =   1572
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   312
         Left            =   2280
         TabIndex        =   2
         Top             =   1200
         Width           =   1572
      End
      Begin VB.Label lblAgenf_AgOrigem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agência de Origem"
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
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1608
      End
   End
End
Attribute VB_Name = "frmRelConcessionarias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdImp_Click()
    
    If cmbAgenf_Agencia.ListIndex = -1 Then
        Beep
        MsgBox "Agencia de coleta não informada!", vbExclamation + vbOKOnly, App.Title
        cmbAgenf_Agencia.SetFocus
        Exit Sub
    End If
    
    Call ImprimeRelatorio

End Sub

Private Sub cmdSair_Click()

    Unload Me
    
End Sub

Private Sub Form_Activate()
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    If Not CarregaAgenciaColetaEmCombo(cmbAgenf_Agencia) Then
        Beep
        MsgBox "Não existe(m) agência de coleta cadastradas, favor verificar!", vbExclamation + vbOKOnly, App.Title
        Unload Me
        Exit Sub
    End If
    
End Sub

Private Sub ImprimeRelatorio()
    
Dim qryVerificarMovto   As rdoQuery
Dim RsVerificarMovto    As rdoResultset
Dim QryTimeOut      As Variant
    
On Error GoTo Err_ImprimeRelatorio

    'Aumenta timeout devido ao processamento demorado da Procedure
    QryTimeOut = Geral.Banco.QueryTimeout
    Geral.Banco.QueryTimeout = 300
    
    Set qryVerificarMovto = Geral.Banco.CreateQuery("", "{call GetRelatConcessionarias(?,?,?)}")
    
    Screen.MousePointer = vbHourglass
    
    With qryVerificarMovto
        .rdoParameters(0).Value = Geral.DataProcessamento
        'Código da agência de origem da capa de Malote/Envelope
        .rdoParameters(1).Value = cmbAgenf_Agencia.ItemData(cmbAgenf_Agencia.ListIndex)
        .rdoParameters(2).Value = 1     'Somente verificação de movto   (0)=Não  (1)=Sim
        Set RsVerificarMovto = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    With Principal.RptGeral
        If qryVerificarMovto.rdoParameters(2).Value > 0 Then
            .Connect = Geral.Banco.Connect
            If Geral.Backup Then
                .ReportFileName = App.path & "\RelConcessionariasBK.rpt"
            Else
                .ReportFileName = App.path & "\RelConcessionarias.rpt"
            End If
            
            .Formulas(0) = "AgProcessadora    = '" & Geral.AgenciaCentral & "'"
            .Formulas(1) = "DataMovto    = '" & DataDD_MM_AAAA(Geral.DataProcessamento) & "'"
            If cmbAgenf_Agencia.ListIndex = 0 Then
                .Formulas(2) = "AgenciaOrigem = ''"
            Else
                .Formulas(2) = "AgenciaOrigem = 'Agencia  " & Trim(cmbAgenf_Agencia.List(cmbAgenf_Agencia.ListIndex)) & "'"
            End If

            .StoredProcParam(0) = Geral.DataProcessamento
            'Relatório com todas agências
            .StoredProcParam(1) = cmbAgenf_Agencia.ItemData(cmbAgenf_Agencia.ListIndex)
            .StoredProcParam(2) = 0
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .WindowTitle = "Relatório de Concessionárias"
            .Action = 1
        Else
            Screen.MousePointer = vbDefault

            MsgBox "Não existe movimento para emissão do Relatório!" & vbCrLf & vbCrLf & _
            IIf(cmbAgenf_Agencia.ItemData(cmbAgenf_Agencia.ListIndex) > 0, _
                "Agência de origem (" & Trim(cmbAgenf_Agencia.List(cmbAgenf_Agencia.ListIndex)) & ")", "") _
                , vbInformation, App.Title & " ( " & "Relatório de Concessionárias" & " )"
        End If
        
        .ReportFileName = Empty
        .Connect = Empty
        .Formulas(0) = Empty
        .Formulas(1) = Empty
        .Formulas(2) = Empty
        .StoredProcParam(0) = Empty
        .StoredProcParam(1) = Empty
        .StoredProcParam(2) = Empty
        Screen.MousePointer = vbDefault
    End With
    
Exit_ImprimeRelatorio:
    Geral.Banco.QueryTimeout = QryTimeOut
    Screen.MousePointer = vbDefault
    qryVerificarMovto.Close
    If Not (RsVerificarMovto Is Nothing) Then Set RsVerificarMovto = Nothing
    
    Exit Sub

Err_ImprimeRelatorio:
    Beep
    With Principal.RptGeral
        .ReportFileName = Empty
        .Formulas(0) = Empty
        .Formulas(1) = Empty
        .Formulas(2) = Empty
        .StoredProcParam(0) = Empty
        .StoredProcParam(1) = Empty
        .StoredProcParam(2) = Empty
    End With
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, App.Title
    GoTo Exit_ImprimeRelatorio
    
    
End Sub


