VERSION 5.00
Begin VB.Form frmRelTotalPorModulo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Total Consolidado"
   ClientHeight    =   3924
   ClientLeft      =   3684
   ClientTop       =   3444
   ClientWidth     =   4812
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3924
   ScaleWidth      =   4812
   Begin VB.Frame fraPrincipal 
      Caption         =   " Movimentos disponíveis "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3492
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4332
      Begin VB.CommandButton cmdImp 
         Caption         =   "&Imprimir"
         Height          =   312
         Left            =   360
         TabIndex        =   4
         Top             =   3000
         Width           =   1212
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   312
         Left            =   3000
         TabIndex        =   3
         Top             =   3000
         Width           =   1212
      End
      Begin VB.ListBox lstDias 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1392
         Left            =   360
         MultiSelect     =   2  'Extended
         TabIndex        =   2
         Top             =   1440
         Width           =   3852
      End
      Begin VB.ComboBox cmbMeses 
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
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   2172
      End
      Begin VB.Label lblDias 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Dias de movimentos"
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
         Left            =   360
         TabIndex        =   6
         Top             =   1200
         Width           =   2052
      End
      Begin VB.Label lblMeses 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Meses de movimentos"
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
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   1932
      End
   End
End
Attribute VB_Name = "frmRelTotalPorModulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdImp_Click()
    
    Dim qryResult   As rdoQuery
    Dim rsResult    As rdoResultset
    Dim QryTimeOut  As Variant
    Dim i           As Integer
    
    On Error GoTo Err_mnuRelPercDoctoModulo
    
    'Aumenta timeout devido ao processamento demorado da Procedure
    QryTimeOut = Geral.Banco.QueryTimeout
    Geral.Banco.QueryTimeout = 300
    
    Set qryResult = Geral.Banco.CreateQuery("", "{call GetPercDoctoModulo (?,?)}")
    
    Screen.MousePointer = vbHourglass
    
    With qryResult
        .rdoParameters(0).Value = Geral.DataProcessamento
        .rdoParameters(1).Value = Geral.DataProcessamento
        
        Set rsResult = .OpenResultset(rdOpenStatic, rdConcurReadOnly)
    End With
    
    If rsResult.RowCount <> 0 Then
        
        RptGeral.ReportFileName = App.path & "\RelPercDoctoModulo.rpt"
        
        With RptGeral
            .Formulas(0) = "AgenciaCentral     = '" & Geral.AgenciaCentral & "'"
            .Formulas(1) = "ILG_CapasTratadas  = '" & Formata(rsResult!ILG_CapasTratadas, "I") & "'"
            .Formulas(2) = "ILG_DoctosTratados = '" & Formata(rsResult!ILG_DoctosTratados, "I") & "'"
            .Formulas(3) = "ILG_DoctosTratados = '" & Formata(rsResult!ILG_DoctosTratados, "I") & "'"
            .Formulas(4) = "ILG_CapasReentrantes = '" & Formata(rsResult!ILG_CapasReentrantes, "I") & "'"
            .Formulas(5) = "ILG_PercCapas = '" & Formata(rsResult!ILG_PercCapas, "P") & "'"
            .Formulas(6) = "ILG_PercCapasReentrantes = '" & Formata(rsResult!ILG_PercCapasReentrantes, "P") & "'"
            .Formulas(7) = "ILG_PercDoctos = '" & Formata(rsResult!ILG_PercDoctos, "P") & "'"
            
            .Formulas(8) = "PRZ_CapasTratadas = '" & Formata(rsResult!PRZ_CapasTratadas, "I") & "'"
            .Formulas(9) = "PRZ_DoctosTratados = '" & Formata(rsResult!PRZ_DoctosTratados, "I") & "'"
            .Formulas(10) = "PRZ_CapasReentrantes = '" & Formata(rsResult!PRZ_CapasReentrantes, "I") & "'"
            .Formulas(11) = "PRZ_PercCapas = '" & Formata(rsResult!PRZ_PercCapas, "P") & "'"
            .Formulas(12) = "PRZ_PercCapasReentrantes = '" & Formata(rsResult!PRZ_PercCapasReentrantes, "P") & "'"
            .Formulas(13) = "PRZ_PercDoctos = '" & Formata(rsResult!PRZ_PercDoctos, "P") & "'"
            
            .Formulas(14) = "REC_CapasTratadas = '" & Formata(rsResult!REC_CapasTratadas, "I") & "'"
            .Formulas(15) = "REC_DoctosTratados = '" & Formata(rsResult!REC_DoctosTratados, "I") & "'"
            .Formulas(16) = "REC_CapasReentrantes = '" & Formata(rsResult!REC_CapasReentrantes, "I") & "'"
            .Formulas(17) = "REC_PercCapas = '" & Formata(rsResult!REC_PercCapas, "P") & "'"
            .Formulas(18) = "REC_PercCapasReentrantes = '" & Formata(rsResult!REC_PercCapasReentrantes, "P") & "'"
            .Formulas(19) = "REC_PercDoctos = '" & Formata(rsResult!REC_PercDoctos, "P") & "'"
            
            .Formulas(20) = "DataMovimento = '" & Right(Geral.DataProcessamento, 2) & "/" & Mid(Geral.DataProcessamento, 5, 2) & "/" & Left(Geral.DataProcessamento, 4) & "'"
            
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .WindowTitle = "Relatório de Percentual de Documentos por Módulo"
            .Action = 1
        End With
    Else
        Screen.MousePointer = vbDefault
        MsgBox "Sistema não possui informações Suficientes para emissão deste Relatório!" & vbCr & "Relatório de Segmentação por Documento / Agência", vbInformation, App.Title
    End If
    
    With RptGeral
        .ReportFileName = Empty
        For i = 0 To 20: .Formulas(i) = Empty: Next
    End With
    
    Screen.MousePointer = vbDefault
    
    'Retorno timeout default
    Geral.Banco.QueryTimeout = QryTimeOut
    
    Exit Sub
    
Err_mnuRelPercDoctoModulo:
    
    Screen.MousePointer = vbDefault
    'Retorno timeout default
    Geral.Banco.QueryTimeout = QryTimeOut
    
    Call TratamentoErro("Não foi possível abrir o relatório.", Err, rdoErrors)
    
End Sub

Private Sub cmdSair_Click()

    Unload Me
    
End Sub

Private Sub Form_Activate()
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
End Sub

Private Sub Form_Unload(Cancel As Integer)


End Sub

