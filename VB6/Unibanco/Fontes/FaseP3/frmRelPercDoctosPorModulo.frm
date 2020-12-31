VERSION 5.00
Object = "{4A8E27C4-EACA-11D3-9FFC-00104BC8688C}#1.0#0"; "dateedit.ocx"
Begin VB.Form frmRelPercDoctosPorModulo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório Percentual de Doctos. por Módulo"
   ClientHeight    =   2244
   ClientLeft      =   2352
   ClientTop       =   3408
   ClientWidth     =   4908
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2244
   ScaleWidth      =   4908
   Begin VB.Frame fraPrincipal 
      Caption         =   "Período de  Movimentação"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1932
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   4452
      Begin VB.CommandButton cmdImp 
         Caption         =   "&Imprimir"
         Height          =   312
         Left            =   600
         TabIndex        =   2
         Top             =   1440
         Width           =   1572
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   312
         Left            =   2280
         TabIndex        =   3
         Top             =   1440
         Width           =   1572
      End
      Begin DATEEDITLib.DateEdit dteDataInicial 
         Height          =   384
         Left            =   360
         TabIndex        =   0
         Top             =   720
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
         Top             =   720
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
         TabIndex        =   7
         Top             =   840
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
         TabIndex        =   6
         Top             =   480
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
         TabIndex        =   5
         Top             =   480
         Width           =   948
      End
   End
End
Attribute VB_Name = "frmRelPercDoctosPorModulo"
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
    
    Dim qryResult   As rdoQuery
    Dim rsResult    As rdoResultset
    Dim QryTimeOut  As Variant
    Dim i           As Integer
    
    On Error GoTo Err_ImprimeRelatorio
    
    'Aumenta timeout devido ao processamento demorado da Procedure
    QryTimeOut = Geral.Banco.QueryTimeout
    Geral.Banco.QueryTimeout = 300
    
    Set qryResult = Geral.Banco.CreateQuery("", "{call GetPercDoctoModulo (?,?)}")
    
    Screen.MousePointer = vbHourglass
    
    With qryResult
        .rdoParameters(0).Value = Mid(dteDataInicial.InverseText, 5, 2) + "-" + Right(dteDataInicial.InverseText, 2) + "-" + Left(dteDataInicial.InverseText, 4)
        .rdoParameters(1).Value = Mid(dteDataFinal.InverseText, 5, 2) + "-" + Right(dteDataFinal.InverseText, 2) + "-" + Left(dteDataFinal.InverseText, 4)
        
        Set rsResult = .OpenResultset(rdOpenStatic, rdConcurReadOnly)
    End With
    
    If rsResult.RowCount <> 0 Then
        
        With Principal.RptGeral
            .ReportFileName = App.path & "\RelPercDoctoModulo.rpt"
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
            .Formulas(20) = "DataInicial = '" & dteDataInicial.MaskText & "'"
            .Formulas(21) = "DataFinal = '" & dteDataFinal.MaskText & "'"
            
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .WindowTitle = "Relatório de Percentual de Documentos por Módulo"
            .Action = 1
        End With
    Else
        Screen.MousePointer = vbDefault
        MsgBox "Sistema não possui informações suficientes para emissão deste relatório!" & vbCr & "Relatório de Segmentação por Documento / Agência", vbInformation, App.Title
    End If
    
    With Principal.RptGeral
        .ReportFileName = Empty
        For i = 0 To 21: .Formulas(i) = Empty: Next
    End With
    
Exit_ImprimeRelatorio:
    Screen.MousePointer = vbDefault
    
    'Retorno timeout default
    Geral.Banco.QueryTimeout = QryTimeOut
    
    qryResult.Close
    If Not (rsResult Is Nothing) Then Set rsResult = Nothing
    
    Exit Sub
    
Err_ImprimeRelatorio:
    
    Screen.MousePointer = vbDefault
    Call TratamentoErro("Não foi possível abrir o relatório.", Err, rdoErrors)
    GoTo Exit_ImprimeRelatorio
    
End Sub
