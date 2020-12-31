VERSION 5.00
Object = "{4A8E27C4-EACA-11D3-9FFC-00104BC8688C}#1.0#0"; "dateedit.ocx"
Begin VB.Form FrmRelPerArrecadacao 
   Caption         =   "Impressão de Docto Tipo Arrecadação por Período"
   ClientHeight    =   1836
   ClientLeft      =   3744
   ClientTop       =   4044
   ClientWidth     =   4728
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1836
   ScaleWidth      =   4728
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdImp 
      Caption         =   "&Imprimir"
      Height          =   336
      Left            =   180
      TabIndex        =   4
      Top             =   1428
      Width           =   972
   End
   Begin VB.CommandButton CmdSair 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   336
      Left            =   3564
      TabIndex        =   6
      Top             =   1428
      Width           =   972
   End
   Begin VB.Frame Frame1 
      Caption         =   "Entre com o período :"
      Height          =   1128
      Left            =   192
      TabIndex        =   0
      Top             =   144
      Width           =   4356
      Begin DATEEDITLib.DateEdit DteInicial 
         Height          =   348
         Left            =   816
         TabIndex        =   3
         Top             =   504
         Width           =   1224
         _Version        =   65537
         _ExtentX        =   2159
         _ExtentY        =   614
         _StockProps     =   93
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
      End
      Begin DATEEDITLib.DateEdit DteFinal 
         Height          =   348
         Left            =   2820
         TabIndex        =   5
         Top             =   504
         Width           =   1224
         _Version        =   65537
         _ExtentX        =   2159
         _ExtentY        =   614
         _StockProps     =   93
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
      End
      Begin VB.Label Label2 
         Caption         =   "Até : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   2244
         TabIndex        =   2
         Top             =   576
         Width           =   504
      End
      Begin VB.Label Label1 
         Caption         =   "De :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   240
         TabIndex        =   1
         Top             =   576
         Width           =   540
      End
   End
End
Attribute VB_Name = "FrmRelPerArrecadacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private qryPesqDadosArrec   As rdoQuery
Private RsRelPerArrecadacao As rdoResultset
Private DataInicial         As String
Private DataFinal           As String

Private Sub CmdCancel_Click()
    Unload Me
End Sub
Private Sub CmdImp_Click()

'Verifica Preenchimento de campos Obrigatórios
    
    If Len(Trim(DteInicial.Text)) = 0 Then
       MsgBox "Campo Obrigatório, preencha! ", vbInformation, App.Title
       DteInicial.SetFocus
       Exit Sub
    ElseIf Len(Trim(DteFinal.Text)) = 0 Then
       MsgBox "Campo Obrigatório, preencha! ", vbInformation, App.Title
       DteFinal.SetFocus
       Exit Sub
    End If
    
    Call Formata_Data

    With qryPesqDadosArrec
        .rdoParameters(0).Value = DataInicial
        .rdoParameters(1).Value = DataFinal
        Set RsRelPerArrecadacao = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsRelPerArrecadacao.RowCount <> 0 Then
            With Principal.RptGeral
                .RptGeral.Connect = Geral.StringConexao
                
                If Geral.Backup Then
                    .ReportFileName = App.path & "\RelPerArrecBk.rpt"
                Else
                    .ReportFileName = App.path & "\RelPerArrecProd.rpt"
                End If
                
                .Formulas(0) = "AgenciaCentral    = '" & Geral.AgenciaCentral & "'"
                .StoredProcParam(0) = DataInicial
                .StoredProcParam(1) = DataFinal
                .Destination = crptToWindow
                .WindowTop = 1
                .WindowLeft = 1
                .WindowState = crptMaximized
                .PrintReport
            End With
    Else
        MsgBox "Não Existem dados para emissão deste Relátorio.", vbInformation, App.Title
        Exit Sub
    End If

    'Limpa Parametros + Caminho Relatório
    Principal.RptGeral.ReportFileName = Empty
    Principal.RptGeral.Formulas(0) = Empty
    Principal.RptGeral.StoredProcParam(0) = Empty
    Principal.RptGeral.StoredProcParam(1) = Empty
    
End Sub

Private Sub CmdSair_Click()
    Unload Me
End Sub

Private Sub DteFinal_LostFocus()

    If (KeyAscii = 13) Then
        If Len(DteFinal.Text) >= 8 Then
           Call CmdImp_Click
       End If

    ElseIf (KeyAscii = 27) Then
        KeyAscii = 0
        CmdSair_Click
    End If
    
End Sub

Private Sub DteInicial_KeyPress(KeyAscii As Integer)

    If (KeyAscii = 13) Then
        If Len(DteInicial.Text) >= 8 Then
        DteFinal.SetFocus
       End If

    ElseIf (KeyAscii = 27) Then
        KeyAscii = 0
        CmdSair_Click
    End If
    
End Sub
Private Sub Form_Load()

    'Pesquisa Dados Para emissão de Relatório
    Set qryPesqDadosArrec = Geral.Banco.CreateQuery("", "{Call GetRelPerArrecadacao (?,?)}")

End Sub
Public Sub Formata_Data()

    'Formata Data Inicial
    DataInicial = Mid(DteInicial.Text, 5, 4)                  'Ano
    DataInicial = DataInicial & Mid(DteInicial.Text, 3, 2)    'Mes
    DataInicial = DataInicial & Mid(DteInicial.Text, 1, 2)    'Dia
    
    'Formata Data Final
    DataFinal = Mid(DteFinal.Text, 5, 4)                      'Ano
    DataFinal = DataFinal + Mid(DteFinal.Text, 3, 2)          'Mes
    DataFinal = DataFinal + Mid(DteFinal.Text, 1, 2)          'Dia
    
    DataInicial = CDbl(DataInicial)                           'Converte string em Número
    DataFinal = CDbl(DataFinal)                               'Converte string em Número

End Sub
